using System;
using System.Collections.Generic;
using System.Linq;

namespace LedgerLens.Core
{
    public sealed class ModelCell
    {
        public string Address { get; set; } = "";
        public string Content { get; set; } = "";
        public bool HasFormula
        {
            get; set;
        }
        public bool IsAnalystInput
        {
            get; set;
        }
        public string Ticker { get; set; } = "";
        public string Metric { get; set; } = "";
        public string Period { get; set; } = "";
    }
    public sealed class CellChange
    {
        public string Address { get; set; } = "";
        public string ExpectedContent { get; set; } = "";
        public decimal NewValue
        {
            get; set;
        }
        public FinancialFact Source { get; set; } = new FinancialFact();
    }
    public sealed class RefreshPlan
    {
        public string Id { get; set; } = Guid.NewGuid().ToString("N");
        public string WorkbookId { get; set; } = "";
        public DateTimeOffset CreatedAt { get; set; } = DateTimeOffset.UtcNow;
        public CellChange[] Changes { get; set; } = Array.Empty<CellChange>();
        public string[] PreservedAddresses { get; set; } = Array.Empty<string>();
        public Dictionary<string, string> Dependencies { get; set; } = new Dictionary<string, string>();
    }
    public static class RefreshPlanner
    {
        public static RefreshPlan Create(string workbookId, IEnumerable<ModelCell> cells, Func<FactKey, FinancialFact> resolve, IDictionary<string, string>? dependencies = null)
        {
            if (string.IsNullOrWhiteSpace(workbookId))
                throw new ArgumentException("Workbook identity is required.", nameof(workbookId));
            if (cells == null)
                throw new ArgumentNullException(nameof(cells));
            if (resolve == null)
                throw new ArgumentNullException(nameof(resolve));
            var captured = dependencies == null ? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase) : new Dictionary<string, string>(dependencies, StringComparer.OrdinalIgnoreCase);
            var changes = new List<CellChange>();
            var preserved = new List<string>();
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var cell in cells)
            {
                if (cell == null || string.IsNullOrWhiteSpace(cell.Address) || !seen.Add(cell.Address))
                    throw new ArgumentException("Refresh addresses must be present and unique.", nameof(cells));
                if (seen.Count > 128)
                    throw new ArgumentException("A preview supports at most 128 model cells.", nameof(cells));
                if (cell.HasFormula || cell.IsAnalystInput)
                {
                    preserved.Add(cell.Address);
                    continue;
                }
                if (captured.TryGetValue(cell.Address, out var prior) && prior != cell.Content)
                    throw new ArgumentException("A model dependency conflicts with its captured value.", nameof(dependencies));
                captured[cell.Address] = cell.Content;
                var key = new FactKey(cell.Ticker, cell.Metric, cell.Period);
                var fact = resolve(key);
                fact.Validate();
                if (!fact.Key.Equals(key))
                    throw new InvalidOperationException("The source fact does not match the requested company, metric, and period.");
                if (cell.Content == EncodeNumber(fact.Value))
                    continue;
                changes.Add(new CellChange { Address = cell.Address, ExpectedContent = cell.Content, NewValue = fact.Value, Source = fact.Copy() });
            }
            if (captured.Count > 128)
                throw new ArgumentException("Too many model dependencies.", nameof(dependencies));
            return new RefreshPlan { WorkbookId = workbookId, Changes = changes.ToArray(), PreservedAddresses = preserved.ToArray(), Dependencies = captured };
        }
        public static string[] Conflicts(RefreshPlan plan, string workbookId, Func<string, string> read)
        {
            if (!string.Equals(plan.WorkbookId, workbookId, StringComparison.Ordinal))
                return new[] { "Workbook identity changed." };
            return plan.Changes.Where(c => !string.Equals(read(c.Address), c.ExpectedContent, StringComparison.Ordinal)).Select(c => c.Address)
                .Concat(plan.Dependencies.Where(d => read(d.Key) != d.Value).Select(d => d.Key)).Distinct(StringComparer.OrdinalIgnoreCase).ToArray();
        }
        public static string[] RollbackConflicts(RefreshPlan plan, Func<string, string> read) =>
            plan.Changes.Where(c => read(c.Address) != EncodeNumber(c.NewValue)).Select(c => c.Address).ToArray();
        public static string EncodeNumber(decimal number) => "n:" + number.ToString("G29", System.Globalization.CultureInfo.InvariantCulture);
    }
}
