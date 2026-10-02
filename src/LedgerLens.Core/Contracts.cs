using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using Newtonsoft.Json;

namespace LedgerLens.Core
{
    public sealed class FactKey : IEquatable<FactKey>
    {
        public string Ticker
        {
            get;
        }
        public string Metric
        {
            get;
        }
        public string Period
        {
            get;
        }
        [JsonConstructor]
        public FactKey(string ticker, string metric, string period)
        {
            Ticker = (ticker ?? "").Trim().ToUpperInvariant();
            if (!Regex.IsMatch(Ticker, @"^[A-Z][A-Z0-9.\-]{0,9}$"))
                throw new ArgumentException("Use a ticker such as MSFT.", nameof(ticker));
            Metric = MetricCatalog.Normalize(metric);
            Period = (period ?? "").Trim().ToUpperInvariant();
            if (Regex.IsMatch(Period, @"^20\d{2}$"))
                Period = "FY" + Period;
            if (!Regex.IsMatch(Period, @"^FY20\d{2}$"))
                throw new ArgumentException("Use an annual fiscal period such as FY2025.", nameof(period));
        }
        public bool Equals(FactKey? other) => other != null && Ticker == other.Ticker && Metric == other.Metric && Period == other.Period;
        public override bool Equals(object? obj) => Equals(obj as FactKey);
        public override int GetHashCode() => StringComparer.Ordinal.GetHashCode(ToString());
        public override string ToString() => Ticker + "/" + Metric + "/" + Period;
    }

    public static class MetricCatalog
    {
        public static readonly IReadOnlyDictionary<string, string> Labels = new Dictionary<string, string>
        {
            ["Revenue"] = "Revenue",
            ["GrossProfit"] = "Gross profit",
            ["OperatingIncome"] = "Operating income",
            ["NetIncome"] = "Net income",
            ["OperatingCashFlow"] = "Operating cash flow",
            ["CapitalExpenditure"] = "Capital expenditure",
            ["DilutedEPS"] = "Diluted EPS",
            ["Assets"] = "Total assets",
            ["Cash"] = "Cash & equivalents"
        };
        public static string Normalize(string? value)
        {
            var compact = Regex.Replace(value ?? "", @"[\s_\-]", "");
            var key = Labels.Keys.FirstOrDefault(k => string.Equals(k, compact, StringComparison.OrdinalIgnoreCase));
            return key ?? throw new ArgumentException("Unknown metric. Choose Revenue, GrossProfit, OperatingIncome, NetIncome, OperatingCashFlow, CapitalExpenditure, DilutedEPS, Assets, or Cash.", nameof(value));
        }
    }

    public sealed class FinancialFact
    {
        public string Ticker { get; set; } = "";
        public string Metric { get; set; } = "";
        public string Label { get; set; } = "";
        public string Period { get; set; } = "";
        public decimal Value
        {
            get; set;
        }
        public decimal RawValue
        {
            get; set;
        }
        public string Unit { get; set; } = "";
        public string? Start
        {
            get; set;
        }
        public string End { get; set; } = "";
        public string Filed { get; set; } = "";
        public string Accession { get; set; } = "";
        public string Concept { get; set; } = "";
        public string SourceUrl { get; set; } = "";
        public string SourceId { get; set; } = "";
        public DateTimeOffset AcquiredAt
        {
            get; set;
        }
        [JsonIgnore] public FactKey Key => new FactKey(Ticker, Metric, Period);
        public FinancialFact Copy() => (FinancialFact)MemberwiseClone();
        public void Validate()
        {
            _ = Key;
            if (Unit != (Metric == "DilutedEPS" ? "USD/share" : "USD millions"))
                throw new InvalidOperationException("Unexpected fact units.");
            if (Value != (Metric == "DilutedEPS" ? RawValue : RawValue / 1000000m))
                throw new InvalidOperationException("Unit conversion does not reconcile.");
            var accession = Accession ?? "";
            if (!Regex.IsMatch(accession, @"^\d{10}-\d{2}-\d{6}$"))
                throw new InvalidOperationException("Invalid SEC accession.");
            if (!Uri.TryCreate(SourceUrl, UriKind.Absolute, out var source) || source.Scheme != "https" || source.Host != "www.sec.gov" || !source.AbsolutePath.StartsWith("/Archives/edgar/data/", StringComparison.Ordinal))
                throw new InvalidOperationException("Unexpected filing source.");
            if (!source.AbsolutePath.EndsWith("/" + accession.Replace("-", "") + "/" + accession + "-index.html", StringComparison.Ordinal) || source.Query.Length != 0 || source.Fragment.Length != 0)
                throw new InvalidOperationException("Filing URL does not match its SEC accession.");
            if (!DateTime.TryParseExact(End, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out var end) || Period != "FY" + end.Year)
                throw new InvalidOperationException("Fiscal period does not reconcile with end date for supported issuers.");
            if (!DateTime.TryParseExact(Filed, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out var filed) || filed < end || filed > AcquiredAt.UtcDateTime.Date)
                throw new InvalidOperationException("Invalid filing date.");
            if (Metric != "Assets" && Metric != "Cash")
            {
                if (!DateTime.TryParseExact(Start, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out var start) || (end - start).TotalDays < 330 || (end - start).TotalDays > 380)
                    throw new InvalidOperationException("Fact is not an annual reporting period.");
            }
            else if (Start != null)
                throw new InvalidOperationException("An instant balance-sheet fact cannot have a duration.");
            if (SourceId != Ticker + "-" + Metric + "-" + Period || string.IsNullOrWhiteSpace(Concept) || !Concept.StartsWith("us-gaap:", StringComparison.Ordinal))
                throw new InvalidOperationException("Fact provenance is incomplete.");
        }
    }

    public sealed class Company
    {
        public string Ticker { get; set; } = "";
        public string Name { get; set; } = "";
        public string Cik { get; set; } = "";
        public string FiscalYearEnd { get; set; } = "";
        public string Color { get; set; } = "";
    }
    public sealed class FinancialDataset
    {
        public int Version
        {
            get; set;
        }
        public string SnapshotDate { get; set; } = "";
        public string Description { get; set; } = "";
        public DateTimeOffset AcquiredAt
        {
            get; set;
        }
        public Company[] Companies { get; set; } = Array.Empty<Company>();
        public FinancialFact[] Facts { get; set; } = Array.Empty<FinancialFact>();
    }
    public sealed class FactResult
    {
        public FinancialFact Fact { get; set; } = new FinancialFact();
        public string Freshness { get; set; } = "snapshot";
        public string Provider { get; set; } = "SEC snapshot";
        public DateTimeOffset ServedAt
        {
            get; set;
        }
        public string Message { get; set; } = "";
    }
    public sealed class ResearchRequest
    {
        public string Ticker { get; set; } = "MSFT";
        public string Question { get; set; } = "What changed in profitability and cash generation?";
        public bool UseAi { get; set; } = true;
    }
    public sealed class ResearchClaim
    {
        public string Text { get; set; } = "";
        public string[] SourceIds { get; set; } = Array.Empty<string>();
    }
    public sealed class ResearchAnswer
    {
        public string Headline { get; set; } = "";
        public string Summary { get; set; } = "";
        public ResearchClaim[] Claims { get; set; } = Array.Empty<ResearchClaim>();
        public string[] Caveats { get; set; } = Array.Empty<string>();
        public FinancialFact[] Sources { get; set; } = Array.Empty<FinancialFact>();
        public string Provider { get; set; } = "";
        public string Model { get; set; } = "";
        public bool IsAiGenerated
        {
            get; set;
        }
        public bool Cached
        {
            get; set;
        }
        public DateTimeOffset GeneratedAt
        {
            get; set;
        }
    }
    public sealed class RuntimeEndpoint
    {
        public string BaseUrl { get; set; } = "";
        public string? OfficeUrl
        {
            get; set;
        }
        public string Token { get; set; } = "";
        public int ProcessId
        {
            get; set;
        }
    }
}
