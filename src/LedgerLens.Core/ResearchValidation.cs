using System;
using System.Collections.Generic;
using System.Linq;

namespace LedgerLens.Core
{
    public static class ResearchValidation
    {
        public static void ValidateExport(ResearchAnswer answer)
        {
            if (answer == null || answer.Sources == null || answer.Sources.Length == 0 || answer.Sources.Length > 81)
                throw new InvalidOperationException("Research export requires its source records.");
            var ids = new HashSet<string>(StringComparer.Ordinal);
            foreach (var source in answer.Sources)
            {
                if (source == null || !ids.Add(source.SourceId))
                    throw new InvalidOperationException("Research export contains missing or duplicate source records.");
                source.Validate();
            }
            Validate(answer, answer.Sources);
        }

        public static void Validate(ResearchAnswer answer, IEnumerable<FinancialFact> suppliedFacts)
        {
            var allowed = new HashSet<string>(suppliedFacts.Select(f => f.SourceId), StringComparer.Ordinal);
            if (string.IsNullOrWhiteSpace(answer.Headline) || answer.Headline.Length > 180)
                throw new InvalidOperationException("Research headline is missing or too long.");
            if (string.IsNullOrWhiteSpace(answer.Summary) || answer.Summary.Length > 3000)
                throw new InvalidOperationException("Research summary is missing or too long.");
            if (answer.Claims == null || answer.Claims.Length < 1 || answer.Claims.Length > 8)
                throw new InvalidOperationException("Research requires one to eight supported claims.");
            foreach (var claim in answer.Claims)
            {
                if (claim == null || string.IsNullOrWhiteSpace(claim.Text) || claim.Text.Length > 2000)
                    throw new InvalidOperationException("A research claim is empty or too long.");
                if (claim.SourceIds == null || claim.SourceIds.Length == 0 || claim.SourceIds.Any(id => !allowed.Contains(id)))
                    throw new InvalidOperationException("Research cited an unknown or missing source.");
            }
            if (answer.Caveats == null || answer.Caveats.Length > 8 || answer.Caveats.Any(c => string.IsNullOrWhiteSpace(c) || c.Length > 1600))
                throw new InvalidOperationException("Research caveats are invalid or too long.");
        }
    }
}
