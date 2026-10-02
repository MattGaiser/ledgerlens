using System.Globalization;
using System.Net;
using LedgerLens.Core;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace LedgerLens.Service;

public sealed class FinancialStore
{
    private Dictionary<FactKey, FinancialFact> facts;
    private readonly FinancialDataset template;
    private readonly string? cachePath;
    private readonly object writeGate = new();
    public FinancialDataset Dataset => new() { Version = template.Version, SnapshotDate = template.SnapshotDate, Description = template.Description, AcquiredAt = template.AcquiredAt, Companies = template.Companies.Select(c => new Company { Ticker = c.Ticker, Name = c.Name, Cik = c.Cik, FiscalYearEnd = c.FiscalYearEnd, Color = c.Color }).ToArray(), Facts = Volatile.Read(ref facts).Values.Select(f => f.Copy()).ToArray() };
    public FinancialStore(string path, string? cachePath = null)
    {
        this.cachePath = cachePath;
        template = JsonConvert.DeserializeObject<FinancialDataset>(File.ReadAllText(path)) ?? throw new InvalidDataException("Empty financial dataset.");
        if (template.Version != 1 || template.Facts.Length == 0) throw new InvalidDataException("Unsupported or empty financial dataset.");
        foreach (var fact in template.Facts) fact.Validate();
        if (template.Facts.Select(f => f.Key).Distinct().Count() != template.Facts.Length) throw new InvalidDataException("Duplicate source facts.");
        facts = template.Facts.ToDictionary(f => f.Key, f => f.Copy());
        if (cachePath != null && File.Exists(cachePath))
        {
            try { Replace(JsonConvert.DeserializeObject<FinancialFact[]>(File.ReadAllText(cachePath)) ?? []); }
            catch (Exception e) when (e is JsonException or InvalidOperationException or ArgumentException) { /* Corrupt local state never replaces the validated bundled evidence. */ }
        }
    }
    public FinancialFact Get(FactKey key) => Volatile.Read(ref facts).TryGetValue(key, out var fact) ? fact.Copy() : throw new KeyNotFoundException($"No reported fact for {key}. Supported annual periods: FY2023–FY2025.");
    public FinancialFact[] ForCompany(string ticker) => Volatile.Read(ref facts).Values.Where(f => f.Ticker == ticker.ToUpperInvariant()).OrderBy(f => f.Period).ThenBy(f => f.Metric).Select(f => f.Copy()).ToArray();
    public int Count => Volatile.Read(ref facts).Count;
    public void Replace(IEnumerable<FinancialFact> replacements)
    {
        if (replacements == null) throw new ArgumentNullException(nameof(replacements));
        var complete = replacements.Select(f => f?.Copy() ?? throw new ArgumentException("A replacement fact is missing.")).ToArray();
        foreach (var fact in complete) fact.Validate();
        if (complete.Select(f => f.Key).Distinct().Count() != complete.Length) throw new ArgumentException("Duplicate replacement facts.");
        lock (writeGate)
        {
            var next = new Dictionary<FactKey, FinancialFact>(facts);
            foreach (var fact in complete)
            {
                if (!next.ContainsKey(fact.Key)) throw new ArgumentException("Replacement is outside the configured research universe.");
                next[fact.Key] = fact;
            }
            if (cachePath != null)
            {
                Directory.CreateDirectory(Path.GetDirectoryName(cachePath)!);
                var temporary = cachePath + ".tmp";
                File.WriteAllText(temporary, JsonConvert.SerializeObject(next.Values));
                File.Move(temporary, cachePath, true);
            }
            Volatile.Write(ref facts, next);
        }
    }
}

public sealed class SecClient(HttpClient http)
{
    public async Task<FinancialFact[]> RefreshAsync(Company company, FinancialFact[] existing, CancellationToken cancellation)
    {
        using var request = new HttpRequestMessage(HttpMethod.Get, $"https://data.sec.gov/api/xbrl/companyfacts/CIK{company.Cik}.json");
        request.Headers.UserAgent.ParseAdd("LedgerLens/1.0 (individual financial research prototype)");
        using var response = await http.SendAsync(request, HttpCompletionOption.ResponseHeadersRead, cancellation);
        response.EnsureSuccessStatusCode();
        var json = JObject.Parse(await response.Content.ReadAsStringAsync(cancellation));
        var retrieved = DateTimeOffset.UtcNow;
        return existing.Select(old => SelectAnnualFact(json, old, company.Cik, retrieved)).ToArray();
    }

    public static FinancialFact SelectAnnualFact(JObject json, FinancialFact previous, string cik, DateTimeOffset acquired)
    {
        var concept = previous.Concept.Split(':')[1];
        var units = previous.Metric == "DilutedEPS" ? "USD/shares" : "USD";
        var candidates = json["facts"]?["us-gaap"]?[concept]?["units"]?[units] as JArray ?? throw new InvalidDataException($"SEC no longer supplies {concept} in {units}.");
        var year = int.Parse(previous.Period.Substring(2), CultureInfo.InvariantCulture);
        var instant = previous.Metric is "Assets" or "Cash";
        var selected = candidates.OfType<JObject>().Where(f =>
        {
            if ((string?)f["form"] != "10-K" || !DateTime.TryParseExact((string?)f["end"], "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out var end) || end.Year != year) return false;
            if (!DateTime.TryParseExact((string?)f["filed"], "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out var filed) || filed > acquired.UtcDateTime.Date) return false;
            if (instant) return f["start"] == null;
            return DateTime.TryParseExact((string?)f["start"], "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out var start) && (end - start).TotalDays >= 330 && (end - start).TotalDays <= 380;
        }).OrderByDescending(f => (string?)f["filed"], StringComparer.Ordinal).ThenByDescending(f => (string?)f["end"], StringComparer.Ordinal).FirstOrDefault() ?? throw new InvalidDataException($"No annual SEC fact for {previous.Key}.");
        var fact = previous.Copy();
        fact.RawValue = selected.Value<decimal>("val");
        fact.Value = units == "USD" ? fact.RawValue / 1000000m : fact.RawValue;
        fact.Start = (string?)selected["start"]; fact.End = selected.Value<string>("end")!;
        fact.Filed = selected.Value<string>("filed")!; fact.Accession = selected.Value<string>("accn")!;
        fact.SourceUrl = $"https://www.sec.gov/Archives/edgar/data/{long.Parse(cik, CultureInfo.InvariantCulture)}/{fact.Accession.Replace("-", "")}/{fact.Accession}-index.html";
        fact.AcquiredAt = acquired;
        fact.Validate();
        return fact;
    }
}
