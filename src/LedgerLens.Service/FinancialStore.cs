using System.Globalization;
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
        if (template.Version != 1 || template.Facts is not { Length: > 0 } || template.Companies is not { Length: > 0 } || template.Companies.Any(c => c == null))
            throw new InvalidDataException("Unsupported or empty financial dataset.");
        foreach (var fact in template.Facts)
            (fact ?? throw new InvalidDataException("A source fact is missing.")).Validate();
        if (template.Facts.Select(f => f.Key).Distinct().Count() != template.Facts.Length)
            throw new InvalidDataException("Duplicate source facts.");
        facts = template.Facts.ToDictionary(f => f.Key, f => f.Copy());
        if (cachePath != null && File.Exists(cachePath))
        {
            try
            {
                Replace(JsonConvert.DeserializeObject<FinancialFact[]>(File.ReadAllText(cachePath)) ?? []);
            }
            catch (Exception e) when (e is JsonException or InvalidOperationException or ArgumentException) { /* Corrupt local state never replaces the validated bundled evidence. */ }
        }
    }
    public FinancialFact Get(FactKey key) => Volatile.Read(ref facts).TryGetValue(key, out var fact) ? fact.Copy() : throw new KeyNotFoundException($"No reported fact for {key}. Supported annual periods: FY2023–FY2025.");
    public FinancialFact[] ForCompany(string ticker) => Volatile.Read(ref facts).Values.Where(f => f.Ticker == ticker.ToUpperInvariant()).OrderBy(f => f.Period).ThenBy(f => f.Metric).Select(f => f.Copy()).ToArray();
    public int Count => Volatile.Read(ref facts).Count;
    public void Replace(IEnumerable<FinancialFact> replacements)
    {
        if (replacements == null)
            throw new ArgumentNullException(nameof(replacements));
        var complete = replacements.Select(f => f?.Copy() ?? throw new ArgumentException("A replacement fact is missing.")).ToArray();
        foreach (var fact in complete)
            fact.Validate();
        if (complete.Select(f => f.Key).Distinct().Count() != complete.Length)
            throw new ArgumentException("Duplicate replacement facts.");
        lock (writeGate)
        {
            var next = new Dictionary<FactKey, FinancialFact>(facts);
            foreach (var fact in complete)
            {
                if (!next.ContainsKey(fact.Key))
                    throw new ArgumentException("Replacement is outside the configured research universe.");
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

public sealed class SecClient(SecHttpClient http)
{
    public async Task<FinancialFact[]> RefreshAsync(Company company, FinancialFact[] existing, CancellationToken cancellation)
    {
        var body = await http.GetCompanyFactsAsync(company.Cik, cancellation);
        try
        {
            var json = JObject.Parse(body);
            if (json["cik"] is not JValue { Type: JTokenType.Integer } identity ||
                !long.TryParse(identity.ToString(Formatting.None), NumberStyles.None, CultureInfo.InvariantCulture, out var cik) ||
                cik != long.Parse(company.Cik, CultureInfo.InvariantCulture))
                throw new InvalidDataException("SEC response does not identify the requested company.");
            var retrieved = DateTimeOffset.UtcNow;
            var replacements = existing.Select(old => SelectAnnualFact(json, old, company.Cik, retrieved)).ToArray();
            cancellation.ThrowIfCancellationRequested();
            return replacements;
        }
        catch (Exception error) when (error is JsonException or InvalidDataException)
        {
            throw new SecUnavailableException("sec_invalid_data", "SEC returned incomplete or invalid financial evidence. Saved evidence is unchanged.");
        }
    }

    public static FinancialFact SelectAnnualFact(JObject json, FinancialFact previous, string cik, DateTimeOffset acquired)
    {
        var concept = previous.Concept.Split(':')[1];
        var units = previous.Metric == "DilutedEPS" ? "USD/shares" : "USD";
        var facts = json["facts"] as JObject;
        var gaap = facts?["us-gaap"] as JObject;
        var metric = gaap?[concept] as JObject;
        var unitValues = metric?["units"] as JObject;
        var candidates = unitValues?[units] as JArray ?? throw new InvalidDataException($"SEC no longer supplies {concept} in {units}.");
        var year = int.Parse(previous.Period.Substring(2), CultureInfo.InvariantCulture);
        var instant = previous.Metric is "Assets" or "Cash";
        var selected = candidates.OfType<JObject>().Where(f =>
        {
            if (Text(f, "form") != "10-K" || !DateTime.TryParseExact(Text(f, "end"), "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out var end) || end.Year != year)
                return false;
            if (!DateTime.TryParseExact(Text(f, "filed"), "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out var filed) || filed > acquired.UtcDateTime.Date)
                return false;
            if (instant)
                return f["start"] == null;
            return DateTime.TryParseExact(Text(f, "start"), "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out var start) && (end - start).TotalDays >= 330 && (end - start).TotalDays <= 380;
        }).OrderByDescending(f => Text(f, "filed"), StringComparer.Ordinal).ThenByDescending(f => Text(f, "end"), StringComparer.Ordinal).FirstOrDefault() ?? throw new InvalidDataException($"No annual SEC fact for {previous.Key}.");
        if (selected["val"] is not JValue number || number.Type is not (JTokenType.Integer or JTokenType.Float) ||
            !decimal.TryParse(number.ToString(Formatting.None), NumberStyles.Float, CultureInfo.InvariantCulture, out var rawValue))
            throw new InvalidDataException("SEC annual fact is missing a valid numeric value.");
        var accession = Text(selected, "accn");
        if (string.IsNullOrWhiteSpace(accession))
            throw new InvalidDataException("SEC annual fact is missing its accession.");
        var fact = previous.Copy();
        fact.RawValue = rawValue;
        fact.Value = units == "USD" ? fact.RawValue / 1000000m : fact.RawValue;
        fact.Start = Text(selected, "start");
        fact.End = Text(selected, "end") ?? throw new InvalidDataException("SEC period end is missing.");
        fact.Filed = Text(selected, "filed") ?? throw new InvalidDataException("SEC filing date is missing.");
        fact.Accession = accession;
        fact.SourceUrl = $"https://www.sec.gov/Archives/edgar/data/{long.Parse(cik, CultureInfo.InvariantCulture)}/{fact.Accession.Replace("-", "")}/{fact.Accession}-index.html";
        fact.AcquiredAt = acquired;
        try
        {
            fact.Validate();
        }
        catch (InvalidOperationException error) { throw new InvalidDataException("SEC fact does not satisfy the financial data contract.", error); }
        return fact;
    }

    private static string? Text(JObject value, string property) => value[property]?.Type == JTokenType.String ? value.Value<string>(property) : null;
}
