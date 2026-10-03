import fs from 'node:fs';
import assert from 'node:assert/strict';
import crypto from 'node:crypto';
const dataset = JSON.parse(fs.readFileSync('data/financials.json', 'utf8'));
const checks = [];
for (const input of dataset.inputs) {
  const bytes = fs.readFileSync(`data/raw/${input.ticker.toLowerCase()}.json`);
  assert.equal(crypto.createHash('sha256').update(bytes).digest('hex'), input.sha256);
  assert.equal(bytes.length, input.bytes);
  const source = JSON.parse(bytes);
  const company = dataset.companies.find((c) => c.ticker === input.ticker);
  assert.equal(Number(source.cik), Number(company.cik));
  for (const fact of dataset.facts.filter((f) => f.ticker === input.ticker)) {
    const concept = fact.concept.split(':')[1];
    const unit = fact.metric === 'DilutedEPS' ? 'USD/shares' : 'USD';
    const rows = source.facts['us-gaap'][concept].units[unit];
    const match = rows.find(
      (r) =>
        r.accn === fact.accession &&
        r.filed === fact.filed &&
        r.end === fact.end &&
        (r.start ?? null) === fact.start &&
        r.val === fact.rawValue &&
        r.form === '10-K',
    );
    assert(match, `Original SEC record must match every source field: ${fact.sourceId}`);
    assert.equal(fact.value, fact.rawValue / (unit === 'USD' ? 1e6 : 1));
    assert.equal(fact.unit, unit === 'USD' ? 'USD millions' : 'USD/share');
    assert.equal(fact.period, `FY${fact.end.slice(0, 4)}`);
    assert(fact.filed <= dataset.snapshotDate);
    const duration = fact.start ? (Date.parse(fact.end) - Date.parse(fact.start)) / 86400000 : null;
    assert(
      ['Assets', 'Cash'].includes(fact.metric)
        ? duration === null
        : duration >= 330 && duration <= 380,
    );
    const latest = rows
      .filter(
        (r) =>
          r.form === '10-K' &&
          r.end === fact.end &&
          (r.start ?? null) === fact.start &&
          r.filed <= dataset.snapshotDate,
      )
      .sort((a, b) => b.filed.localeCompare(a.filed))[0];
    assert.equal(
      latest.filed,
      fact.filed,
      'Use the latest eligible comparative filing for this concept and period',
    );
    assert.equal(
      fact.sourceUrl,
      `https://www.sec.gov/Archives/edgar/data/${Number(company.cik)}/${fact.accession.replaceAll('-', '')}/${fact.accession}-index.html`,
    );
    checks.push({
      sourceId: fact.sourceId,
      status: 'PASS',
      value: fact.value,
      unit: fact.unit,
      end: fact.end,
      filed: fact.filed,
      accession: fact.accession,
    });
  }
}
assert.equal(checks.length, 81);
assert.equal(new Set(checks.map((c) => c.sourceId)).size, 81);
fs.mkdirSync('artifacts/test-results', { recursive: true });
fs.writeFileSync(
  'artifacts/test-results/source-reconciliation.json',
  JSON.stringify(
    {
      checkedAt: new Date().toISOString(),
      status: 'PASS',
      method:
        'Independent field-by-field comparison with three hashed SEC companyfacts responses; original records, scaling, annual duration, filing cutoff, latest same-concept comparative, source URL.',
      checks,
    },
    null,
    2,
  ),
);
console.log('PASS: all 81 facts reconcile to hashed original SEC responses.');
