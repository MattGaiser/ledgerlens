import fs from 'node:fs';
import path from 'node:path';
import crypto from 'node:crypto';
import { fileURLToPath } from 'node:url';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const companies = [
  {
    ticker: 'MSFT',
    name: 'Microsoft',
    cik: '0000789019',
    fiscalYearEnd: 'June 30',
    color: '#21b7a8',
  },
  {
    ticker: 'AAPL',
    name: 'Apple',
    cik: '0000320193',
    fiscalYearEnd: 'September',
    color: '#6f91ef',
  },
  { ticker: 'NVDA', name: 'NVIDIA', cik: '0001045810', fiscalYearEnd: 'January', color: '#97bd5d' },
];
const metrics = [
  [
    'Revenue',
    'Revenue',
    ['RevenueFromContractWithCustomerExcludingAssessedTax', 'Revenues', 'SalesRevenueNet'],
  ],
  ['GrossProfit', 'Gross profit', ['GrossProfit']],
  ['OperatingIncome', 'Operating income', ['OperatingIncomeLoss']],
  ['NetIncome', 'Net income', ['NetIncomeLoss']],
  ['OperatingCashFlow', 'Operating cash flow', ['NetCashProvidedByUsedInOperatingActivities']],
  [
    'CapitalExpenditure',
    'Capital expenditure',
    ['PaymentsToAcquirePropertyPlantAndEquipment', 'PaymentsToAcquireProductiveAssets'],
  ],
  ['DilutedEPS', 'Diluted EPS', ['EarningsPerShareDiluted']],
  ['Assets', 'Total assets', ['Assets']],
  ['Cash', 'Cash & equivalents', ['CashAndCashEquivalentsAtCarryingValue']],
];
const acquiredAt = new Date().toISOString();
const snapshotDate = acquiredAt.slice(0, 10);
const facts = [];
const inputs = [];
for (const company of companies) {
  const raw = fs.readFileSync(path.join(root, 'data/raw', `${company.ticker.toLowerCase()}.json`));
  const source = JSON.parse(raw);
  inputs.push({
    ticker: company.ticker,
    url: `https://data.sec.gov/api/xbrl/companyfacts/CIK${company.cik}.json`,
    sha256: crypto.createHash('sha256').update(raw).digest('hex'),
    bytes: raw.length,
    acquiredAt,
  });
  for (const [metric, label, concepts] of metrics) {
    for (const year of [2023, 2024, 2025]) {
      const instant = ['Assets', 'Cash'].includes(metric);
      const unit = metric === 'DilutedEPS' ? 'USD/shares' : 'USD';
      const candidates = concepts.flatMap((concept, priority) =>
        (source.facts['us-gaap'][concept]?.units[unit] ?? [])
          .filter(
            (f) => f.form === '10-K' && f.end.startsWith(`${year}-`) && f.filed <= snapshotDate,
          )
          .filter((f) =>
            instant
              ? !f.start
              : (Date.parse(f.end) - Date.parse(f.start)) / 86400000 >= 330 &&
                (Date.parse(f.end) - Date.parse(f.start)) / 86400000 <= 380,
          )
          .map((f) => ({ ...f, concept, priority })),
      );
      candidates.sort(
        (a, b) =>
          b.filed.localeCompare(a.filed) || a.priority - b.priority || b.end.localeCompare(a.end),
      );
      if (!candidates.length)
        throw new Error(`No annual fact: ${company.ticker}/${metric}/${year}`);
      const f = candidates[0];
      facts.push({
        ticker: company.ticker,
        metric,
        label,
        period: `FY${year}`,
        value: metric === 'DilutedEPS' ? f.val : f.val / 1e6,
        rawValue: f.val,
        unit: metric === 'DilutedEPS' ? 'USD/share' : 'USD millions',
        start: f.start ?? null,
        end: f.end,
        filed: f.filed,
        accession: f.accn,
        concept: `us-gaap:${f.concept}`,
        sourceUrl: `https://www.sec.gov/Archives/edgar/data/${Number(company.cik)}/${f.accn.replaceAll('-', '')}/${f.accn}-index.html`,
        sourceId: `${company.ticker}-${metric}-FY${year}`,
        acquiredAt,
      });
    }
  }
}
const dataset = {
  version: 1,
  snapshotDate,
  acquiredAt,
  description:
    'Annual reported SEC XBRL facts, latest available 10-K comparative values as of acquisition. Currency values in USD millions; EPS in USD/share. Fiscal years differ across issuers. Historical snapshot, not live market prices.',
  companies,
  inputs,
  facts,
};
fs.writeFileSync(path.join(root, 'data/financials.json'), JSON.stringify(dataset, null, 2) + '\n');
console.log(
  JSON.stringify(
    {
      facts: facts.length,
      snapshotDate,
      revenue2025: facts
        .filter((f) => f.period === 'FY2025' && f.metric === 'Revenue')
        .map((f) => ({ ticker: f.ticker, value: f.value, filed: f.filed })),
    },
    null,
    2,
  ),
);
