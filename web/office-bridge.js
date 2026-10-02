// A deliberately small Office.js adapter. Native XLL formulas and the synchronous,
// guarded multi-cell refresh remain Windows capabilities; this bridge inserts a
// sourced snapshot and exports research in current Office.js hosts.
export function createOfficeBridge(office, excel, request) {
  function requireWriteAccess() {
    if (office.context.document.mode === office.DocumentMode.ReadOnly)
      throw new Error('This workbook is read-only. Save an editable copy first.');
  }
  async function insertValue(payload) {
    requireWriteAccess();
    const key = [payload.ticker, payload.metric, payload.period].map(encodeURIComponent).join('/');
    const result = await request('/facts/' + key);
    if (!Number.isFinite(result.fact.value))
      throw new Error('The source returned an invalid numeric value.');
    return excel.run(async (context) => {
      const cell = context.workbook.getSelectedRange();
      cell.load(['address', 'rowCount', 'columnCount', 'values', 'formulas']);
      const merged = cell.getMergedAreasOrNullObject();
      cell.worksheet.protection.load('protected');
      await context.sync();
      if (
        cell.rowCount !== 1 ||
        cell.columnCount !== 1 ||
        !merged.isNullObject ||
        !['', null].includes(cell.values[0][0]) ||
        String(cell.formulas[0][0]).startsWith('=')
      )
        throw new Error(
          'Select one empty, unmerged cell. Existing content is never overwritten by Insert.',
        );
      if (cell.worksheet.protection.protected)
        throw new Error('The selected worksheet is protected.');
      cell.values = [[result.fact.value]];
      cell.numberFormat = [[result.fact.metric === 'DilutedEPS' ? '0.00' : '#,##0;[Red](#,##0);–']];
      await context.sync();
      return { address: cell.address, mode: 'snapshot', sourceId: result.fact.sourceId };
    });
  }
  async function saveResearch({ answer }) {
    requireWriteAccess();
    if (
      !answer ||
      !Array.isArray(answer.claims) ||
      answer.claims.length > 8 ||
      !Array.isArray(answer.sources)
    )
      throw new Error('The research answer is incomplete.');
    const allowed = new Set(answer.sources.map((source) => source.sourceId));
    if (
      answer.claims.some(
        (claim) => !claim.sourceIds?.length || claim.sourceIds.some((id) => !allowed.has(id)),
      )
    )
      throw new Error('A research citation is missing.');
    for (const source of answer.sources) {
      const url = new URL(source.sourceUrl);
      if (url.protocol !== 'https:' || url.hostname !== 'www.sec.gov')
        throw new Error('A research source is not an SEC filing.');
    }
    const rows = [
      ['LEDGERLENS RESEARCH', answer.provider],
      [answer.headline, answer.model],
      [answer.summary, ''],
      ['Claim', 'Evidence IDs'],
      ...answer.claims.map((claim) => [claim.text, claim.sourceIds.join('; ')]),
      ['CAVEATS', ''],
      ...answer.caveats.map((text) => [text, '']),
      ['SOURCE', 'SEC filing'],
      ...answer.sources.map((source) => [source.sourceId, source.sourceUrl]),
    ];
    return excel.run(async (context) => {
      const name =
        'Research ' +
        new Date().toISOString().slice(11, 19).replaceAll(':', '') +
        ' ' +
        crypto.randomUUID().slice(0, 4);
      const sheet = context.workbook.worksheets.add(name);
      const range = sheet.getRangeByIndexes(0, 0, rows.length, 2);
      range.numberFormat = rows.map(() => ['@', '@']);
      range.values = rows.map((row) => row.map((value) => "'" + String(value ?? '')));
      range.format.columnWidth = 360;
      range.format.wrapText = true;
      range.format.autofitRows();
      const header = sheet.getRange('A1:B1');
      header.format.fill.color = '#142f3c';
      header.format.font.color = '#ffffff';
      header.format.font.bold = true;
      sheet.activate();
      await context.sync();
      return { sheet: name };
    });
  }
  return {
    execute(command, payload) {
      if (command === 'insertFormula') return insertValue(payload);
      if (command === 'saveResearch') return saveResearch(payload);
      return Promise.reject(
        new Error(
          'This action is available in the Windows native add-in. The Office.js preview supports sourced values and research export.',
        ),
      );
    },
  };
}

export async function initializeOffice(request) {
  if (!globalThis.Office) {
    await new Promise((resolve, reject) => {
      const script = document.createElement('script');
      script.src = 'https://appsforoffice.microsoft.com/lib/1/hosted/office.js';
      script.onload = resolve;
      script.onerror = () =>
        reject(new Error('Office.js could not load. Check the Office CDN connection.'));
      document.head.append(script);
    });
  }
  const info = await Promise.race([
    Office.onReady(),
    new Promise((_, reject) =>
      setTimeout(() => reject(new Error('Open this page as an Excel Office add-in.')), 15000),
    ),
  ]);
  if (
    info.host !== Office.HostType.Excel ||
    !Office.context.requirements.isSetSupported('ExcelApi', '1.13')
  )
    throw new Error('The Office.js preview requires an Excel host supporting ExcelApi 1.13.');
  return createOfficeBridge(Office, Excel, request);
}
