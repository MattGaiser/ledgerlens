// A deliberately small Office.js adapter. Native XLL formulas and the synchronous,
// guarded multi-cell refresh remain Windows capabilities; this bridge inserts a
// sourced snapshot and exports research in current Office.js hosts.
export function createOfficeBridge(office, excel, request) {
  function requireWriteAccess() {
    if (office.context.document.mode === office.DocumentMode.ReadOnly)
      throw new Error('This workbook is read-only. Save an editable copy first.');
  }
  function loadDestination(cell) {
    cell.load(['address', 'rowCount', 'columnCount', 'values', 'formulas']);
    cell.worksheet.protection.load('protected');
    return cell.getMergedAreasOrNullObject();
  }
  function validateDestination(cell, merged) {
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
  }
  async function insertValue(payload) {
    requireWriteAccess();
    const key = [payload.ticker, payload.metric, payload.period].map(encodeURIComponent).join('/');
    return excel.run(async (context) => {
      const cell = context.workbook.getSelectedRange();
      let merged = loadDestination(cell);
      await context.sync();
      validateDestination(cell, merged);
      const originalAddress = cell.address;
      const result = await request('/facts/' + key);
      if (
        !Number.isFinite(result?.fact?.value) ||
        result.fact.ticker !== payload.ticker ||
        result.fact.metric !== payload.metric ||
        result.fact.period !== payload.period ||
        result.fact.sourceId !== `${payload.ticker}-${payload.metric}-${payload.period}`
      )
        throw new Error('The source did not return the requested financial fact.');
      requireWriteAccess();
      // Keep the original range in this Excel.run context, then reload its state
      // after the network wait. A new selection must not redirect the write.
      const selected = context.workbook.getSelectedRange();
      selected.load('address');
      merged = loadDestination(cell);
      await context.sync();
      requireWriteAccess();
      if (cell.address !== originalAddress || selected.address !== originalAddress)
        throw new Error(
          'The selection changed while data loaded. Select the intended cell and try again.',
        );
      validateDestination(cell, merged);
      cell.values = [[result.fact.value]];
      cell.numberFormat = [[result.fact.metric === 'DilutedEPS' ? '0.00' : '#,##0;[Red](#,##0);–']];
      await context.sync();
      return { address: cell.address, mode: 'snapshot', sourceId: result.fact.sourceId };
    });
  }
  async function saveResearch({ answer }) {
    requireWriteAccess();
    validateResearch(answer);
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

function validateResearch(answer) {
  const text = (value, maximum) =>
    typeof value === 'string' && value.trim().length > 0 && value.length <= maximum;
  if (
    !answer ||
    !text(answer.headline, 180) ||
    !text(answer.summary, 3000) ||
    !Array.isArray(answer.claims) ||
    answer.claims.length < 1 ||
    answer.claims.length > 8 ||
    !Array.isArray(answer.sources) ||
    answer.sources.length < 1 ||
    answer.sources.length > 81 ||
    !Array.isArray(answer.caveats) ||
    answer.caveats.length > 8 ||
    answer.caveats.some((value) => !text(value, 1600))
  )
    throw new Error('The research answer is incomplete.');
  if (answer.sources.some((source) => !source || !text(source.sourceId, 128)))
    throw new Error('A research source is invalid.');
  const allowed = new Set(answer.sources.map((source) => source.sourceId));
  if (allowed.size !== answer.sources.length) throw new Error('Research sources must be unique.');
  if (
    answer.claims.some(
      (claim) =>
        !claim ||
        !text(claim.text, 2000) ||
        !Array.isArray(claim.sourceIds) ||
        claim.sourceIds.length < 1 ||
        claim.sourceIds.length > 81 ||
        claim.sourceIds.some((id) => !allowed.has(id)),
    )
  )
    throw new Error('A research citation is missing.');
  for (const source of answer.sources) {
    const url = new URL(source.sourceUrl);
    if (
      url.protocol !== 'https:' ||
      url.hostname !== 'www.sec.gov' ||
      !url.pathname.startsWith('/Archives/edgar/data/')
    )
      throw new Error('A research source is not an SEC filing.');
  }
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
  let timer;
  let info;
  try {
    info = await Promise.race([
      Office.onReady(),
      new Promise((_, reject) => {
        timer = setTimeout(
          () => reject(new Error('Open this page as an Excel Office add-in.')),
          15000,
        );
      }),
    ]);
  } finally {
    clearTimeout(timer);
  }
  if (
    info.host !== Office.HostType.Excel ||
    !Office.context.requirements.isSetSupported('ExcelApi', '1.13')
  )
    throw new Error('The Office.js preview requires an Excel host supporting ExcelApi 1.13.');
  return createOfficeBridge(Office, Excel, request);
}
