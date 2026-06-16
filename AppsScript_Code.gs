const SHEET_NAME = 'UploadHistory';

function doGet(e) {
  try {
    const action = e?.parameter?.action;
    if (action === 'clear')         return handleClear();
    if (action === 'write')         return handleWrite(e.parameter.data);
    if (action === 'storeDiffRows') return handleStoreDiffRows(e.parameter);
    if (action === 'getDiffRows')   return handleGetDiffRows(e.parameter);
    return handleRead();
  } catch(err) {
    return jsonResponse({ error: err.message });
  }
}

function doPost(e) {
  return jsonResponse({ error: 'Use GET' });
}

// ─── Read all history rows ────────────────────────────────────────────────────
function handleRead() {
  const sheet = getSheet();
  const [headers, ...rows] = sheet.getDataRange().getValues();
  const result = rows.map(row => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = row[i] ?? '');
    return obj;
  });
  return jsonResponse(result);
}

// ─── Write / upsert a session row ─────────────────────────────────────────────
function handleWrite(dataStr) {
  const session = JSON.parse(dataStr || '{}');
  const sheet   = getSheet();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const data    = sheet.getDataRange().getValues();
  const flat    = sessionToRow(session, headers);
  const idCol   = headers.indexOf('id');
  let found     = false;
  for (let i = 1; i < data.length; i++) {
    if (data[i][idCol] === session.id) {
      sheet.getRange(i + 1, 1, 1, headers.length).setValues([flat]);
      found = true;
      break;
    }
  }
  if (!found) sheet.appendRow(flat);
  return jsonResponse({ ok: true });
}

// ─── Store diff/error rows in a Sheet tab ────────────────────────────────────
// Each request appends a batch of rows tagged with sessionId as first column.
function handleStoreDiffRows(params) {
  const { sessionId, fileType, rows } = params;
  if (!sessionId || !fileType || !rows) return jsonResponse({ error: 'Missing params' });
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const tabName = fileType === 'error' ? 'ErrorData' : 'DiffData';
  const tab     = ss.getSheetByName(tabName) || ss.insertSheet(tabName);
  JSON.parse(rows).forEach(row => tab.appendRow([sessionId, ...row]));
  return jsonResponse({ ok: true });
}

// ─── Fetch stored diff/error rows for a session ───────────────────────────────
function handleGetDiffRows(params) {
  const { sessionId, fileType } = params;
  if (!sessionId || !fileType) return jsonResponse({ error: 'Missing params' });
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const tabName = fileType === 'error' ? 'ErrorData' : 'DiffData';
  const tab     = ss.getSheetByName(tabName);
  if (!tab) return jsonResponse({ rows: [] });
  const all  = tab.getDataRange().getValues();
  const rows = all.filter(r => String(r[0]) === String(sessionId)).map(r => r.slice(1));
  return jsonResponse({ rows });
}

// ─── Clear all history rows (keep header) ─────────────────────────────────────
function handleClear() {
  const sheet   = getSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.deleteRows(2, lastRow - 1);
  return jsonResponse({ ok: true });
}

// ─── Helpers ──────────────────────────────────────────────────────────────────
function getSheet() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let   sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow([
      'id','timestamp','storeName',
      'systemFileName','systemFileSize','systemFileRows','uniqueBarcodes',
      'physFileName','physFileSize','physFileRows',
      'varianceRows','varianceLocs','varianceTotalQty','varianceGeneratedAt',
      'matched','withDiff','notScanned','errorRows','diffGeneratedAt'
    ]);
  }
  return sheet;
}

function sessionToRow(s, headers) {
  const map = {
    id:                  s.id || '',
    timestamp:           s.timestamp || '',
    storeName:           s.storeName || '',
    systemFileName:      s.systemFile?.name || '',
    systemFileSize:      s.systemFile?.size || '',
    systemFileRows:      s.systemFile?.rows || '',
    uniqueBarcodes:      s.systemFile?.uniqueBarcodes ?? '',
    physFileName:        s.physicalScan?.name || '',
    physFileSize:        s.physicalScan?.size || '',
    physFileRows:        s.physicalScan?.rows || '',
    varianceRows:        s.varianceReport?.rows || '',
    varianceLocs:        s.varianceReport?.locations || '',
    varianceTotalQty:    s.varianceReport?.totalQty || '',
    varianceGeneratedAt: s.varianceReport?.generatedAt || '',
    matched:             s.differenceReport?.matched ?? '',
    withDiff:            s.differenceReport?.withDiff ?? '',
    notScanned:          s.differenceReport?.notScanned ?? '',
    errorRows:           s.differenceReport?.errorRows ?? '',
    diffGeneratedAt:     s.differenceReport?.generatedAt || '',
  };
  return headers.map(h => map[h] ?? '');
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
