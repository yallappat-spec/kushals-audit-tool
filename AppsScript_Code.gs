// ─── Kushals Audit Tool — Google Apps Script Backend ──────────────────────────
// Paste this entire file into your Google Apps Script project and re-deploy
// as a Web App (Execute as: Me, Who has access: Anyone).
//
// Required Drive folder for uploaded reports:
//   Create a folder in Google Drive called "Kushals Audit Reports"
//   and paste its folder ID into REPORT_FOLDER_ID below.

const SHEET_NAME        = 'UploadHistory';
const REPORT_FOLDER_ID  = 'YOUR_DRIVE_FOLDER_ID'; // <-- replace this

// ─── GET: all operations go through doGet to avoid Apps Script POST redirect issues ───
function doGet(e) {
  try {
    const action = e?.parameter?.action;
    if (action === 'clear')          return handleClear();
    if (action === 'write')          return handleWrite(e.parameter.data);
    if (action === 'storeChunk')     return handleStoreChunk(e.parameter);
    if (action === 'finalizeUpload') return handleFinalizeUpload(e.parameter);
    return handleRead();
  } catch(err) {
    return jsonResponse({ error: err.message });
  }
}

function doPost(e) {
  // Not used — Apps Script 302 redirect drops POST body.
  return jsonResponse({ error: 'Use GET' });
}

// ─── Read all rows ────────────────────────────────────────────────────────────
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

  // Build a flat row from the session object
  const flat = sessionToRow(session, headers);

  // Find existing row by id
  const idCol = headers.indexOf('id');
  let found   = false;
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

// ─── Store one base64 chunk in Script Properties ─────────────────────────────
function handleStoreChunk(params) {
  const { sessionId, fileType, idx, data } = params;
  if (!sessionId || !fileType || idx === undefined || !data) return jsonResponse({ error: 'Missing params' });
  const key = 'chunk_' + sessionId + '_' + fileType + '_' + idx;
  PropertiesService.getScriptProperties().setProperty(key, data);
  return jsonResponse({ ok: true });
}

// ─── Assemble all stored chunks and save to Drive ────────────────────────────
function handleFinalizeUpload(params) {
  const { sessionId, fileType, total } = params;
  if (!sessionId || !fileType || !total) return jsonResponse({ error: 'Missing params' });
  const props = PropertiesService.getScriptProperties();
  let b64 = '';
  for (let i = 0; i < Number(total); i++) {
    const key   = 'chunk_' + sessionId + '_' + fileType + '_' + i;
    const chunk = props.getProperty(key);
    if (!chunk) return jsonResponse({ error: 'Missing chunk ' + i });
    b64 += chunk;
    props.deleteProperty(key);
  }
  return handleSaveFile({ sessionId, fileType, data: b64 });
}

// ─── Save XLSX to Google Drive, return download URL ───────────────────────────
function handleSaveFile(params) {
  const { sessionId, fileType, data } = params;
  if (!sessionId || !fileType || !data) return jsonResponse({ error: 'Missing params' });

  const bytes    = Utilities.base64Decode(data);
  const filename = `Kushals_${fileType === 'error' ? 'Error' : 'Difference'}_Report_${sessionId}.xlsx`;
  const blob     = Utilities.newBlob(
    bytes,
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    filename
  );

  let folder;
  try {
    folder = DriveApp.getFolderById(REPORT_FOLDER_ID);
  } catch(_) {
    folder = DriveApp.getRootFolder();
  }

  // Remove old file with same name if exists
  const existing = folder.getFilesByName(filename);
  while (existing.hasNext()) existing.next().setTrashed(true);

  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  // Direct download URL
  const url = 'https://drive.google.com/uc?export=download&id=' + file.getId();

  // Update the session row in the sheet with the file URL
  const urlField = fileType === 'error' ? 'errorFileUrl' : 'diffFileUrl';
  const sheet    = getSheet();
  const headers  = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const col      = headers.indexOf(urlField);
  if (col >= 0) {
    const idCol = headers.indexOf('id');
    const data  = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][idCol] === sessionId) {
        sheet.getRange(i + 1, col + 1).setValue(url);
        break;
      }
    }
  }

  return jsonResponse({ url });
}

// ─── Clear all rows (keep header) ────────────────────────────────────────────
function handleClear() {
  const sheet    = getSheet();
  const lastRow  = sheet.getLastRow();
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
      'matched','withDiff','notScanned','errorRows','diffGeneratedAt',
      'diffFileUrl','errorFileUrl'
    ]);
  }
  return sheet;
}

function sessionToRow(s, headers) {
  const map = {
    id:                 s.id || '',
    timestamp:          s.timestamp || '',
    storeName:          s.storeName || '',
    systemFileName:     s.systemFile?.name || '',
    systemFileSize:     s.systemFile?.size || '',
    systemFileRows:     s.systemFile?.rows || '',
    uniqueBarcodes:     s.systemFile?.uniqueBarcodes ?? '',
    physFileName:       s.physicalScan?.name || '',
    physFileSize:       s.physicalScan?.size || '',
    physFileRows:       s.physicalScan?.rows || '',
    varianceRows:       s.varianceReport?.rows || '',
    varianceLocs:       s.varianceReport?.locations || '',
    varianceTotalQty:   s.varianceReport?.totalQty || '',
    varianceGeneratedAt:s.varianceReport?.generatedAt || '',
    matched:            s.differenceReport?.matched ?? '',
    withDiff:           s.differenceReport?.withDiff ?? '',
    notScanned:         s.differenceReport?.notScanned ?? '',
    errorRows:          s.differenceReport?.errorRows ?? '',
    diffGeneratedAt:    s.differenceReport?.generatedAt || '',
    diffFileUrl:        s.differenceReport?.diffFileUrl || '',
    errorFileUrl:       s.differenceReport?.errorFileUrl || '',
  };
  return headers.map(h => map[h] ?? '');
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
