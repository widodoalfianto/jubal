/**
 * Developer-only diagnostics helpers.
 */

function getDeveloperSheetNames() {
  return ['Execution Logs', 'Debug Responses'];
}

function isDeveloperDiagnosticsEnabled() {
  try {
    return PropertiesService.getScriptProperties().getProperty('jubalDeveloperDiagnostics') === 'true';
  } catch (error) {
    console.error('Unable to read jubalDeveloperDiagnostics property: ' + error.message);
    return false;
  }
}

function hideDeveloperSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let hiddenCount = 0;

  getDeveloperSheetNames().forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet || sheet.isSheetHidden()) return;
    try {
      sheet.hideSheet();
      hiddenCount++;
    } catch (error) {
      console.warn(`Could not hide developer sheet '${sheetName}': ${error.message}`);
    }
  });

  return { status: 'hidden', count: hiddenCount };
}

function showDeveloperSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let shownCount = 0;

  getDeveloperSheetNames().forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet || !sheet.isSheetHidden()) return;
    try {
      sheet.showSheet();
      shownCount++;
    } catch (error) {
      console.warn(`Could not show developer sheet '${sheetName}': ${error.message}`);
    }
  });

  return { status: 'shown', count: shownCount };
}

function enableDeveloperDiagnostics() {
  PropertiesService.getScriptProperties().setProperty('jubalDeveloperDiagnostics', 'true');
  return showDeveloperSheets();
}

function disableDeveloperDiagnostics() {
  PropertiesService.getScriptProperties().deleteProperty('jubalDeveloperDiagnostics');
  return hideDeveloperSheets();
}

/**
 * Centralized debug logger. Writes to console and a lightweight sheet for persistent logs.
 * level: 'info' | 'warn' | 'error'
 */
function logDebug(level, msg, data) {
  try {
    const payload = { ts: new Date().toISOString(), level: level, message: msg, data: data || null };
    console.log(JSON.stringify(payload));

    if (!isDeveloperDiagnosticsEnabled()) return;

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let logSheet = ss.getSheetByName('Execution Logs');
    if (!logSheet) {
      logSheet = ss.insertSheet('Execution Logs');
      logSheet.appendRow(['timestamp', 'level', 'message', 'data']);
    }

    const dataString = data ? (typeof data === 'string' ? data : JSON.stringify(data)) : '';
    logSheet.appendRow([payload.ts, payload.level, payload.message, dataString]);
  } catch (e) {
    console.error('logDebug failed: ' + e.message);
  }
}

/**
 * Append the incoming form event to a Debug Responses sheet for tracing.
 * Columns: timestamp | formId | responseRow | namedValues (JSON)
 */
function logFormResponse(e) {
  if (!isDeveloperDiagnosticsEnabled()) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let dbg = ss.getSheetByName('Debug Responses');
  if (!dbg) {
    dbg = ss.insertSheet('Debug Responses');
    dbg.appendRow(['timestamp', 'formId', 'responseRow', 'namedValues']);
  }

  const ts = new Date().toISOString();

  let formId = '';
  const meta = ss.getSheetByName(CONFIG.sheetNames.formMetadata);
  if (meta) {
    formId = (meta.getRange('B2').getValue() || meta.getRange('B1').getValue() || '').toString();
  }

  let responseRow = '';
  try {
    if (e && e.range && typeof e.range.getRow === 'function') responseRow = e.range.getRow();
  } catch (ignored) {}

  const namedValues = e && e.namedValues ? JSON.stringify(e.namedValues) : '';
  dbg.appendRow([ts, formId, responseRow, namedValues]);
}
