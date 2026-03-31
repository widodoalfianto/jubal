/**
 * 
 * Jubal: Google Apps Script for automating church music ministry scheduling.
 * - Updates availability matrix from form responses.
 * - Ensures all members are included.
 * - Highlights missing responses.
 * - Creates a new sheet every month for form responses.
 * - Supports multiple roles per minister.
 * - Tracks number of times each minister is willing to serve per month.
 * - Updates the database file whenever a form response is submitted.
 * - Auto-fills the bottom portion of the availability matrix based on responses.
 * 
 * - AUTHOR: Alfianto Widodo
 * 
 */

function onFormSubmit(e) {
  // Snapshot the raw form response for diagnostics before any mutation
  try {
    logFormResponse(e);
  } catch (err) {
    console.error('Failed to log form response: ' + err.message);
  }

  logDebug('info', 'onFormSubmit invoked', { namedValues: e && e.namedValues ? e.namedValues : null });
  updateDatabase(e);
}

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
    // Console log for quick inspection in executions
    console.log(JSON.stringify(payload));

    if (!isDeveloperDiagnosticsEnabled()) return;

    // Also append to an 'Execution Logs' sheet for persisted diagnostics
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let logSheet = ss.getSheetByName('Execution Logs');
    if (!logSheet) {
      logSheet = ss.insertSheet('Execution Logs');
      logSheet.appendRow(['timestamp', 'level', 'message', 'data']);
    }
    // Keep data small to avoid huge cells
    const dataString = data ? (typeof data === 'string' ? data : JSON.stringify(data)) : '';
    logSheet.appendRow([payload.ts, payload.level, payload.message, dataString]);
  } catch (e) {
    // If logging fails, fall back to console only.
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

  // Try to fetch a formId from the metadata sheet if present
  let formId = '';
  const meta = ss.getSheetByName(CONFIG.sheetNames.formMetadata);
  if (meta) {
    // Prefer B2 (current form) but fall back to B1
    formId = (meta.getRange('B2').getValue() || meta.getRange('B1').getValue() || '').toString();
  }

  // If the event includes a range (spreadsheet onFormSubmit), capture row
  let responseRow = '';
  try {
    if (e && e.range && typeof e.range.getRow === 'function') responseRow = e.range.getRow();
  } catch (ignored) {}

  const namedValues = e && e.namedValues ? JSON.stringify(e.namedValues) : '';
  dbg.appendRow([ts, formId, responseRow, namedValues]);
}

/**
 * Parse a free-form unavailable dates string into normalized MM/dd tokens.
 * Accepts inputs like "3/29", "03/29/2026", "Mar 29", "March 29 - afternoon", or ranges "3/29 - 4/5".
 * Returns an array of unique MM/dd strings.
 */
function parseUnavailableDates(raw) {
  const errors = [];
  if (!raw) return { parsed: [], errors };
  if (Array.isArray(raw)) raw = raw.join(',');
  raw = raw.toString();

  const parts = raw.split(',').map(s => s.trim()).filter(Boolean);
  const results = [];

  parts.forEach(part => {
    const rangeKeys = extractDateRangeKeys(part);
    if (rangeKeys) {
      results.push(rangeKeys[0]);
      results.push(rangeKeys[1]);
      return;
    }

    const dateKey = extractDateKey(part);
    if (dateKey) {
      results.push(dateKey);
    } else {
      errors.push(part);
    }
  });

  // Deduplicate while preserving order
  const parsed = results.filter((v, i, a) => a.indexOf(v) === i && v);
  return { parsed, errors };
}

function parseSingleDate(s) {
  if (!s) return null;
  s = s.toString().trim();

  // Match MM/DD or M/D or MM/DD/YYYY
  const m = s.match(/^(\d{1,2})\/(\d{1,2})(?:\/(\d{2,4}))?$/);
  if (m) {
    const month = parseInt(m[1], 10) - 1;
    const day = parseInt(m[2], 10);
    const year = m[3] ? parseInt(m[3], 10) : new Date().getFullYear();
    const dt = createStrictDate(year, month, day);
    return isNaN(dt.getTime()) ? null : dt;
  }

  // Match ISO YYYY-MM-DD
  const iso = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (iso) {
    const dt = createStrictDate(parseInt(iso[1], 10), parseInt(iso[2], 10) - 1, parseInt(iso[3], 10));
    return isNaN(dt.getTime()) ? null : dt;
  }

  // Try parsing with current year appended (for 'Mar 29' etc.)
  const withYear = s + ' ' + new Date().getFullYear();
  const parsed = new Date(withYear);
  if (!isNaN(parsed.getTime())) return parsed;

  // Last resort: try direct Date parse
  const direct = new Date(s);
  if (!isNaN(direct.getTime())) return direct;

  return null;
}

function formatDateMMDD(d) {
  const mm = ('0' + (d.getMonth() + 1)).slice(-2);
  const dd = ('0' + d.getDate()).slice(-2);
  return mm + '/' + dd;
}

function extractDateRangeKeys(value) {
  if (value === null || value === undefined || value === '') return null;
  const raw = String(value).trim();
  const rangeMatch = raw.match(/^(.+?)\s*-\s*(.+)$/);
  if (!rangeMatch) return null;

  const startKey = extractDateKey(rangeMatch[1]);
  const endKey = extractDateKey(rangeMatch[2]);
  if (!startKey || !endKey) return null;

  return [startKey, endKey];
}

function extractDateKey(value) {
  if (value === null || value === undefined || value === '') return '';
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    return formatDateMMDD(value);
  }

  const raw = String(value).trim();
  if (!raw) return '';

  const patterns = [
    /\b(\d{4}-\d{1,2}-\d{1,2})\b/,
    /\b(\d{1,2}\/\d{1,2}(?:\/\d{2,4})?)\b/,
    /\b([A-Za-z]+ \d{1,2}(?:, \d{4})?)\b/
  ];

  for (let i = 0; i < patterns.length; i++) {
    const match = raw.match(patterns[i]);
    if (!match) continue;

    const parsed = parseSingleDate(match[1]);
    if (parsed) return formatDateMMDD(parsed);
  }

  const parsed = parseSingleDate(raw);
  return parsed ? formatDateMMDD(parsed) : '';
}

function extractDateLabel(value) {
  if (value === null || value === undefined || value === '') return '';
  const raw = String(value).trim();
  if (!raw) return '';

  const patterns = [
    /^\d{4}-\d{1,2}-\d{1,2}\s*[-:]\s*(.+)$/,
    /^\d{1,2}\/\d{1,2}(?:\/\d{2,4})?\s*[-:]\s*(.+)$/,
    /^[A-Za-z]+ \d{1,2}(?:, \d{4})?\s*[-:]\s*(.+)$/
  ];

  for (let i = 0; i < patterns.length; i++) {
    const match = raw.match(patterns[i]);
    if (match && match[1]) return match[1].trim();
  }

  return '';
}

function normalizeDateChoice(value) {
  const dateKey = extractDateKey(value);
  if (!dateKey) return '';

  const label = extractDateLabel(value);
  return label ? `${dateKey} - ${label}` : dateKey;
}

function mergeDateChoices(choices) {
  const byDate = {};

  (choices || []).forEach(choice => {
    const normalized = normalizeDateChoice(choice);
    const dateKey = extractDateKey(normalized);
    if (!dateKey) return;

    if (!byDate[dateKey]) {
      byDate[dateKey] = {
        dateKey: dateKey,
        labels: []
      };
    }

    const label = extractDateLabel(normalized);
    if (label && byDate[dateKey].labels.indexOf(label) === -1) {
      byDate[dateKey].labels.push(label);
    }
  });

  return sortDateChoices(Object.keys(byDate).map(dateKey => {
    const labels = byDate[dateKey].labels;
    return labels.length ? `${dateKey} - ${labels.join(', ')}` : dateKey;
  }));
}

function sortDateChoices(choices) {
  return (choices || [])
    .map((choice, index) => ({ choice: normalizeDateChoice(choice) || String(choice).trim(), index: index, key: extractDateKey(choice) }))
    .sort((a, b) => {
      if (a.key && b.key && a.key !== b.key) return a.key.localeCompare(b.key);
      if (a.key && !b.key) return -1;
      if (!a.key && b.key) return 1;
      return a.index - b.index;
    })
    .map(item => item.choice);
}

function safeGetScriptTimeZone() {
  try {
    return Session.getScriptTimeZone() || CONFIG.defaults.timeZone;
  } catch (err) {
    return CONFIG.defaults.timeZone;
  }
}

function parseCsv(value) {
  if (Array.isArray(value)) return value.map(v => String(v).trim()).filter(Boolean);
  if (value === null || value === undefined) return [];
  return String(value)
    .split(',')
    .map(part => part.trim())
    .filter(Boolean);
}

function normalizeEmailList(values) {
  const seen = {};
  return parseCsv(values)
    .map(email => String(email || '').trim())
    .filter(email => {
      const normalized = email.toLowerCase();
      if (!normalized || seen[normalized]) return false;
      seen[normalized] = true;
      return true;
    });
}

function normalizeRoleList(values) {
  const seen = {};
  return parseCsv(values)
    .map(role => String(role || '').trim())
    .filter(role => {
      const normalized = role.toUpperCase();
      if (!normalized || seen[normalized]) return false;
      seen[normalized] = true;
      return true;
    });
}

function parseBooleanLike(value, fallback) {
  if (value === null || value === undefined || value === '') return fallback;
  if (typeof value === 'boolean') return value;
  if (typeof value === 'number') return value !== 0;

  const normalized = String(value).trim().toLowerCase();
  if (['true', 'yes', 'y', '1', 'on'].includes(normalized)) return true;
  if (['false', 'no', 'n', '0', 'off'].includes(normalized)) return false;
  return fallback;
}

function toIntegerOrDefault(value, fallback) {
  const parsed = parseInt(value, 10);
  return isNaN(parsed) ? fallback : parsed;
}

function clampDayOfMonthSetting(value, fallback) {
  const parsed = toIntegerOrDefault(value, fallback);
  if (parsed < 1) return 1;
  if (parsed > 28) return 28;
  return parsed;
}

function createStrictDate(year, monthIndex, dayOfMonth) {
  const date = new Date(year, monthIndex, dayOfMonth);
  if (isNaN(date.getTime())) return null;
  if (date.getFullYear() !== year || date.getMonth() !== monthIndex || date.getDate() !== dayOfMonth) return null;
  return date;
}

function isBlankRow(row) {
  return !row || row.every(cell => String(cell === null || cell === undefined ? '' : cell).trim() === '');
}

function buildHeaderMap(headerRow) {
  const map = {};
  (headerRow || []).forEach((header, index) => {
    const key = String(header || '').trim().toLowerCase();
    if (key) map[key] = index;
  });
  return map;
}

function getValueByHeader(row, headerMap, names, fallback) {
  for (let i = 0; i < names.length; i++) {
    const key = String(names[i]).trim().toLowerCase();
    if (Object.prototype.hasOwnProperty.call(headerMap, key)) {
      return row[headerMap[key]];
    }
  }
  return fallback;
}

function loadKeyValueSheet(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() < 2) return {};

  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, Math.min(2, sheet.getLastColumn())).getValues();
  const out = {};
  data.forEach(row => {
    const key = String(row[0] || '').trim().toLowerCase();
    if (!key) return;
    out[key] = row[1];
  });
  return out;
}

function loadAdminEmailsFromSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.sheetNames.admins);
  if (!sheet || sheet.getLastRow() < 2) return [];

  const headerMap = getSheetHeaderMap(sheet);
  if (!Object.prototype.hasOwnProperty.call(headerMap, 'email')) return [];

  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  const emails = rows
    .map(row => {
      const enabled = Object.prototype.hasOwnProperty.call(headerMap, 'enabled')
        ? parseBooleanLike(getValueByHeader(row, headerMap, ['Enabled'], false), false)
        : true;
      const email = String(getValueByHeader(row, headerMap, ['Email'], '') || '').trim();
      return enabled && email ? email : '';
    })
    .filter(Boolean);

  return normalizeEmailList(emails);
}

function loadRolesFromSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.sheetNames.rolesConfig);
  if (!sheet || sheet.getLastRow() < 2) return [];

  const headerMap = getSheetHeaderMap(sheet);
  if (!Object.prototype.hasOwnProperty.call(headerMap, 'role')) return [];

  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  const roles = rows
    .map(row => {
      const enabled = Object.prototype.hasOwnProperty.call(headerMap, 'enabled')
        ? parseBooleanLike(getValueByHeader(row, headerMap, ['Enabled'], false), false)
        : true;
      const role = String(getValueByHeader(row, headerMap, ['Role'], '') || '').trim();
      return enabled && role ? role : '';
    })
    .filter(Boolean);

  return normalizeRoleList(roles);
}

function getDefaultRuntimeSettings() {
  return {
    churchName: CONFIG.defaults.churchName,
    timeZone: safeGetScriptTimeZone(),
    formsFolder: CONFIG.ids.formsFolder || '',
    adminEmails: CONFIG.ids.adminEmails.slice(),
    roles: CONFIG.roles.slice(),
    formCreationDay: CONFIG.defaults.formCreationDay,
    timesChoices: CONFIG.defaults.timesChoices.slice(),
    adminReminderEnabled: CONFIG.defaults.adminReminderEnabled,
    adminReminderDay: CONFIG.defaults.adminReminderDay,
    eventsArchiveFrequency: CONFIG.defaults.eventsArchiveFrequency
  };
}

function loadRuntimeSettings() {
  const defaults = getDefaultRuntimeSettings();
  const raw = loadKeyValueSheet(CONFIG.sheetNames.settings);
  const sheetAdminEmails = loadAdminEmailsFromSheet();
  const sheetRoles = loadRolesFromSheet();
  if (!Object.keys(raw).length) {
    if (sheetAdminEmails.length) {
      const fallbackDefaults = Object.assign({}, defaults);
      fallbackDefaults.adminEmails = sheetAdminEmails;
      if (sheetRoles.length) fallbackDefaults.roles = sheetRoles;
      return fallbackDefaults;
    }
    if (sheetRoles.length) {
      const fallbackDefaults = Object.assign({}, defaults);
      fallbackDefaults.roles = sheetRoles;
      return fallbackDefaults;
    }
    return defaults;
  }

  const configuredAdminEmails = normalizeEmailList(raw.admin_emails);
  const configuredRoles = normalizeRoleList(raw.roles);

  const settings = {
    churchName: String(raw.church_name || defaults.churchName).trim() || defaults.churchName,
    timeZone: String(raw.time_zone || defaults.timeZone).trim() || defaults.timeZone,
    formsFolder: String(raw.forms_folder_id || defaults.formsFolder).trim() || defaults.formsFolder,
    adminEmails: sheetAdminEmails.length ? sheetAdminEmails : configuredAdminEmails,
    roles: sheetRoles.length ? sheetRoles : configuredRoles,
    formCreationDay: clampDayOfMonthSetting(raw.form_creation_day, defaults.formCreationDay),
    timesChoices: parseCsv(raw.times_choices),
    adminReminderEnabled: parseBooleanLike(raw.admin_reminder_enabled, defaults.adminReminderEnabled),
    adminReminderDay: clampDayOfMonthSetting(raw.admin_reminder_day, Math.max(clampDayOfMonthSetting(raw.form_creation_day, defaults.formCreationDay) - 3, 1)),
    eventsArchiveFrequency: String(raw.events_archive_frequency || defaults.eventsArchiveFrequency).trim() || defaults.eventsArchiveFrequency
  };

  if (!settings.adminEmails.length) settings.adminEmails = defaults.adminEmails.slice();
  if (!settings.roles.length) settings.roles = defaults.roles.slice();
  if (!settings.timesChoices.length) settings.timesChoices = defaults.timesChoices.slice();
  if (!settings.eventsArchiveFrequency) settings.eventsArchiveFrequency = defaults.eventsArchiveFrequency;

  return settings;
}

function getMinistryMemberNames() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const databaseSheet = ss.getSheetByName(CONFIG.sheetNames.ministryMembers);
  if (!databaseSheet || databaseSheet.getLastRow() < 2) return [];

  const memberColumns = getMinistryMembersColumnMap(databaseSheet);
  return databaseSheet
    .getRange(2, memberColumns.name, databaseSheet.getLastRow() - 1, 1)
    .getDisplayValues()
    .flat()
    .map(name => String(name || '').trim())
    .filter(Boolean);
}

function getCurrentFormMetadata(metadataSheet) {
  if (!metadataSheet) return { formName: '', formId: '' };

  const row2FormId = String(metadataSheet.getRange('B2').getValue() || '').trim();
  if (row2FormId) {
    return {
      formName: String(metadataSheet.getRange('A2').getValue() || '').trim(),
      formId: row2FormId
    };
  }

  const row1FormId = String(metadataSheet.getRange('B1').getValue() || '').trim();
  if (row1FormId && row1FormId.toLowerCase() !== 'form id') {
    return {
      formName: String(metadataSheet.getRange('A1').getValue() || '').trim(),
      formId: row1FormId
    };
  }

  return { formName: '', formId: '' };
}

function writeCurrentFormMetadata(metadataSheet, formName, formId) {
  if (!metadataSheet) return;

  const current = getCurrentFormMetadata(metadataSheet);
  metadataSheet.clearContents();

  if (current.formId && current.formId !== String(formId)) {
    metadataSheet.getRange(1, 1, 1, 2).setValues([[current.formName || 'Previous Form', current.formId]]);
  }

  metadataSheet.getRange(2, 1, 1, 2).setValues([[formName, String(formId)]]);
  applySheetTheme(metadataSheet);
  fitSheetToContent(metadataSheet);
  applyTableBordersToDataRange(metadataSheet);
}

function getAdminRecipientList(settings) {
  const runtimeSettings = settings || loadRuntimeSettings();
  return normalizeEmailList(runtimeSettings.adminEmails || []);
}

function getAdminRecipientString(settings) {
  return getAdminRecipientList(settings).join(',');
}

function sendEmailToAdmins(subject, body, settings) {
  const recipients = getAdminRecipientString(settings);
  if (!recipients) return false;
  MailApp.sendEmail(recipients, subject, body);
  return true;
}

function onOpen() {
  if (!isDeveloperDiagnosticsEnabled()) {
    try {
      hideDeveloperSheets();
    } catch (error) {
      console.warn('Could not hide developer sheets on open: ' + error.message);
    }
  }

  SpreadsheetApp.getUi()
    .createMenu('Scheduling')
    .addItem('Add Special Event', 'showAddEventDialog')
    .addSeparator()
    .addItem('Apply Event Changes to Next Month', 'menuApplyEventChangesToPlanningMonth')
    .addItem('Refresh Form Dates', 'menuSyncCurrentFormWithAvailability')
    .addItem('Refresh Availability Sheet', 'menuUpdateAvailability')
    .addToUi();
}

function getAddEventDialogContext() {
  const runtimeSettings = loadRuntimeSettings();
  const recurringEvents = getRecurringEventDropdownValues();
  const commonEvents = ['Good Friday', 'Easter', 'Christmas', 'Christmas Eve', 'Worship Night', 'Prayer Night'];
  const mergedCommonEvents = commonEvents.concat(recurringEvents).filter((value, index, values) => values.indexOf(value) === index);
  const today = new Date();

  return {
    today: Utilities.formatDate(today, runtimeSettings.timeZone || safeGetScriptTimeZone(), 'yyyy-MM-dd'),
    recurringEvents: recurringEvents,
    commonEvents: mergedCommonEvents
  };
}

function showAddEventDialog() {
  const template = HtmlService.createTemplateFromFile('AddEventDialog');
  template.dialogContext = JSON.stringify(getAddEventDialogContext());

  const html = template.evaluate()
    .setWidth(420)
    .setHeight(560)
    .setTitle('Add Special Event');

  SpreadsheetApp.getUi().showModalDialog(html, 'Add Special Event');
}

function addEventFromDialog(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let eventsSheet = ss.getSheetByName(CONFIG.sheetNames.events);

  if (!eventsSheet || !sheetUsesFriendlyEventsLayout(eventsSheet)) {
    initializeProject();
    eventsSheet = ss.getSheetByName(CONFIG.sheetNames.events);
  }

  if (!eventsSheet || !sheetUsesFriendlyEventsLayout(eventsSheet)) {
    throw new Error('The Events sheet is not ready yet. Please run initializeProject() and try again.');
  }

  const action = 'ADD';

  const dateValue = String((payload && payload.date) || '').trim();
  const eventName = String((payload && payload.event) || '').trim();
  const recurringEvent = String((payload && payload.recurringEvent) || '').trim();
  const notes = String((payload && payload.notes) || '').trim();
  const parsedDate = parseSingleDate(dateValue);

  if (!parsedDate) {
    throw new Error('Please choose a valid event date from the date field.');
  }
  if (!eventName) {
    throw new Error('Please enter an event name.');
  }

  const row = [
    true,
    parsedDate,
    eventName,
    action,
    recurringEvent,
    parseBooleanLike(payload && payload.includeInForm, true),
    parseBooleanLike(payload && payload.includeInSchedule, true),
    notes
  ];

  eventsSheet.appendRow(row);
  configureEventsSheetUi(eventsSheet);

  const rowIndex = eventsSheet.getLastRow();
  const endColumnLetter = columnToLetter(eventsSheet.getLastColumn());
  const rowRange = eventsSheet.getRange(rowIndex, 1, 1, eventsSheet.getLastColumn());
  ss.setActiveSheet(eventsSheet);
  ss.setActiveRange(rowRange);

  return {
    rowIndex: rowIndex,
    eventName: eventName,
    dateDisplay: Utilities.formatDate(parsedDate, safeGetScriptTimeZone(), 'EEE, MMM d, yyyy'),
    sheetUrl: getSheetRangeUrl(eventsSheet, rowIndex, 'A', endColumnLetter)
  };
}

function getPlanningMonthContext(referenceDate, settings) {
  const runtimeSettings = settings || loadRuntimeSettings();
  const baseDate = referenceDate ? new Date(referenceDate) : new Date();
  const planDate = new Date(baseDate);
  const timeZone = runtimeSettings.timeZone || safeGetScriptTimeZone();
  planDate.setMonth(baseDate.getMonth() + 1);

  return {
    referenceDate: baseDate,
    planDate: planDate,
    planYear: planDate.getFullYear(),
    planMonth: planDate.getMonth(),
    planMonthName: Utilities.formatDate(planDate, timeZone, 'MMMM'),
    sheetName: getAvailabilitySheetName(planDate.getFullYear(), planDate.getMonth(), runtimeSettings),
    timeZone: timeZone
  };
}

function captureScheduleAssignments(sheet, roles) {
  const assignments = {};
  if (!sheet || !roles || !roles.length || sheet.getLastRow() < 2 || sheet.getLastColumn() <= 1) return assignments;

  const headers = sheet.getRange(CONFIG.layout.dateRowIndex, 2, 1, sheet.getLastColumn() - 1).getDisplayValues()[0];
  const rowCount = Math.min(roles.length, Math.max(sheet.getLastRow() - 1, 0));
  if (rowCount <= 0) return assignments;

  const rows = sheet.getRange(2, 1, rowCount, sheet.getLastColumn()).getDisplayValues();
  rows.forEach((row, rowIndex) => {
    const roleName = String(row[0] || roles[rowIndex] || '').trim().toUpperCase();
    if (!roleName) return;

    headers.forEach((header, columnIndex) => {
      const dateKey = extractDateKey(header);
      const value = String(row[columnIndex + 1] || '').trim();
      if (!dateKey || !value) return;
      assignments[`${roleName}|${dateKey}`] = value;
    });
  });

  return assignments;
}

function restoreScheduleAssignments(sheet, assignments, roles) {
  if (!sheet || !roles || !roles.length || !assignments || !Object.keys(assignments).length || sheet.getLastRow() < 2 || sheet.getLastColumn() <= 1) {
    return 0;
  }

  const rowCount = Math.min(roles.length, Math.max(sheet.getLastRow() - 1, 0));
  if (rowCount <= 0) return 0;

  const headers = sheet.getRange(CONFIG.layout.dateRowIndex, 2, 1, sheet.getLastColumn() - 1).getDisplayValues()[0];
  const rows = sheet.getRange(2, 1, rowCount, sheet.getLastColumn()).getValues();
  let restoredCount = 0;

  rows.forEach((row, rowIndex) => {
    const roleName = String(row[0] || roles[rowIndex] || '').trim().toUpperCase();
    if (!roleName) return;

    headers.forEach((header, columnIndex) => {
      const dateKey = extractDateKey(header);
      if (!dateKey) return;
      const preservedValue = assignments[`${roleName}|${dateKey}`];
      if (!preservedValue) return;
      row[columnIndex + 1] = preservedValue;
      restoredCount++;
    });
  });

  sheet.getRange(2, 1, rowCount, sheet.getLastColumn()).setValues(rows);
  return restoredCount;
}

function buildMenuAlertLines(result) {
  if (!result) return ['The action finished without returning details.'];

  switch (result.status) {
    case 'missing_sheet':
      return [
        `Next month has not been generated yet.`,
        `The sheet "${result.sheetName}" does not exist yet. Run the normal monthly setup first, then use this action for later date changes.`
      ];
    case 'missing_metadata':
      return [
        'The Form Metadata sheet was not found.',
        'The workbook was updated, but the form could not be refreshed automatically.'
      ];
    case 'missing_form_id':
      return [
        'No current form was found in Form Metadata.',
        'The workbook was updated, but the form dates could not be refreshed automatically.'
      ];
    case 'missing_dates_question':
      return [
        'The current form does not have the unavailable-dates question.',
        'Please check the form structure before trying again.'
      ];
    case 'missing_database':
      return [
        'The Ministry Members sheet was not found.',
        'Please run initializeProject() or restore the sheet before trying again.'
      ];
    case 'updated':
      return [
        `Updated the availability list in "${result.sheetName}".`,
        `${result.dateCount || 0} date column(s) and ${result.roleCount || 0} role row(s) were refreshed.`
      ];
    case 'synced':
      return [
        `Updated the date choices in the current form from "${result.sheetName}".`,
        `${result.choiceCount || 0} date choice(s) are now on the form.`
      ];
    case 'completed': {
      const lines = [
        `Updated "${result.sheetName}" from Recurring and Events.`,
        `Preserved ${result.restoredAssignments || 0} scheduled cell(s) for dates that still exist.`
      ];

      if (result.formSync) {
        if (result.formSync.status === 'synced') {
          lines.push(`Refreshed the current form with ${result.formSync.choiceCount || 0} date choice(s).`);
        } else if (result.formSync.status !== 'skipped') {
          lines.push(`The workbook was updated, but the form could not be refreshed automatically.`);
        }
      }

      if (result.availabilityUpdate && result.availabilityUpdate.status === 'updated') {
        lines.push(`Refreshed the availability section for ${result.availabilityUpdate.roleCount || 0} role row(s).`);
      }

      if (result.addedDates && result.addedDates.length) {
        lines.push(`New dates added: ${result.addedDates.join(', ')}. Ask members to confirm availability for these new dates.`);
      }
      if (result.removedDates && result.removedDates.length) {
        lines.push(`Removed dates: ${result.removedDates.join(', ')}.`);
      }

      return lines;
    }
    default:
      return [String(result.message || 'The action finished.')];
  }
}

function showMenuAlert(title, result) {
  const ui = SpreadsheetApp.getUi();
  ui.alert(title, buildMenuAlertLines(result).join('\n\n'), ui.ButtonSet.OK);
}

function menuSyncCurrentFormWithAvailability() {
  const ui = SpreadsheetApp.getUi();
  try {
    const result = syncCurrentFormWithAvailability();
    showMenuAlert('Form Dates Refreshed', result);
    return result;
  } catch (error) {
    ui.alert('Could Not Refresh Form Dates', error.message, ui.ButtonSet.OK);
    throw error;
  }
}

function menuUpdateAvailability() {
  const ui = SpreadsheetApp.getUi();
  try {
    const result = updateAvailability();
    showMenuAlert('Availability Sheet Refreshed', result);
    return result;
  } catch (error) {
    ui.alert('Could Not Refresh Availability Sheet', error.message, ui.ButtonSet.OK);
    throw error;
  }
}

function menuApplyEventChangesToPlanningMonth() {
  const ui = SpreadsheetApp.getUi();
  try {
    const result = applyEventChangesToPlanningMonth();
    showMenuAlert('Next Month Updated', result);
    return result;
  } catch (error) {
    ui.alert('Could Not Apply Event Changes', error.message, ui.ButtonSet.OK);
    throw error;
  }
}

function formatHumanDateTime(value, timeZone) {
  if (!value) return '';
  const parsed = Object.prototype.toString.call(value) === '[object Date]' ? value : new Date(value);
  if (isNaN(parsed.getTime())) return String(value);
  return Utilities.formatDate(parsed, timeZone || safeGetScriptTimeZone(), "MMMM d, yyyy 'at' h:mm a z");
}

function formatDayOfMonthHuman(day) {
  const parsed = toIntegerOrDefault(day, 0);
  if (parsed <= 0) return '';
  const mod100 = parsed % 100;
  if (mod100 >= 11 && mod100 <= 13) return `${parsed}th`;
  const mod10 = parsed % 10;
  if (mod10 === 1) return `${parsed}st`;
  if (mod10 === 2) return `${parsed}nd`;
  if (mod10 === 3) return `${parsed}rd`;
  return `${parsed}th`;
}

function getDayOfMonthDropdownValues() {
  const values = [];
  for (let i = 1; i <= 28; i++) values.push(String(i));
  return values;
}

function getMinistryMembersColumnMap(sheet) {
  const defaults = {
    name: 1,
    dates: 2,
    times: 3,
    comments: 4,
    roles: 5,
    canonicalName: 6,
    firstRoleCheckbox: 6
  };

  if (!sheet || sheet.getLastColumn() < 1) return defaults;

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  const headerMap = buildHeaderMap(headers);
  const getColumn = (header, fallback) => {
    const key = String(header || '').trim().toLowerCase();
    return Object.prototype.hasOwnProperty.call(headerMap, key) ? headerMap[key] + 1 : fallback;
  };

  const rolesColumn = getColumn(CONFIG.sheetHeaders.roles, defaults.roles);
  const canonicalColumn = getColumn(CONFIG.sheetHeaders.canonicalName, Math.max(sheet.getLastColumn(), defaults.canonicalName));

  return {
    name: getColumn(CONFIG.sheetHeaders.name, defaults.name),
    dates: getColumn(CONFIG.sheetHeaders.dates, defaults.dates),
    times: getColumn(CONFIG.sheetHeaders.times, defaults.times),
    comments: getColumn(CONFIG.sheetHeaders.comments, defaults.comments),
    roles: rolesColumn,
    canonicalName: canonicalColumn,
    firstRoleCheckbox: rolesColumn + 1
  };
}

function fitSheetToContent(sheet) {
  if (!sheet || sheet.getLastRow() < 1 || sheet.getLastColumn() < 1) return;
  const range = sheet.getDataRange();
  range.setWrap(false);
  sheet.autoResizeColumns(1, sheet.getLastColumn());
  sheet.autoResizeRows(1, sheet.getLastRow());
}

function applyTableBorder(sheet, startRow, startColumn, numRows, numColumns) {
  if (!sheet || numRows <= 0 || numColumns <= 0) return;
  sheet
    .getRange(startRow, startColumn, numRows, numColumns)
    .setBorder(true, true, true, true, true, true, '#666666', SpreadsheetApp.BorderStyle.SOLID);
}

function applyTableBordersToDataRange(sheet) {
  if (!sheet || sheet.getLastRow() < 1 || sheet.getLastColumn() < 1) return;
  applyTableBorder(sheet, 1, 1, sheet.getLastRow(), sheet.getLastColumn());
}

function getMonthSheetNames() {
  return [
    'January', 'February', 'March', 'April', 'May', 'June',
    'July', 'August', 'September', 'October', 'November', 'December'
  ];
}

function isMonthSheetName(sheetName) {
  return getMonthSheetNames().indexOf(String(sheetName || '').trim()) !== -1;
}

function getMonthSheetSortMetadata(sheetName, referenceDate) {
  const monthIndex = getMonthSheetNames().indexOf(String(sheetName || '').trim());
  const today = referenceDate ? new Date(referenceDate) : new Date();
  const currentMonth = today.getMonth();
  const forwardDistance = (monthIndex - currentMonth + 12) % 12;
  const backwardDistance = (currentMonth - monthIndex + 12) % 12;

  if (forwardDistance === 0) {
    return { group: 1, rank: 0 };
  }

  if (forwardDistance > 0 && forwardDistance <= 6) {
    return { group: 0, rank: forwardDistance };
  }

  return { group: 2, rank: backwardDistance };
}

function sortMonthSheetsByRecency(sheets, referenceDate) {
  return (sheets || []).slice().sort((left, right) => {
    const a = getMonthSheetSortMetadata(left.getName(), referenceDate);
    const b = getMonthSheetSortMetadata(right.getName(), referenceDate);

    if (a.group !== b.group) return a.group - b.group;
    if (a.rank !== b.rank) return a.rank - b.rank;
    return left.getName().localeCompare(right.getName());
  });
}

function reorderWorkbookSheets(referenceDate) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const originalActiveSheet = ss.getActiveSheet();
  const allSheets = ss.getSheets();
  const orderedNames = [];

  const pushSheet = sheet => {
    if (!sheet) return;
    const name = sheet.getName();
    if (orderedNames.indexOf(name) === -1) {
      orderedNames.push(name);
    }
  };

  pushSheet(ss.getSheetByName(CONFIG.sheetNames.ministryMembers));
  sortMonthSheetsByRecency(allSheets.filter(sheet => isMonthSheetName(sheet.getName())), referenceDate).forEach(pushSheet);
  pushSheet(ss.getSheetByName(CONFIG.sheetNames.recurring));
  pushSheet(ss.getSheetByName(CONFIG.sheetNames.events));
  allSheets
    .filter(sheet => sheet.getName().startsWith('Form Responses'))
    .forEach(pushSheet);
  pushSheet(ss.getSheetByName(CONFIG.sheetNames.rolesConfig));
  pushSheet(ss.getSheetByName(CONFIG.sheetNames.admins));
  pushSheet(ss.getSheetByName(CONFIG.sheetNames.settings));
  pushSheet(ss.getSheetByName(CONFIG.sheetNames.formMetadata));

  allSheets.forEach(pushSheet);

  orderedNames.forEach((sheetName, index) => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return;
    ss.setActiveSheet(sheet);
    ss.moveActiveSheet(index + 1);
  });

  if (originalActiveSheet) {
    try {
      ss.setActiveSheet(originalActiveSheet);
    } catch (error) {
      console.warn('Unable to restore active sheet after reordering: ' + error.message);
    }
  }
}

function highlightExampleRows(sheet, notesColumnNumber) {
  if (!sheet || !notesColumnNumber || sheet.getLastRow() < 2 || sheet.getLastColumn() < 1) return;

  const width = sheet.getLastColumn();
  const noteValues = sheet.getRange(2, notesColumnNumber, sheet.getLastRow() - 1, 1).getDisplayValues().flat();

  noteValues.forEach((value, index) => {
    const rowNumber = index + 2;
    const isExample = String(value || '').trim().toLowerCase().indexOf('example') === 0;
    const rowRange = sheet.getRange(rowNumber, 1, 1, width);
    if (isExample) {
      rowRange
        .setBackground('#fff2cc')
        .setFontStyle('italic');
    } else {
      rowRange
        .setBackground(null)
        .setFontStyle('normal');
    }
  });
}

function highlightAdminRows(sheet, emailColumnNumber, notesColumnNumber) {
  if (!sheet || sheet.getLastRow() < 2 || sheet.getLastColumn() < 1) return;

  const width = sheet.getLastColumn();
  const starterEmails = normalizeEmailList(CONFIG.ids.adminEmails || []);
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, width).getDisplayValues();

  rows.forEach((row, index) => {
    const rowNumber = index + 2;
    const email = String(row[emailColumnNumber - 1] || '').trim();
    const note = String(row[notesColumnNumber - 1] || '').trim().toLowerCase();
    const isExample = note.indexOf('example') === 0 || starterEmails.indexOf(email) !== -1;
    const rowRange = sheet.getRange(rowNumber, 1, 1, width);

    if (isExample) {
      rowRange
        .setBackground('#fff2cc')
        .setFontStyle('italic');
    } else {
      rowRange
        .setBackground(null)
        .setFontStyle('normal');
    }
  });
}

function promoteExampleRows(sheet, notesColumnNumber) {
  if (!sheet || !notesColumnNumber || sheet.getLastRow() < 3 || sheet.getLastColumn() < 1) return;

  const width = sheet.getLastColumn();
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, width).getValues();
  const exampleRows = [];
  const otherRows = [];

  rows.forEach(row => {
    if (isBlankRow(row)) return;
    const isExample = String(row[notesColumnNumber - 1] || '').trim().toLowerCase().indexOf('example') === 0;
    if (isExample) {
      exampleRows.push(row);
    } else {
      otherRows.push(row);
    }
  });

  const orderedRows = exampleRows.concat(otherRows);
  if (!orderedRows.length) return;

  sheet.getRange(2, 1, Math.max(sheet.getMaxRows() - 1, 1), width).clearContent();
  sheet.getRange(2, 1, orderedRows.length, width).setValues(orderedRows);
}

function getTimeZoneOptions() {
  return normalizeRoleList([
    safeGetScriptTimeZone(),
    'UTC',
    'America/Los_Angeles',
    'America/Denver',
    'America/Phoenix',
    'America/Chicago',
    'America/New_York',
    'America/Anchorage',
    'Pacific/Honolulu',
    'America/Toronto',
    'America/Vancouver',
    'America/Mexico_City',
    'America/Sao_Paulo',
    'Europe/London',
    'Europe/Dublin',
    'Europe/Paris',
    'Europe/Berlin',
    'Europe/Madrid',
    'Europe/Rome',
    'Africa/Johannesburg',
    'Asia/Dubai',
    'Asia/Kolkata',
    'Asia/Bangkok',
    'Asia/Singapore',
    'Asia/Manila',
    'Asia/Hong_Kong',
    'Asia/Tokyo',
    'Asia/Seoul',
    'Australia/Sydney',
    'Pacific/Auckland'
  ]);
}

function getThemeForSheet(sheetName) {
  if (!sheetName) return null;
  const themeKey = Object.keys(CONFIG.sheetNames).find(key => CONFIG.sheetNames[key] === sheetName);
  return themeKey && CONFIG.themes[themeKey] ? CONFIG.themes[themeKey] : null;
}

function applySheetTheme(sheet) {
  if (!sheet) return;
  const theme = getThemeForSheet(sheet.getName());
  if (!theme) return;

  try {
    sheet.setTabColor(theme.tab || null);
  } catch (err) {
    console.warn('Unable to set tab color for ' + sheet.getName() + ': ' + err.message);
  }

  if (sheet.getLastColumn() < 1) return;
  const headerRange = sheet.getRange(1, 1, 1, sheet.getLastColumn());
  headerRange
    .setBackground(theme.header || null)
    .setFontColor(theme.text || '#000000')
    .setFontWeight('bold');
}

function getSheetRangeUrl(sheet, rowIndex, startColumnLetter, endColumnLetter) {
  if (!sheet) return '';
  const ss = sheet.getParent();
  const start = `${startColumnLetter || 'A'}${rowIndex || 1}`;
  const end = `${endColumnLetter || startColumnLetter || 'A'}${rowIndex || 1}`;
  return `${ss.getUrl()}#gid=${sheet.getSheetId()}&range=${start}:${end}`;
}

function getSheetUrlByName(sheetName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  return sheet ? getSheetRangeUrl(sheet, 1, 'A', 'A') : SpreadsheetApp.getActiveSpreadsheet().getUrl();
}

function getAvailabilitySheetNameForMonthName(monthName, settings) {
  return String(monthName || '').trim();
}

function getAvailabilitySheetName(year, month, settings) {
  const monthName = new Date(year, month, 1).toLocaleString('default', { month: 'long' });
  return getAvailabilitySheetNameForMonthName(monthName, settings);
}

function normalizeWeekday(value) {
  if (value === null || value === undefined || value === '') return null;
  const weekdayMap = {
    sunday: 0,
    sun: 0,
    monday: 1,
    mon: 1,
    tuesday: 2,
    tue: 2,
    tues: 2,
    wednesday: 3,
    wed: 3,
    thursday: 4,
    thu: 4,
    thur: 4,
    thurs: 4,
    friday: 5,
    fri: 5,
    saturday: 6,
    sat: 6
  };
  const normalized = String(value).trim().toLowerCase();
  return Object.prototype.hasOwnProperty.call(weekdayMap, normalized) ? weekdayMap[normalized] : null;
}

function normalizeOrdinal(value) {
  if (value === null || value === undefined || value === '') return null;
  const normalized = String(value).trim().toLowerCase();
  const map = {
    first: 1,
    second: 2,
    third: 3,
    fourth: 4,
    fifth: 5,
    last: -1,
    every: 'every'
  };
  if (Object.prototype.hasOwnProperty.call(map, normalized)) return map[normalized];

  const parsed = parseInt(normalized, 10);
  return isNaN(parsed) ? null : parsed;
}

function parseMonthNumber(value) {
  if (value === null || value === undefined || value === '') return null;
  if (typeof value === 'number' && !isNaN(value)) return value;

  const raw = String(value).trim();
  const asNumber = parseInt(raw, 10);
  if (!isNaN(asNumber)) return asNumber;

  const monthName = new Date(Date.parse(raw + ' 1, 2000'));
  return isNaN(monthName.getTime()) ? null : monthName.getMonth() + 1;
}

function monthMatchesFilter(filter, month) {
  if (filter === null || filter === undefined || filter === '') return true;
  const normalized = String(filter).trim().toLowerCase();
  if (!normalized || normalized === 'all') return true;

  const parts = normalized.split(',').map(part => part.trim()).filter(Boolean);
  return parts.some(part => parseMonthNumber(part) === month + 1);
}

function normalizeRecurrence(value, fallback) {
  if (value === null || value === undefined || value === '') return fallback || 'monthly';
  const normalized = String(value).trim().toLowerCase();
  if (normalized === 'weekly' || normalized === 'monthly' || normalized === 'yearly') return normalized;
  return fallback || 'monthly';
}

function inferRuleRecurrence(ruleType, monthFilter) {
  if (ruleType === 'easter_offset') return 'yearly';
  const normalizedMonthFilter = String(monthFilter || '').trim().toLowerCase();
  if (normalizedMonthFilter && normalizedMonthFilter !== 'all') return 'yearly';
  return 'monthly';
}

function normalizeFrequency(value) {
  if (value === null || value === undefined || value === '') return '';
  const normalized = String(value).trim().toLowerCase();
  if (['weekly', 'monthly', 'yearly', 'easter'].includes(normalized)) return normalized;
  return '';
}

function normalizeArchiveFrequency(value) {
  if (value === null || value === undefined || value === '') return 'yearly';
  const normalized = String(value).trim().toLowerCase();
  if (['off', 'monthly', 'quarterly', 'yearly'].includes(normalized)) return normalized;
  return 'yearly';
}

function getArchiveTriggerMonths(frequency) {
  const normalizedFrequency = normalizeArchiveFrequency(frequency);

  if (normalizedFrequency === 'off') return [];
  if (normalizedFrequency === 'monthly') return [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12];
  if (normalizedFrequency === 'quarterly') return [1, 4, 7, 10];
  return [1];
}

function shouldArchiveEventsNow(referenceDate, settings) {
  const runtimeSettings = settings || loadRuntimeSettings();
  const triggerMonths = getArchiveTriggerMonths(runtimeSettings.eventsArchiveFrequency);
  const now = referenceDate || new Date();
  const currentMonth = now.getMonth() + 1;
  return now.getDate() === 1 && triggerMonths.indexOf(currentMonth) !== -1;
}

function slugifyRuleId(value, fallback) {
  const normalized = String(value || '')
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '_')
    .replace(/^_+|_+$/g, '');

  return normalized || fallback;
}

function getSheetHeaderMap(sheet) {
  if (!sheet || sheet.getLastColumn() < 1) return {};
  const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  return buildHeaderMap(headerRow);
}

function sheetLooksLikeRecurring(sheet) {
  if (!sheet) return false;
  const headerMap = getSheetHeaderMap(sheet);
  if (Object.prototype.hasOwnProperty.call(headerMap, 'frequency')) return true;
  if (Object.prototype.hasOwnProperty.call(headerMap, 'rule type')) return true;
  if (Object.prototype.hasOwnProperty.call(headerMap, 'week of month')) return true;
  if (Object.prototype.hasOwnProperty.call(headerMap, 'weekday') &&
      !Object.prototype.hasOwnProperty.call(headerMap, 'action')) return true;
  return false;
}

function sheetLooksLikeEventOverrides(sheet) {
  if (!sheet) return false;
  const headerMap = getSheetHeaderMap(sheet);
  if (!Object.prototype.hasOwnProperty.call(headerMap, 'date')) return false;
  return (
    Object.prototype.hasOwnProperty.call(headerMap, 'action') ||
    Object.prototype.hasOwnProperty.call(headerMap, 'event') ||
    Object.prototype.hasOwnProperty.call(headerMap, 'label') ||
    Object.prototype.hasOwnProperty.call(headerMap, 'year') ||
    Object.prototype.hasOwnProperty.call(headerMap, 'month')
  );
}

function sheetUsesFriendlyRecurringLayout(sheet) {
  const headerMap = getSheetHeaderMap(sheet);
  return Object.prototype.hasOwnProperty.call(headerMap, 'event') &&
    Object.prototype.hasOwnProperty.call(headerMap, 'frequency') &&
    Object.prototype.hasOwnProperty.call(headerMap, 'week of month');
}

function sheetUsesFriendlyEventsLayout(sheet) {
  const headerMap = getSheetHeaderMap(sheet);
  return Object.prototype.hasOwnProperty.call(headerMap, 'event') &&
    Object.prototype.hasOwnProperty.call(headerMap, 'action') &&
    Object.prototype.hasOwnProperty.call(headerMap, 'recurring event');
}

function sheetUsesFriendlyAdminsLayout(sheet) {
  const headerMap = getSheetHeaderMap(sheet);
  return Object.prototype.hasOwnProperty.call(headerMap, 'enabled') &&
    Object.prototype.hasOwnProperty.call(headerMap, 'email');
}

function sheetUsesFriendlyRolesLayout(sheet) {
  const headerMap = getSheetHeaderMap(sheet);
  return Object.prototype.hasOwnProperty.call(headerMap, 'enabled') &&
    Object.prototype.hasOwnProperty.call(headerMap, 'role');
}

function getConfiguredRecurringSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const recurringSheet = ss.getSheetByName(CONFIG.sheetNames.recurring);
  const oldEventsSheet = ss.getSheetByName(CONFIG.sheetNames.events);
  const legacySheet = ss.getSheetByName(CONFIG.sheetNames.recurringEvents);

  if (recurringSheet && recurringSheet.getLastRow() > 1) return recurringSheet;
  if (oldEventsSheet && oldEventsSheet.getLastRow() > 1 && sheetLooksLikeRecurring(oldEventsSheet)) return oldEventsSheet;
  if (legacySheet && legacySheet.getLastRow() > 1) return legacySheet;
  return recurringSheet || (sheetLooksLikeRecurring(oldEventsSheet) ? oldEventsSheet : null) || legacySheet || null;
}

function getConfiguredEventsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const preferredSheet = ss.getSheetByName(CONFIG.sheetNames.events);
  const legacySheet = ss.getSheetByName(CONFIG.sheetNames.monthlyEvents);

  if (preferredSheet && preferredSheet.getLastRow() > 1 && sheetLooksLikeEventOverrides(preferredSheet)) return preferredSheet;
  if (legacySheet && legacySheet.getLastRow() > 1) return legacySheet;
  if (preferredSheet && preferredSheet.getLastRow() <= 1) return preferredSheet;
  return legacySheet || null;
}

function getNthWeekdayOfMonth(year, month, weekday, ordinal) {
  if (weekday === null || weekday === undefined || ordinal === null || ordinal === undefined || ordinal === 'every') {
    return null;
  }

  if (ordinal === -1) {
    const date = new Date(year, month + 1, 0);
    while (date.getDay() !== weekday) {
      date.setDate(date.getDate() - 1);
    }
    return date;
  }

  const firstMatch = new Date(year, month, 1);
  while (firstMatch.getDay() !== weekday) {
    firstMatch.setDate(firstMatch.getDate() + 1);
  }

  firstMatch.setDate(firstMatch.getDate() + ((ordinal - 1) * 7));
  return firstMatch.getMonth() === month ? firstMatch : null;
}

function getAllWeekdaysInMonth(year, month, weekday) {
  if (weekday === null || weekday === undefined) return [];

  const dates = [];
  const current = new Date(year, month, 1);
  while (current.getMonth() === month) {
    if (current.getDay() === weekday) {
      dates.push(new Date(current));
    }
    current.setDate(current.getDate() + 1);
  }
  return dates;
}

function toIsoDate(date, timeZone) {
  return Utilities.formatDate(date, timeZone || safeGetScriptTimeZone(), 'yyyy-MM-dd');
}

function toEventObject(date, rule, timeZone, source) {
  const isoDate = toIsoDate(date, timeZone);
  const keySuffix = rule.ruleId || rule.label || rule.ruleType || 'event';

  return {
    key: `${isoDate}|${keySuffix}`,
    isoDate: isoDate,
    mmdd: formatDateMMDD(date),
    label: rule.label || '',
    type: rule.type || '',
    ruleId: rule.ruleId || '',
    includeInForm: rule.includeInForm !== false,
    includeInSchedule: rule.includeInSchedule !== false,
    sortOrder: toIntegerOrDefault(rule.sortOrder, 100),
    source: source || 'recurring'
  };
}

function mergeEventObjects(events) {
  const map = {};
  (events || []).forEach(event => {
    if (!event || !event.key) return;
    map[event.key] = event;
  });
  return Object.keys(map).map(key => map[key]);
}

function sortEvents(events) {
  return (events || []).slice().sort((a, b) => {
    if (a.isoDate !== b.isoDate) return a.isoDate.localeCompare(b.isoDate);
    if (a.sortOrder !== b.sortOrder) return a.sortOrder - b.sortOrder;
    return String(a.label || '').localeCompare(String(b.label || ''));
  });
}

function formatEventChoice(event) {
  return normalizeDateChoice(event.mmdd + (event.label ? ' - ' + event.label : ''));
}

function getAutomaticSpecialEvents(year, month, timeZone) {
  const events = [];
  const easter = computeEaster(year);
  if (easter.getMonth() === month) {
    events.push(toEventObject(easter, {
      ruleId: 'easter',
      label: 'Easter',
      type: 'special',
      includeInForm: true,
      includeInSchedule: true,
      sortOrder: 30
    }, timeZone, 'auto'));
  }

  if (month === 11) {
    events.push(toEventObject(new Date(year, 11, 25), {
      ruleId: 'christmas',
      label: 'Christmas',
      type: 'special',
      includeInForm: true,
      includeInSchedule: true,
      sortOrder: 40
    }, timeZone, 'auto'));
  }

  return events;
}

function getDefaultRecurringRules() {
  return [
    {
      enabled: true,
      ruleId: 'sunday_service',
      label: '',
      recurrence: 'monthly',
      ruleType: 'every_weekday',
      weekday: 0,
      ordinal: 'every',
      monthFilter: 'all',
      dayOfMonth: null,
      offsetDays: 0,
      includeInForm: true,
      includeInSchedule: true,
      sortOrder: 20,
      type: 'service'
    }
  ];
}

function hasConfiguredRecurringRules() {
  const sheet = getConfiguredRecurringSheet();
  return !!(sheet && sheet.getLastRow() > 1);
}

function normalizeRecurringRule(row, headerMap, fallbackRuleId) {
  let ruleType = String(getValueByHeader(row, headerMap, ['Rule Type'], '') || '').trim().toLowerCase();
  let recurrence;
  let monthFilter = String(getValueByHeader(row, headerMap, ['Month', 'Month Filter'], 'all') || 'all').trim() || 'all';
  let weekday = normalizeWeekday(getValueByHeader(row, headerMap, ['Weekday'], ''));
  let ordinal = normalizeOrdinal(getValueByHeader(row, headerMap, ['Week Of Month', 'Ordinal'], ''));
  let dayOfMonth = toIntegerOrDefault(getValueByHeader(row, headerMap, ['Day', 'Day Of Month'], ''), null);
  let offsetDays = toIntegerOrDefault(getValueByHeader(row, headerMap, ['Offset Days'], 0), 0);

  if (!ruleType) {
    const frequency = normalizeFrequency(getValueByHeader(row, headerMap, ['Frequency', 'Recurrence'], ''));
    if (!frequency) return null;

    if (frequency === 'weekly') {
      ruleType = 'every_weekday';
      recurrence = 'weekly';
      ordinal = 'every';
      monthFilter = 'all';
    } else if (frequency === 'monthly') {
      recurrence = 'monthly';
      if (dayOfMonth) {
        ruleType = 'fixed_date';
      } else if (weekday !== null && ordinal && ordinal !== 'every') {
        ruleType = 'nth_weekday';
      } else if (weekday !== null) {
        ruleType = 'every_weekday';
        ordinal = 'every';
      }
    } else if (frequency === 'yearly') {
      recurrence = 'yearly';
      if (dayOfMonth) {
        ruleType = 'fixed_date';
      } else if (weekday !== null && ordinal) {
        ruleType = 'nth_weekday';
      }
    } else if (frequency === 'easter') {
      ruleType = 'easter_offset';
      recurrence = 'yearly';
      monthFilter = 'all';
    }
  }

  if (!ruleType) return null;

  recurrence = normalizeRecurrence(
    recurrence || getValueByHeader(row, headerMap, ['Recurrence'], ''),
    inferRuleRecurrence(ruleType, monthFilter)
  );

  const label = String(getValueByHeader(row, headerMap, ['Event', 'Label'], '') || '').trim();
  const providedRuleId = String(getValueByHeader(row, headerMap, ['Rule ID'], fallbackRuleId) || fallbackRuleId).trim();
  const inferredType = String(getValueByHeader(row, headerMap, ['Type'], '') || '').trim();
  const defaultSortOrder = label ? 10 : 20;

  return {
    enabled: parseBooleanLike(getValueByHeader(row, headerMap, ['Enabled'], true), true),
    ruleId: slugifyRuleId(providedRuleId || label, fallbackRuleId),
    label: label,
    recurrence: recurrence,
    ruleType: ruleType,
    weekday: weekday,
    ordinal: ordinal,
    monthFilter: monthFilter,
    dayOfMonth: dayOfMonth,
    offsetDays: offsetDays,
    includeInForm: parseBooleanLike(getValueByHeader(row, headerMap, ['Include In Form'], true), true),
    includeInSchedule: parseBooleanLike(getValueByHeader(row, headerMap, ['Include In Schedule'], true), true),
    sortOrder: toIntegerOrDefault(getValueByHeader(row, headerMap, ['Sort Order'], defaultSortOrder), defaultSortOrder),
    type: inferredType || (label ? 'special' : 'service')
  };
}

function loadRecurringRules() {
  const sheet = getConfiguredRecurringSheet();
  if (!sheet || sheet.getLastRow() < 2) return getDefaultRecurringRules();

  const rows = sheet.getDataRange().getValues();
  const headerMap = buildHeaderMap(rows[0] || []);
  const rules = [];

  for (let i = 1; i < rows.length; i++) {
    if (isBlankRow(rows[i])) continue;
    const rule = normalizeRecurringRule(rows[i], headerMap, `rule_${i}`);
    if (rule && rule.enabled) rules.push(rule);
  }

  return rules.length ? rules : getDefaultRecurringRules();
}

function buildEventsForRule(year, month, rule, timeZone) {
  if (!rule || !rule.enabled || !monthMatchesFilter(rule.monthFilter, month)) return [];

  let dates = [];
  switch (rule.ruleType) {
    case 'every_weekday':
      dates = getAllWeekdaysInMonth(year, month, rule.weekday);
      break;
    case 'nth_weekday': {
      const nthDate = getNthWeekdayOfMonth(year, month, rule.weekday, rule.ordinal);
      if (nthDate) dates = [nthDate];
      break;
    }
    case 'fixed_date': {
      if (!rule.dayOfMonth) break;
      const fixedDate = new Date(year, month, rule.dayOfMonth);
      if (fixedDate.getMonth() === month) dates = [fixedDate];
      break;
    }
    case 'easter_offset': {
      const easter = computeEaster(year);
      easter.setDate(easter.getDate() + toIntegerOrDefault(rule.offsetDays, 0));
      if (easter.getMonth() === month) dates = [easter];
      break;
    }
    default:
      logDebug('warn', 'Unknown recurring rule type encountered', { ruleType: rule.ruleType, ruleId: rule.ruleId });
      return [];
  }

  return dates.map(date => toEventObject(date, rule, timeZone, 'recurring'));
}

function buildRecurringEvents(year, month, rules, timeZone) {
  let events = [];
  (rules || []).forEach(rule => {
    events = events.concat(buildEventsForRule(year, month, rule, timeZone));
  });
  return mergeEventObjects(events);
}

function normalizeMonthlyOverride(row, headerMap, fallbackRuleId, timeZone) {
  const rawDate = getValueByHeader(row, headerMap, ['Date'], '');
  const parsedDate = parseSingleDate(rawDate);
  if (!parsedDate) return null;

  const action = String(getValueByHeader(row, headerMap, ['Action'], 'ADD') || 'ADD').trim().toUpperCase();
  if (action !== 'ADD' && action !== 'REMOVE') return null;

  const label = String(getValueByHeader(row, headerMap, ['Event', 'Label'], '') || '').trim();
  const recurringEvent = String(getValueByHeader(row, headerMap, ['Recurring Event', 'Rule ID'], '') || '').trim();

  return {
    enabled: parseBooleanLike(getValueByHeader(row, headerMap, ['Enabled'], true), true),
    action: action,
    dateObject: parsedDate,
    isoDate: toIsoDate(parsedDate, timeZone),
    label: label,
    ruleId: recurringEvent ? slugifyRuleId(recurringEvent, fallbackRuleId) : '',
    includeInForm: parseBooleanLike(getValueByHeader(row, headerMap, ['Include In Form'], true), true),
    includeInSchedule: parseBooleanLike(getValueByHeader(row, headerMap, ['Include In Schedule'], true), true),
    sortOrder: toIntegerOrDefault(getValueByHeader(row, headerMap, ['Sort Order'], 100), 100),
    type: String(getValueByHeader(row, headerMap, ['Type'], '') || '').trim()
  };
}

function loadMonthlyOverrides(year, month, timeZone) {
  const sheet = getConfiguredEventsSheet();
  if (!sheet || sheet.getLastRow() < 2) return [];

  const rows = sheet.getDataRange().getValues();
  const headerMap = buildHeaderMap(rows[0] || []);
  if (!Object.prototype.hasOwnProperty.call(headerMap, 'action')) return [];

  const overrides = [];
  for (let i = 1; i < rows.length; i++) {
    if (isBlankRow(rows[i])) continue;
    const override = normalizeMonthlyOverride(rows[i], headerMap, `override_${i}`, timeZone);
    if (!override || !override.enabled) continue;
    if (override.dateObject.getFullYear() !== year || override.dateObject.getMonth() !== month) continue;
    overrides.push(override);
  }

  return overrides;
}

function monthlyOverrideMatchesEvent(override, event) {
  if (!override || !event || override.isoDate !== event.isoDate) return false;
  if (override.ruleId) return override.ruleId === event.ruleId;
  if (override.label) return normalizeName(override.label) === normalizeName(event.label || '');
  return true;
}

function applyMonthlyOverrides(events, overrides, timeZone) {
  const next = (events || []).slice();

  (overrides || []).forEach(override => {
    if (override.action === 'REMOVE') {
      for (let i = next.length - 1; i >= 0; i--) {
        if (monthlyOverrideMatchesEvent(override, next[i])) {
          next.splice(i, 1);
        }
      }
      return;
    }

    const event = toEventObject(override.dateObject, {
      ruleId: override.ruleId || `override_${override.isoDate}`,
      label: override.label || '',
      ruleType: 'override',
      includeInForm: override.includeInForm,
      includeInSchedule: override.includeInSchedule,
      sortOrder: override.sortOrder,
      type: override.type || 'override'
    }, timeZone, 'override');

    const existingIndex = next.findIndex(existing => existing.key === event.key);
    if (existingIndex >= 0) {
      next[existingIndex] = event;
    } else {
      next.push(event);
    }
  });

  return mergeEventObjects(next);
}

function getLegacyMonthlyEventChoices(year, month, timeZone) {
  const sheet = getConfiguredEventsSheet();
  if (!sheet || sheet.getLastRow() < 2) return [];

  const rows = sheet.getDataRange().getValues();
  const headerMap = buildHeaderMap(rows[0] || []);
  if (Object.prototype.hasOwnProperty.call(headerMap, 'action')) return [];
  if (!Object.prototype.hasOwnProperty.call(headerMap, 'year') ||
      !Object.prototype.hasOwnProperty.call(headerMap, 'month') ||
      !Object.prototype.hasOwnProperty.call(headerMap, 'date')) {
    return [];
  }

  const events = [];
  for (let i = 1; i < rows.length; i++) {
    if (isBlankRow(rows[i])) continue;

    const parsedDate = parseSingleDate(getValueByHeader(rows[i], headerMap, ['Date'], ''));
    if (!parsedDate) continue;

    const rowYear = toIntegerOrDefault(getValueByHeader(rows[i], headerMap, ['Year'], parsedDate.getFullYear()), parsedDate.getFullYear());
    const rowMonth = parseMonthNumber(getValueByHeader(rows[i], headerMap, ['Month'], parsedDate.getMonth() + 1));

    if (rowYear !== year) continue;
    if (rowMonth !== null && rowMonth !== month + 1 && rowMonth !== month) continue;

    events.push(toEventObject(parsedDate, {
      ruleId: `legacy_${i}`,
      label: String(getValueByHeader(rows[i], headerMap, ['Label'], '') || '').trim(),
      ruleType: 'legacy',
      includeInForm: true,
      includeInSchedule: true,
      sortOrder: 100,
      type: String(getValueByHeader(rows[i], headerMap, ['Type'], '') || '').trim()
    }, timeZone, 'legacy'));
  }

  if (!events.length) return [];

  return mergeDateChoices(
    sortEvents(mergeEventObjects(events.concat(getAutomaticSpecialEvents(year, month, timeZone))))
      .map(formatEventChoice)
  );
}

function getAvailabilitySheetHeaderChoices(year, month, settings) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const runtimeSettings = settings || loadRuntimeSettings();
  const sheet = ss.getSheetByName(getAvailabilitySheetName(year, month, runtimeSettings));
  if (!sheet) return [];

  const lastCol = sheet.getLastColumn();
  if (lastCol <= 1) return [];

  const headers = sheet.getRange(CONFIG.layout.dateRowIndex, 2, 1, lastCol - 1)
    .getDisplayValues()[0]
    .map(normalizeDateChoice)
    .filter(Boolean);

  if (!headers.length) return [];

  return mergeDateChoices(headers);
}

function getBuiltInFallbackServiceDates(year, month, timeZone) {
  return mergeDateChoices(
    getAllWeekdaysInMonth(year, month, 0).map(date => Utilities.formatDate(date, timeZone, 'MM/dd'))
  );
}

/**
 * Add a reconciliation entry for admin review.
 * Columns: timestamp | formId | submittedName | matchedCanonical | parseErrors | actionRequired | alerted
 */
function getFriendlyReconciliationAction(actionRequired, name) {
  const normalized = String(actionRequired || '').trim();
  if (!normalized) return 'Please review this item.';
  if (normalized.indexOf('New member added') === 0) {
    return `${name || 'This person'} was added as a new row because the submission did not match anyone already listed in Ministry Members. Please check whether this is a brand new person or a duplicate/spelling variation.`;
  }
  return normalized;
}

function buildReconciliationAlertPayload(rowsToAlert, sheet, settings) {
  const runtimeSettings = settings || loadRuntimeSettings();
  const sheetLink = getSheetRangeUrl(sheet, 1, 'A', 'G');
  const membersLink = getSheetUrlByName(CONFIG.sheetNames.ministryMembers);
  const items = rowsToAlert.map((item, index) => {
    const row = item.row;
    const submittedAt = formatHumanDateTime(row[0], runtimeSettings.timeZone);
    const name = String(row[2] || '').trim() || 'Unknown name';
    const parseErrors = String(row[4] || '').trim();
    const action = getFriendlyReconciliationAction(row[5], name);
    const rowLink = getSheetRangeUrl(sheet, item.idx, 'A', 'G');

    const lines = [
      `${index + 1}. ${name}`,
      `Submitted: ${submittedAt}`,
      `What needs review: ${action}`
    ];
    if (parseErrors) lines.push(`Dates we could not understand: ${parseErrors}`);
    lines.push(`Open this item: ${rowLink}`);
    return lines.join('\n');
  });

  const itemLabel = rowsToAlert.length === 1 ? 'item' : 'items';
  return {
    subject: `${runtimeSettings.churchName}: ${rowsToAlert.length} reconciliation ${itemLabel} need${rowsToAlert.length === 1 ? 's' : ''} review`,
    body: [
      `Church: ${runtimeSettings.churchName}`,
      '',
      `Please review the following ${itemLabel} in Jubal:`,
      '',
      items.join('\n\n'),
      '',
      `Reconciliation sheet: ${sheetLink}`,
      `Ministry Members sheet: ${membersLink}`
    ].join('\n')
  };
}

function addReconciliationEntry(e, submittedName, matchedCanonical, parseErrors, actionRequired) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(CONFIG.sheetNames.reconciliation);
    if (!sheet) {
      sheet = ss.insertSheet(CONFIG.sheetNames.reconciliation);
      sheet.appendRow(['timestamp', 'formId', 'submittedName', 'matchedCanonical', 'parseErrors', 'actionRequired', 'alerted']);
      applySheetTheme(sheet);
    }

    const ts = new Date().toISOString();
    const meta = ss.getSheetByName(CONFIG.sheetNames.formMetadata);
    let formId = '';
    if (meta) formId = (meta.getRange('B2').getValue() || meta.getRange('B1').getValue() || '').toString();

    const errorsStr = Array.isArray(parseErrors) ? parseErrors.join('; ') : (parseErrors || '');
    const rowIndex = sheet.getLastRow() + 1;
    sheet.appendRow([ts, formId, submittedName || '', matchedCanonical || '', errorsStr, actionRequired || '', '']);
    fitSheetToContent(sheet);
    applyTableBordersToDataRange(sheet);

    // Send immediate alert for this reconciliation item (can be changed to digest later)
    sendReconciliationAlert(sheet, rowIndex);
  } catch (err) {
    console.error('addReconciliationEntry failed: ' + err.message);
  }
}

/**
 * Send an email alert for reconciliation entries. Marks the 'alerted' column with timestamp.
 * If a sheet & specific rowIndex is provided, only alert for that row. Otherwise alert for all un-alerted rows.
 */
function sendReconciliationAlert(sheet, rowIndex) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!sheet) sheet = ss.getSheetByName(CONFIG.sheetNames.reconciliation);
    if (!sheet) return;

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return; // no entries

    const startRow = rowIndex ? rowIndex : 2;
    const range = sheet.getRange(startRow, 1, lastRow - startRow + 1, sheet.getLastColumn());
    const values = range.getValues();

    const rowsToAlert = [];
    for (let i = 0; i < values.length; i++) {
      const row = values[i];
      const alerted = (row[6] || '').toString().trim();
      if (!alerted) {
        rowsToAlert.push({ idx: startRow + i, row });
      }
    }

    if (!rowsToAlert.length) return;

    const email = buildReconciliationAlertPayload(rowsToAlert, sheet, loadRuntimeSettings());
    const sent = sendEmailToAdmins(email.subject, email.body);
    if (!sent) return;

    // Mark alerted timestamp
    rowsToAlert.forEach(r => {
      try {
        sheet.getRange(r.idx, 7).setValue(new Date().toISOString());
      } catch (inner) {
        console.error('Failed to mark alerted for row ' + r.idx + ': ' + inner.message);
      }
    });
  } catch (err) {
    console.error('sendReconciliationAlert failed: ' + err.message);
  }
}

/**
 * Normalize a person's name for canonical matching.
 * - trims, collapses spaces, removes diacritics, and lowercases
 */
function normalizeName(name) {
  if (!name) return "";
  try {
    const s = name.toString().trim().replace(/\s+/g, " ");
    const normalized = s.normalize ? s.normalize("NFD") : s;
    return normalized.replace(/[\u0300-\u036f]/g, "").toLowerCase();
  } catch (e) {
    return name.toString().toLowerCase();
  }
}

function formatAvailabilityDisplayName(name) {
  const normalized = String(name || '').trim().replace(/\s+/g, ' ');
  if (!normalized) return '';

  const nameParts = normalized.split(' ');
  if (nameParts.length > 1) {
    return nameParts[0] + ' ' + nameParts[1].charAt(0).toUpperCase() + '.';
  }

  return normalized;
}

function updateDatabase(e) {
  try {
    syncConfiguredMemberRoles();
    const databaseSS = SpreadsheetApp.getActiveSpreadsheet();
    const databaseSheet = databaseSS.getSheetByName(CONFIG.sheetNames.ministryMembers);
    const databaseData = databaseSheet.getDataRange().getValues();
    const memberColumns = getMinistryMembersColumnMap(databaseSheet);

    console.log('--- STARTING UPDATE DATABASE ---');
    console.log(`Event Data: ${JSON.stringify(e)}`);

    // Extract form responses using namedValues
    const responses = e.namedValues;
    // namedValues returns an array of strings for each key
    const name = responses[CONFIG.formHeaders.name] ? responses[CONFIG.formHeaders.name][0] : null;
    const timesWilling = responses[CONFIG.formHeaders.times] ? responses[CONFIG.formHeaders.times][0] : "";

    let unavailableDates = [];
    const rawDates = responses[CONFIG.formHeaders.dates];
    let parseErrors = [];

    if (rawDates && rawDates.length > 0) {
      console.log(`Raw Unavailable Dates from Form: ${rawDates[0]}`);
      const parseResult = parseUnavailableDates(rawDates[0]);
      unavailableDates = parseResult.parsed;
      parseErrors = parseResult.errors;
    }

    const unavailableDatesString = unavailableDates.join(','); // Prepare for storage
    const comments = responses[CONFIG.formHeaders.comments] ? responses[CONFIG.formHeaders.comments][0] : "";

    // Log final extracted values before database write
    console.log(`Name: ${name}`);
    console.log(`Times Willing to Serve: ${timesWilling}`);
    console.log(`Parsed Unavailable Dates: ${unavailableDatesString}`);
    console.log(`Comments: ${comments}`);
    
    // --- Database Update Logic with canonical name matching (locked) ---
    const lock = LockService.getScriptLock();
    let lockAcquired = false;
    try {
      // Wait up to 30 seconds to acquire the script lock
      lock.waitLock(30000);
      lockAcquired = true;

      // Re-read the database to get the latest data under lock
      const freshDatabaseData = databaseSheet.getDataRange().getValues();
      const incomingCanonical = normalizeName(name || '');

      let found = false;
      for (let i = 1; i < freshDatabaseData.length; i++) {
        const dbName = freshDatabaseData[i][memberColumns.name - 1] ? freshDatabaseData[i][memberColumns.name - 1].toString() : '';
        let dbCanonical = freshDatabaseData[i][memberColumns.canonicalName - 1] ? freshDatabaseData[i][memberColumns.canonicalName - 1].toString() : '';

        // If canonical missing, compute and persist it
        if (!dbCanonical && dbName) {
          try {
            dbCanonical = normalizeName(dbName);
            databaseSheet.getRange(i + 1, memberColumns.canonicalName).setValue(dbCanonical);
          } catch (err) {
            console.error('Failed to persist canonical name for row ' + (i + 1) + ': ' + err.message);
          }
        }

        if (dbCanonical && incomingCanonical && dbCanonical === incomingCanonical) {
          // Update the corresponding row
          databaseSheet.getRange(i + 1, memberColumns.times).setValue(timesWilling);
          databaseSheet.getRange(i + 1, memberColumns.dates).setValue(unavailableDatesString); // Use the joined string
          databaseSheet.getRange(i + 1, memberColumns.comments).setValue(comments);
          // Ensure canonical is stored
          databaseSheet.getRange(i + 1, memberColumns.canonicalName).setValue(incomingCanonical);
          found = true;
          console.log('Updated existing row ' + (i + 1) + ' for ' + name + ' (canonical: ' + incomingCanonical + ')');
          break;
        }
      }

      if (!found) {
        // If no match is found, append a new row with the canonical name in its configured column
        const lastRow = databaseSheet.getLastRow() + 1;
        databaseSheet.getRange(lastRow, memberColumns.name).setValue(name);
        databaseSheet.getRange(lastRow, memberColumns.times).setValue(timesWilling);
        databaseSheet.getRange(lastRow, memberColumns.dates).setValue(unavailableDatesString);
        databaseSheet.getRange(lastRow, memberColumns.comments).setValue(comments);
        databaseSheet.getRange(lastRow, memberColumns.canonicalName).setValue(incomingCanonical);
        ensureRolesFormulaForRow(databaseSheet, lastRow, loadRuntimeSettings().roles);
        console.log('Added new row ' + lastRow + ' for ' + name + ' (canonical: ' + incomingCanonical + ')');

        // Record reconciliation entry for admin review since this submission didn't match an existing member
        try {
          addReconciliationEntry(e, name, incomingCanonical, parseErrors, 'New member added — verify and merge if necessary');
        } catch (recErr) {
          console.error('Failed to add reconciliation entry: ' + recErr.message);
        }
      }
    } catch (err) {
      logDebug('warn', 'Could not acquire lock for updateDatabase; queuing submission', { error: err && err.message ? err.message : err });
      // Queue this submission for manual replay by appending the raw event
      try { logFormResponse(e); } catch (qerr) { console.error('Failed to queue submission: ' + qerr.message); }
      return;
    } finally {
      if (lockAcquired) {
        try { lock.releaseLock(); } catch (releaseErr) { console.error('Failed to release lock: ' + releaseErr.message); }
      }
    }
  } catch (error) {
    console.error(`!!! ERROR in updateDatabase: ${error.message}`);
  }
  updateAvailability();
  console.log('Updated availability and finished execution.');
}

function getServiceDates(year, month) {
  const settings = loadRuntimeSettings();
  const timeZone = settings.timeZone || safeGetScriptTimeZone();

  try {
    const legacyChoices = getLegacyMonthlyEventChoices(year, month, timeZone);
    if (legacyChoices.length) return legacyChoices;

    if (!hasConfiguredRecurringRules()) {
      const headerChoices = getAvailabilitySheetHeaderChoices(year, month, settings);
      if (headerChoices.length) return headerChoices;
    }

    const recurringRules = loadRecurringRules();
    let events = buildRecurringEvents(year, month, recurringRules, timeZone);
    const monthlyOverrides = loadMonthlyOverrides(year, month, timeZone);
    if (monthlyOverrides.length) {
      events = applyMonthlyOverrides(events, monthlyOverrides, timeZone);
    }

    const scheduleChoices = sortEvents(events)
      .filter(event => event.includeInSchedule)
      .map(formatEventChoice);

    if (scheduleChoices.length) return mergeDateChoices(scheduleChoices);
  } catch (err) {
    console.error('getServiceDates failed, falling back to built-in defaults: ' + err.message);
  }

  return getBuiltInFallbackServiceDates(year, month, timeZone);
}

/**
 * Compute Easter date for given year (Western/Gregorian) using Anonymous Gregorian algorithm.
 */
function computeEaster(year) {
  const a = year % 19;
  const b = Math.floor(year / 100);
  const c = year % 100;
  const d = Math.floor(b / 4);
  const e = b % 4;
  const f = Math.floor((b + 8) / 25);
  const g = Math.floor((b - f + 1) / 3);
  const h = (19 * a + b - d - g + 15) % 30;
  const i = Math.floor(c / 4);
  const k = c % 4;
  const l = (32 + 2 * e + 2 * i - h - k) % 7;
  const m = Math.floor((a + 11 * h + 22 * l) / 451);
  const month = Math.floor((h + l - 7 * m + 114) / 31) - 1; // 0-based month
  const day = ((h + l - 7 * m + 114) % 31) + 1;
  return new Date(year, month, day);
}

function sendNewFormCreatedEmail(monthName, responderUrl, editUrl, settings) {
  const runtimeSettings = settings || loadRuntimeSettings();
  const emailSubject = `${runtimeSettings.churchName}: New availability form created for ${monthName}`;
  const emailBody = `Church: ${runtimeSettings.churchName}\n\n` +
                  "A new Music Ministry Availability Form has been created for the month of " + monthName + ".\n\n" +
                  "You can access and fill out the form using the following link:\n" + responderUrl + "\n\n" +
                  "If you need to edit the form, use the following link:\n" + editUrl + "\n\n" +
                  "Please submit your availability as soon as possible.";

  sendEmailToAdmins(emailSubject, emailBody, runtimeSettings);
}

function createNewFormForMonth(month, year, monthName, options) {
  const opts = options || {};
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const metadataSheet = ss.getSheetByName(CONFIG.sheetNames.formMetadata) || ss.insertSheet(CONFIG.sheetNames.formMetadata);
  const runtimeSettings = opts.settings || loadRuntimeSettings();
  const availabilitySheetName = opts.sheetName || getAvailabilitySheetName(year, month, runtimeSettings);
  const memberNames = getMinistryMemberNames();

  if (!memberNames.length) {
    throw new Error(`Cannot create the ${monthName} form because Ministry Members does not have any names yet.`);
  }

  // Create a new form for the upcoming month
  const formTitle = `${runtimeSettings.churchName} Availability - ${monthName}`;
  const form = FormApp.create(formTitle);

  // Name Dropdown (ListItem)
  const nameDropdown = form.addListItem().setTitle(CONFIG.formHeaders.name).setRequired(true);
  nameDropdown.setChoiceValues(memberNames);

  const numDropdown = form.addListItem()
  .setTitle(CONFIG.formHeaders.times)
  .setChoiceValues(runtimeSettings.timesChoices) // Set the dropdown options
  .setRequired(true); // Make the question required
  
  // Add the service dates to the form for unavailable dates selection
  const serviceDates = getServiceDates(year, month);
  const dateChoices = serviceDates;

  const availMC = form.addCheckboxItem();
  availMC.setTitle(CONFIG.formHeaders.dates)
    .setChoices(dateChoices.map(date => availMC.createChoice(date)));

  // Optional comments section
  form.addTextItem().setTitle(CONFIG.formHeaders.comments).setRequired(false);

  if (!opts.deferDestination) {
    form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());
    console.log("Linked form responses to new sheet");
  }

  // Get the links for the edit and responder URLs
  const editUrl = form.getEditUrl(); // Edit link for the form owner
  const responderUrl = form.getPublishedUrl(); // Responder link for the participants

  if (runtimeSettings.formsFolder) {
    try {
      const file = DriveApp.getFileById(form.getId());
      const targetFolder = DriveApp.getFolderById(runtimeSettings.formsFolder);
      file.moveTo(targetFolder);
    } catch (err) {
      console.error('Failed to move form to configured folder: ' + err.message);
    }
  }

  // Sync form's date choices with the availability sheet (in case headers were edited)
  const syncResult = syncFormWithSheet(form.getId(), availabilitySheetName);
  if (!syncResult || syncResult.status !== 'synced') {
    throw new Error(`Failed to sync the new form with "${availabilitySheetName}".`);
  }

  if (!opts.skipMetadata) {
    writeCurrentFormMetadata(metadataSheet, monthName + " Form", form.getId());
  }

  if (!opts.skipEmail) {
    sendNewFormCreatedEmail(monthName, responderUrl, editUrl, runtimeSettings);
  }

  return {
    formId: form.getId(),
    formName: formTitle,
    metadataLabel: monthName + " Form",
    editUrl: editUrl,
    responderUrl: responderUrl,
    monthName: monthName,
    syncResult: syncResult
  };
}

function updateFormDropdown(formIdOverride) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const databaseSheet = ss.getSheetByName(CONFIG.sheetNames.ministryMembers);
  const metadataSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.sheetNames.formMetadata);

  if (!formIdOverride && !metadataSheet) {
    console.log("Form Metadata sheet missing.");
    return { status: 'missing_metadata' };
  }

  const formId = formIdOverride || metadataSheet.getRange("B2").getValue();
  if (!formId) {
    console.log("No Form ID found.");
    return { status: 'missing_form_id' };
  }

  // Retrieve the list of names from the "Ministry Members" sheet
  const memberNames = getMinistryMemberNames();
  if (!memberNames.length) {
    console.log("No names found in Ministry Members.");
    return { status: 'no_names' };
  }

  // Open the form using the Form ID
  const form = FormApp.openById(formId);

  // Locate the dropdown question by its title
  const items = form.getItems(FormApp.ItemType.LIST);
  const dropdownTitle = CONFIG.formHeaders.name; // Adjust this to match your question title
  let dropdownItem = null;

  for (let i = 0; i < items.length; i++) {
    if (items[i].getTitle() === dropdownTitle) {
      dropdownItem = items[i].asListItem();
      break;
    }
  }

  if (dropdownItem) {
    // Update the dropdown choices
    dropdownItem.setChoiceValues(memberNames);
    console.log("Dropdown updated with names from the sheet.");
    return {
      status: 'updated',
      formId: formId.toString(),
      choiceCount: memberNames.length
    };
  } else {
    console.log("Dropdown question not found.");
    return { status: 'missing_question', formId: formId.toString() };
  }
}

/**
 * Sync the date choices of a form with the header values from an availability sheet.
 * formId: the id of the Form to update
 * sheetName: availability sheet name to read header from
 */
function syncFormWithSheet(formId, sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    console.log('syncFormWithSheet: sheet not found: ' + sheetName);
    return { status: 'missing_sheet', sheetName: sheetName };
  }

  const lastCol = sheet.getLastColumn();
  if (lastCol <= 1) {
    console.log('syncFormWithSheet: no date columns in sheet');
    return { status: 'missing_date_columns', sheetName: sheetName };
  }

  const dateHeaders = sheet.getRange(CONFIG.layout.dateRowIndex, 2, 1, lastCol - 1).getDisplayValues()[0].map(h => String(h).trim()).filter(Boolean);
  const normalizedHeaders = mergeDateChoices(dateHeaders);

  if (!normalizedHeaders.length) {
    console.log('syncFormWithSheet: no date headers found');
    return { status: 'missing_date_headers', sheetName: sheetName };
  }

  const form = FormApp.openById(formId);
  const items = form.getItems(FormApp.ItemType.CHECKBOX);
  let target = null;
  for (let i = 0; i < items.length; i++) {
    if (items[i].getTitle() === CONFIG.formHeaders.dates) {
      target = items[i].asCheckboxItem();
      break;
    }
  }

  if (!target) {
    console.log('syncFormWithSheet: checkbox item for dates not found in form');
    return { status: 'missing_dates_question', sheetName: sheetName, formId: formId };
  }

  const choices = normalizedHeaders.map(d => target.createChoice(d));
  target.setChoices(choices);
  console.log('syncFormWithSheet: updated form choices from sheet ' + sheetName);
  return {
    status: 'synced',
    sheetName: sheetName,
    formId: formId,
    choiceCount: normalizedHeaders.length
  };
}

/**
 * Convenience wrapper: sync the current open form (from metadata) with the planned availability sheet for next month.
 */
function syncCurrentFormWithAvailability(referenceDate) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const meta = ss.getSheetByName(CONFIG.sheetNames.formMetadata);
  if (!meta) {
    console.log('No metadata sheet');
    return { status: 'missing_metadata' };
  }
  const formId = meta.getRange('B2').getValue() || meta.getRange('B1').getValue();
  if (!formId) {
    console.log('No form id in metadata');
    return { status: 'missing_form_id' };
  }

  const context = getPlanningMonthContext(referenceDate);
  return syncFormWithSheet(formId.toString(), context.sheetName);
}

function applyEventChangesToPlanningMonth(options) {
  const opts = options || {};
  const runtimeSettings = loadRuntimeSettings();
  const context = getPlanningMonthContext(opts.referenceDate, runtimeSettings);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const existingSheet = ss.getSheetByName(context.sheetName);

  if (!existingSheet) {
    return {
      status: 'missing_sheet',
      sheetName: context.sheetName,
      planMonthName: context.planMonthName,
      planYear: context.planYear
    };
  }

  const previousChoices = getAvailabilitySheetHeaderChoices(context.planYear, context.planMonth, runtimeSettings);
  const preservedAssignments = captureScheduleAssignments(existingSheet, runtimeSettings.roles);

  setupAvailability(context.sheetName, context.planYear, context.planMonth);

  const rebuiltSheet = ss.getSheetByName(context.sheetName);
  const restoredAssignments = restoreScheduleAssignments(rebuiltSheet, preservedAssignments, runtimeSettings.roles);
  const updatedChoices = getAvailabilitySheetHeaderChoices(context.planYear, context.planMonth, runtimeSettings);
  const previousDateKeys = previousChoices.map(extractDateKey).filter(Boolean);
  const updatedDateKeys = updatedChoices.map(extractDateKey).filter(Boolean);
  const addedDates = updatedChoices.filter(choice => previousDateKeys.indexOf(extractDateKey(choice)) === -1);
  const removedDates = previousChoices.filter(choice => updatedDateKeys.indexOf(extractDateKey(choice)) === -1);

  let formSync = { status: 'skipped' };
  if (!opts.skipFormSync) {
    formSync = syncCurrentFormWithAvailability(opts.referenceDate);
  }

  let availabilityUpdate = { status: 'skipped' };
  if (!opts.skipMatrixRefresh) {
    availabilityUpdate = updateAvailability(opts.referenceDate);
  }

  reorderWorkbookSheets(opts.referenceDate);

  return {
    status: 'completed',
    sheetName: context.sheetName,
    planMonthName: context.planMonthName,
    planYear: context.planYear,
    restoredAssignments: restoredAssignments,
    addedDates: addedDates,
    removedDates: removedDates,
    formSync: formSync,
    availabilityUpdate: availabilityUpdate
  };
}

function getAdminReminderPropertyKey(referenceDate) {
  const date = referenceDate ? new Date(referenceDate) : new Date();
  const planDate = new Date(date);
  planDate.setMonth(planDate.getMonth() + 1);
  return `adminReminderSent:${planDate.getFullYear()}-${('0' + (planDate.getMonth() + 1)).slice(-2)}`;
}

function getMonthlySetupPropertyKey(referenceDate) {
  const date = referenceDate ? new Date(referenceDate) : new Date();
  const planDate = new Date(date);
  planDate.setMonth(planDate.getMonth() + 1);
  return `monthlySetupCompleted:${planDate.getFullYear()}-${('0' + (planDate.getMonth() + 1)).slice(-2)}`;
}

function shouldRunMonthlySetupToday(referenceDate, settings) {
  const runtimeSettings = settings || loadRuntimeSettings();
  const today = referenceDate ? new Date(referenceDate) : new Date();
  return today.getDate() === runtimeSettings.formCreationDay;
}

function buildAdminPlanningReminder(referenceDate, settings) {
  const runtimeSettings = settings || loadRuntimeSettings();
  const today = referenceDate ? new Date(referenceDate) : new Date();
  const planDate = new Date(today);
  planDate.setMonth(planDate.getMonth() + 1);

  const planMonthName = Utilities.formatDate(planDate, runtimeSettings.timeZone, 'MMMM yyyy');
  const monthlySetupDay = formatDayOfMonthHuman(runtimeSettings.formCreationDay);
  const subject = `${runtimeSettings.churchName}: ${planMonthName} Events Reminder`;
  const bodyLines = [
    `Church: ${runtimeSettings.churchName}`,
    '',
    `Please review the schedule setup for ${planMonthName}.`,
    '',
    'Before the monthly form is created, please check these areas:',
    `- Recurring schedule: ${getSheetUrlByName(CONFIG.sheetNames.recurring)}`,
    `- One-time events and changes: ${getSheetUrlByName(CONFIG.sheetNames.events)}`,
    `- Ministry roles: ${getSheetUrlByName(CONFIG.sheetNames.rolesConfig)}`,
    `- Ministry members and role updates: ${getSheetUrlByName(CONFIG.sheetNames.ministryMembers)}`,
    `- Admin contacts and notifications: ${getSheetUrlByName(CONFIG.sheetNames.admins)}`,
    '',
    'Recommended checklist:',
    '- Confirm your normal recurring events are correct.',
    '- Add any special events, cancellations, or moved dates for next month.',
    '- Add or disable roles in the Roles sheet if your ministry roster changed.',
    '- Add any new members or update role checkboxes as needed.',
    '- Add or remove admin recipients in the Admins sheet if notification recipients need to change.',
    '',
    `Monthly setup day: the ${monthlySetupDay} of each month.`,
    'Once that looks right, the next month form and availability sheet can be generated by the monthly setup trigger.',
    `Reminder date: ${formatHumanDateTime(today, runtimeSettings.timeZone)}`
  ];

  return {
    subject: subject,
    body: bodyLines.join('\n')
  };
}

function sendAdminPlanningReminderIfDue() {
  const runtimeSettings = loadRuntimeSettings();
  if (!runtimeSettings.adminReminderEnabled) return { status: 'disabled' };
  if (!getAdminRecipientList(runtimeSettings).length) return { status: 'no_recipients' };

  const today = new Date();
  if (today.getDate() !== runtimeSettings.adminReminderDay) {
    return { status: 'not_due_today', day: today.getDate(), reminderDay: runtimeSettings.adminReminderDay };
  }

  const propertyKey = getAdminReminderPropertyKey(today);
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty(propertyKey)) {
    return { status: 'already_sent', key: propertyKey };
  }

  const reminder = buildAdminPlanningReminder(today, runtimeSettings);
  const sent = sendEmailToAdmins(reminder.subject, reminder.body, runtimeSettings);
  if (!sent) return { status: 'no_recipients' };

  props.setProperty(propertyKey, new Date().toISOString());
  return { status: 'sent', key: propertyKey };
}

function sendAdminPlanningReminderNow() {
  const runtimeSettings = loadRuntimeSettings();
  const reminder = buildAdminPlanningReminder(new Date(), runtimeSettings);
  const sent = sendEmailToAdmins(reminder.subject, reminder.body, runtimeSettings);
  return { status: sent ? 'sent' : 'no_recipients' };
}

/**
 * Legacy helper: ensure an old-style Events/Monthly Events sheet contains
 * automatic special events (Easter/Christmas) for the provided year/month.
 */
function ensureMonthlyEventsFor(year, month) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  // Defensive defaults: if year not provided or invalid, use current year
  year = parseInt(year, 10);
  if (isNaN(year)) year = new Date().getFullYear();
  // Normalize month to 0-based integer
  month = parseInt(month, 10);
  if (isNaN(month)) month = new Date().getMonth();
  let me = getConfiguredEventsSheet();
  if (!me) return;

  const data = me.getDataRange().getValues();
  const header = data[0] || [];
  if (header.indexOf('Action') >= 0) {
    logDebug('info', 'Skipping ensureMonthlyEventsFor because Events uses action-based overrides');
    return;
  }
  const yCol = header.indexOf('Year');
  const mCol = header.indexOf('Month');
  const dCol = header.indexOf('Date');
  const lCol = header.indexOf('Label');

  const toEnsure = [];
  const eas = computeEaster(year);
  if (eas.getMonth() === month) {
    // Use ISO date for clarity
    const iso = Utilities.formatDate(eas, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    toEnsure.push({ date: iso, label: 'Easter' });
  }
  if (month === 11) {
    const iso = Utilities.formatDate(new Date(year, 11, 25), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    toEnsure.push({ date: iso, label: 'Christmas' });
  }

  for (const ev of toEnsure) {
    let found = false;
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row) continue;
      const ry = parseInt(row[yCol], 10);
      const rm = row[mCol];
      const rd = row[dCol] ? row[dCol].toString() : '';

      // Accept month as 1-based or 0-based number or name; normalize
      let rmNum = null;
      if (typeof rm === 'number' && !isNaN(rm)) rmNum = rm; else if (typeof rm === 'string') {
        const asNum = parseInt(rm, 10);
        if (!isNaN(asNum)) rmNum = asNum; else {
          const mn = new Date(Date.parse(rm + ' 1, 2000'));
          if (!isNaN(mn.getTime())) rmNum = mn.getMonth() + 1;
        }
      }

      if (ry === year && (rmNum === month || rmNum === (month + 1))) {
        // Compare normalized date keys to avoid format and timezone drift.
        if (extractDateKey(rd) === extractDateKey(ev.date)) {
          found = true;
          break;
        }
      }
    }

    if (!found) {
      // Append Year, Month (1-based), Date, Label, Type
      me.appendRow([year, month + 1, ev.date, ev.label, 'auto']);
      logDebug('info', 'ensureMonthlyEventsFor added event', { year, month, event: ev });
    }
  }
}

/**
 * Convenience: populate automatic events (Easter/Christmas) for the whole year.
 */
function populateAnnualEvents(year) {
  // Default to current year if none provided
  year = parseInt(year, 10);
  if (isNaN(year)) year = new Date().getFullYear();
  for (let m = 0; m < 12; m++) ensureMonthlyEventsFor(year, m);
}

function getTrackedFormIds(metadataSheet) {
  if (!metadataSheet || metadataSheet.getLastRow() < 1 || metadataSheet.getLastColumn() < 2) return [];

  const values = metadataSheet.getRange(1, 2, metadataSheet.getLastRow(), 1).getValues().flat();
  return values
    .map(value => String(value || '').trim())
    .filter(value => value && value.toLowerCase() !== 'form id')
    .filter((value, index, arr) => arr.indexOf(value) === index);
}

function getFormResponseSheetIds(ss) {
  return (ss || SpreadsheetApp.getActiveSpreadsheet())
    .getSheets()
    .filter(sheet => sheet.getName().startsWith("Form Responses"))
    .map(sheet => sheet.getSheetId());
}

function unlinkTrackedForms(formIds) {
  (formIds || []).forEach(formId => {
    try {
      FormApp.openById(formId).removeDestination();
      console.log("De-linked form with ID: " + formId);
    } catch (e) {
      console.log("Could not de-link or find form " + formId + ": " + e.message);
    }
  });
}

function deleteFormResponseSheetsById(ss, sheetIds) {
  const spreadsheet = ss || SpreadsheetApp.getActiveSpreadsheet();
  const targetIds = {};
  (sheetIds || []).forEach(sheetId => {
    targetIds[String(sheetId)] = true;
  });

  spreadsheet.getSheets().forEach(sheet => {
    if (!sheet.getName().startsWith("Form Responses")) return;
    if (!Object.prototype.hasOwnProperty.call(targetIds, String(sheet.getSheetId()))) return;

    const toDelete = sheet.getName();
    try {
      spreadsheet.deleteSheet(sheet);
      console.log("Deleted old Form Responses tab: " + toDelete);
    } catch (e) {
      console.log("Skipped deleting Form Responses tab '" + toDelete + "': " + e.message);
    }
  });
}

function deleteFormResponseSheets(ss) {
  ss.getSheets().forEach(sheet => {
    if (!sheet.getName().startsWith("Form Responses")) return;

    const toDelete = sheet.getName();
    try {
      ss.deleteSheet(sheet);
      console.log("Deleted old Form Responses tab: " + toDelete);
    } catch (e) {
      console.log("Skipped deleting Form Responses tab '" + toDelete + "': " + e.message);
    }
  });
}

function applyCheckboxColumn(sheet, column, numRows) {
  if (!sheet || numRows <= 0) return;
  sheet.getRange(2, column, numRows, 1).insertCheckboxes();
}

function applyEmailColumn(sheet, column, numRows) {
  if (!sheet || numRows <= 0) return;
  const rule = SpreadsheetApp.newDataValidation()
    .requireTextIsEmail()
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, column, numRows, 1).setDataValidation(rule);
}

function applyDropdownColumn(sheet, column, values, numRows, allowInvalid) {
  if (!sheet || numRows <= 0) return;
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(values, true)
    .setAllowInvalid(allowInvalid === true)
    .build();
  sheet.getRange(2, column, numRows, 1).setDataValidation(rule);
}

function applyDateColumn(sheet, column, numRows, formatPattern) {
  if (!sheet || numRows <= 0) return;
  const rule = SpreadsheetApp.newDataValidation()
    .requireDate()
    .setHelpText('Click the cell and use the calendar picker, or type a date like Apr 5, 2026.')
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, column, numRows, 1).setDataValidation(rule);
  sheet.getRange(2, column, numRows, 1).setNumberFormat(formatPattern || 'ddd, mmm d, yyyy');
}

function setHeaderNotes(sheet, notes) {
  if (!sheet || !notes || !notes.length) return;
  sheet.getRange(1, 1, 1, notes.length).setNotes([notes]);
}

function styleConfigHeader(sheet, backgrounds) {
  if (!sheet || !backgrounds || !backgrounds.length) return;
  const range = sheet.getRange(1, 1, 1, backgrounds.length);
  range.setBackgrounds([backgrounds]);
  range.setFontWeight('bold');
}

function buildRichTextWithBoldPhrases(text, phrases) {
  const value = String(text || '');
  const builder = SpreadsheetApp.newRichTextValue().setText(value);
  const boldStyle = SpreadsheetApp.newTextStyle().setBold(true).build();

  (phrases || []).forEach(phrase => {
    if (!phrase) return;
    let startIndex = value.indexOf(phrase);
    while (startIndex !== -1) {
      builder.setTextStyle(startIndex, startIndex + phrase.length, boldStyle);
      startIndex = value.indexOf(phrase, startIndex + phrase.length);
    }
  });

  return builder.build();
}

function getRecurringEventDropdownValues() {
  const sheet = getConfiguredRecurringSheet();
  if (!sheet || sheet.getLastRow() < 2) return [];

  const headerMap = getSheetHeaderMap(sheet);
  const eventIndex = headerMap.event;
  if (eventIndex === undefined) return [];

  const values = sheet
    .getRange(2, eventIndex + 1, sheet.getLastRow() - 1, 1)
    .getDisplayValues()
    .flat()
    .map(value => String(value || '').trim())
    .filter(Boolean);

  return values.filter((value, index) => values.indexOf(value) === index);
}

function configureRecurringSheetUi(sheet) {
  if (!sheet) return;
  const maxRows = Math.max(sheet.getMaxRows() - 1, 1);
  const headerMap = getSheetHeaderMap(sheet);
  const hasLegacyYearlyColumns = Object.prototype.hasOwnProperty.call(headerMap, 'month') || Object.prototype.hasOwnProperty.call(headerMap, 'day');
  sheet.setFrozenRows(1);
  if (hasLegacyYearlyColumns) {
    setHeaderNotes(sheet, [
      'Check this to use the row.',
      'Name shown on the form and schedule. Leave blank for plain Sunday dates.',
      'Weekly = every week. Monthly = like first Friday. Yearly = fixed date each year. Easter = Easter Sunday.',
      'Pick the weekday for weekly or monthly patterns.',
      'Use "every" for weekly rows, or 1/2/3/4/5/last for monthly patterns.',
      'Use "all" for every month, or pick a specific month for yearly events.',
      'Day of month for yearly dates like Christmas on 25.',
      'Check to show this event on the availability form.',
      'Check to show this event on the schedule sheet.',
      'Optional reminder for admins.'
    ]);
  } else {
    setHeaderNotes(sheet, [
      'Check this to use the row.',
      'Name shown on the form and schedule. Leave blank for plain Sunday dates.',
      'Weekly = every week. Monthly = like first Friday. Use the Events sheet for dated specials like Easter, Christmas, and Good Friday.',
      'Pick the weekday for weekly or monthly patterns.',
      'Use "every" for weekly rows, or 1/2/3/4/5/last for monthly patterns.',
      'Check to show this event on the availability form.',
      'Check to show this event on the schedule sheet.',
      'Optional reminder for admins.'
    ]);
  }

  if (headerMap.enabled !== undefined) applyCheckboxColumn(sheet, headerMap.enabled + 1, maxRows);
  if (headerMap.frequency !== undefined) {
    const values = hasLegacyYearlyColumns ? ['Weekly', 'Monthly', 'Yearly', 'Easter'] : ['Weekly', 'Monthly'];
    applyDropdownColumn(sheet, headerMap.frequency + 1, values, maxRows);
  }
  if (headerMap.weekday !== undefined) {
    applyDropdownColumn(sheet, headerMap.weekday + 1, ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'], maxRows);
  }
  if (headerMap['week of month'] !== undefined) {
    applyDropdownColumn(sheet, headerMap['week of month'] + 1, ['every', '1', '2', '3', '4', '5', 'last'], maxRows);
  }
  if (headerMap.month !== undefined) {
    applyDropdownColumn(sheet, headerMap.month + 1, ['all', '1', '2', '3', '4', '5', '6', '7', '8', '9', '10', '11', '12', 'January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'], maxRows);
  }
  if (headerMap.day !== undefined) {
    sheet.getRange(2, headerMap.day + 1, maxRows, 1).setNumberFormat('0');
  }
  if (headerMap['include in form'] !== undefined) applyCheckboxColumn(sheet, headerMap['include in form'] + 1, maxRows);
  if (headerMap['include in schedule'] !== undefined) applyCheckboxColumn(sheet, headerMap['include in schedule'] + 1, maxRows);
  applySheetTheme(sheet);
  fitSheetToContent(sheet);
  applyTableBordersToDataRange(sheet);
}

function renderEventsInstructionBanner(sheet) {
  if (!sheet) return;

  const startColumn = 10; // J
  const width = 5; // J:N
  const titleRow = 2;
  const bodyRow = 3;
  const bodyHeight = 4;
  const lastNeededColumn = startColumn + width - 1;

  if (sheet.getMaxColumns() < lastNeededColumn) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), lastNeededColumn - sheet.getMaxColumns());
  }

  const titleRange = sheet.getRange(titleRow, startColumn, 1, width);
  const bodyRange = sheet.getRange(bodyRow, startColumn, bodyHeight, width);

  titleRange.breakApart().clearContent().clearFormat();
  bodyRange.breakApart().clearContent().clearFormat();

  titleRange
    .merge()
    .setValue('How to Add Events')
    .setFontWeight('bold')
    .setFontColor('#1f1f1f')
    .setBackground('#d9ead3')
    .setHorizontalAlignment('left');

  bodyRange
    .merge()
    .setValue(
      'Option 1: DOUBLE-CLICK a Date cell to open the calendar.\n' +
      'OR\n' +
      'Option 2: use Add Special Event from the spreadsheet menu.\n' +
      'Use ADD for one-time events and REMOVE to skip one recurring date.'
    )
    .setBackground('#f3f8ee')
    .setFontColor('#1f1f1f')
    .setVerticalAlignment('top')
    .setHorizontalAlignment('left')
    .setWrap(true);

  sheet
    .getRange(bodyRow, startColumn)
    .setRichTextValue(
      buildRichTextWithBoldPhrases(
        'Option 1: DOUBLE-CLICK a Date cell to open the calendar.\n' +
        'OR\n' +
        'Option 2: use Add Special Event from the spreadsheet menu.\n' +
        'Use ADD for one-time events and REMOVE to skip one recurring date.',
        ['DOUBLE-CLICK', 'Date', 'OR', 'Add Special Event']
      )
    );

  applyTableBorder(sheet, titleRow, startColumn, 1, width);
  applyTableBorder(sheet, bodyRow, startColumn, bodyHeight, width);
  sheet.autoResizeRows(titleRow, bodyHeight + 1);
}

function applyEventsInstructionRichText(sheet) {
  if (!sheet || sheet.getLastRow() < 2) return;
  const notesValues = sheet.getRange(2, 8, sheet.getLastRow() - 1, 1).getDisplayValues().flat();
  notesValues.forEach((value, index) => {
    const text = String(value || '');
    if (text.indexOf('DOUBLE-CLICK') === -1) return;
    sheet.getRange(index + 2, 8).setRichTextValue(buildRichTextWithBoldPhrases(text, ['DOUBLE-CLICK', 'Date', 'OR', 'Add Special Event']));
  });
}

function configureEventsSheetUi(sheet) {
  if (!sheet) return;
  const maxRows = Math.max(sheet.getMaxRows() - 1, 1);
  promoteExampleRows(sheet, 8);
  sheet.setFrozenRows(1);
  setHeaderNotes(sheet, [
    'Check this to use the row.',
    'Click the cell and use the calendar picker. For the easiest workflow, use the Add Special Event menu option.',
    'Name shown on the form and schedule for this one-time event.',
    'ADD creates a one-time event. REMOVE cancels one date from the normal schedule.',
    'Optional. Use the same event name as the Recurring sheet when moving or cancelling a recurring event.',
    'Check to show this event on the availability form.',
    'Check to show this event on the schedule sheet.',
    'Optional reminder for admins.'
  ]);
  applyCheckboxColumn(sheet, 1, maxRows);
  applyDateColumn(sheet, 2, maxRows, 'ddd, mmm d, yyyy');
  applyDropdownColumn(sheet, 4, ['ADD', 'REMOVE'], maxRows);
  const recurringEventValues = getRecurringEventDropdownValues();
  if (recurringEventValues.length) {
    applyDropdownColumn(sheet, 5, recurringEventValues, maxRows, true);
  }
  applyCheckboxColumn(sheet, 6, maxRows);
  applyCheckboxColumn(sheet, 7, maxRows);
  applySheetTheme(sheet);
  sheet.getRange(1, 2).setBackground('#9fc5e8');
  sheet.getRange(2, 2, maxRows, 1).setBackground('#eef4ff');
  highlightExampleRows(sheet, 8);
  applyEventsInstructionRichText(sheet);
  fitSheetToContent(sheet);
  applyTableBorder(sheet, 1, 1, sheet.getLastRow(), 8);
  renderEventsInstructionBanner(sheet);
}

function getAdminsSeedRows(emails) {
  const adminEmails = normalizeEmailList(emails && emails.length ? emails : CONFIG.ids.adminEmails);
  const starterEmails = normalizeEmailList(CONFIG.ids.adminEmails || []);
  const rows = [
    ['Enabled', 'Email', 'Notes']
  ];

  if (adminEmails.length) {
    adminEmails.forEach(email => {
      rows.push([
        true,
        email,
        starterEmails.indexOf(email) !== -1 ? 'Example row' : ''
      ]);
    });
  } else {
    rows.push([true, '', 'Enter the first admin email here.']);
  }

  rows.push([false, 'admin@example.com', 'Example row']);
  return rows;
}

function normalizeAdminsSheetLayout(sheet) {
  if (!sheet) return;

  const desiredHeaders = ['Enabled', 'Email', 'Notes'];
  const lastRow = Math.max(sheet.getLastRow(), 1);
  const lastColumn = Math.max(sheet.getLastColumn(), 3);

  if (sheet.getMaxColumns() < lastColumn) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), lastColumn - sheet.getMaxColumns());
  }

  const existingHeaders = sheet.getRange(1, 1, 1, lastColumn).getDisplayValues()[0];
  const headerMap = buildHeaderMap(existingHeaders);
  const rows = [desiredHeaders];

  if (lastRow > 1) {
    const existingRows = sheet.getRange(2, 1, lastRow - 1, lastColumn).getValues();
    existingRows.forEach(row => {
      const enabled = getValueByHeader(row, headerMap, ['Enabled'], '');
      const email = getValueByHeader(row, headerMap, ['Email'], '');
      const notes = getValueByHeader(row, headerMap, ['Notes'], '') || getValueByHeader(row, headerMap, ['Name'], '');
      rows.push([enabled, email, notes]);
    });
  }

  sheet.clearContents();
  sheet.getRange(1, 1, rows.length, desiredHeaders.length).setValues(rows);
  sheet.getRange(1, 1, 1, desiredHeaders.length).setFontWeight('bold');
}

function configureAdminsSheetUi(sheet) {
  if (!sheet) return;
  const maxRows = Math.max(sheet.getMaxRows() - 1, 1);
  sheet.setFrozenRows(1);
  setHeaderNotes(sheet, [
    'Check this when the admin should receive reminders and alerts.',
    'Email address used for admin reminders and alerts.',
    'Optional note for the team.'
  ]);
  applyCheckboxColumn(sheet, 1, maxRows);
  applyEmailColumn(sheet, 2, maxRows);
  applySheetTheme(sheet);
  highlightAdminRows(sheet, 2, 3);
  fitSheetToContent(sheet);
  applyTableBordersToDataRange(sheet);
}

function getRolesSeedRows(roles) {
  const configuredRoles = normalizeRoleList(roles && roles.length ? roles : CONFIG.roles);
  const rows = [
    ['Enabled', 'Role', 'Notes'],
    [false, 'MEDIA', 'Example row']
  ];

  configuredRoles.forEach(role => {
    rows.push([
      true,
      role,
      ''
    ]);
  });

  return rows;
}

function configureRolesSheetUi(sheet) {
  if (!sheet) return;
  const maxRows = Math.max(sheet.getMaxRows() - 1, 1);
  promoteExampleRows(sheet, 3);
  sheet.setFrozenRows(1);
  setHeaderNotes(sheet, [
    'Check this when the role should be active in scheduling.',
    'Role name shown in the schedule and on member role checkboxes. To add a new role, add a new row, type the role name, and check Enabled.',
    'Optional note for admins.'
  ]);
  applyCheckboxColumn(sheet, 1, maxRows);
  applySheetTheme(sheet);
  highlightExampleRows(sheet, 3);
  fitSheetToContent(sheet);
  applyTableBordersToDataRange(sheet);
}

function getRecurringSeedRows() {
  return [
    ['Enabled', 'Event', 'Frequency', 'Weekday', 'Week Of Month', 'Include In Form', 'Include In Schedule', 'Notes'],
    [true, '', 'Weekly', 'Sunday', 'every', true, true, 'Default weekly Sunday schedule. Leave Event blank to show plain dates.'],
    [false, 'Corporate Prayer', 'Monthly', 'Friday', 1, true, true, 'Enable if your church has a monthly prayer gathering']
  ];
}

function getUpcomingSpecialEventExampleDates(referenceDate) {
  const today = referenceDate ? new Date(referenceDate) : new Date();
  const todayStart = new Date(today.getFullYear(), today.getMonth(), today.getDate());

  let easter = computeEaster(todayStart.getFullYear());
  if (easter.getTime() < todayStart.getTime()) {
    easter = computeEaster(todayStart.getFullYear() + 1);
  }

  const goodFriday = new Date(easter);
  goodFriday.setDate(goodFriday.getDate() - 2);

  let christmas = new Date(todayStart.getFullYear(), 11, 25);
  if (christmas.getTime() < todayStart.getTime()) {
    christmas = new Date(todayStart.getFullYear() + 1, 11, 25);
  }

  return {
    goodFriday: goodFriday,
    easter: easter,
    christmas: christmas
  };
}

function getEventsSeedRows() {
  const exampleDates = getUpcomingSpecialEventExampleDates();
  return [
    ['Enabled', 'Date', 'Event', 'Action', 'Recurring Event', 'Include In Form', 'Include In Schedule', 'Notes'],
    [false, exampleDates.easter, 'Easter', 'ADD', '', true, true, 'Example row - DOUBLE-CLICK a Date cell OR use Add Special Event'],
    [false, exampleDates.goodFriday, 'Good Friday', 'ADD', '', true, true, 'Example row - dated special events like Good Friday belong in Events.'],
    [false, exampleDates.christmas, 'Christmas', 'ADD', '', true, true, 'Example row - dated special events like Christmas belong in Events.'],
    [false, exampleDates.goodFriday, 'Corporate Prayer', 'REMOVE', 'Corporate Prayer', true, true, 'Example row - use REMOVE when a recurring event should not happen on one date.']
  ];
}

function seedSheetRowsIfEmpty(sheet, rows) {
  if (!sheet || !rows || rows.length < 2) return false;

  if (sheet.getLastRow() > 1) {
    const existingRows = sheet.getRange(2, 1, sheet.getLastRow() - 1, Math.min(sheet.getLastColumn(), rows[0].length)).getDisplayValues();
    const hasNonBlankData = existingRows.some(row => !isBlankRow(row));
    if (hasNonBlankData) return false;
  }

  const bodyRows = rows.slice(1);
  const requiredRows = bodyRows.length;
  if (sheet.getMaxRows() < requiredRows + 1) {
    sheet.insertRowsAfter(sheet.getMaxRows(), requiredRows + 1 - sheet.getMaxRows());
  }

  sheet.getRange(2, 1, requiredRows, rows[0].length).setValues(bodyRows);
  return true;
}

function getSettingsSeedRows() {
  return [
    ['Key', 'Value', 'Notes'],
    ['church_name', CONFIG.defaults.churchName, 'Used in form titles and notifications'],
    ['time_zone', safeGetScriptTimeZone(), 'Time zone used for event generation and reminder emails'],
    ['form_creation_day', CONFIG.defaults.formCreationDay, 'Day of month when the daily monthlySetup trigger should create the next month form and availability sheet.'],
    ['admin_reminder_enabled', CONFIG.defaults.adminReminderEnabled, 'TRUE or FALSE. When TRUE, send planning reminders to admins.'],
    ['admin_reminder_day', CONFIG.defaults.adminReminderDay, 'Day of month to send the admin planning reminder for next month.'],
    ['times_choices', CONFIG.defaults.timesChoices.join(','), 'Choices shown in the form question for how many times someone is willing to serve this month.'],
    ['events_archive_frequency', CONFIG.defaults.eventsArchiveFrequency, 'Off, Monthly, Quarterly, or Yearly. Cleanup runs automatically on the first day of the new period.'],
    ['forms_folder_id', CONFIG.ids.formsFolder, 'Only change this if you want new monthly forms to be stored in a different Drive folder.']
  ];
}

function ensureSettingsSheetRows(sheet, rows) {
  if (!sheet || !rows || rows.length < 2) return;

  const existing = {};
  if (sheet.getLastRow() >= 2) {
    const currentRows = sheet.getRange(2, 1, sheet.getLastRow() - 1, Math.min(3, sheet.getLastColumn())).getValues();
    currentRows.forEach((row, index) => {
      const key = String(row[0] || '').trim().toLowerCase();
      if (key) existing[key] = index + 2;
    });
  }

  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    const key = String(row[0] || '').trim().toLowerCase();
    if (!key) continue;

    if (existing[key]) {
      if (!sheet.getRange(existing[key], 3).getValue()) {
        sheet.getRange(existing[key], 3).setValue(row[2]);
      }
      continue;
    }

    sheet.appendRow(row);
  }
}

function rewriteSettingsSheetRows(sheet, seedRows) {
  if (!sheet || !seedRows || seedRows.length < 2) return;

  const existing = {};
  if (sheet.getLastRow() >= 2) {
    const currentRows = sheet.getRange(2, 1, sheet.getLastRow() - 1, Math.min(3, sheet.getLastColumn())).getValues();
    currentRows.forEach(row => {
      const key = String(row[0] || '').trim().toLowerCase();
      if (!key) return;
      existing[key] = {
        value: row[1],
        notes: row[2]
      };
    });
  }

  const orderedRows = seedRows.slice(1).map(seedRow => {
    const key = String(seedRow[0] || '').trim().toLowerCase();
    const existingRow = existing[key];
    const hasExistingValue = existingRow && existingRow.value !== '' && existingRow.value !== null && existingRow.value !== undefined;
    const hasExistingNotes = existingRow && String(existingRow.notes || '').trim();
    let value = hasExistingValue ? existingRow.value : seedRow[1];

    if (key === 'form_creation_day' || key === 'admin_reminder_day') {
      value = String(clampDayOfMonthSetting(value, seedRow[1]));
    } else if (key === 'admin_reminder_enabled') {
      value = parseBooleanLike(value, seedRow[1]) ? 'TRUE' : 'FALSE';
    } else if (key === 'events_archive_frequency') {
      const normalizedFrequency = normalizeArchiveFrequency(value);
      value = normalizedFrequency.charAt(0).toUpperCase() + normalizedFrequency.slice(1);
    }

    return [
      seedRow[0],
      value,
      hasExistingNotes ? existingRow.notes : seedRow[2]
    ];
  });

  if (sheet.getMaxRows() < orderedRows.length + 1) {
    sheet.insertRowsAfter(sheet.getMaxRows(), orderedRows.length + 1 - sheet.getMaxRows());
  }

  sheet.getRange(2, 1, Math.max(sheet.getMaxRows() - 1, 1), 3).clearContent();
  sheet.getRange(2, 1, orderedRows.length, 3).setValues(orderedRows);
}

function configureSettingsSheetUi(sheet) {
  if (!sheet || sheet.getLastRow() < 2) return;

  sheet.setFrozenRows(1);
  applySheetTheme(sheet);
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, Math.min(2, sheet.getLastColumn())).getValues();
  const formsFolderHighlight = '#fce5cd';
  rows.forEach((row, index) => {
    const key = String(row[0] || '').trim().toLowerCase();
    const valueRange = sheet.getRange(index + 2, 2);
    valueRange.clearDataValidations();
    sheet.getRange(index + 2, 1, 1, Math.min(3, sheet.getLastColumn())).setBackground(null);

    if (key === 'time_zone') {
      const timeZoneOptions = normalizeRoleList([String(valueRange.getValue() || '').trim()].concat(getTimeZoneOptions()));
      valueRange.setDataValidation(
        SpreadsheetApp.newDataValidation()
          .requireValueInList(timeZoneOptions, true)
          .setAllowInvalid(false)
          .build()
      );
    } else if (key === 'events_archive_frequency') {
      valueRange.setDataValidation(
        SpreadsheetApp.newDataValidation()
          .requireValueInList(['Off', 'Monthly', 'Quarterly', 'Yearly'], true)
          .setAllowInvalid(false)
          .build()
      );
    } else if (key === 'form_creation_day') {
      valueRange.setDataValidation(
        SpreadsheetApp.newDataValidation()
          .requireValueInList(getDayOfMonthDropdownValues(), true)
          .setAllowInvalid(false)
          .build()
      );
    } else if (key === 'admin_reminder_enabled') {
      valueRange.setDataValidation(
        SpreadsheetApp.newDataValidation()
          .requireValueInList(['TRUE', 'FALSE'], true)
          .setAllowInvalid(false)
          .build()
      );
    } else if (key === 'admin_reminder_day') {
      valueRange.setDataValidation(
        SpreadsheetApp.newDataValidation()
          .requireValueInList(getDayOfMonthDropdownValues(), true)
          .setAllowInvalid(false)
          .build()
      );
    }

    if (key === 'forms_folder_id') {
      sheet.getRange(index + 2, 1, 1, Math.min(3, sheet.getLastColumn())).setBackground(formsFolderHighlight);
      sheet.getRange(index + 2, 1).setFontWeight('bold');
    }
  });

  fitSheetToContent(sheet);
  applyTableBordersToDataRange(sheet);
}

function ensureEventsArchiveSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.sheetNames.eventsArchive);
  const headers = ['Enabled', 'Date', 'Event', 'Action', 'Recurring Event', 'Include In Form', 'Include In Schedule', 'Notes', 'Archived At'];

  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.sheetNames.eventsArchive);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  } else if (sheet.getLastRow() < 1) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }

  applySheetTheme(sheet);
  fitSheetToContent(sheet);
  applyTableBordersToDataRange(sheet);

  return sheet;
}

function isExampleEventsRow(row, headerMap) {
  const note = String(getValueByHeader(row, headerMap, ['Notes'], '') || '').trim().toLowerCase();
  const enabled = parseBooleanLike(getValueByHeader(row, headerMap, ['Enabled'], false), false);
  return !enabled && note.indexOf('example') === 0;
}

function archivePastEvents(referenceDate) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const eventsSheet = ss.getSheetByName(CONFIG.sheetNames.events);
  if (!eventsSheet || !sheetUsesFriendlyEventsLayout(eventsSheet) || eventsSheet.getLastRow() < 2) {
    return { archivedRows: 0, keptRows: 0, examplesPreserved: 0 };
  }

  const rows = eventsSheet.getDataRange().getValues();
  const header = rows[0];
  const headerMap = buildHeaderMap(header);
  const now = referenceDate ? new Date(referenceDate) : new Date();
  const cutoff = new Date(now.getFullYear(), now.getMonth(), 1);
  const archivedAt = new Date();

  const keepRows = [];
  const archiveRows = [];
  let examplesPreserved = 0;

  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    if (isBlankRow(row)) continue;

    if (isExampleEventsRow(row, headerMap)) {
      keepRows.push(row);
      examplesPreserved++;
      continue;
    }

    const parsedDate = parseSingleDate(getValueByHeader(row, headerMap, ['Date'], ''));
    if (parsedDate && parsedDate < cutoff) {
      archiveRows.push(row.concat([archivedAt]));
    } else {
      keepRows.push(row);
    }
  }

  if (!archiveRows.length) {
    seedSheetRowsIfEmpty(eventsSheet, getEventsSeedRows());
    configureEventsSheetUi(eventsSheet);
    return { archivedRows: 0, keptRows: keepRows.length, examplesPreserved: examplesPreserved };
  }

  const archiveSheet = ensureEventsArchiveSheet();
  archiveSheet.getRange(archiveSheet.getLastRow() + 1, 1, archiveRows.length, archiveRows[0].length).setValues(archiveRows);
  archiveSheet.getRange(2, 2, Math.max(archiveSheet.getLastRow() - 1, 1), 1).setNumberFormat('yyyy-mm-dd');

  const maxColumns = Math.max(eventsSheet.getLastColumn(), header.length);
  if (eventsSheet.getMaxRows() < keepRows.length + 1) {
    eventsSheet.insertRowsAfter(eventsSheet.getMaxRows(), keepRows.length + 1 - eventsSheet.getMaxRows());
  }
  if (eventsSheet.getMaxColumns() < maxColumns) {
    eventsSheet.insertColumnsAfter(eventsSheet.getMaxColumns(), maxColumns - eventsSheet.getMaxColumns());
  }

  eventsSheet.getRange(2, 1, Math.max(eventsSheet.getMaxRows() - 1, 1), maxColumns).clearContent();
  if (keepRows.length) {
    eventsSheet.getRange(2, 1, keepRows.length, header.length).setValues(keepRows);
  }

  seedSheetRowsIfEmpty(eventsSheet, getEventsSeedRows());
  configureEventsSheetUi(eventsSheet);

  console.log(`Archived ${archiveRows.length} past event row(s) from Events to ${CONFIG.sheetNames.eventsArchive}.`);
  return {
    archivedRows: archiveRows.length,
    keptRows: keepRows.length,
    examplesPreserved: examplesPreserved
  };
}

function archivePastEventsIfDue(referenceDate, settings) {
  const runtimeSettings = settings || loadRuntimeSettings();
  if (!shouldArchiveEventsNow(referenceDate, runtimeSettings)) {
    return { status: 'skipped_schedule', archivedRows: 0 };
  }

  const result = archivePastEvents(referenceDate);
  return {
    status: result.archivedRows ? 'archived' : 'no_changes',
    archivedRows: result.archivedRows,
    keptRows: result.keptRows,
    examplesPreserved: result.examplesPreserved
  };
}

function columnToLetter(columnNumber) {
  let column = columnNumber;
  let letter = '';
  while (column > 0) {
    const temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = Math.floor((column - temp - 1) / 26);
  }
  return letter;
}

function getRoleCheckboxStartColumn(sheet) {
  return getMinistryMembersColumnMap(sheet).firstRoleCheckbox;
}

function getMinistryMembersHeaderRow() {
  return [
    CONFIG.sheetHeaders.name,
    CONFIG.sheetHeaders.dates,
    CONFIG.sheetHeaders.times,
    CONFIG.sheetHeaders.comments,
    CONFIG.sheetHeaders.roles
  ];
}

function ensureMinistryMembersLayout(sheet) {
  if (!sheet) return;

  const baseHeaders = getMinistryMembersHeaderRow();
  const canonicalHeader = CONFIG.sheetHeaders.canonicalName;
  const lastRow = Math.max(sheet.getLastRow(), 1);
  const lastColumn = Math.max(sheet.getLastColumn(), baseHeaders.length + 1);

  if (sheet.getMaxColumns() < lastColumn) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), lastColumn - sheet.getMaxColumns());
  }

  const existingHeaders = sheet.getRange(1, 1, 1, lastColumn).getDisplayValues()[0];
  const headerMap = buildHeaderMap(existingHeaders);
  const knownHeaders = baseHeaders.concat([canonicalHeader]).map(header => String(header || '').trim().toLowerCase());
  const extraHeaders = existingHeaders
    .map(header => String(header || '').trim())
    .filter(header => header && knownHeaders.indexOf(header.toLowerCase()) === -1);
  const desiredHeaders = baseHeaders.concat(extraHeaders).concat([canonicalHeader]);
  const desiredWidth = desiredHeaders.length;
  const output = [desiredHeaders];

  if (lastRow > 1) {
    const rows = sheet.getRange(2, 1, lastRow - 1, lastColumn).getValues();
    rows.forEach(row => {
      output.push(desiredHeaders.map(header => {
        const headerKey = String(header || '').trim().toLowerCase();
        const existingIndex = headerMap[headerKey];
        return existingIndex === undefined ? '' : row[existingIndex];
      }));
    });
  }

  sheet.getRange(1, 1, output.length, desiredWidth).setValues(output);
  sheet.getRange(1, 1, 1, desiredWidth).setFontWeight('bold');
}

function getConfiguredMemberRoleColumnMap(sheet) {
  const map = {};
  if (!sheet) return map;

  const memberColumns = getMinistryMembersColumnMap(sheet);
  const startColumn = getRoleCheckboxStartColumn(sheet);
  const endColumn = Math.min(sheet.getLastColumn(), memberColumns.canonicalName - 1);
  const width = endColumn - startColumn + 1;
  if (width <= 0) return map;

  const headers = sheet.getRange(1, startColumn, 1, width).getDisplayValues()[0];
  headers.forEach((header, index) => {
    const role = String(header || '').trim();
    if (!role) return;
    map[role.toUpperCase()] = startColumn + index;
  });
  return map;
}

function hasAnyRoleCheckboxHeaders(sheet) {
  return Object.keys(getConfiguredMemberRoleColumnMap(sheet)).length > 0;
}

function ensureRoleCheckboxColumns(sheet, roles) {
  if (!sheet || !roles || !roles.length) return;

  const startColumn = getRoleCheckboxStartColumn(sheet);
  const roleColumnMap = getConfiguredMemberRoleColumnMap(sheet);
  let memberColumns = getMinistryMembersColumnMap(sheet);

  roles.forEach(role => {
    const normalized = String(role || '').trim().toUpperCase();
    if (!normalized || roleColumnMap[normalized]) return;

    memberColumns = getMinistryMembersColumnMap(sheet);
    const canonicalColumn = memberColumns.canonicalName;
    let targetColumn = canonicalColumn;

    if (canonicalColumn >= 1 && canonicalColumn <= sheet.getMaxColumns()) {
      sheet.insertColumnsBefore(canonicalColumn, 1);
      targetColumn = canonicalColumn;
    } else {
      const nextColumn = Math.max(sheet.getLastColumn() + 1, startColumn);
      if (sheet.getMaxColumns() < nextColumn) {
        sheet.insertColumnsAfter(sheet.getMaxColumns(), nextColumn - sheet.getMaxColumns());
      }
      targetColumn = nextColumn;
    }

    sheet.getRange(1, targetColumn).setValue(role).setFontWeight('bold');
    if (sheet.getMaxRows() > 1) {
      sheet.getRange(2, targetColumn, sheet.getMaxRows() - 1, 1).insertCheckboxes();
    }
    roleColumnMap[normalized] = targetColumn;
  });

  const finalRoleColumnMap = getConfiguredMemberRoleColumnMap(sheet);
  const finalColumns = Object.keys(finalRoleColumnMap).map(key => finalRoleColumnMap[key]);
  if (finalColumns.length) {
    const firstColumn = Math.min.apply(null, finalColumns);
    const lastColumn = Math.max.apply(null, finalColumns);
    sheet.autoResizeColumns(firstColumn, lastColumn - firstColumn + 1);
  }
}

function normalizeMinistryMembersColumnValidation(sheet, roles) {
  if (!sheet || sheet.getMaxRows() < 2) return;

  const memberColumns = getMinistryMembersColumnMap(sheet);
  const dataRowCount = sheet.getMaxRows() - 1;
  const plainTextColumns = [
    memberColumns.name,
    memberColumns.dates,
    memberColumns.times,
    memberColumns.comments,
    memberColumns.roles,
    memberColumns.canonicalName
  ].filter((column, index, columns) => column && columns.indexOf(column) === index);

  plainTextColumns.forEach(columnNumber => {
    sheet.getRange(2, columnNumber, dataRowCount, 1).clearDataValidations();
  });

  const roleColumnMap = getConfiguredMemberRoleColumnMap(sheet);
  const roleColumnNumbers = Object.keys(roleColumnMap).map(key => roleColumnMap[key]);
  roleColumnNumbers.forEach(columnNumber => {
    sheet.getRange(2, columnNumber, dataRowCount, 1).insertCheckboxes();
  });
}

function syncRoleCheckboxesFromRolesColumn(sheet, roles) {
  if (!sheet || !roles || !roles.length || sheet.getLastRow() < 2) return;

  const roleColumnMap = getConfiguredMemberRoleColumnMap(sheet);
  const lastRow = sheet.getLastRow();
  const memberColumns = getMinistryMembersColumnMap(sheet);
  const rolesText = sheet.getRange(2, memberColumns.roles, lastRow - 1, 1).getDisplayValues().flat();
  const activeRoleColumns = roles
    .map(role => roleColumnMap[String(role || '').trim().toUpperCase()])
    .filter(Boolean);

  for (let i = 0; i < rolesText.length; i++) {
    const rowNumber = i + 2;
    const hasAnyChecked = activeRoleColumns.some(columnNumber => sheet.getRange(rowNumber, columnNumber).getValue() === true);
    if (hasAnyChecked) continue;

    const parsedRoles = String(rolesText[i] || '')
      .split(',')
      .map(role => role.trim().toUpperCase())
      .filter(Boolean);
    if (!parsedRoles.length) continue;

    roles.forEach(role => {
      const normalized = String(role || '').trim().toUpperCase();
      const columnNumber = roleColumnMap[normalized];
      if (!columnNumber) return;
      if (parsedRoles.indexOf(normalized) !== -1) {
        sheet.getRange(rowNumber, columnNumber).setValue(true);
      }
    });
  }
}

function ensureRolesFormulaColumn(sheet, roles) {
  if (!sheet || !roles || !roles.length) return;

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const roleColumnMap = getConfiguredMemberRoleColumnMap(sheet);
  const roleColumns = roles
    .map(role => roleColumnMap[String(role || '').trim().toUpperCase()])
    .filter(Boolean);
  if (!roleColumns.length) return;

  const formulas = [];
  for (let row = 2; row <= lastRow; row++) {
    const headerRefs = roleColumns.map(column => `${columnToLetter(column)}$1`);
    const rowRefs = roleColumns.map(column => `${columnToLetter(column)}${row}`);
    formulas.push([
      `=IF(COUNTIF({${rowRefs.join(',')}},TRUE)=0,"",TEXTJOIN(", ",TRUE,FILTER({${headerRefs.join(',')}},{${rowRefs.join(',')}}=TRUE)))`
    ]);
  }

  sheet.getRange(2, getMinistryMembersColumnMap(sheet).roles, formulas.length, 1).setFormulas(formulas);
}

function configureMinistryMembersRoles(sheet, roles) {
  if (!sheet || !roles || !roles.length) return;
  ensureRoleCheckboxColumns(sheet, roles);
  normalizeMinistryMembersColumnValidation(sheet, roles);
  syncRoleCheckboxesFromRolesColumn(sheet, roles);
  ensureRolesFormulaColumn(sheet, roles);
  applySheetTheme(sheet);
  fitSheetToContent(sheet);
  applyTableBordersToDataRange(sheet);
}

function hasRoleCheckboxColumnsConfigured(sheet, roles) {
  if (!sheet || !roles || !roles.length) return false;
  const roleColumnMap = getConfiguredMemberRoleColumnMap(sheet);
  return roles.every(role => roleColumnMap[String(role || '').trim().toUpperCase()]);
}

function ensureRolesFormulaForRow(sheet, rowNumber, roles) {
  if (!sheet || !roles || !roles.length || rowNumber < 2) return;
  if (!hasRoleCheckboxColumnsConfigured(sheet, roles)) return;

  const roleColumnMap = getConfiguredMemberRoleColumnMap(sheet);
  const roleColumns = roles
    .map(role => roleColumnMap[String(role || '').trim().toUpperCase()])
    .filter(Boolean);
  if (!roleColumns.length) return;

  const headerRefs = roleColumns.map(column => `${columnToLetter(column)}$1`);
  const rowRefs = roleColumns.map(column => `${columnToLetter(column)}${rowNumber}`);
  const formula = `=IF(COUNTIF({${rowRefs.join(',')}},TRUE)=0,"",TEXTJOIN(", ",TRUE,FILTER({${headerRefs.join(',')}},{${rowRefs.join(',')}}=TRUE)))`;
  sheet.getRange(rowNumber, getMinistryMembersColumnMap(sheet).roles).setFormula(formula);
}

function summarizeRoleMigration(sheet, roles) {
  const summary = {
    roleCount: roles.length,
    memberRows: 0,
    rowsWithLegacyRoles: 0,
    conflictingCells: 0,
    conflictSamples: []
  };

  if (!sheet || sheet.getLastRow() < 2 || !roles.length) return summary;

  const lastRow = sheet.getLastRow();
  const memberColumns = getMinistryMembersColumnMap(sheet);
  const startColumn = getRoleCheckboxStartColumn(sheet);
  const requiredLastColumn = startColumn + roles.length - 1;
  summary.memberRows = lastRow - 1;

  const rolesText = sheet.getRange(2, memberColumns.roles, lastRow - 1, 1).getDisplayValues().flat();
  summary.rowsWithLegacyRoles = rolesText.filter(value => String(value || '').trim()).length;

  if (sheet.getLastColumn() >= startColumn) {
    const inspectWidth = Math.max(0, Math.min(requiredLastColumn, memberColumns.canonicalName - 1, sheet.getLastColumn()) - startColumn + 1);
    if (inspectWidth > 0) {
      const existing = sheet.getRange(1, startColumn, lastRow, inspectWidth).getDisplayValues();
      for (let r = 0; r < existing.length; r++) {
        for (let c = 0; c < existing[r].length; c++) {
          const value = String(existing[r][c] || '').trim();
          if (!value) continue;
          if (r === 0 && value === String(roles[c] || '').trim()) continue;
          summary.conflictingCells++;
          if (summary.conflictSamples.length < 5) {
            summary.conflictSamples.push(`${columnToLetter(startColumn + c)}${r + 1}=${value}`);
          }
        }
      }
    }
  }

  return summary;
}

function migrateMemberRolesToCheckboxes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.sheetNames.ministryMembers);
  if (!sheet) throw new Error('Ministry Members sheet not found.');
  ensureMinistryMembersLayout(sheet);

  const roles = loadRuntimeSettings().roles;
  if (!roles.length) throw new Error(`No roles configured in ${CONFIG.sheetNames.rolesConfig}.`);

  if (hasRoleCheckboxColumnsConfigured(sheet, roles)) {
    configureMinistryMembersRoles(sheet, roles);
    SpreadsheetApp.flush();
    console.log('Role checkbox migration skipped because the sheet is already configured.');
    return {
      status: 'already_configured',
      roles: roles,
      startColumn: columnToLetter(getRoleCheckboxStartColumn(sheet)),
      migratedRows: 0,
      memberRows: Math.max(sheet.getLastRow() - 1, 0)
    };
  }

  if (hasAnyRoleCheckboxHeaders(sheet)) {
    configureMinistryMembersRoles(sheet, roles);
    SpreadsheetApp.flush();
    console.log('Role checkbox columns refreshed from the Roles sheet.');
    return {
      status: 'updated_roles',
      roles: roles,
      startColumn: columnToLetter(getRoleCheckboxStartColumn(sheet)),
      migratedRows: 0,
      memberRows: Math.max(sheet.getLastRow() - 1, 0)
    };
  }

  const summary = summarizeRoleMigration(sheet, roles);
  if (summary.conflictingCells > 0) {
    throw new Error(
      `Role checkbox migration stopped because columns ${columnToLetter(getRoleCheckboxStartColumn(sheet))}+ already contain data. ` +
      `Conflicts found: ${summary.conflictingCells}. Examples: ${summary.conflictSamples.join(', ')}`
    );
  }

  configureMinistryMembersRoles(sheet, roles);
  SpreadsheetApp.flush();

  console.log(
    `Role migration complete. Added ${roles.length} checkbox columns starting at ${columnToLetter(getRoleCheckboxStartColumn(sheet))}. ` +
    `Migrated ${summary.rowsWithLegacyRoles} member row(s) from the Roles column.`
  );

  return {
    status: 'migrated',
    roles: roles,
    startColumn: columnToLetter(getRoleCheckboxStartColumn(sheet)),
    migratedRows: summary.rowsWithLegacyRoles,
    memberRows: summary.memberRows
  };
}

function maybeMigrateMemberRolesDuringSetup() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.sheetNames.ministryMembers);
  if (!sheet) return { status: 'skipped', reason: 'missing_sheet' };

  const roles = loadRuntimeSettings().roles;
  if (!roles.length) return { status: 'skipped', reason: 'missing_roles' };

  try {
    return migrateMemberRolesToCheckboxes();
  } catch (error) {
    console.warn('Skipping automatic role checkbox migration during initializeProject: ' + error.message);
    return {
      status: 'skipped_conflict',
      reason: error.message
    };
  }
}

function syncConfiguredMemberRoles() {
  const result = maybeMigrateMemberRolesDuringSetup();
  if (result && result.status === 'skipped_conflict') {
    console.warn('Role checkbox sync skipped: ' + result.reason);
  }
  return result;
}

function getMemberRolesFromRow(row, configuredRoles, roleColumnMap, memberColumns) {
  const roles = (configuredRoles || []).map(role => String(role || '').trim()).filter(Boolean);
  if (!row || !roles.length) return [];
  const resolvedMemberColumns = memberColumns || getMinistryMembersColumnMap();

  const normalizedColumnMap = roleColumnMap || {};
  const checkboxRoles = roles.filter(role => {
    const columnNumber = normalizedColumnMap[String(role || '').trim().toUpperCase()];
    return columnNumber && row[columnNumber - 1] === true;
  });
  if (checkboxRoles.length) {
    return checkboxRoles.map(role => role.toUpperCase());
  }

  return row[resolvedMemberColumns.roles - 1]
    ? row[resolvedMemberColumns.roles - 1].toString().split(",").map(role => role.trim().toUpperCase()).filter(Boolean)
    : [];
}

function setupAvailability(sheetName, year, month) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const runtimeSettings = loadRuntimeSettings();
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName); // If the sheet doesn't exist, create it
  } else {
    sheet.clear(); // Clear the existing sheet if it exists
  }

  // Get next month's Sundays dynamically
  const serviceDates = getServiceDates(year, month);

  const headerRow = ["Schedule"].concat(serviceDates);
  sheet.appendRow(headerRow); // Adding the header row to the sheet

  // Select the header row range and make it bold
  const headerRange = sheet.getRange(1, 1, 1, headerRow.length);
  headerRange.setFontWeight("bold"); // Make the header text bold

  // Define the roles (without any members for now)
  const roles = runtimeSettings.roles;

  // Add each role with empty cells under each Sunday
  roles.forEach(role => {
    const roleRow = [role];
    serviceDates.forEach(() => {
      roleRow.push(""); // Adding empty cells for each Sunday
    });
    sheet.appendRow(roleRow); // Add the row for the role
  });

  // Apply bold formatting to all the rows with roles
  const lastRow = sheet.getLastRow();
  sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).setFontWeight("bold"); // Make role rows bold

  // Add 3 empty rows of space before the availability section
  const insertionRow = lastRow + 1;
  sheet.insertRowsAfter(insertionRow, 3);

  // Add "Availability" above the role section
  const availabilityHeaderRow = insertionRow + 3;
  sheet.getRange(availabilityHeaderRow, 1).setValue("Availability").setFontWeight("bold");
  const availabilityRange = sheet.getRange(sheet.getLastRow(), 1, 1, sheet.getLastColumn());
  availabilityRange.setFontWeight("bold"); // Make the "Availability" text bold

  // Auto-resize the columns to fit the content
  sheet.autoResizeColumns(1, sheet.getLastColumn());

  // Set up empty data below the "Availability" section for each role
  const emptyData = roles.map(role => [role, ...Array(serviceDates.length).fill("")]);

  // Add empty data under the "Availability" heading for each role
  emptyData.forEach(dataRow => {
    sheet.appendRow(dataRow); // Add the empty data row for the role
  });

  fitSheetToContent(sheet);
  applyTableBorder(sheet, 1, 1, roles.length + 1, headerRow.length);
  applyTableBorder(sheet, availabilityHeaderRow, 1, roles.length + 1, headerRow.length);
}

function clearByHeader(header) {
    // Open the spreadsheet by its ID
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Access the "Ministry Members" sheet
  const sheet = ss.getSheetByName(CONFIG.sheetNames.ministryMembers);
  
  if (!sheet) {
    console.log("Sheet not found: " + CONFIG.sheetNames.ministryMembers);
    return;
  }
  
  // Get the data range of the sheet
  const dataRange = sheet.getDataRange();
  
  // Get the values in the first row to find the "Not Available Dates" column
  const headers = dataRange.getValues()[0];
  
  // Find the index of the provided header column
  const colIndex = headers.indexOf(header) + 1; // +1 to convert to 1-based index
  
  if (colIndex === 0) {
    console.log(header + ' column not found.');
    return;
  }
  
  // Determine the range to clear: from row 2 to the last row in the identified column
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    console.log("No data to clear.");
    return;
  }
  
  // Clear the contents of the column, starting from row 2
  sheet.getRange(2, colIndex, lastRow - 1).clearContent();
  
  console.log(header + ' column cleared.');
}

function cleanupPreparedMonthlyArtifacts(ss, stagingSheetName, preparedFormId, existingResponseSheetIds) {
  const spreadsheet = ss || SpreadsheetApp.getActiveSpreadsheet();

  if (preparedFormId) {
    try {
      FormApp.openById(preparedFormId).removeDestination();
    } catch (unlinkError) {
      console.log('Could not remove destination from staged form ' + preparedFormId + ': ' + unlinkError.message);
    }

    try {
      DriveApp.getFileById(preparedFormId).setTrashed(true);
      console.log('Deleted staged form ' + preparedFormId);
    } catch (trashError) {
      console.log('Could not delete staged form ' + preparedFormId + ': ' + trashError.message);
    }
  }

  const existingIds = {};
  (existingResponseSheetIds || []).forEach(sheetId => {
    existingIds[String(sheetId)] = true;
  });

  spreadsheet.getSheets().forEach(sheet => {
    if (!sheet.getName().startsWith("Form Responses")) return;
    if (Object.prototype.hasOwnProperty.call(existingIds, String(sheet.getSheetId()))) return;

    try {
      spreadsheet.deleteSheet(sheet);
      console.log("Deleted staged Form Responses tab: " + sheet.getName());
    } catch (error) {
      console.log("Could not delete staged Form Responses tab '" + sheet.getName() + "': " + error.message);
    }
  });

  const stagingSheet = stagingSheetName ? spreadsheet.getSheetByName(stagingSheetName) : null;
  if (stagingSheet) {
    try {
      spreadsheet.deleteSheet(stagingSheet);
      console.log("Deleted staging tab: " + stagingSheetName);
    } catch (error) {
      console.log("Could not delete staging tab '" + stagingSheetName + "': " + error.message);
    }
  }
}

function runMonthlySetupInternal(options) {
  const opts = options || {};
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  syncConfiguredMemberRoles();
  const metadataSheet = ss.getSheetByName("Form Metadata") || ss.insertSheet("Form Metadata");
  const runtimeSettings = loadRuntimeSettings();
  const trackedFormIds = getTrackedFormIds(metadataSheet);
  const existingResponseSheetIds = getFormResponseSheetIds(ss);
  const today = opts.referenceDate ? new Date(opts.referenceDate) : new Date();
  const propertyKey = getMonthlySetupPropertyKey(today);
  const props = PropertiesService.getScriptProperties();

  if (!opts.force && !shouldRunMonthlySetupToday(today, runtimeSettings)) {
    return {
      status: 'not_due_today',
      day: today.getDate(),
      formCreationDay: runtimeSettings.formCreationDay
    };
  }

  if (!opts.force && props.getProperty(propertyKey)) {
    return {
      status: 'already_ran',
      key: propertyKey
    };
  }

  try {
    const archiveResult = archivePastEventsIfDue(today, runtimeSettings);
    if (archiveResult.status === 'archived') {
      console.log(`Archived ${archiveResult.archivedRows} past Events row(s) before monthly setup.`);
    }
  } catch (err) {
    console.error('Failed to archive past Events rows: ' + err.message);
  }

  const planDate = new Date(today);
  planDate.setMonth(today.getMonth() + 1);

  const oldDate = new Date(today);
  oldDate.setMonth(today.getMonth() - 1);

  const todayMonthName = today.toLocaleString('default', { month: 'long' });
  const todayMonth = today.getMonth();
  const todayYear = today.getFullYear();

  const planMonthName = planDate.toLocaleString('default', { month: 'long' });
  const planMonth = planDate.getMonth();
  const planYear = planDate.getFullYear();

  const oldMonthName = oldDate.toLocaleString('default', { month: 'long' });
  const oldMonth = oldDate.getMonth();
  const oldYear = oldDate.getFullYear();

  const newTabName = getAvailabilitySheetNameForMonthName(planMonthName, runtimeSettings);
  const deleteTabName = getAvailabilitySheetNameForMonthName(oldMonthName, runtimeSettings);
  const stagingSheetName = `${newTabName} (Staging)`;
  let preparedForm = null;
  let setupCommitted = false;

  if (!opts.force && ss.getSheetByName(newTabName)) {
    throw new Error(`The ${newTabName} sheet already exists. Please review it before running monthlySetup again.`);
  }

  if (ss.getSheetByName(stagingSheetName)) {
    ss.deleteSheet(ss.getSheetByName(stagingSheetName));
  }

  try {
    setupAvailability(stagingSheetName, planYear, planMonth);

    preparedForm = createNewFormForMonth(planMonth, planYear, planMonthName, {
      settings: runtimeSettings,
      sheetName: stagingSheetName,
      skipMetadata: true,
      skipEmail: true,
      deferDestination: true
    });

    FormApp.openById(preparedForm.formId).setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());
    console.log("Linked staged form responses to spreadsheet");

    const existingNewTab = ss.getSheetByName(newTabName);
    if (existingNewTab) {
      ss.deleteSheet(existingNewTab);
      console.log("Replaced existing tab: " + newTabName);
    }

    const stagingSheet = ss.getSheetByName(stagingSheetName);
    if (!stagingSheet) {
      throw new Error("Staging availability tab was not found during monthly setup.");
    }
    stagingSheet.setName(newTabName);

    clearByHeader(CONFIG.sheetHeaders.times);
    clearByHeader(CONFIG.sheetHeaders.dates);
    clearByHeader(CONFIG.sheetHeaders.comments);

    writeCurrentFormMetadata(metadataSheet, preparedForm.metadataLabel, preparedForm.formId);

    unlinkTrackedForms(trackedFormIds);
    deleteFormResponseSheetsById(ss, existingResponseSheetIds);

    const oldSheet = ss.getSheetByName(deleteTabName);
    if (oldSheet) {
      ss.deleteSheet(oldSheet);
      console.log("Deleted old tab: " + deleteTabName);
    }

    setupCommitted = true;
  } catch (error) {
    if (!setupCommitted) {
      cleanupPreparedMonthlyArtifacts(ss, stagingSheetName, preparedForm ? preparedForm.formId : '', existingResponseSheetIds);
    }
    throw error;
  }

  try {
    sendNewFormCreatedEmail(planMonthName, preparedForm.responderUrl, preparedForm.editUrl, runtimeSettings);
  } catch (error) {
    console.error('Failed to send new form email: ' + error.message);
  }
  console.log(`Created new form for ${planMonthName}`);

  // Legacy support: populate automatic events only for old-style Events/Monthly Events sheets.
  try {
    ensureMonthlyEventsFor(planYear, planMonth);
  } catch (err) {
    console.error('Failed to ensure legacy Events entries: ' + err.message);
  }

  if (!isDeveloperDiagnosticsEnabled()) {
    hideDeveloperSheets();
  }

  reorderWorkbookSheets(today);
  props.setProperty(propertyKey, new Date().toISOString());
  return {
    status: 'completed',
    key: propertyKey,
    planMonthName: planMonthName,
    planYear: planYear,
    planMonth: planMonth + 1
  };
}

function monthlySetup() {
  return runMonthlySetupInternal();
}

function runMonthlySetupNow() {
  return runMonthlySetupInternal({ force: true });
}

function findFormResponseSheet() {
  // Open the active spreadsheet
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Get all sheets in the spreadsheet
  const sheets = ss.getSheets();
  
  // Iterate through each sheet to find the form response sheet
  for (let i = 0; i < sheets.length; i++) {
    const sheet = sheets[i];
    const sheetName = sheet.getName();
    
    // Check if the sheet name starts with "Form Responses"
    if (sheetName.startsWith("Form Responses")) {
      // Further verification can be done here, such as checking specific headers
      return sheet; // Return the identified sheet
    }
  }
  
  // If no sheet is found, return null or handle accordingly
  return null;
}

function updateAvailability(referenceDate) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  syncConfiguredMemberRoles();
  const runtimeSettings = loadRuntimeSettings();
  console.log('--- STARTING updateAvailability ---');

  const planningContext = getPlanningMonthContext(referenceDate, runtimeSettings);
  const sheetName = planningContext.sheetName;

  const matrixSheet = ss.getSheetByName(sheetName);
  const databaseSheet = ss.getSheetByName(CONFIG.sheetNames.ministryMembers);

  if (!matrixSheet || !databaseSheet) {
    console.log("Error: One or more required sheets are missing.");
    return {
      status: !matrixSheet ? 'missing_sheet' : 'missing_database',
      sheetName: sheetName
    };
  }

  const memberColumns = getMinistryMembersColumnMap(databaseSheet);

  const databaseData = databaseSheet.getDataRange().getValues();

  if (!databaseData.length) {
    console.log("No data found in the Ministry Members sheet.");
    return { status: 'missing_database_rows', sheetName: sheetName };
  }

  // Get Date Headers from the sheet (Row 1, starting from column 2)
  let lastCol = matrixSheet.getLastColumn();
  if (lastCol <= 1) {
    console.log("Error: Availability matrix has no date columns.");
    return { status: 'missing_date_columns', sheetName: sheetName };
  }
  const headerRowValues = matrixSheet.getRange(1, 2, 1, lastCol - 1).getValues();
  let dateHeaders = headerRowValues[0];
  console.log('Matrix Date Headers (Raw): ' + dateHeaders.join(', '));

  lastCol = matrixSheet.getLastColumn();

  dateHeaders = matrixSheet
    .getRange(CONFIG.layout.dateRowIndex, 2, 1, lastCol - 1)
    .getDisplayValues()[0];

  const serviceDateKeys = dateHeaders.map(extractDateKey);

  console.log('Standardized Date Keys: ' + serviceDateKeys.join(', '));

  // Initialize the availability object
  const availability = {};
  let roleOrder = runtimeSettings.roles;
  const roleColumnMap = getConfiguredMemberRoleColumnMap(databaseSheet);

  // Standardize roleOrder to uppercase for case-insensitive matching
  roleOrder = roleOrder.map(role => role.toUpperCase());

  // Process each row in the Ministry Members sheet
  for (let i = 1; i < databaseData.length; i++) {
    const row = databaseData[i];
    let name = row[memberColumns.name - 1] ? row[memberColumns.name - 1].trim() : "";
    if (!name) continue;

    // Ensure canonical name exists in the canonical-name column
    try {
      const existingCanonical = row[memberColumns.canonicalName - 1] ? row[memberColumns.canonicalName - 1].toString() : '';
      if (!existingCanonical && name) {
        const canon = normalizeName(name);
        databaseSheet.getRange(i + 1, memberColumns.canonicalName).setValue(canon);
        // Update local copy so further logic can use it if needed
        row[memberColumns.canonicalName - 1] = canon;
      }
    } catch (err) {
      console.error('Failed to persist canonical in updateAvailability for row ' + (i + 1) + ': ' + err.message);
    }
    
    const roles = getMemberRolesFromRow(row, runtimeSettings.roles, roleColumnMap, memberColumns);
    const timesWilling = row[memberColumns.times - 1] ? row[memberColumns.times - 1].toString().trim() : "";
    const rawUnavailableDates = row[memberColumns.dates - 1] ? row[memberColumns.dates - 1].toString() : "";
    
    const unavailableDates = parseUnavailableDates(rawUnavailableDates).parsed;

    if (!name || !roles.length) continue;

    name = formatAvailabilityDisplayName(name);

    // If "Times Willing to Serve" is blank, mark unavailable for all dates
    const isUnavailableAllMonth = timesWilling === "";

    roles.forEach(role => {
      if (!availability[role]) availability[role] = {};
      
      // *** Use serviceDateKeys for reliable matching ***
      serviceDateKeys.forEach(dateKey => {
        
        // dateKey is the clean "MM/dd" string (e.g., "12/14")
        const date = dateKey; 
        
        if (!availability[role][date]) availability[role][date] = [];
        
        // Add name if not marked unavailable for all dates and not in unavailableDates
        // Comparison is now guaranteed to work: "12/14" === "12/14"
        if (!isUnavailableAllMonth && !unavailableDates.includes(date)) {
          availability[role][date].push(name);
        }
      });
    });
  } // End of main database iteration loop

  // Clear the old values from the matrix (excluding headers) - Run once
  const numRoles = roleOrder.length;
  // Use serviceDateKeys.length to define the range width
  const clearRange = matrixSheet.getRange(CONFIG.layout.headerRowIndex, 2, numRoles, serviceDateKeys.length); 
  clearRange.clearContent();

  // Update the availability matrix in the sheet
  let roleRowIndex = CONFIG.layout.headerRowIndex;
  roleOrder.forEach(role => {
    const roleData = availability[role];
    if (roleData) {
      const namesRow = serviceDateKeys.map(dateKey => {
        // dateKey is the clean "MM/dd" string, used as the lookup key
        return roleData[dateKey] ? roleData[dateKey].join("\n") : "";
      });
      
      // Set values in the sheet
      const range = matrixSheet.getRange(roleRowIndex, 2, 1, namesRow.length);
      range.setValues([namesRow]);
      range.setWrap(false); // Disable text wrapping for the range
      roleRowIndex++;
    }
  });

  matrixSheet.autoResizeColumns(1, matrixSheet.getLastColumn() - 1);
  matrixSheet.autoResizeRows(CONFIG.layout.headerRowIndex, roleRowIndex - CONFIG.layout.headerRowIndex + 1);
  console.log("Availability matrix updated in sheet: " + sheetName);
  console.log('--- FINISHED updateAvailability ---');
  return {
    status: 'updated',
    sheetName: sheetName,
    roleCount: roleOrder.length,
    dateCount: serviceDateKeys.length
  };
}

/**
 * Run this function once to set up the spreadsheet for a new user.
 * It creates the database sheet and metadata sheet with dummy data.
 */
function initializeProject() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Create Ministry Members Database if it doesn't exist
  let dbSheet = ss.getSheetByName(CONFIG.sheetNames.ministryMembers);
  if (!dbSheet) {
    dbSheet = ss.insertSheet(CONFIG.sheetNames.ministryMembers);
    // Add Headers
    const headers = getMinistryMembersHeaderRow().concat([CONFIG.sheetHeaders.canonicalName]);
    dbSheet.appendRow(headers);
    dbSheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");

    console.log("Created Ministry Members sheet.");
  }

  ensureMinistryMembersLayout(dbSheet);
  dbSheet.getRange(1, 1, 1, 6).setFontWeight("bold");
  applySheetTheme(dbSheet);
  fitSheetToContent(dbSheet);
  applyTableBordersToDataRange(dbSheet);

  // 2. Create Form Metadata sheet if it doesn't exist
  let metaSheet = ss.getSheetByName(CONFIG.sheetNames.formMetadata);
  if (!metaSheet) {
    metaSheet = ss.insertSheet(CONFIG.sheetNames.formMetadata);
    metaSheet.appendRow(["Form Name", "Form ID"]);
    console.log("Created Form Metadata sheet.");
  }
  applySheetTheme(metaSheet);
  fitSheetToContent(metaSheet);
  applyTableBordersToDataRange(metaSheet);

  // 3. Create Settings sheet if it doesn't exist
  let settingsSheet = ss.getSheetByName(CONFIG.sheetNames.settings);
  if (!settingsSheet) {
    settingsSheet = ss.insertSheet(CONFIG.sheetNames.settings);
    settingsSheet.getRange(1, 1, 1, 3).setValues([getSettingsSeedRows()[0]]);
    console.log("Created Settings sheet.");
  }
  settingsSheet.getRange(1, 1, 1, 3).setFontWeight('bold');
  ensureSettingsSheetRows(settingsSheet, getSettingsSeedRows());
  configureSettingsSheetUi(settingsSheet);

  // 4. Create Admins sheet if it doesn't exist
  let adminsSheet = ss.getSheetByName(CONFIG.sheetNames.admins);
  if (!adminsSheet) {
    adminsSheet = ss.insertSheet(CONFIG.sheetNames.admins);
    adminsSheet.getRange(1, 1, 1, 3).setValues([getAdminsSeedRows(loadRuntimeSettings().adminEmails)[0]]);
    console.log("Created Admins sheet.");
  }
  normalizeAdminsSheetLayout(adminsSheet);
  adminsSheet.getRange(1, 1, 1, 3).setFontWeight('bold');
  seedSheetRowsIfEmpty(adminsSheet, getAdminsSeedRows(loadRuntimeSettings().adminEmails));
  if (sheetUsesFriendlyAdminsLayout(adminsSheet)) {
    configureAdminsSheetUi(adminsSheet);
  }

  // 5. Create Roles sheet if it doesn't exist
  let rolesSheet = ss.getSheetByName(CONFIG.sheetNames.rolesConfig);
  if (!rolesSheet) {
    rolesSheet = ss.insertSheet(CONFIG.sheetNames.rolesConfig);
    rolesSheet.getRange(1, 1, 1, 3).setValues([getRolesSeedRows(loadRuntimeSettings().roles)[0]]);
    console.log("Created Roles sheet.");
  }
  rolesSheet.getRange(1, 1, 1, 3).setFontWeight('bold');
  seedSheetRowsIfEmpty(rolesSheet, getRolesSeedRows(loadRuntimeSettings().roles));
  configureRolesSheetUi(rolesSheet);

  rewriteSettingsSheetRows(settingsSheet, getSettingsSeedRows());
  configureSettingsSheetUi(settingsSheet);

  const roleMigrationResult = maybeMigrateMemberRolesDuringSetup();
  if (roleMigrationResult && roleMigrationResult.status === 'migrated') {
    console.log('Initialized role checkbox workflow in Ministry Members.');
  } else if (roleMigrationResult && roleMigrationResult.status === 'already_configured') {
    console.log('Role checkbox workflow already configured in Ministry Members.');
  } else if (roleMigrationResult && roleMigrationResult.status === 'updated_roles') {
    console.log('Role checkbox workflow refreshed from the Roles sheet.');
  }

  // 6. Migrate legacy sheet names to the friendlier go-forward names when possible.
  let recurringSheet = ss.getSheetByName(CONFIG.sheetNames.recurring);
  let eventsSheet = ss.getSheetByName(CONFIG.sheetNames.events);
  let legacyRecurringSheet = ss.getSheetByName(CONFIG.sheetNames.recurringEvents);
  let legacyMonthlyEventsSheet = ss.getSheetByName(CONFIG.sheetNames.monthlyEvents);

  if (!recurringSheet && eventsSheet && sheetLooksLikeRecurring(eventsSheet)) {
    eventsSheet.setName(CONFIG.sheetNames.recurring);
    recurringSheet = eventsSheet;
    eventsSheet = ss.getSheetByName(CONFIG.sheetNames.events);
    console.log("Renamed legacy Events sheet to Recurring.");
  }

  if (!recurringSheet && legacyRecurringSheet) {
    legacyRecurringSheet.setName(CONFIG.sheetNames.recurring);
    recurringSheet = legacyRecurringSheet;
    legacyRecurringSheet = ss.getSheetByName(CONFIG.sheetNames.recurringEvents);
    console.log("Renamed Recurring Events sheet to Recurring.");
  }

  if (!eventsSheet && legacyMonthlyEventsSheet) {
    legacyMonthlyEventsSheet.setName(CONFIG.sheetNames.events);
    eventsSheet = legacyMonthlyEventsSheet;
    legacyMonthlyEventsSheet = ss.getSheetByName(CONFIG.sheetNames.monthlyEvents);
    console.log("Renamed Monthly Events sheet to Events.");
  }

  // 7. Create the new Recurring sheet if no recurring configuration exists.
  if (!recurringSheet && !legacyRecurringSheet) {
    recurringSheet = ss.insertSheet(CONFIG.sheetNames.recurring);
    recurringSheet.getRange(1, 1, 1, 8).setValues([getRecurringSeedRows()[0]]);
    seedSheetRowsIfEmpty(recurringSheet, getRecurringSeedRows());
    configureRecurringSheetUi(recurringSheet);
    console.log("Created Recurring sheet.");
  } else if (recurringSheet && sheetUsesFriendlyRecurringLayout(recurringSheet)) {
    seedSheetRowsIfEmpty(recurringSheet, getRecurringSeedRows());
    configureRecurringSheetUi(recurringSheet);
  }

  // 8. Create the new Events sheet for month-specific additions/removals.
  if (!eventsSheet) {
    eventsSheet = ss.insertSheet(CONFIG.sheetNames.events);
    eventsSheet.getRange(1, 1, 1, 8).setValues([getEventsSeedRows()[0]]);
    seedSheetRowsIfEmpty(eventsSheet, getEventsSeedRows());
    eventsSheet.getRange(1, 1, 1, 8).setFontWeight('bold');
    configureEventsSheetUi(eventsSheet);
    console.log("Created Events sheet.");
  } else if (sheetUsesFriendlyEventsLayout(eventsSheet)) {
    seedSheetRowsIfEmpty(eventsSheet, getEventsSeedRows());
    configureEventsSheetUi(eventsSheet);
  }

  if (!isDeveloperDiagnosticsEnabled()) {
    hideDeveloperSheets();
  }

  reorderWorkbookSheets();
  
  console.log("Initialization complete. You can now run monthlySetup().");
}
