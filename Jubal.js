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
 * - If you would like to report any issues with this script, email: widodoalfianto94@gmail.com
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

/**
 * Centralized debug logger. Writes to console and a lightweight sheet for persistent logs.
 * level: 'info' | 'warn' | 'error'
 */
function logDebug(level, msg, data) {
  try {
    const payload = { ts: new Date().toISOString(), level: level, message: msg, data: data || null };
    // Console log for quick inspection in executions
    console.log(JSON.stringify(payload));

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
    const dt = new Date(year, month, day);
    return isNaN(dt.getTime()) ? null : dt;
  }

  // Match ISO YYYY-MM-DD
  const iso = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (iso) {
    const dt = new Date(parseInt(iso[1], 10), parseInt(iso[2], 10) - 1, parseInt(iso[3], 10));
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

function getDefaultRuntimeSettings() {
  return {
    churchName: CONFIG.defaults.churchName,
    timeZone: safeGetScriptTimeZone(),
    formsFolder: CONFIG.ids.formsFolder || '',
    adminEmails: CONFIG.ids.adminEmails.slice(),
    roles: CONFIG.roles.slice(),
    formCreationDay: CONFIG.defaults.formCreationDay,
    timesChoices: CONFIG.defaults.timesChoices.slice(),
    availabilitySheetSuffix: CONFIG.defaults.availabilitySheetSuffix,
    adminReminderEnabled: CONFIG.defaults.adminReminderEnabled,
    adminReminderDay: CONFIG.defaults.adminReminderDay,
    eventsArchiveFrequency: CONFIG.defaults.eventsArchiveFrequency,
    eventsArchiveMonth: CONFIG.defaults.eventsArchiveMonth
  };
}

function loadRuntimeSettings() {
  const defaults = getDefaultRuntimeSettings();
  const raw = loadKeyValueSheet(CONFIG.sheetNames.settings);
  const sheetAdminEmails = loadAdminEmailsFromSheet();
  if (!Object.keys(raw).length) {
    if (sheetAdminEmails.length) {
      const fallbackDefaults = Object.assign({}, defaults);
      fallbackDefaults.adminEmails = sheetAdminEmails;
      return fallbackDefaults;
    }
    return defaults;
  }

  const configuredAdminEmails = normalizeEmailList(raw.admin_emails);

  const settings = {
    churchName: String(raw.church_name || defaults.churchName).trim() || defaults.churchName,
    timeZone: String(raw.time_zone || defaults.timeZone).trim() || defaults.timeZone,
    formsFolder: String(raw.forms_folder_id || defaults.formsFolder).trim() || defaults.formsFolder,
    adminEmails: sheetAdminEmails.length ? sheetAdminEmails : configuredAdminEmails,
    roles: parseCsv(raw.roles),
    formCreationDay: toIntegerOrDefault(raw.form_creation_day, defaults.formCreationDay),
    timesChoices: parseCsv(raw.times_choices),
    availabilitySheetSuffix: String(raw.availability_sheet_suffix || defaults.availabilitySheetSuffix).trim() || defaults.availabilitySheetSuffix,
    adminReminderEnabled: parseBooleanLike(raw.admin_reminder_enabled, defaults.adminReminderEnabled),
    adminReminderDay: toIntegerOrDefault(raw.admin_reminder_day, Math.max(toIntegerOrDefault(raw.form_creation_day, defaults.formCreationDay) - 3, 1)),
    eventsArchiveFrequency: String(raw.events_archive_frequency || defaults.eventsArchiveFrequency).trim() || defaults.eventsArchiveFrequency,
    eventsArchiveMonth: String(raw.events_archive_month || defaults.eventsArchiveMonth).trim() || defaults.eventsArchiveMonth
  };

  if (!settings.adminEmails.length) settings.adminEmails = defaults.adminEmails.slice();
  if (!settings.roles.length) settings.roles = defaults.roles.slice();
  if (!settings.timesChoices.length) settings.timesChoices = defaults.timesChoices.slice();
  if (settings.adminReminderDay < 1) settings.adminReminderDay = 1;
  if (!settings.eventsArchiveFrequency) settings.eventsArchiveFrequency = defaults.eventsArchiveFrequency;
  if (!settings.eventsArchiveMonth) settings.eventsArchiveMonth = defaults.eventsArchiveMonth;

  return settings;
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
  const runtimeSettings = settings || loadRuntimeSettings();
  const suffix = String(runtimeSettings.availabilitySheetSuffix || CONFIG.defaults.availabilitySheetSuffix).trim();
  return suffix ? `${monthName} ${suffix}` : monthName;
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

function getArchiveTriggerMonths(frequency, archiveMonth) {
  const normalizedFrequency = normalizeArchiveFrequency(frequency);
  const monthNumber = parseMonthNumber(archiveMonth);

  if (normalizedFrequency === 'off') return [];
  if (normalizedFrequency === 'monthly') return [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12];
  if (normalizedFrequency === 'quarterly') return [1, 4, 7, 10];
  return [monthNumber || 1];
}

function shouldArchiveEventsNow(referenceDate, settings) {
  const runtimeSettings = settings || loadRuntimeSettings();
  const triggerMonths = getArchiveTriggerMonths(runtimeSettings.eventsArchiveFrequency, runtimeSettings.eventsArchiveMonth);
  const currentMonth = (referenceDate || new Date()).getMonth() + 1;
  return triggerMonths.indexOf(currentMonth) !== -1;
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
    subject: `Jubal: ${rowsToAlert.length} ${itemLabel} need${rowsToAlert.length === 1 ? 's' : ''} review`,
    body: [
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
    }

    const ts = new Date().toISOString();
    const meta = ss.getSheetByName(CONFIG.sheetNames.formMetadata);
    let formId = '';
    if (meta) formId = (meta.getRange('B2').getValue() || meta.getRange('B1').getValue() || '').toString();

    const errorsStr = Array.isArray(parseErrors) ? parseErrors.join('; ') : (parseErrors || '');
    const rowIndex = sheet.getLastRow() + 1;
    sheet.appendRow([ts, formId, submittedName || '', matchedCanonical || '', errorsStr, actionRequired || '', '']);

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
    const databaseSS = SpreadsheetApp.getActiveSpreadsheet();
    const databaseSheet = databaseSS.getSheetByName(CONFIG.sheetNames.ministryMembers);
    const databaseData = databaseSheet.getDataRange().getValues();

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
        const dbName = freshDatabaseData[i][0] ? freshDatabaseData[i][0].toString() : '';
        let dbCanonical = freshDatabaseData[i][5] ? freshDatabaseData[i][5].toString() : '';

        // If canonical missing, compute and persist it
        if (!dbCanonical && dbName) {
          try {
            dbCanonical = normalizeName(dbName);
            databaseSheet.getRange(i + 1, 6).setValue(dbCanonical);
          } catch (err) {
            console.error('Failed to persist canonical name for row ' + (i + 1) + ': ' + err.message);
          }
        }

        if (dbCanonical && incomingCanonical && dbCanonical === incomingCanonical) {
          // Update the corresponding row
          databaseSheet.getRange(i + 1, 3).setValue(timesWilling);
          databaseSheet.getRange(i + 1, 4).setValue(unavailableDatesString); // Use the joined string
          databaseSheet.getRange(i + 1, 5).setValue(comments);
          // Ensure canonical is stored
          databaseSheet.getRange(i + 1, 6).setValue(incomingCanonical);
          found = true;
          console.log('Updated existing row ' + (i + 1) + ' for ' + name + ' (canonical: ' + incomingCanonical + ')');
          break;
        }
      }

      if (!found) {
        // If no match is found, append a new row with canonical name in column 6
        const lastRow = databaseSheet.getLastRow() + 1;
        databaseSheet.getRange(lastRow, 1).setValue(name);
        databaseSheet.getRange(lastRow, 3).setValue(timesWilling);
        databaseSheet.getRange(lastRow, 4).setValue(unavailableDatesString);
        databaseSheet.getRange(lastRow, 5).setValue(comments);
        databaseSheet.getRange(lastRow, 6).setValue(incomingCanonical);
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

function createNewFormForMonth(month, year, monthName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const metadataSheet = ss.getSheetByName(CONFIG.sheetNames.formMetadata) || ss.insertSheet(CONFIG.sheetNames.formMetadata);
  const runtimeSettings = loadRuntimeSettings();

  // Create a new form for the upcoming month
  const formTitle = `${runtimeSettings.churchName} Availability - ${monthName}`;
  const form = FormApp.create(formTitle);

  // Name Dropdown (ListItem)
  const nameDropdown = form.addListItem().setTitle(CONFIG.formHeaders.name).setRequired(true);
  nameDropdown.setChoiceValues(["Loading..."]);

  const numDropdown = form.addListItem()
  .setTitle(CONFIG.formHeaders.times)
  .setChoiceValues(runtimeSettings.timesChoices) // Set the dropdown options
  .setRequired(true); // Make the question required
  
  // Add next month's form metadata
  const lastRow = metadataSheet.getLastRow();
  metadataSheet.getRange(lastRow + 1, 1).setValue(monthName + " Form");
  metadataSheet.getRange(lastRow + 1, 2).setValue(form.getId());

  // Update the dropdown with real names
  updateFormDropdown();

  // Add the service dates to the form for unavailable dates selection
  const serviceDates = getServiceDates(year, month);
  const dateChoices = serviceDates;

  const availMC = form.addCheckboxItem();
  availMC.setTitle(CONFIG.formHeaders.dates)
    .setChoices(dateChoices.map(date => availMC.createChoice(date)));

  // Optional comments section
  form.addTextItem().setTitle(CONFIG.formHeaders.comments).setRequired(false);

  // Link the form responses to a new sheet in the current spreadsheet
  form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());  // Link the form to the new response sheet
  console.log("Linked form responses to new sheet");

  // Get the links for the edit and responder URLs
  const editUrl = form.getEditUrl(); // Edit link for the form owner
  const responderUrl = form.getPublishedUrl(); // Responder link for the participants

  // Send email notification about the new form
  const emailSubject = "New Music Ministry Availability Form Created";
  const emailBody = "A new Music Ministry Availability Form has been created for the month of " + monthName + ".\n\n" +
                  "You can access and fill out the form using the following link:\n" + responderUrl + "\n\n" +
                  "If you need to edit the form, use the following link:\n" + editUrl + "\n\n" +
                  "Please submit your availability as soon as possible.";

  // Send email
  sendEmailToAdmins(emailSubject, emailBody, runtimeSettings);

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
  try {
    const availSheetName = getAvailabilitySheetName(year, month, runtimeSettings);
    syncFormWithSheet(form.getId(), availSheetName);
  } catch (err) {
    console.error('Failed to sync form with sheet after creation: ' + err.message);
  }
}

function updateFormDropdown() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const databaseSheet = ss.getSheetByName(CONFIG.sheetNames.ministryMembers);
  const metadataSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.sheetNames.formMetadata);

  if (!metadataSheet) {
    console.log("Form Metadata sheet missing.");
    return;
  }

  const formId = metadataSheet.getRange("B2").getValue();
  if (!formId) {
    console.log("No Form ID found.");
    return;
  }

  // Retrieve the list of names from the "Ministry Members" sheet
  let names = databaseSheet.getRange("A2:A" + databaseSheet.getLastRow()).getValues();
  names = names.flat().filter(String); // Flatten the array and remove any empty strings

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
    dropdownItem.setChoiceValues(names);
    console.log("Dropdown updated with names from the sheet.");
  } else {
    console.log("Dropdown question not found.");
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
    return;
  }

  const lastCol = sheet.getLastColumn();
  if (lastCol <= 1) {
    console.log('syncFormWithSheet: no date columns in sheet');
    return;
  }

  const dateHeaders = sheet.getRange(CONFIG.layout.dateRowIndex, 2, 1, lastCol - 1).getDisplayValues()[0].map(h => String(h).trim()).filter(Boolean);
  const normalizedHeaders = mergeDateChoices(dateHeaders);

  if (!normalizedHeaders.length) {
    console.log('syncFormWithSheet: no date headers found');
    return;
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
    return;
  }

  const choices = normalizedHeaders.map(d => target.createChoice(d));
  target.setChoices(choices);
  console.log('syncFormWithSheet: updated form choices from sheet ' + sheetName);
}

/**
 * Convenience wrapper: sync the current open form (from metadata) with the planned availability sheet for next month.
 */
function syncCurrentFormWithAvailability() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const meta = ss.getSheetByName(CONFIG.sheetNames.formMetadata);
  if (!meta) { console.log('No metadata sheet'); return; }
  const formId = meta.getRange('B2').getValue() || meta.getRange('B1').getValue();
  if (!formId) { console.log('No form id in metadata'); return; }

  const today = new Date();
  const planDate = new Date(today);
  planDate.setMonth(today.getMonth() + 1);
  const planMonthName = planDate.toLocaleString('default', { month: 'long' });
  const sheetName = getAvailabilitySheetNameForMonthName(planMonthName);
  syncFormWithSheet(formId.toString(), sheetName);
}

function getAdminReminderPropertyKey(referenceDate) {
  const date = referenceDate ? new Date(referenceDate) : new Date();
  const planDate = new Date(date);
  planDate.setMonth(planDate.getMonth() + 1);
  return `adminReminderSent:${planDate.getFullYear()}-${('0' + (planDate.getMonth() + 1)).slice(-2)}`;
}

function buildAdminPlanningReminder(referenceDate, settings) {
  const runtimeSettings = settings || loadRuntimeSettings();
  const today = referenceDate ? new Date(referenceDate) : new Date();
  const planDate = new Date(today);
  planDate.setMonth(planDate.getMonth() + 1);

  const planMonthName = Utilities.formatDate(planDate, runtimeSettings.timeZone, 'MMMM yyyy');
  const monthlySetupDay = formatDayOfMonthHuman(runtimeSettings.formCreationDay);
  const subject = `Jubal Reminder: Please review ${planMonthName}`;
  const bodyLines = [
    `Please review the schedule setup for ${planMonthName}.`,
    '',
    'Before the monthly form is created, please check these areas:',
    `- Recurring schedule: ${getSheetUrlByName(CONFIG.sheetNames.recurring)}`,
    `- One-time events and changes: ${getSheetUrlByName(CONFIG.sheetNames.events)}`,
    `- Ministry members and role updates: ${getSheetUrlByName(CONFIG.sheetNames.ministryMembers)}`,
    `- Admin contacts and notifications: ${getSheetUrlByName(CONFIG.sheetNames.admins)}`,
    '',
    'Recommended checklist:',
    '- Confirm your normal recurring events are correct.',
    '- Add any special events, cancellations, or moved dates for next month.',
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
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, column, numRows, 1).setDataValidation(rule);
  sheet.getRange(2, column, numRows, 1).setNumberFormat(formatPattern || 'yyyy-mm-dd');
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
  sheet.setFrozenRows(1);
  styleConfigHeader(sheet, ['#d9ead3', '#fce5cd', '#cfe2f3', '#cfe2f3', '#cfe2f3', '#fff2cc', '#fff2cc', '#d9ead3', '#d9ead3', '#ead1dc']);
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
  applyCheckboxColumn(sheet, 1, maxRows);
  applyDropdownColumn(sheet, 3, ['Weekly', 'Monthly', 'Yearly', 'Easter'], maxRows);
  applyDropdownColumn(sheet, 4, ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'], maxRows);
  applyDropdownColumn(sheet, 5, ['every', '1', '2', '3', '4', '5', 'last'], maxRows);
  applyDropdownColumn(sheet, 6, ['all', '1', '2', '3', '4', '5', '6', '7', '8', '9', '10', '11', '12', 'January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'], maxRows);
  sheet.getRange(2, 7, maxRows, 1).setNumberFormat('0');
  applyCheckboxColumn(sheet, 8, maxRows);
  applyCheckboxColumn(sheet, 9, maxRows);
  sheet.autoResizeColumns(1, Math.min(10, sheet.getLastColumn()));
}

function configureEventsSheetUi(sheet) {
  if (!sheet) return;
  const maxRows = Math.max(sheet.getMaxRows() - 1, 1);
  sheet.setFrozenRows(1);
  styleConfigHeader(sheet, ['#d9ead3', '#f4cccc', '#fce5cd', '#cfe2f3', '#cfe2f3', '#d9ead3', '#d9ead3', '#ead1dc']);
  setHeaderNotes(sheet, [
    'Check this to use the row.',
    'Use a real date cell. Recommended format: yyyy-mm-dd.',
    'Name shown on the form and schedule for this one-time event.',
    'ADD creates a one-time event. REMOVE cancels one date from the normal schedule.',
    'Optional. Use the same event name as the Recurring sheet when moving or cancelling a recurring event.',
    'Check to show this event on the availability form.',
    'Check to show this event on the schedule sheet.',
    'Optional reminder for admins.'
  ]);
  applyCheckboxColumn(sheet, 1, maxRows);
  applyDateColumn(sheet, 2, maxRows, 'yyyy-mm-dd');
  applyDropdownColumn(sheet, 4, ['ADD', 'REMOVE'], maxRows);
  const recurringEventValues = getRecurringEventDropdownValues();
  if (recurringEventValues.length) {
    applyDropdownColumn(sheet, 5, recurringEventValues, maxRows, true);
  }
  applyCheckboxColumn(sheet, 6, maxRows);
  applyCheckboxColumn(sheet, 7, maxRows);
  sheet.autoResizeColumns(1, Math.min(8, sheet.getLastColumn()));
}

function getAdminsSeedRows(emails) {
  const adminEmails = normalizeEmailList(emails && emails.length ? emails : CONFIG.ids.adminEmails);
  const rows = [
    ['Enabled', 'Name', 'Email', 'Notes']
  ];

  if (adminEmails.length) {
    adminEmails.forEach((email, index) => {
      rows.push([
        true,
        index === 0 ? 'Primary Admin' : '',
        email,
        index === 0
          ? 'Add more rows here when you want more admins to receive reminders and alerts.'
          : ''
      ]);
    });
  } else {
    rows.push([true, 'Primary Admin', '', 'Enter the first admin email here.']);
  }

  rows.push([false, '', '', 'Example extra admin row. Check Enabled when ready to use it.']);
  return rows;
}

function configureAdminsSheetUi(sheet) {
  if (!sheet) return;
  const maxRows = Math.max(sheet.getMaxRows() - 1, 1);
  sheet.setFrozenRows(1);
  styleConfigHeader(sheet, ['#d9ead3', '#fce5cd', '#cfe2f3', '#ead1dc']);
  setHeaderNotes(sheet, [
    'Check this when the admin should receive reminders and alerts.',
    'Optional display name for your team.',
    'Email address used for admin reminders and alerts.',
    'Optional note for the team.'
  ]);
  applyCheckboxColumn(sheet, 1, maxRows);
  applyEmailColumn(sheet, 3, maxRows);
  sheet.autoResizeColumns(1, Math.min(4, sheet.getLastColumn()));
}

function getRecurringSeedRows() {
  return [
    ['Enabled', 'Event', 'Frequency', 'Weekday', 'Week Of Month', 'Month', 'Day', 'Include In Form', 'Include In Schedule', 'Notes'],
    [true, '', 'Weekly', 'Sunday', 'every', 'all', '', true, true, 'Default weekly Sunday schedule. Leave Event blank to show plain dates.'],
    [false, 'Corporate Prayer', 'Monthly', 'Friday', 1, 'all', '', true, true, 'Enable if your church has a monthly prayer gathering'],
    [false, 'Easter', 'Easter', '', '', 'all', '', true, true, 'Enable to include Easter automatically each year'],
    [false, 'Christmas', 'Yearly', '', '', '12', 25, true, true, 'Enable to include Christmas automatically each year']
  ];
}

function getEventsSeedRows() {
  return [
    ['Enabled', 'Date', 'Event', 'Action', 'Recurring Event', 'Include In Form', 'Include In Schedule', 'Notes'],
    [false, new Date(2026, 3, 3), 'Good Friday', 'ADD', '', true, true, 'Example one-time event. Change the date, then check Enabled to use it.'],
    [false, new Date(2026, 3, 3), 'Corporate Prayer', 'REMOVE', 'Corporate Prayer', true, true, 'Example: remove one recurring event date for a specific month.']
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
    ['time_zone', safeGetScriptTimeZone(), 'IANA timezone for event generation'],
    ['forms_folder_id', CONFIG.ids.formsFolder, 'Drive folder where forms are moved after creation'],
    ['admin_emails', CONFIG.ids.adminEmails.join(','), 'Legacy fallback only. Prefer the Admins sheet for a friendlier admin list.'],
    ['roles', CONFIG.roles.join(','), 'Comma-separated ministry roles'],
    ['form_creation_day', CONFIG.defaults.formCreationDay, 'Monthly setup day shown to admins in reminder emails. Keep your Apps Script trigger aligned with this day.'],
    ['admin_reminder_enabled', CONFIG.defaults.adminReminderEnabled, 'TRUE or FALSE. When TRUE, send planning reminders to admins.'],
    ['admin_reminder_day', CONFIG.defaults.adminReminderDay, 'Day of month to send the admin planning reminder for next month.'],
    ['times_choices', CONFIG.defaults.timesChoices.join(','), 'Comma-separated willingness choices'],
    ['availability_sheet_suffix', CONFIG.defaults.availabilitySheetSuffix, 'Suffix used for monthly availability tabs'],
    ['events_archive_frequency', CONFIG.defaults.eventsArchiveFrequency, 'Off, Monthly, Quarterly, or Yearly'],
    ['events_archive_month', CONFIG.defaults.eventsArchiveMonth, 'Month that yearly archiving should run, such as January']
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

function configureSettingsSheetUi(sheet) {
  if (!sheet || sheet.getLastRow() < 2) return;

  sheet.setFrozenRows(1);
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, Math.min(2, sheet.getLastColumn())).getValues();
  rows.forEach((row, index) => {
    const key = String(row[0] || '').trim().toLowerCase();
    const valueRange = sheet.getRange(index + 2, 2);
    valueRange.clearDataValidations();

    if (key === 'events_archive_frequency') {
      valueRange.setDataValidation(
        SpreadsheetApp.newDataValidation()
          .requireValueInList(['Off', 'Monthly', 'Quarterly', 'Yearly'], true)
          .setAllowInvalid(false)
          .build()
      );
    } else if (key === 'form_creation_day') {
      valueRange.setDataValidation(
        SpreadsheetApp.newDataValidation()
          .requireNumberBetween(1, 28)
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
          .requireNumberBetween(1, 28)
          .setAllowInvalid(false)
          .build()
      );
    } else if (key === 'events_archive_month') {
      valueRange.setDataValidation(
        SpreadsheetApp.newDataValidation()
          .requireValueInList(['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'], true)
          .setAllowInvalid(false)
          .build()
      );
    }
  });
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

function getRoleCheckboxStartColumn() {
  return 7; // A-F are core member fields, role checkboxes begin in column G.
}

function ensureRoleCheckboxColumns(sheet, roles) {
  if (!sheet || !roles || !roles.length) return;

  const startColumn = getRoleCheckboxStartColumn();
  const requiredLastColumn = startColumn + roles.length - 1;
  if (sheet.getMaxColumns() < requiredLastColumn) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), requiredLastColumn - sheet.getMaxColumns());
  }

  sheet.getRange(1, startColumn, 1, roles.length).setValues([roles]);
  sheet.getRange(1, startColumn, 1, roles.length).setFontWeight('bold');
  if (sheet.getMaxRows() > 1) {
    sheet.getRange(2, startColumn, sheet.getMaxRows() - 1, roles.length).insertCheckboxes();
  }
  sheet.autoResizeColumns(startColumn, roles.length);
}

function syncRoleCheckboxesFromRolesColumn(sheet, roles) {
  if (!sheet || !roles || !roles.length || sheet.getLastRow() < 2) return;

  const startColumn = getRoleCheckboxStartColumn();
  const lastRow = sheet.getLastRow();
  const rolesText = sheet.getRange(2, 2, lastRow - 1, 1).getDisplayValues().flat();
  const checkboxRange = sheet.getRange(2, startColumn, lastRow - 1, roles.length);
  const checkboxValues = checkboxRange.getValues();

  for (let i = 0; i < checkboxValues.length; i++) {
    const existingChecks = checkboxValues[i];
    const hasAnyChecked = existingChecks.some(value => value === true);
    if (hasAnyChecked) continue;

    const parsedRoles = String(rolesText[i] || '')
      .split(',')
      .map(role => role.trim().toUpperCase())
      .filter(Boolean);
    if (!parsedRoles.length) continue;

    checkboxValues[i] = roles.map(role => parsedRoles.indexOf(String(role).trim().toUpperCase()) !== -1);
  }

  checkboxRange.setValues(checkboxValues);
}

function ensureRolesFormulaColumn(sheet, roles) {
  if (!sheet || !roles || !roles.length) return;

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const startColumn = getRoleCheckboxStartColumn();
  const endColumn = startColumn + roles.length - 1;
  const headerRange = `${columnToLetter(startColumn)}$1:${columnToLetter(endColumn)}$1`;

  const formulas = [];
  for (let row = 2; row <= lastRow; row++) {
    const rowRange = `${columnToLetter(startColumn)}${row}:${columnToLetter(endColumn)}${row}`;
    formulas.push([
      `=IF(COUNTIF(${rowRange},TRUE)=0,"",TEXTJOIN(", ",TRUE,FILTER(${headerRange},${rowRange}=TRUE)))`
    ]);
  }

  sheet.getRange(2, 2, formulas.length, 1).setFormulas(formulas);
}

function configureMinistryMembersRoles(sheet, roles) {
  if (!sheet || !roles || !roles.length) return;
  ensureRoleCheckboxColumns(sheet, roles);
  syncRoleCheckboxesFromRolesColumn(sheet, roles);
  ensureRolesFormulaColumn(sheet, roles);
}

function hasRoleCheckboxColumnsConfigured(sheet, roles) {
  if (!sheet || !roles || !roles.length) return false;

  const startColumn = getRoleCheckboxStartColumn();
  const requiredLastColumn = startColumn + roles.length - 1;
  if (sheet.getLastColumn() < requiredLastColumn) return false;

  const headers = sheet.getRange(1, startColumn, 1, roles.length).getDisplayValues()[0];
  return roles.every((role, index) => String(headers[index] || '').trim() === String(role || '').trim());
}

function ensureRolesFormulaForRow(sheet, rowNumber, roles) {
  if (!sheet || !roles || !roles.length || rowNumber < 2) return;
  if (!hasRoleCheckboxColumnsConfigured(sheet, roles)) return;

  const startColumn = getRoleCheckboxStartColumn();
  const endColumn = startColumn + roles.length - 1;
  const headerRange = `${columnToLetter(startColumn)}$1:${columnToLetter(endColumn)}$1`;
  const rowRange = `${columnToLetter(startColumn)}${rowNumber}:${columnToLetter(endColumn)}${rowNumber}`;
  const formula = `=IF(COUNTIF(${rowRange},TRUE)=0,"",TEXTJOIN(", ",TRUE,FILTER(${headerRange},${rowRange}=TRUE)))`;
  sheet.getRange(rowNumber, 2).setFormula(formula);
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
  const startColumn = getRoleCheckboxStartColumn();
  const requiredLastColumn = startColumn + roles.length - 1;
  summary.memberRows = lastRow - 1;

  const rolesText = sheet.getRange(2, 2, lastRow - 1, 1).getDisplayValues().flat();
  summary.rowsWithLegacyRoles = rolesText.filter(value => String(value || '').trim()).length;

  if (sheet.getLastColumn() >= startColumn) {
    const inspectWidth = Math.max(0, Math.min(requiredLastColumn, sheet.getLastColumn()) - startColumn + 1);
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

  const roles = loadRuntimeSettings().roles;
  if (!roles.length) throw new Error('No roles configured in Settings.');

  if (hasRoleCheckboxColumnsConfigured(sheet, roles)) {
    configureMinistryMembersRoles(sheet, roles);
    SpreadsheetApp.flush();
    console.log('Role checkbox migration skipped because the sheet is already configured.');
    return {
      status: 'already_configured',
      roles: roles,
      startColumn: columnToLetter(getRoleCheckboxStartColumn()),
      migratedRows: 0,
      memberRows: Math.max(sheet.getLastRow() - 1, 0)
    };
  }

  const summary = summarizeRoleMigration(sheet, roles);
  if (summary.conflictingCells > 0) {
    throw new Error(
      `Role checkbox migration stopped because columns ${columnToLetter(getRoleCheckboxStartColumn())}+ already contain data. ` +
      `Conflicts found: ${summary.conflictingCells}. Examples: ${summary.conflictSamples.join(', ')}`
    );
  }

  configureMinistryMembersRoles(sheet, roles);
  SpreadsheetApp.flush();

  console.log(
    `Role migration complete. Added ${roles.length} checkbox columns starting at ${columnToLetter(getRoleCheckboxStartColumn())}. ` +
    `Migrated ${summary.rowsWithLegacyRoles} member row(s) from the Roles column.`
  );

  return {
    status: 'migrated',
    roles: roles,
    startColumn: columnToLetter(getRoleCheckboxStartColumn()),
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

function getMemberRolesFromRow(row, configuredRoles) {
  const roles = (configuredRoles || []).map(role => String(role || '').trim()).filter(Boolean);
  if (!row || !roles.length) return [];

  const checkboxStartIndex = getRoleCheckboxStartColumn() - 1;
  const checkboxRoles = roles.filter((role, index) => row[checkboxStartIndex + index] === true);
  if (checkboxRoles.length) {
    return checkboxRoles.map(role => role.toUpperCase());
  }

  return row[1]
    ? row[1].toString().split(",").map(role => role.trim().toUpperCase()).filter(Boolean)
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
  sheet.getRange(insertionRow + 3, 1).setValue("Availability").setFontWeight("bold");
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

function monthlySetup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const metadataSheet = ss.getSheetByName("Form Metadata") || ss.insertSheet("Form Metadata");
  const runtimeSettings = loadRuntimeSettings();
  const trackedFormIds = getTrackedFormIds(metadataSheet);

  const today = new Date();
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
  setupAvailability(newTabName, planYear, planMonth);

  clearByHeader(CONFIG.sheetHeaders.times);
  clearByHeader(CONFIG.sheetHeaders.dates);
  clearByHeader(CONFIG.sheetHeaders.comments);

  if (!ss.getSheetByName(newTabName)) {
    ss.insertSheet(newTabName);
    console.log("Created new tab: " + newTabName);
  }

  const oldSheet = ss.getSheetByName(deleteTabName);
  if (oldSheet) {
    ss.deleteSheet(oldSheet);
    console.log("Deleted old tab: " + deleteTabName);
  }

  // Store the form ID in the "Form Metadata" sheet in the specified structure
  // Clear any old form metadata if we have more than 2 entries
  const lastRow = metadataSheet.getLastRow();
  if (lastRow > 1) {
    const currentMonthFormLabel = metadataSheet.getRange(2, 1).getValue();  // Get the current month's form label
    const currentMonthFormId = metadataSheet.getRange(2, 2).getValue();  // Get the current month's form ID

    // Move the current month's form metadata to row 1
    metadataSheet.getRange(1, 1).setValue(currentMonthFormLabel);  // Move label to row 1
    metadataSheet.getRange(1, 2).setValue(currentMonthFormId);  // Move form ID to row 1
    metadataSheet.deleteRow(2);
  }

  unlinkTrackedForms(trackedFormIds);
  deleteFormResponseSheets(ss);

  createNewFormForMonth(planMonth, planYear, planMonthName);
  console.log(`Created new form for ${planMonthName}`);

  // Legacy support: populate automatic events only for old-style Events/Monthly Events sheets.
  try {
    ensureMonthlyEventsFor(planYear, planMonth);
  } catch (err) {
    console.error('Failed to ensure legacy Events entries: ' + err.message);
  }
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

function updateAvailability() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const runtimeSettings = loadRuntimeSettings();

  console.log('--- STARTING updateAvailability ---');

  const today = new Date();
  const planDate = new Date(today);
  planDate.setMonth(today.getMonth() + 1);
  const planMonthName = planDate.toLocaleString('default', { month: 'long' });
  const sheetName = getAvailabilitySheetNameForMonthName(planMonthName, runtimeSettings);

  const matrixSheet = ss.getSheetByName(sheetName);
  const databaseSheet = ss.getSheetByName(CONFIG.sheetNames.ministryMembers);

  if (!matrixSheet || !databaseSheet) {
    console.log("Error: One or more required sheets are missing.");
    return;
  }

  const databaseData = databaseSheet.getDataRange().getValues();

  if (!databaseData.length) {
    console.log("No data found in the Ministry Members sheet.");
    return;
  }

  // Get Date Headers from the sheet (Row 1, starting from column 2)
  let lastCol = matrixSheet.getLastColumn();
  if (lastCol <= 1) {
    console.log("Error: Availability matrix has no date columns.");
    return;
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

  // Standardize roleOrder to uppercase for case-insensitive matching
  roleOrder = roleOrder.map(role => role.toUpperCase());

  // Process each row in the Ministry Members sheet
  for (let i = 1; i < databaseData.length; i++) {
    const row = databaseData[i];
    let name = row[0] ? row[0].trim() : "";
    if (!name) continue;

    // Ensure canonical name exists in column 6
    try {
      const existingCanonical = row[5] ? row[5].toString() : '';
      if (!existingCanonical && name) {
        const canon = normalizeName(name);
        databaseSheet.getRange(i + 1, 6).setValue(canon);
        // Update local copy so further logic can use it if needed
        row[5] = canon;
      }
    } catch (err) {
      console.error('Failed to persist canonical in updateAvailability for row ' + (i + 1) + ': ' + err.message);
    }
    
    const roles = getMemberRolesFromRow(row, runtimeSettings.roles);
    const timesWilling = row[2] ? row[2].toString().trim() : "";
    const rawUnavailableDates = row[3] ? row[3].toString() : "";
    
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
    const headers = [
      CONFIG.sheetHeaders.name, 
      CONFIG.sheetHeaders.roles, 
      CONFIG.sheetHeaders.times, 
      CONFIG.sheetHeaders.dates, 
      CONFIG.sheetHeaders.comments,
      CONFIG.sheetHeaders.canonicalName
    ];
    dbSheet.appendRow(headers);
    dbSheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");
    
    // Add Dummy Data
    const dummyRow = ["John Doe", "WL, ACOUSTIC", "4", "", "Excited to serve!", normalizeName("John Doe")];
    dbSheet.appendRow(dummyRow);
    console.log("Created Ministry Members sheet with dummy data.");
  }

  // Ensure the canonical-name column is labeled for existing spreadsheets too.
  if (dbSheet.getRange(1, 6).getValue() !== CONFIG.sheetHeaders.canonicalName) {
    dbSheet.getRange(1, 6).setValue(CONFIG.sheetHeaders.canonicalName);
  }
  dbSheet.getRange(1, 1, 1, 6).setFontWeight("bold");

  // 2. Create Form Metadata sheet if it doesn't exist
  let metaSheet = ss.getSheetByName(CONFIG.sheetNames.formMetadata);
  if (!metaSheet) {
    metaSheet = ss.insertSheet(CONFIG.sheetNames.formMetadata);
    metaSheet.appendRow(["Form Name", "Form ID"]);
    console.log("Created Form Metadata sheet.");
  }

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
    adminsSheet.getRange(1, 1, 1, 4).setValues([getAdminsSeedRows(loadRuntimeSettings().adminEmails)[0]]);
    console.log("Created Admins sheet.");
  }
  adminsSheet.getRange(1, 1, 1, 4).setFontWeight('bold');
  seedSheetRowsIfEmpty(adminsSheet, getAdminsSeedRows(loadRuntimeSettings().adminEmails));
  if (sheetUsesFriendlyAdminsLayout(adminsSheet)) {
    configureAdminsSheetUi(adminsSheet);
  }

  const roleMigrationResult = maybeMigrateMemberRolesDuringSetup();
  if (roleMigrationResult && roleMigrationResult.status === 'migrated') {
    console.log('Initialized role checkbox workflow in Ministry Members.');
  } else if (roleMigrationResult && roleMigrationResult.status === 'already_configured') {
    console.log('Role checkbox workflow already configured in Ministry Members.');
  }

  // 5. Migrate legacy sheet names to the friendlier go-forward names when possible.
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

  // 6. Create the new Recurring sheet if no recurring configuration exists.
  if (!recurringSheet && !legacyRecurringSheet) {
    recurringSheet = ss.insertSheet(CONFIG.sheetNames.recurring);
    recurringSheet.getRange(1, 1, 1, 10).setValues([getRecurringSeedRows()[0]]);
    seedSheetRowsIfEmpty(recurringSheet, getRecurringSeedRows());
    configureRecurringSheetUi(recurringSheet);
    console.log("Created Recurring sheet.");
  } else if (recurringSheet && sheetUsesFriendlyRecurringLayout(recurringSheet)) {
    seedSheetRowsIfEmpty(recurringSheet, getRecurringSeedRows());
    configureRecurringSheetUi(recurringSheet);
  }

  // 7. Create the new Events sheet for month-specific additions/removals.
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
  
  console.log("Initialization complete. You can now run monthlySetup().");
}
