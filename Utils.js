/**
 * Shared parsing, normalization, and runtime settings helpers.
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

  const parsed = results.filter((v, i, a) => a.indexOf(v) === i && v);
  return { parsed, errors };
}

function parseSingleDate(s) {
  if (!s) return null;
  s = s.toString().trim();

  const m = s.match(/^(\d{1,2})\/(\d{1,2})(?:\/(\d{2,4}))?$/);
  if (m) {
    const month = parseInt(m[1], 10) - 1;
    const day = parseInt(m[2], 10);
    const year = m[3] ? parseInt(m[3], 10) : new Date().getFullYear();
    const dt = createStrictDate(year, month, day);
    return isNaN(dt.getTime()) ? null : dt;
  }

  const iso = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (iso) {
    const dt = createStrictDate(parseInt(iso[1], 10), parseInt(iso[2], 10) - 1, parseInt(iso[3], 10));
    return isNaN(dt.getTime()) ? null : dt;
  }

  const withYear = s + ' ' + new Date().getFullYear();
  const parsed = new Date(withYear);
  if (!isNaN(parsed.getTime())) return parsed;

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
