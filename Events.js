/**
 * Scheduling engine and event generation helpers.
 */

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
  const month = Math.floor((h + l - 7 * m + 114) / 31) - 1;
  const day = ((h + l - 7 * m + 114) % 31) + 1;
  return new Date(year, month, day);
}

function ensureMonthlyEventsFor(year, month) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  year = parseInt(year, 10);
  if (isNaN(year)) year = new Date().getFullYear();
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

  const toEnsure = [];
  const eas = computeEaster(year);
  if (eas.getMonth() === month) {
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

      let rmNum = null;
      if (typeof rm === 'number' && !isNaN(rm)) rmNum = rm; else if (typeof rm === 'string') {
        const asNum = parseInt(rm, 10);
        if (!isNaN(asNum)) rmNum = asNum; else {
          const mn = new Date(Date.parse(rm + ' 1, 2000'));
          if (!isNaN(mn.getTime())) rmNum = mn.getMonth() + 1;
        }
      }

      if (ry === year && (rmNum === month || rmNum === (month + 1))) {
        if (extractDateKey(rd) === extractDateKey(ev.date)) {
          found = true;
          break;
        }
      }
    }

    if (!found) {
      me.appendRow([year, month + 1, ev.date, ev.label, 'auto']);
      logDebug('info', 'ensureMonthlyEventsFor added event', { year, month, event: ev });
    }
  }
}

function populateAnnualEvents(year) {
  year = parseInt(year, 10);
  if (isNaN(year)) year = new Date().getFullYear();
  for (let m = 0; m < 12; m++) ensureMonthlyEventsFor(year, m);
}
