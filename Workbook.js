/**
 * Workbook formatting, setup, and admin-sheet UI helpers.
 */

function getDayOfMonthDropdownValues() {
  const values = [];
  for (let i = 1; i <= 28; i++) values.push(String(i));
  return values;
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

  const startColumn = 10;
  const width = 5;
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
    ['Enabled', 'Email', 'Notes'],
    [false, 'admin@example.com', 'Example row']
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
