/**
 * Spreadsheet menu actions and dialog helpers.
 */

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
    .addItem('Initialize Project', 'menuInitializeProject')
    .addSeparator()
    .addItem('Add Special Event', 'showAddEventDialog')
    .addSeparator()
    .addItem('Apply Event Changes to Next Month', 'menuApplyEventChangesToPlanningMonth')
    .addItem('Refresh Form Dates', 'menuSyncCurrentFormWithAvailability')
    .addItem('Refresh Availability Sheet', 'menuUpdateAvailability')
    .addSeparator()
    .addItem('Sync Automation Triggers', 'menuSyncAutomationTriggers')
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
        'Next month has not been generated yet.',
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
          lines.push('The workbook was updated, but the form could not be refreshed automatically.');
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

function menuInitializeProject() {
  const ui = SpreadsheetApp.getUi();
  try {
    initializeProject();
    ui.alert(
      'Project Initialized',
      'The setup sheets, formatting, and menu configuration have been refreshed.',
      ui.ButtonSet.OK
    );
  } catch (error) {
    ui.alert('Could Not Initialize Project', error.message, ui.ButtonSet.OK);
    throw error;
  }
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

function menuSyncAutomationTriggers() {
  const ui = SpreadsheetApp.getUi();
  try {
    const result = syncAutomationTriggers();
    const createdCount = result && result.created ? result.created.length : 0;
    const removedCount = result && result.removed ? result.removed.length : 0;
    const createdHandlers = (result && result.created ? result.created : [])
      .map(item => item.handler)
      .join(', ');

    const lines = [
      'Automation triggers were refreshed for this spreadsheet.',
      `${createdCount} trigger(s) created.`,
      `${removedCount} older managed trigger(s) removed.`
    ];

    if (createdHandlers) {
      lines.push(`Current managed triggers: ${createdHandlers}.`);
    }

    ui.alert('Automation Triggers Synced', lines.join('\n\n'), ui.ButtonSet.OK);
    return result;
  } catch (error) {
    ui.alert('Could Not Sync Automation Triggers', error.message, ui.ButtonSet.OK);
    throw error;
  }
}
