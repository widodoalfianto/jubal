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
  const result = {
    deleted: [],
    skipped: []
  };
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
      result.deleted.push(toDelete);
    } catch (e) {
      const reason = e && e.message ? e.message : 'Unknown error';
      const likelyCause = reason.indexOf('linked form') !== -1
        ? 'This response sheet is probably still linked to an older Google Form that is no longer tracked in Form Metadata.'
        : '';
      console.log("Skipped deleting Form Responses tab '" + toDelete + "': " + reason);
      result.skipped.push({
        sheetName: toDelete,
        reason: reason,
        likelyCause: likelyCause
      });
    }
  });

  return result;
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
  let formResponseCleanup = { deleted: [], skipped: [] };

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
    formResponseCleanup = deleteFormResponseSheetsById(ss, existingResponseSheetIds);

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
    planMonth: planMonth + 1,
    formResponseCleanup: formResponseCleanup
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
