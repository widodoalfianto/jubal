/**
 * End-to-End Testing Suite
 * 
 * This file contains tests to verify the full workflow of Jubal.
 * It is intended to be excluded from production deployments via .gitignore.
 * 
 * Usage: Run runFullSystemTest() from the Apps Script editor.
 */

function runFullSystemTest() {
  console.log("🧪 STARTING FULL SYSTEM TEST");
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // --- Step 1: Initial Setup ---
  console.log("\n--- Step 1: Initial Setup ---");
  // Ensure the project is initialized (creates sheets if missing)
  initializeProject();
  
  // Verify sheets exist
  if (!ss.getSheetByName(CONFIG.sheetNames.ministryMembers)) throw new Error("Ministry Members sheet missing");
  if (!ss.getSheetByName(CONFIG.sheetNames.formMetadata)) throw new Error("Form Metadata sheet missing");
  console.log("✅ Initial setup verified");

  // --- Step 2: Monthly Setup ---
  console.log("\n--- Step 2: Monthly Setup ---");
  // Run the monthly setup to create the form and availability sheet
  monthlySetup();
  
  // Calculate expected sheet name for next month
  const today = new Date();
  const planDate = new Date(today);
  planDate.setMonth(today.getMonth() + 1);
  const planMonthName = planDate.toLocaleString('default', { month: 'long' });
  const availabilitySheetName = getAvailabilitySheetName(planDate.getFullYear(), planDate.getMonth());
  
  const availSheet = ss.getSheetByName(availabilitySheetName);
  if (!availSheet) throw new Error(`Availability sheet '${availabilitySheetName}' not created`);
  console.log("✅ Monthly setup verified");

  // --- Step 3: Simulating Form Submission ---
  console.log("\n--- Step 3: Simulating Form Submission ---");
  
  // Determine a valid date from the created sheet to verify against
  // Headers are in the row defined by CONFIG.layout.dateRowIndex
  // Note: getValues returns 2D array. Row index is 0-based relative to range, but we want absolute row.
  // Actually, let's just grab the header row from the sheet directly.
  const headerRowValues = availSheet.getRange(CONFIG.layout.dateRowIndex, 1, 1, availSheet.getLastColumn()).getDisplayValues()[0];
  
  // Find the first date column (skipping "Schedule" or other labels)
  // We look for a date-like pattern or just pick the second column (index 1) if available
  let targetDate = "";
  let targetColIndex = -1;
  
  for (let i = 1; i < headerRowValues.length; i++) {
    const val = headerRowValues[i];
    const extracted = extractDateKey(val);
    if (extracted) {
      targetDate = extracted;
      targetColIndex = i;
      break;
    }
  }
  
  if (!targetDate) throw new Error("Could not find a valid date column in the availability sheet");
  console.log(`Targeting date: ${targetDate} at column index ${targetColIndex}`);

  const testName = "Test User " + new Date().getTime();
  const testRole = loadRuntimeSettings().roles[0]; // e.g., "WL"
  
  // Mock Event Object mimicking a Google Form submission
  const e = {
    namedValues: {
      [CONFIG.formHeaders.name]: [testName],
      [CONFIG.formHeaders.times]: ["4"], // Willing to serve
      [CONFIG.formHeaders.dates]: [""], // Available all dates
      [CONFIG.formHeaders.comments]: ["Automated Test Comment"]
    }
  };
  
  // Run the database update
  updateDatabase(e);
  
  // MANUAL INTERVENTION FOR TEST:
  // The system requires a Role to be assigned in the "Ministry Members" sheet for the user to appear in the matrix.
  // Since the form submission doesn't include Role (it's an admin field), we must inject it for this new test user.
  const dbSheet = ss.getSheetByName(CONFIG.sheetNames.ministryMembers);
  const dbData = dbSheet.getDataRange().getValues();
  let userRowIndex = -1;
  
  for (let i = 1; i < dbData.length; i++) {
    if (dbData[i][0] == testName) {
      userRowIndex = i + 1; // Convert 0-based array index to 1-based sheet row
      break;
    }
  }
  
  if (userRowIndex === -1) throw new Error("Test user was not added to the database");
  
  // Set the role
  dbSheet.getRange(userRowIndex, 2).setValue(testRole);
  SpreadsheetApp.flush(); // Ensure the role is saved before the script reads it again
  console.log(`✅ Form submission processed. Assigned role '${testRole}' to '${testName}'.`);

  // --- Step 4: Verify Availability ---
  console.log("\n--- Step 4: Verifying Availability Matrix ---");
  
  // Trigger the matrix update
  updateAvailability();
  
  // Fetch updated data
  const updatedAvailData = availSheet.getDataRange().getValues();
  
  // Find the row for the test role, ensuring we look in the Availability section (bottom)
  // CONFIG.layout.headerRowIndex is 1-based (e.g., 13), so we subtract 1 for 0-based array index
  let roleRowIndex = -1;
  const startSearchIndex = CONFIG.layout.headerRowIndex - 1;
  for (let i = startSearchIndex; i < updatedAvailData.length; i++) {
    if (updatedAvailData[i][0] == testRole) {
      roleRowIndex = i;
      break;
    }
  }
  
  if (roleRowIndex === -1) throw new Error(`Row for role '${testRole}' not found in availability sheet`);
  
  // Check the specific cell
  const cellValue = updatedAvailData[roleRowIndex][targetColIndex];
  
  const expectedDisplayName = formatAvailabilityDisplayName(testName);

  if (String(cellValue).includes(expectedDisplayName)) {
    console.log(`✅ SUCCESS: Found '${expectedDisplayName}' in the '${testRole}' row for ${targetDate}.`);
  } else {
    throw new Error(`❌ FAILURE: Expected '${expectedDisplayName}' in cell [${roleRowIndex+1}, ${targetColIndex+1}] but found: '${cellValue}'`);
  }
  
  console.log("\n🎉 FULL SYSTEM TEST PASSED");
  
  // --- Step 5: Cleanup ---
  cleanupTestArtifacts(ss, testName, availabilitySheetName, planMonthName);
}

/**
 * Cleans up test data to leave the spreadsheet in a clean state.
 */
function cleanupTestArtifacts(ss, testName, createdSheetName, monthName) {
  console.log("\n--- Step 5: Cleanup ---");
  
  // 1. Remove the test user from the database
  const dbSheet = ss.getSheetByName(CONFIG.sheetNames.ministryMembers);
  const dbData = dbSheet.getDataRange().getValues();
  
  for (let i = dbData.length - 1; i >= 0; i--) {
    if (dbData[i][0] === testName) {
      dbSheet.deleteRow(i + 1);
      console.log(`✅ Deleted test user '${testName}' from database.`);
      break;
    }
  }
  
  // 2. Delete the created availability sheet
  // Note: In a real scenario, you might want to keep it to inspect, 
  // but for a pure automated test, we clean it up.
  const createdSheet = ss.getSheetByName(createdSheetName);
  if (createdSheet) {
    ss.deleteSheet(createdSheet);
    console.log(`✅ Deleted test sheet '${createdSheetName}'.`);
  }

  // 3. Delete the Test Form and Metadata Entry
  const metaSheet = ss.getSheetByName(CONFIG.sheetNames.formMetadata);
  if (metaSheet) {
    const lastRow = metaSheet.getLastRow();
    // Check if the last entry matches our test month
    if (lastRow > 1) {
      const formName = metaSheet.getRange(lastRow, 1).getValue();
      const formId = metaSheet.getRange(lastRow, 2).getValue();
      
      if (formName.includes(monthName)) {
        try {
          try {
            FormApp.openById(formId).removeDestination();
          } catch (unlinkError) {
            console.log(`⚠️ Could not remove destination from test form '${formName}': ${unlinkError.message}`);
          }
          const file = DriveApp.getFileById(formId);
          file.setTrashed(true);
          console.log(`✅ Deleted test form '${formName}' from Drive.`);
          
          metaSheet.deleteRow(lastRow);
          console.log(`✅ Removed metadata entry for '${formName}'.`);
        } catch (e) {
          console.log(`⚠️ Could not delete form file: ${e.message}`);
        }
      }
    }
  }

  // 4. Delete the empty Form Responses sheet created by the test
  const sheets = ss.getSheets();
  for (const sheet of sheets) {
    if (sheet.getName().startsWith("Form Responses") && sheet.getLastRow() <= 1) {
      try {
        ss.deleteSheet(sheet);
        console.log(`✅ Deleted empty test sheet '${sheet.getName()}'.`);
      } catch (e) {
        console.log(`⚠️ Could not delete linked test sheet '${sheet.getName()}': ${e.message}`);
      }
    }
  }
}

/** Unit test harness and tests for parsing/normalization utilities */
function runUnitTests() {
  initializeProject();
  resetTestResultsSheet();
  deleteSheetIfExists(CONFIG.sheetNames.recurringEvents);
  replaceSheetContents(CONFIG.sheetNames.settings, getDefaultSettingsSheetRows());
  replaceSheetContents(CONFIG.sheetNames.events, getDefaultEventsSheetRows());
  replaceSheetContents(CONFIG.sheetNames.monthlyEvents, getDefaultMonthlyEventsSheetRows());

  // parseUnavailableDates tests
  try {
    const r1 = parseUnavailableDates('3/29');
    assertEqual(r1.parsed[0], '03/29', 'parse 3/29 -> 03/29');
    recordResult('parseUnavailableDates: simple MM/DD', true, 'Parsed: ' + JSON.stringify(r1));
  } catch (e) { recordResult('parseUnavailableDates: simple MM/DD', false, e.message); }

  try {
    const r2 = parseUnavailableDates('March 5');
    assertEqual(r2.parsed[0], formatDateMMDD(new Date(new Date().getFullYear(), 2, 5)), 'parse March 5');
    recordResult('parseUnavailableDates: "March 5"', true, 'Parsed: ' + JSON.stringify(r2));
  } catch (e) { recordResult('parseUnavailableDates: "March 5"', false, e.message); }

  try {
    const r3 = parseUnavailableDates('2026-03-29');
    assertEqual(r3.parsed[0], '03/29', 'parse ISO date');
    recordResult('parseUnavailableDates: ISO', true, 'Parsed: ' + JSON.stringify(r3));
  } catch (e) { recordResult('parseUnavailableDates: ISO', false, e.message); }

  try {
    const r4 = parseUnavailableDates('3/29 - 4/5');
    assertEqual(r4.parsed[0], '03/29', 'range start parsed');
    assertEqual(r4.parsed[1], '04/05', 'range end parsed');
    recordResult('parseUnavailableDates: range', true, 'Parsed: ' + JSON.stringify(r4));
  } catch (e) { recordResult('parseUnavailableDates: range', false, e.message); }

  try {
    const r4b = parseUnavailableDates('04/03 - Good Friday, 4/5/2026 - Easter');
    assertEqual(r4b.parsed[0], '04/03', 'labeled date parsed');
    assertEqual(r4b.parsed[1], '04/05', 'labeled ISO-like date parsed');
    recordResult('parseUnavailableDates: labeled choices', true, 'Parsed: ' + JSON.stringify(r4b));
  } catch (e) { recordResult('parseUnavailableDates: labeled choices', false, e.message); }

  try {
    const r5 = parseUnavailableDates('foobar');
    // Expect a fallback and at least one parse error
    if (!r5.errors || r5.errors.length === 0) throw new Error('Expected parseErrors for invalid input');
    recordResult('parseUnavailableDates: invalid input', true, 'Parsed: ' + JSON.stringify(r5));
  } catch (e) { recordResult('parseUnavailableDates: invalid input', false, e.message); }

  try {
    assertEqual(extractDateKey('04/03 - Good Friday'), '04/03', 'extractDateKey for labeled MM/DD');
    assertEqual(extractDateKey('4/5/2026 - Easter'), '04/05', 'extractDateKey for labeled M/D/YYYY');
    assertEqual(extractDateKey('04/12'), '04/12', 'extractDateKey for plain MM/DD');
    assertEqual(extractDateKey('Service on April 26, 2026'), '04/26', 'extractDateKey from embedded month name');
    recordResult('extractDateKey: header normalization', true, 'Mixed header formats normalize correctly');
  } catch (e) { recordResult('extractDateKey: header normalization', false, e.message); }

  try {
    assertEqual(normalizeDateChoice('4/5/2026 - Easter'), '04/05 - Easter', 'normalizeDateChoice with label');
    assertEqual(normalizeDateChoice('04/12'), '04/12', 'normalizeDateChoice without label');
    assertEqual(extractDateLabel('04/03 - Good Friday'), 'Good Friday', 'extractDateLabel');
    recordResult('normalizeDateChoice: standard display', true, 'Dates display as MM/DD with optional labels');
  } catch (e) { recordResult('normalizeDateChoice: standard display', false, e.message); }

  try {
    const sorted = sortDateChoices(['04/12', '4/5/2026 - Easter', '04/03 - Good Friday']);
    assertEqual(sorted[0], '04/03 - Good Friday', 'sorted first date');
    assertEqual(sorted[1], '04/05 - Easter', 'sorted second date');
    assertEqual(sorted[2], '04/12', 'sorted third date');
    recordResult('sortDateChoices: chronological order', true, 'Date choices sort chronologically');
  } catch (e) { recordResult('sortDateChoices: chronological order', false, e.message); }

  try {
    const merged = mergeDateChoices(['04/05', '4/5/2026 - Easter', '04/03 - Good Friday']);
    assertEqual(merged[0], '04/03 - Good Friday', 'merged first date');
    assertEqual(merged[1], '04/05 - Easter', 'special event label should replace plain same-day display');
    recordResult('mergeDateChoices: labeled same-day merge', true, 'Same-day choices merge to labeled display');
  } catch (e) { recordResult('mergeDateChoices: labeled same-day merge', false, e.message); }

  // normalizeName tests
  try {
    const n1 = normalizeName(' John   Doe ');
    assertEqual(n1, 'john doe', 'normalize spaces and lowercase');
    recordResult('normalizeName: spacing', true, 'Result: ' + n1);
  } catch (e) { recordResult('normalizeName: spacing', false, e.message); }

  try {
    const n2 = normalizeName('José Álvarez');
    assertEqual(n2, 'jose alvarez', 'remove diacritics');
    recordResult('normalizeName: diacritics', true, 'Result: ' + n2);
  } catch (e) { recordResult('normalizeName: diacritics', false, e.message); }

  // getServiceDates tests
  try {
    const sd = getServiceDates(2026, 2); // March 2026
    const expected = ['03/01', '03/08', '03/15', '03/22', '03/29'];
    if (JSON.stringify(sd) !== JSON.stringify(expected)) throw new Error('Expected Sunday-only default schedule, got ' + JSON.stringify(sd));
    for (let i = 0; i < sd.length; i++) {
      if (!sd[i].match(/^\d{2}\/\d{2}/)) throw new Error('Service date not in MM/dd format: ' + sd[i]);
    }
    recordResult('getServiceDates: March 2026 default Sundays', true, 'Service dates: ' + JSON.stringify(sd));
  } catch (e) { recordResult('getServiceDates: March 2026 default Sundays', false, e.message); }

  try {
    const eventRows = getDefaultEventsSheetRows();
    eventRows[2][0] = true; // Enable corporate prayer
    replaceSheetContents(CONFIG.sheetNames.events, eventRows);
    const sd = getServiceDates(2026, 2);
    if (!sd.some(date => date.indexOf('03/06 - Corporate Prayer') !== -1)) throw new Error('Corporate Prayer should appear when monthly recurring event is enabled');
    recordResult('getServiceDates:monthlyRecurringEvent', true, 'Monthly recurring event enabled successfully');
  } catch (e) { recordResult('getServiceDates:monthlyRecurringEvent', false, e.message); }

  try {
    const eventRows = getDefaultEventsSheetRows();
    eventRows[2][0] = true; // Enable corporate prayer for override test
    replaceSheetContents(CONFIG.sheetNames.events, eventRows);
    replaceSheetContents(CONFIG.sheetNames.monthlyEvents, [
      ['Enabled', 'Year', 'Month', 'Date', 'Action', 'Label', 'Rule ID', 'Include In Form', 'Include In Schedule', 'Sort Order', 'Type', 'Notes'],
      [true, 2026, 3, '2026-03-06', 'REMOVE', 'Corporate Prayer', 'corporate_prayer', true, true, 10, 'override', 'Moved off first Friday'],
      [true, 2026, 3, '2026-03-13', 'ADD', 'Corporate Prayer', 'corporate_prayer', true, true, 10, 'override', 'Moved to second Friday']
    ]);
    const sd = getServiceDates(2026, 2);
    if (sd.some(date => date.indexOf('03/06 - Corporate Prayer') !== -1)) throw new Error('Corporate Prayer was not removed from first Friday');
    if (!sd.some(date => date.indexOf('03/13 - Corporate Prayer') !== -1)) throw new Error('Corporate Prayer was not added to second Friday');
    recordResult('getServiceDates:actionOverrides', true, 'ADD/REMOVE override flow works');
  } catch (e) { recordResult('getServiceDates:actionOverrides', false, e.message); }

  try {
    const eventRows = getDefaultEventsSheetRows();
    eventRows[3][0] = true; // Enable Easter
    eventRows[4][0] = true; // Enable Christmas
    replaceSheetContents(CONFIG.sheetNames.events, eventRows);
    replaceSheetContents(CONFIG.sheetNames.monthlyEvents, getDefaultMonthlyEventsSheetRows());
    const aprilDates = getServiceDates(2026, 3); // April 2026
    const decemberDates = getServiceDates(2026, 11); // December 2026
    if (!aprilDates.includes('04/05 - Easter')) throw new Error('Easter should appear when yearly event is enabled');
    if (!decemberDates.includes('12/25 - Christmas')) throw new Error('Christmas should appear when yearly event is enabled');
    recordResult('getServiceDates:yearlyRecurringEvents', true, 'Yearly recurring events enabled successfully');
  } catch (e) { recordResult('getServiceDates:yearlyRecurringEvents', false, e.message); }

  try {
    replaceSheetContents(CONFIG.sheetNames.monthlyEvents, [
      ['Year', 'Month', 'Date', 'Label', 'Type'],
      [2026, 5, '2026-05-15', 'LegacyEvent', 'legacy']
    ]);
    const sd = getServiceDates(2026, 4); // May 2026
    if (!sd.some(date => date.indexOf('05/15 - LegacyEvent') !== -1)) throw new Error('Legacy Monthly Events did not remain authoritative');
    recordResult('getServiceDates:legacyMonthlyEvents', true, 'Legacy Monthly Events compatibility preserved');
  } catch (e) { recordResult('getServiceDates:legacyMonthlyEvents', false, e.message); }

  try {
    const eventRows = getDefaultEventsSheetRows();
    eventRows[2][0] = true; // Enable corporate prayer
    eventRows[3][0] = true; // Enable Easter
    replaceSheetContents(CONFIG.sheetNames.events, eventRows);
    replaceSheetContents(CONFIG.sheetNames.monthlyEvents, getDefaultMonthlyEventsSheetRows());
    const sd = getServiceDates(2026, 3); // April 2026
    if (!sd.includes('04/03 - Corporate Prayer')) throw new Error('Corporate Prayer should be labeled');
    if (!sd.includes('04/05 - Easter')) throw new Error('Easter should be labeled');
    if (sd.includes('4/5/2026 - Easter')) throw new Error('Displayed dates must be standardized to MM/DD');
    if (sd.includes('04/05')) throw new Error('Plain same-day Easter date should not also appear');
    recordResult('getServiceDates:standardizedDisplay', true, 'Special events retain labels with standardized dates');
  } catch (e) { recordResult('getServiceDates:standardizedDisplay', false, e.message); }

  replaceSheetContents(CONFIG.sheetNames.monthlyEvents, getDefaultMonthlyEventsSheetRows());
  logDebug('info', 'Unit tests completed');
}

function assertEqual(actual, expected, label) {
  if (actual !== expected) throw new Error(`${label} — expected '${expected}', got '${actual}'`);
}

function recordResult(testName, passed, message) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let tr = ss.getSheetByName('Test Results');
  if (!tr) {
    tr = ss.insertSheet('Test Results');
    tr.appendRow(['timestamp', 'testName', 'status', 'message']);
  }
  tr.appendRow([new Date().toISOString(), testName, passed ? 'PASS' : 'FAIL', message || '']);
}

/** Run all unit and integration tests. WARNING: run on a spreadsheet copy. */
function runAllTests() {
  logDebug('info', 'Starting full test run');
  runUnitTests();
  try {
    runIntegrationTests();
  } catch (e) {
    recordResult('runIntegrationTests', false, e.message);
  }
  logDebug('info', 'All tests finished');
  return summarizeTestResults();
}

function runIntegrationTests() {
  // WARNING: integration tests modify sheets and may create forms/files. Run on a copy.
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // test: computeEaster correctness
  try {
    const eas = computeEaster(2026);
    assertEqual(formatDateMMDD(eas), '04/05', 'computeEaster for 2026');
    recordResult('computeEaster:2026', true, formatDateMMDD(eas));
  } catch (e) { recordResult('computeEaster:2026', false, e.message); }

  // test: initializeProject seeds the new configuration sheets
  try {
    initializeProject();
    if (!ss.getSheetByName(CONFIG.sheetNames.settings)) throw new Error('Settings sheet missing');
    if (!ss.getSheetByName(CONFIG.sheetNames.events) && !ss.getSheetByName(CONFIG.sheetNames.recurringEvents)) {
      throw new Error('Events sheet missing');
    }
    if (!ss.getSheetByName(CONFIG.sheetNames.monthlyEvents)) throw new Error('Monthly Events sheet missing');
    recordResult('initializeProject:configSheets', true, 'Configuration sheets created');
  } catch (e) { recordResult('initializeProject:configSheets', false, e.message); }

  // test: ensureMonthlyEventsFor remains usable for legacy installs
  try {
    replaceSheetContents(CONFIG.sheetNames.monthlyEvents, [['Year', 'Month', 'Date', 'Label', 'Type']]);
    ensureMonthlyEventsFor(2026, 3); // April 2026
    const me = ss.getSheetByName(CONFIG.sheetNames.monthlyEvents);
    const rows = me.getDataRange().getValues();
    const found = rows.some(r => String(r.join('|')).indexOf('Easter') !== -1 && String(r.join('|')).indexOf('2026') !== -1);
    if (!found) throw new Error('Legacy Monthly Events did not receive Easter');
    recordResult('ensureMonthlyEventsFor:legacyMode', true, 'Legacy Monthly Events auto-population still works');
  } catch (e) { recordResult('ensureMonthlyEventsFor:legacyMode', false, e.message); }

  // test: monthlySetup creates availability sheet and form metadata (smoke)
  try {
    replaceSheetContents(CONFIG.sheetNames.settings, getDefaultSettingsSheetRows());
    replaceSheetContents(CONFIG.sheetNames.events, getDefaultEventsSheetRows());
    replaceSheetContents(CONFIG.sheetNames.monthlyEvents, getDefaultMonthlyEventsSheetRows());
    monthlySetup();
    recordResult('monthlySetup:smoke', true, 'monthlySetup executed');
  } catch (e) { recordResult('monthlySetup:smoke', false, e.message); }
}

function resetTestResultsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let tr = ss.getSheetByName('Test Results');
  if (!tr) {
    tr = ss.insertSheet('Test Results');
  }
  tr.clearContents();
  tr.appendRow(['timestamp', 'testName', 'status', 'message']);
}

function summarizeTestResults() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tr = ss.getSheetByName('Test Results');
  if (!tr || tr.getLastRow() < 2) {
    return { pass: 0, fail: 0, total: 0 };
  }

  const rows = tr.getRange(2, 1, tr.getLastRow() - 1, 4).getValues();
  const summary = rows.reduce((acc, row) => {
    const status = String(row[2] || '').toUpperCase();
    if (status === 'PASS') acc.pass++;
    if (status === 'FAIL') acc.fail++;
    return acc;
  }, { pass: 0, fail: 0, total: rows.length });

  return summary;
}

function replaceSheetContents(sheetName, rows) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  } else {
    sheet.clearContents();
  }

  if (!rows || !rows.length) return sheet;
  sheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
  return sheet;
}

function deleteSheetIfExists(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (sheet) ss.deleteSheet(sheet);
}

function getDefaultSettingsSheetRows() {
  return [
    ['Key', 'Value', 'Notes'],
    ['church_name', CONFIG.defaults.churchName, 'Used in form titles and notifications'],
    ['time_zone', safeGetScriptTimeZone(), 'IANA timezone for event generation'],
    ['forms_folder_id', CONFIG.ids.formsFolder, 'Drive folder where forms are moved after creation'],
    ['admin_emails', CONFIG.ids.adminEmails.join(','), 'Comma-separated admin recipients'],
    ['roles', CONFIG.roles.join(','), 'Comma-separated ministry roles'],
    ['form_creation_day', CONFIG.defaults.formCreationDay, 'Reserved for future time-driven setup'],
    ['times_choices', CONFIG.defaults.timesChoices.join(','), 'Comma-separated willingness choices'],
    ['availability_sheet_suffix', CONFIG.defaults.availabilitySheetSuffix, 'Suffix used for monthly availability tabs']
  ];
}

function getDefaultEventsSheetRows() {
  return [
    ['Enabled', 'Rule ID', 'Label', 'Recurrence', 'Rule Type', 'Month', 'Weekday', 'Ordinal', 'Day Of Month', 'Offset Days', 'Include In Form', 'Include In Schedule', 'Sort Order', 'Type', 'Notes'],
    [true, 'sunday_service', '', 'monthly', 'every_weekday', 'all', 'Sunday', 'every', '', 0, true, true, 20, 'service', 'Default weekly Sunday schedule'],
    [false, 'corporate_prayer', 'Corporate Prayer', 'monthly', 'nth_weekday', 'all', 'Friday', 1, '', 0, true, true, 10, 'prayer', 'Enable if your church has a monthly prayer gathering'],
    [false, 'easter', 'Easter', 'yearly', 'easter_offset', 'all', '', '', '', 0, true, true, 30, 'special', 'Enable to include Easter automatically each year'],
    [false, 'christmas', 'Christmas', 'yearly', 'fixed_date', '12', '', '', 25, 0, true, true, 40, 'special', 'Enable to include Christmas automatically each year']
  ];
}

function getDefaultMonthlyEventsSheetRows() {
  return [['Enabled', 'Year', 'Month', 'Date', 'Action', 'Label', 'Rule ID', 'Include In Form', 'Include In Schedule', 'Sort Order', 'Type', 'Notes']];
}
