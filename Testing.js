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
  const seededMemberName = "Existing Test Member";

  // --- Step 1: Initial Setup ---
  console.log("\n--- Step 1: Initial Setup ---");
  // Ensure the project is initialized (creates sheets if missing)
  initializeProject();
  
  // Verify sheets exist
  if (!ss.getSheetByName(CONFIG.sheetNames.ministryMembers)) throw new Error("Ministry Members sheet missing");
  if (!ss.getSheetByName(CONFIG.sheetNames.formMetadata)) throw new Error("Form Metadata sheet missing");
  ensureTestMemberExists(seededMemberName, loadRuntimeSettings().roles[0]);
  console.log("✅ Initial setup verified");

  // --- Step 2: Monthly Setup ---
  console.log("\n--- Step 2: Monthly Setup ---");
  // Run the monthly setup to create the form and availability sheet
  runMonthlySetupNow();
  
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
  
  const roleHeaders = dbSheet.getRange(1, 1, 1, dbSheet.getLastColumn()).getDisplayValues()[0];
  const roleColumnIndex = roleHeaders.indexOf(testRole) + 1;
  if (roleColumnIndex <= 0) throw new Error(`Role checkbox column for '${testRole}' not found`);

  dbSheet.getRange(userRowIndex, roleColumnIndex).setValue(true);
  SpreadsheetApp.flush(); // Ensure the checkbox and derived Roles formula are saved before the script reads again

  const derivedRoles = dbSheet.getRange(userRowIndex, getMinistryMembersColumnMap(dbSheet).roles).getDisplayValue();
  if (derivedRoles.indexOf(testRole) === -1) {
    throw new Error(`Derived Roles column did not include '${testRole}'. Found '${derivedRoles}'`);
  }

  console.log(`✅ Form submission processed. Checked role '${testRole}' for '${testName}'.`);

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
  cleanupTestArtifacts(ss, testName, availabilitySheetName, planMonthName, [seededMemberName]);
}

/**
 * Cleans up test data to leave the spreadsheet in a clean state.
 */
function cleanupTestArtifacts(ss, testName, createdSheetName, monthName, additionalMemberNames) {
  console.log("\n--- Step 5: Cleanup ---");
  
  // 1. Remove the test user from the database
  const dbSheet = ss.getSheetByName(CONFIG.sheetNames.ministryMembers);
  const dbData = dbSheet.getDataRange().getValues();
  const namesToDelete = [testName].concat(additionalMemberNames || []);
  
  for (let i = dbData.length - 1; i >= 0; i--) {
    if (namesToDelete.indexOf(dbData[i][0]) !== -1) {
      dbSheet.deleteRow(i + 1);
      console.log(`✅ Deleted test user '${dbData[i][0]}' from database.`);
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
  const responseSheetNames = ss.getSheets()
    .map(sheet => {
      try {
        return sheet.getName();
      } catch (e) {
        return '';
      }
    })
    .filter(name => name.startsWith("Form Responses"));

  for (const sheetName of responseSheetNames) {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) continue;

    try {
      if (sheet.getLastRow() <= 1) {
        ss.deleteSheet(sheet);
        console.log(`✅ Deleted empty test sheet '${sheetName}'.`);
      }
    } catch (e) {
      console.log(`⚠️ Could not delete linked test sheet '${sheetName}': ${e.message}`);
    }
  }
}

/** Unit test harness and tests for parsing/normalization utilities */
function runUnitTests() {
  initializeProject();
  resetTestResultsSheet();
  deleteSheetIfExists(CONFIG.sheetNames.monthlyEvents);
  deleteSheetIfExists(CONFIG.sheetNames.recurringEvents);
  replaceSheetContents(CONFIG.sheetNames.recurring, getDefaultRecurringSheetRows());
  replaceSheetContents(CONFIG.sheetNames.settings, getDefaultSettingsSheetRows());
  replaceSheetContents(CONFIG.sheetNames.events, getDefaultEventsSheetRows());

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
    if (parseSingleDate('2026-02-31') !== null) throw new Error('Invalid ISO dates should be rejected');
    if (parseSingleDate('2/31/2026') !== null) throw new Error('Invalid slash dates should be rejected');
    recordResult('parseSingleDate: invalid calendar dates', true, 'Impossible dates are rejected');
  } catch (e) { recordResult('parseSingleDate: invalid calendar dates', false, e.message); }

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
    const recurringRows = getDefaultRecurringSheetRows();
    recurringRows[3][0] = true; // Enable corporate prayer
    replaceSheetContents(CONFIG.sheetNames.recurring, recurringRows);
    const sd = getServiceDates(2026, 2);
    if (!sd.some(date => date.indexOf('03/06 - Corporate Prayer') !== -1)) throw new Error('Corporate Prayer should appear when monthly recurring event is enabled');
    recordResult('getServiceDates:monthlyRecurringEvent', true, 'Monthly recurring event enabled successfully');
  } catch (e) { recordResult('getServiceDates:monthlyRecurringEvent', false, e.message); }

  try {
    const recurringRows = getDefaultRecurringSheetRows();
    recurringRows[3][0] = true; // Enable corporate prayer for override test
    replaceSheetContents(CONFIG.sheetNames.recurring, recurringRows);
    replaceSheetContents(CONFIG.sheetNames.events, [
      ['Enabled', 'Date', 'Event', 'Action', 'Recurring Event', 'Include In Form', 'Include In Schedule', 'Notes'],
      [true, '2026-03-06', 'Corporate Prayer', 'REMOVE', 'corporate_prayer', true, true, 'Moved off first Friday'],
      [true, '2026-03-13', 'Corporate Prayer', 'ADD', 'corporate_prayer', true, true, 'Moved to second Friday']
    ]);
    const sd = getServiceDates(2026, 2);
    if (sd.some(date => date.indexOf('03/06 - Corporate Prayer') !== -1)) throw new Error('Corporate Prayer was not removed from first Friday');
    if (!sd.some(date => date.indexOf('03/13 - Corporate Prayer') !== -1)) throw new Error('Corporate Prayer was not added to second Friday');
    recordResult('getServiceDates:actionOverrides', true, 'ADD/REMOVE override flow works');
  } catch (e) { recordResult('getServiceDates:actionOverrides', false, e.message); }

  try {
    replaceSheetContents(CONFIG.sheetNames.recurring, getDefaultRecurringSheetRows());
    replaceSheetContents(CONFIG.sheetNames.events, [
      ['Enabled', 'Date', 'Event', 'Action', 'Recurring Event', 'Include In Form', 'Include In Schedule', 'Notes'],
      [true, '2026-03-15', '', 'REMOVE', '', true, true, 'Cancel one Sunday']
    ]);
    const sd = getServiceDates(2026, 2);
    if (sd.includes('03/15')) throw new Error('Date-only REMOVE should remove a plain Sunday');
    recordResult('getServiceDates:removePlainSunday', true, 'Date-only Sunday removal works');
  } catch (e) { recordResult('getServiceDates:removePlainSunday', false, e.message); }

  try {
    const eventRows = getDefaultEventsSheetRows().map(row => row.slice());
    eventRows.forEach(row => {
      if (row[2] === 'Easter' || row[2] === 'Christmas') row[0] = true;
    });
    replaceSheetContents(CONFIG.sheetNames.recurring, getDefaultRecurringSheetRows());
    replaceSheetContents(CONFIG.sheetNames.events, eventRows);
    const aprilDates = getServiceDates(2026, 3); // April 2026
    const decemberDates = getServiceDates(2026, 11); // December 2026
    if (!aprilDates.includes('04/05 - Easter')) throw new Error('Easter should appear when enabled in Events');
    if (!decemberDates.includes('12/25 - Christmas')) throw new Error('Christmas should appear when enabled in Events');
    recordResult('getServiceDates:datedSpecialEvents', true, 'Dated special events are driven by Events');
  } catch (e) { recordResult('getServiceDates:datedSpecialEvents', false, e.message); }

  try {
    deleteSheetIfExists(CONFIG.sheetNames.events);
    replaceSheetContents(CONFIG.sheetNames.monthlyEvents, [
      ['Year', 'Month', 'Date', 'Label', 'Type'],
      [2026, 5, '2026-05-15', 'LegacyEvent', 'legacy']
    ]);
    const sd = getServiceDates(2026, 4); // May 2026
    if (!sd.some(date => date.indexOf('05/15 - LegacyEvent') !== -1)) throw new Error('Legacy Monthly Events did not remain authoritative');
    recordResult('getServiceDates:legacyMonthlyEvents', true, 'Legacy Monthly Events compatibility preserved');
  } catch (e) { recordResult('getServiceDates:legacyMonthlyEvents', false, e.message); }

  try {
    deleteSheetIfExists(CONFIG.sheetNames.monthlyEvents);
    const recurringRows = getDefaultRecurringSheetRows();
    recurringRows[3][0] = true; // Enable corporate prayer
    replaceSheetContents(CONFIG.sheetNames.recurring, recurringRows);
    const eventRows = getDefaultEventsSheetRows().map(row => row.slice());
    eventRows.forEach(row => {
      if (row[2] === 'Easter') row[0] = true;
    });
    replaceSheetContents(CONFIG.sheetNames.events, eventRows);
    const sd = getServiceDates(2026, 3); // April 2026
    if (!sd.includes('04/03 - Corporate Prayer')) throw new Error('Corporate Prayer should be labeled');
    if (!sd.includes('04/05 - Easter')) throw new Error('Easter should be labeled');
    if (sd.includes('4/5/2026 - Easter')) throw new Error('Displayed dates must be standardized to MM/DD');
    if (sd.includes('04/05')) throw new Error('Plain same-day Easter date should not also appear');
    recordResult('getServiceDates:standardizedDisplay', true, 'Special events retain labels with standardized dates');
  } catch (e) { recordResult('getServiceDates:standardizedDisplay', false, e.message); }

  try {
    if (!shouldArchiveEventsNow(new Date(2026, 0, 1), { eventsArchiveFrequency: 'Yearly' })) {
      throw new Error('Yearly archive should run on January 1');
    }
    if (shouldArchiveEventsNow(new Date(2026, 0, 2), { eventsArchiveFrequency: 'Yearly' })) {
      throw new Error('Yearly archive should not run after January 1');
    }
    if (!shouldArchiveEventsNow(new Date(2026, 3, 1), { eventsArchiveFrequency: 'Quarterly' })) {
      throw new Error('Quarterly archive should run on the first day of a new quarter');
    }
    if (shouldArchiveEventsNow(new Date(2026, 3, 2), { eventsArchiveFrequency: 'Quarterly' })) {
      throw new Error('Quarterly archive should not run after the first day of the quarter');
    }
    if (shouldArchiveEventsNow(new Date(2026, 4, 15), { eventsArchiveFrequency: 'Off' })) {
      throw new Error('Archive should not run when frequency is Off');
    }
    recordResult('eventsArchive:schedule', true, 'Archive cadence settings behave as expected');
  } catch (e) { recordResult('eventsArchive:schedule', false, e.message); }

  try {
    if (getAdminRecipientString({ adminEmails: ['admin1@example.com', ' admin2@example.com '] }) !== 'admin1@example.com,admin2@example.com') {
      throw new Error('Admin email list did not normalize correctly');
    }
    replaceSheetContents(CONFIG.sheetNames.admins, getDefaultAdminsSheetRows());
    const settingsFromSheet = loadRuntimeSettings();
    if (getAdminRecipientString(settingsFromSheet) !== 'admin1@example.com,admin2@example.com') {
      throw new Error('Admins sheet should be the preferred source for admin recipients');
    }
    const reminder = buildAdminPlanningReminder(new Date(2026, 2, 5), {
      churchName: 'Jubal Test',
      timeZone: safeGetScriptTimeZone(),
      adminEmails: ['admin@example.com'],
      formCreationDay: 8
    });
    if (reminder.subject !== 'Jubal Test: April 2026 Events Reminder') throw new Error('Reminder subject should use the church name and month-forward reminder format');
    if (reminder.subject.indexOf('Jubal Test') === -1) throw new Error('Reminder subject should include the church name');
    if (reminder.body.indexOf('Church: Jubal Test') === -1) throw new Error('Reminder body should include the church name');
    if (reminder.body.indexOf('Recurring schedule') === -1 || reminder.body.indexOf('One-time events and changes') === -1) {
      throw new Error('Reminder body should guide admins to Recurring and Events');
    }
    if (reminder.body.indexOf('Monthly setup day: the 8th of each month.') === -1) {
      throw new Error('Reminder should mention the configured monthly setup day');
    }
    if (reminder.body.indexOf('Admin contacts and notifications') === -1) {
      throw new Error('Reminder should link admins to the Admins sheet');
    }
    recordResult('adminReminder:content', true, reminder.subject);
  } catch (e) { recordResult('adminReminder:content', false, e.message); }

  try {
    if (!shouldRunMonthlySetupToday(new Date(2026, 2, 8), { formCreationDay: 8 })) {
      throw new Error('Monthly setup should run on the configured setup day');
    }
    if (shouldRunMonthlySetupToday(new Date(2026, 2, 7), { formCreationDay: 8 })) {
      throw new Error('Monthly setup should not run before the configured setup day');
    }
    if (getMonthlySetupPropertyKey(new Date(2026, 2, 8)) !== 'monthlySetupCompleted:2026-04') {
      throw new Error('Monthly setup property key should target the next month');
    }
    recordResult('monthlySetup:guardLogic', true, 'Daily trigger guard uses form_creation_day and next-month keys');
  } catch (e) { recordResult('monthlySetup:guardLogic', false, e.message); }

  try {
    const invalidSettingsRows = getDefaultSettingsSheetRows().map(row => row.slice());
    invalidSettingsRows.forEach(row => {
      if (row[0] === 'form_creation_day') row[1] = 31;
      if (row[0] === 'admin_reminder_day') row[1] = 0;
    });
    replaceSheetContents(CONFIG.sheetNames.settings, invalidSettingsRows);
    const clampedSettings = loadRuntimeSettings();
    assertEqual(clampedSettings.formCreationDay, 28, 'form_creation_day should clamp to 28');
    assertEqual(clampedSettings.adminReminderDay, 1, 'admin_reminder_day should clamp to 1');
    recordResult('settings:clampInvalidDays', true, 'Invalid day-of-month settings clamp to the supported 1-28 range');
  } catch (e) { recordResult('settings:clampInvalidDays', false, e.message); }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const tempSheetName = 'Admins Layout Test';
    deleteSheetIfExists(tempSheetName);
    const sheet = ss.insertSheet(tempSheetName);
    sheet.getRange(1, 1, 3, 4).setValues([
      ['Enabled', 'Name', 'Email', 'Notes'],
      [true, 'Widodo', 'widodo@example.com', 'Primary admin'],
      [false, 'Legacy Name', 'legacy@example.com', '']
    ]);
    normalizeAdminsSheetLayout(sheet);

    const headers = sheet.getRange(1, 1, 1, 3).getDisplayValues()[0];
    if (headers.join('|') !== 'Enabled|Email|Notes') {
      throw new Error('Admins layout should normalize to Enabled, Email, Notes');
    }
    if (sheet.getRange(2, 2).getValue() !== 'widodo@example.com') {
      throw new Error('Email should remain in the Email column after Admins migration');
    }
    if (sheet.getRange(2, 3).getValue() !== 'Primary admin') {
      throw new Error('Notes should be preserved during Admins migration');
    }
    if (sheet.getRange(3, 3).getValue() !== 'Legacy Name') {
      throw new Error('Legacy Name column should fall back into Notes when Notes is blank');
    }
    ss.deleteSheet(sheet);
    recordResult('admins:normalizeLayout', true, 'Legacy Admins layouts migrate cleanly');
  } catch (e) {
    deleteSheetIfExists('Admins Layout Test');
    recordResult('admins:normalizeLayout', false, e.message);
  }

  try {
    const ordered = sortMonthSheetsByRecency([
      { getName: () => 'January' },
      { getName: () => 'March' },
      { getName: () => 'April' },
      { getName: () => 'May' }
    ], new Date(2026, 2, 30)).map(sheet => sheet.getName());
    if (ordered.join('|') !== 'April|May|March|January') {
      throw new Error('Expected month tabs ordered by nearest upcoming month first, got ' + ordered.join(', '));
    }
    recordResult('sheetOrder:monthSorting', true, ordered.join(' -> '));
  } catch (e) { recordResult('sheetOrder:monthSorting', false, e.message); }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    deleteSheetIfExists('Execution Logs');
    deleteSheetIfExists('Debug Responses');
    disableDeveloperDiagnostics();
    logDebug('info', 'Testing default diagnostics mode');
    logFormResponse({ namedValues: { Name: ['Test'] } });
    if (ss.getSheetByName('Execution Logs') || ss.getSheetByName('Debug Responses')) {
      throw new Error('Developer sheets should not be created when diagnostics are disabled');
    }
    recordResult('developerDiagnostics:disabledByDefault', true, 'Developer sheets stay out of admin-facing workbooks by default');
  } catch (e) { recordResult('developerDiagnostics:disabledByDefault', false, e.message); }

  try {
    deleteSheetIfExists(CONFIG.sheetNames.eventsArchive);
    replaceSheetContents(CONFIG.sheetNames.events, getDefaultEventsSheetRows().concat([
      [true, new Date(2026, 11, 24), 'Christmas Eve', 'ADD', '', true, true, 'Past real event'],
      [true, new Date(2027, 0, 15), 'Vision Night', 'ADD', '', true, true, 'Current month event'],
      [true, new Date(2027, 1, 10), 'Prayer Night', 'ADD', '', true, true, 'Future event']
    ]));
    const archiveResult = archivePastEventsIfDue(new Date(2027, 0, 1), { eventsArchiveFrequency: 'Yearly' });
    if (archiveResult.archivedRows !== 1) throw new Error('Expected exactly one past real event to be archived');

    const eventsRows = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.sheetNames.events).getDataRange().getDisplayValues();
    const archiveRows = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.sheetNames.eventsArchive).getDataRange().getDisplayValues();

    if (!eventsRows.some(row => row.join('|').indexOf('Good Friday') !== -1)) throw new Error('Example Good Friday row should remain in Events');
    if (!eventsRows.some(row => row.join('|').indexOf('Corporate Prayer') !== -1 && row.join('|').indexOf('REMOVE') !== -1)) throw new Error('Example Corporate Prayer row should remain in Events');
    if (eventsRows.some(row => row.join('|').indexOf('Christmas Eve') !== -1)) throw new Error('Past real event should have been removed from Events');
    if (!archiveRows.some(row => row.join('|').indexOf('Christmas Eve') !== -1)) throw new Error('Past real event should have been copied to Events Archive');
    recordResult('eventsArchive:preserveExamples', true, JSON.stringify(archiveResult));
  } catch (e) { recordResult('eventsArchive:preserveExamples', false, e.message); }

  deleteSheetIfExists(CONFIG.sheetNames.monthlyEvents);
  deleteSheetIfExists(CONFIG.sheetNames.eventsArchive);
  replaceSheetContents(CONFIG.sheetNames.events, getDefaultEventsSheetRows());
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
    deleteSheetIfExists('Execution Logs');
    deleteSheetIfExists('Debug Responses');
    disableDeveloperDiagnostics();
    const executionLogs = ss.insertSheet('Execution Logs');
    executionLogs.appendRow(['timestamp', 'level', 'message', 'data']);
    const debugResponses = ss.insertSheet('Debug Responses');
    debugResponses.appendRow(['timestamp', 'formId', 'responseRow', 'namedValues']);
    initializeProject();
    const dbSheet = ss.getSheetByName(CONFIG.sheetNames.ministryMembers);
    const runtimeSettings = loadRuntimeSettings();
    const memberColumns = getMinistryMembersColumnMap(dbSheet);
    const settingsSheet = ss.getSheetByName(CONFIG.sheetNames.settings);
    const adminsSheet = ss.getSheetByName(CONFIG.sheetNames.admins);
    const rolesSheet = ss.getSheetByName(CONFIG.sheetNames.rolesConfig);
    const recurringSheet = ss.getSheetByName(CONFIG.sheetNames.recurring) || ss.getSheetByName(CONFIG.sheetNames.recurringEvents);
    const eventsSheet = ss.getSheetByName(CONFIG.sheetNames.events) || ss.getSheetByName(CONFIG.sheetNames.monthlyEvents);
    if (!settingsSheet) throw new Error('Settings sheet missing');
    if (!adminsSheet) throw new Error('Admins sheet missing');
    if (!rolesSheet) throw new Error('Roles sheet missing');
    const settingsValues = settingsSheet.getRange(2, 1, settingsSheet.getLastRow() - 1, 1).getDisplayValues().flat();
    if (!recurringSheet) {
      throw new Error('Recurring sheet missing');
    }
    if (!eventsSheet) {
      throw new Error('Events sheet missing');
    }
    if (eventsSheet.getRange(2, 10).getDisplayValue() !== 'How to Add Events') {
      throw new Error('Events sheet should include the visible instruction banner');
    }
    const eventsBannerBody = eventsSheet.getRange(3, 10).getDisplayValue();
    if (eventsBannerBody.indexOf('DOUBLE-CLICK') === -1 || eventsBannerBody.indexOf('OR') === -1 || eventsBannerBody.indexOf('Add Special Event') === -1) {
      throw new Error('Events instruction banner should clearly separate sheet entry from the Add Special Event dialog');
    }
    const recurringHeaders = recurringSheet.getRange(1, 1, 1, recurringSheet.getLastColumn()).getDisplayValues()[0];
    if (recurringHeaders.indexOf('Month') !== -1 || recurringHeaders.indexOf('Day') !== -1) {
      throw new Error('Recurring should use the simplified weekly/monthly layout without Month or Day columns');
    }
    if (!eventsSheet.getDataRange().getDisplayValues().some(row => row.join('|').indexOf('Easter') !== -1)) {
      throw new Error('Events sheet should include an Easter example row');
    }
    if (!eventsSheet.getDataRange().getDisplayValues().some(row => row.join('|').indexOf('Christmas') !== -1)) {
      throw new Error('Events sheet should include a Christmas example row');
    }
    if (!dbSheet || dbSheet.getRange(1, memberColumns.canonicalName).getValue() !== CONFIG.sheetHeaders.canonicalName) {
      throw new Error('Canonical Name header missing from Ministry Members');
    }
    if (dbSheet.getLastRow() !== 1) {
      throw new Error('initializeProject should not seed a dummy ministry member row');
    }
    if (memberColumns.canonicalName !== dbSheet.getLastColumn()) {
      throw new Error('Canonical Name should be the last column in Ministry Members');
    }
    if (dbSheet.getRange(1, memberColumns.dates).getValue() !== CONFIG.sheetHeaders.dates) {
      throw new Error('Unavailable Dates header missing from Ministry Members');
    }
    if (dbSheet.getRange(1, memberColumns.times).getValue() !== CONFIG.sheetHeaders.times) {
      throw new Error('Times Willing to Serve header missing from Ministry Members');
    }
    if (dbSheet.getRange(1, memberColumns.roles).getValue() !== CONFIG.sheetHeaders.roles) {
      throw new Error('Roles header missing from Ministry Members');
    }
    if (dbSheet.getRange(1, getRoleCheckboxStartColumn(dbSheet)).getValue() !== runtimeSettings.roles[0]) {
      throw new Error('Role checkbox headers were not created in Ministry Members');
    }
    const firstRoleValidation = dbSheet.getRange(2, getRoleCheckboxStartColumn(dbSheet)).getDataValidation();
    if (!firstRoleValidation || firstRoleValidation.getCriteriaType() !== SpreadsheetApp.DataValidationCriteria.CHECKBOX) {
      throw new Error('Role checkbox columns should use checkbox validation');
    }
    if (dbSheet.getRange(2, memberColumns.canonicalName).getDataValidation() !== null) {
      throw new Error('Canonical Name column should not keep checkbox validation');
    }
    if (settingsValues.indexOf('events_archive_frequency') === -1) {
      throw new Error('Archive settings were not created in Settings');
    }
    if (settingsValues.indexOf('admin_reminder_enabled') === -1 || settingsValues.indexOf('admin_reminder_day') === -1) {
      throw new Error('Admin reminder settings were not created in Settings');
    }
    if (settingsValues.indexOf('admin_emails') !== -1 || settingsValues.indexOf('roles') !== -1 || settingsValues.indexOf('events_archive_month') !== -1) {
      throw new Error('Deprecated Settings rows should have been removed');
    }
    if (!sheetUsesFriendlyAdminsLayout(adminsSheet)) {
      throw new Error('Admins sheet does not use the friendly admin layout');
    }
    if (adminsSheet.getRange(1, 1, 1, 3).getDisplayValues()[0].join('|') !== 'Enabled|Email|Notes') {
      throw new Error('Admins sheet headers should normalize to Enabled, Email, Notes');
    }
    if (String(adminsSheet.getRange(2, 2).getDisplayValue() || '').trim() !== '') {
      throw new Error('Fresh setup should not seed a real admin email into the Admins sheet');
    }
    if (!sheetUsesFriendlyRolesLayout(rolesSheet)) {
      throw new Error('Roles sheet does not use the friendly role layout');
    }
    if (!ss.getSheetByName('Execution Logs').isSheetHidden() || !ss.getSheetByName('Debug Responses').isSheetHidden()) {
      throw new Error('Developer-only sheets should be hidden from admins by default');
    }
    recordResult('initializeProject:configSheets', true, 'Configuration sheets created');
  } catch (e) { recordResult('initializeProject:configSheets', false, e.message); }

  try {
    const referenceDate = new Date(2026, 2, 30);
    const planningContext = getPlanningMonthContext(referenceDate, loadRuntimeSettings());
    const dbSheet = ss.getSheetByName(CONFIG.sheetNames.ministryMembers);
    if (dbSheet.getLastRow() > 1) {
      dbSheet.getRange(2, 1, dbSheet.getLastRow() - 1, dbSheet.getLastColumn()).clearContent();
    }

    let failedAsExpected = false;
    try {
      runMonthlySetupInternal({ force: true, referenceDate: referenceDate });
    } catch (error) {
      failedAsExpected = error.message.indexOf('does not have any names yet') !== -1;
    }

    if (!failedAsExpected) {
      throw new Error('monthlySetup should fail cleanly when Ministry Members has no names');
    }
    if (ss.getSheetByName(planningContext.sheetName)) {
      throw new Error('monthlySetup should clean up the staged availability sheet when form creation fails');
    }
    if (ss.getSheetByName(planningContext.sheetName + ' (Staging)')) {
      throw new Error('monthlySetup should remove the staging availability sheet after a failure');
    }
    if (getTrackedFormIds(ss.getSheetByName(CONFIG.sheetNames.formMetadata)).length) {
      throw new Error('Failed monthlySetup should not commit a new form into Form Metadata');
    }

    recordResult('monthlySetup:noMembersFailure', true, 'monthlySetup fails safely when the roster is empty');
  } catch (e) { recordResult('monthlySetup:noMembersFailure', false, e.message); }

  try {
    const result = migrateMemberRolesToCheckboxes();
    if (!result || (result.status !== 'already_configured' && result.status !== 'migrated')) {
      throw new Error('Expected migrateMemberRolesToCheckboxes to be idempotent');
    }
    recordResult('migrateMemberRolesToCheckboxes:smoke', true, JSON.stringify(result));
  } catch (e) { recordResult('migrateMemberRolesToCheckboxes:smoke', false, e.message); }

  // test: ensureMonthlyEventsFor remains usable for legacy installs
  try {
    deleteSheetIfExists(CONFIG.sheetNames.events);
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
    replaceSheetContents(CONFIG.sheetNames.rolesConfig, getDefaultRolesSheetRows());
    replaceSheetContents(CONFIG.sheetNames.recurring, getDefaultRecurringSheetRows());
    replaceSheetContents(CONFIG.sheetNames.events, getDefaultEventsSheetRows());
    deleteSheetIfExists(CONFIG.sheetNames.monthlyEvents);
    ensureTestMemberExists('Monthly Setup Smoke Test', 'WL');
    runMonthlySetupNow();
    recordResult('monthlySetup:smoke', true, 'runMonthlySetupNow executed');
  } catch (e) { recordResult('monthlySetup:smoke', false, e.message); }

  // test: applying event changes refreshes the next month sheet without losing same-date schedule assignments
  try {
    replaceSheetContents(CONFIG.sheetNames.settings, getDefaultSettingsSheetRows());
    replaceSheetContents(CONFIG.sheetNames.rolesConfig, getDefaultRolesSheetRows());
    replaceSheetContents(CONFIG.sheetNames.recurring, getDefaultRecurringSheetRows());
    replaceSheetContents(CONFIG.sheetNames.events, getDefaultEventsSheetRows());

    const referenceDate = new Date(2026, 2, 30); // March 30, 2026 -> April 2026 planning month
    const planningContext = getPlanningMonthContext(referenceDate, loadRuntimeSettings());
    setupAvailability(planningContext.sheetName, planningContext.planYear, planningContext.planMonth);

    const availabilitySheet = ss.getSheetByName(planningContext.sheetName);
    availabilitySheet.getRange(2, 2).setValue('Leader A');

    const eventRows = getDefaultEventsSheetRows().map(row => row.slice());
    eventRows.forEach(row => {
      if (row[2] === 'Easter') row[0] = true;
    });
    replaceSheetContents(CONFIG.sheetNames.events, eventRows);
    configureEventsSheetUi(ss.getSheetByName(CONFIG.sheetNames.events));

    const result = applyEventChangesToPlanningMonth({
      referenceDate: referenceDate,
      skipFormSync: true,
      skipMatrixRefresh: true
    });

    const refreshedSheet = ss.getSheetByName(planningContext.sheetName);
    const refreshedHeader = refreshedSheet.getRange(1, 2).getDisplayValue();
    const preservedValue = refreshedSheet.getRange(2, 2).getDisplayValue();

    if (refreshedHeader.indexOf('04/05 - Easter') === -1) {
      throw new Error('Expected Easter label to be applied to the April 5 header');
    }
    if (preservedValue !== 'Leader A') {
      throw new Error(`Expected existing schedule assignment to be preserved, found '${preservedValue}'`);
    }
    if (result.restoredAssignments !== 1) {
      throw new Error('Expected exactly one preserved schedule assignment');
    }
    if (result.addedDates.length !== 0 || result.removedDates.length !== 0) {
      throw new Error('Relabeling an existing date should not be treated as an added or removed date');
    }

    recordResult('applyEventChangesToPlanningMonth:preserveAssignments', true, JSON.stringify(result));
  } catch (e) { recordResult('applyEventChangesToPlanningMonth:preserveAssignments', false, e.message); }
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

function ensureTestMemberExists(name, role) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.sheetNames.ministryMembers);
  if (!sheet) throw new Error('Ministry Members sheet missing');

  const memberColumns = getMinistryMembersColumnMap(sheet);
  const existingNames = sheet.getLastRow() >= 2
    ? sheet.getRange(2, memberColumns.name, sheet.getLastRow() - 1, 1).getDisplayValues().flat().map(value => String(value || '').trim())
    : [];

  const existingIndex = existingNames.indexOf(name);
  if (existingIndex !== -1) return existingIndex + 2;

  const width = Math.max(sheet.getLastColumn(), memberColumns.canonicalName);
  const row = Array(width).fill('');
  row[memberColumns.name - 1] = name;
  row[memberColumns.times - 1] = '4';
  row[memberColumns.comments - 1] = 'Seeded test member';
  row[memberColumns.canonicalName - 1] = normalizeName(name);
  sheet.appendRow(row);

  const rowNumber = sheet.getLastRow();
  const roleColumnMap = getConfiguredMemberRoleColumnMap(sheet);
  const roleColumn = roleColumnMap[String(role || '').trim().toUpperCase()];
  if (roleColumn) {
    sheet.getRange(rowNumber, roleColumn).setValue(true);
  }
  ensureRolesFormulaForRow(sheet, rowNumber, loadRuntimeSettings().roles);
  return rowNumber;
}

function getDefaultSettingsSheetRows() {
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

function getDefaultAdminsSheetRows() {
  return [
    ['Enabled', 'Email', 'Notes'],
    [true, 'admin1@example.com', 'Main admin'],
    [true, 'admin2@example.com', 'Also receives alerts'],
    [false, 'ignore@example.com', 'Example row']
  ];
}

function getDefaultRolesSheetRows() {
  return [
    ['Enabled', 'Role', 'Notes'],
    [false, 'MEDIA', 'Example row'],
    [true, 'WL', 'Worship leader'],
    [true, 'SINGER', 'Vocals'],
    [true, 'DRUMS', 'Drum kit']
  ];
}

function getDefaultRecurringSheetRows() {
  return [
    ['Enabled', 'Event', 'Frequency', 'Weekday', 'Week Of Month', 'Include In Form', 'Include In Schedule', 'Notes'],
    [true, '', 'Weekly', 'Sunday', 'every', true, true, 'Default weekly Sunday schedule. Leave Event blank to show plain dates.'],
    [false, 'Midweek Rehearsal', 'Weekly', 'Wednesday', 'every', false, false, 'Example weekly event that stays off the form by default'],
    [false, 'Corporate Prayer', 'Monthly', 'Friday', 1, true, true, 'Enable if your church has a monthly prayer gathering']
  ];
}

function getDefaultEventsSheetRows() {
  return [
    ['Enabled', 'Date', 'Event', 'Action', 'Recurring Event', 'Include In Form', 'Include In Schedule', 'Notes'],
    [false, new Date(2026, 3, 3), 'Good Friday', 'ADD', '', true, true, 'Example row'],
    [false, new Date(2026, 3, 5), 'Easter', 'ADD', '', true, true, 'Example row'],
    [false, new Date(2026, 11, 25), 'Christmas', 'ADD', '', true, true, 'Example row'],
    [false, new Date(2026, 3, 3), 'Corporate Prayer', 'REMOVE', 'Corporate Prayer', true, true, 'Example row']
  ];
}
