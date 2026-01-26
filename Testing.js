/**
 * End-to-End Testing Suite
 * 
 * This file contains tests to verify the full workflow of the Ministry Scheduler.
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
  const availabilitySheetName = `${planMonthName} Availability`;
  
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
    // Simple check: assumes headers are dates like "MM/dd" or "MM/dd - Label"
    const val = headerRowValues[i];
    if (val.match(/\d{2}\/\d{2}/)) {
      targetDate = val.split(' - ')[0].trim();
      targetColIndex = i;
      break;
    }
  }
  
  if (!targetDate) throw new Error("Could not find a valid date column in the availability sheet");
  console.log(`Targeting date: ${targetDate} at column index ${targetColIndex}`);

  const testName = "Test User " + new Date().getTime();
  const testRole = CONFIG.roles[0]; // e.g., "WL"
  
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
  
  if (String(cellValue).includes(testName)) {
    console.log(`✅ SUCCESS: Found '${testName}' in the '${testRole}' row for ${targetDate}.`);
  } else {
    throw new Error(`❌ FAILURE: Expected '${testName}' in cell [${roleRowIndex+1}, ${targetColIndex+1}] but found: '${cellValue}'`);
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
      ss.deleteSheet(sheet);
      console.log(`✅ Deleted empty test sheet '${sheet.getName()}'.`);
    }
  }
}