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
  updateDatabase(e);
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

    if (rawDates && rawDates.length > 0) {
      console.log(`Raw Unavailable Dates from Form: ${rawDates[0]}`);

      // 1. Split the single string response (rawDates[0]) into an array of individual date strings.
      unavailableDates = rawDates[0].split(',').map((dateStr) => {
        // 2. Use regex to remove ' - ' followed by any string to the end of the date item.
        // Also trim any leading/trailing spaces from the split process.
        const cleanStr = dateStr.replace(/\s-\s.*$/, '').trim();

        // 3. Convert to Date object to ensure format is MM/dd (and handle potential variations)
        const date = new Date(cleanStr);
        if (!isNaN(date.getTime())) {
          const mm = ('0' + (date.getMonth() + 1)).slice(-2);
          const dd = ('0' + date.getDate()).slice(-2);
          return `${mm}/${dd}`;
        }

        // Fallback for unparseable strings (should be rare)
        console.log(`  -> Fallback returned (unparseable): ${cleanStr}`);
        return cleanStr;
      });
    }

    const unavailableDatesString = unavailableDates.join(','); // Prepare for storage
    const comments = responses[CONFIG.formHeaders.comments] ? responses[CONFIG.formHeaders.comments][0] : "";

    // Log final extracted values before database write
    console.log(`Name: ${name}`);
    console.log(`Times Willing to Serve: ${timesWilling}`);
    console.log(`Parsed Unavailable Dates: ${unavailableDatesString}`);
    console.log(`Comments: ${comments}`);
    
    // --- Database Update Logic ---

    let found = false;
    for (let i = 1; i < databaseData.length; i++) {
      if (databaseData[i][0] == name) {
        // Update the corresponding row
        databaseSheet.getRange(i + 1, 3).setValue(timesWilling);
        databaseSheet.getRange(i + 1, 4).setValue(unavailableDatesString); // Use the joined string
        databaseSheet.getRange(i + 1, 5).setValue(comments);
        found = true;
        console.log(`Updated existing row ${i + 1} for ${name}`);
        break;
      }
    }

    if (!found) {
      // If no match is found, append a new row
      const lastRow = databaseSheet.getLastRow() + 1;
      databaseSheet.getRange(lastRow, 1).setValue(name);
      databaseSheet.getRange(lastRow, 3).setValue(timesWilling);
      databaseSheet.getRange(lastRow, 4).setValue(unavailableDatesString);
      databaseSheet.getRange(lastRow, 5).setValue(comments);
      console.log(`Added new row ${lastRow} for ${name}`);
    }
  } catch (error) {
    console.error(`!!! ERROR in updateDatabase: ${error.message}`);
  }
  updateAvailability();
  console.log('Updated availability and finished execution.');
}

function getServiceDates(year, month) {
  const serviceDates = [];
  
  // Get the first day of the month
  const firstDay = new Date(year, month, 1);
  
  // Find the first Friday of the month
  const firstFriday = new Date(firstDay);
  while (firstFriday.getDay() !== 5) { // 5 represents Friday
    firstFriday.setDate(firstFriday.getDate() + 1);
  }
  serviceDates.push(Utilities.formatDate(firstFriday, Session.getScriptTimeZone(), 'MM/dd') + ' - Corporate Prayer');
  
  // Iterate through the days of the month to find all Sundays
  let currentDate = new Date(firstDay);
  while (currentDate.getMonth() === month) {
    if (currentDate.getDay() === 0) { // 0 represents Sunday
      serviceDates.push(Utilities.formatDate(currentDate, Session.getScriptTimeZone(), 'MM/dd'));
    }
    currentDate.setDate(currentDate.getDate() + 1);
  }
  return serviceDates;
}

function createNewFormForMonth(month, year, monthName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const metadataSheet = ss.getSheetByName(CONFIG.sheetNames.formMetadata) || ss.insertSheet(CONFIG.sheetNames.formMetadata);

  // Create a new form for the upcoming month
  const formTitle = `Music Ministry Availability - ${monthName}`;
  const form = FormApp.create(formTitle);

  // Name Dropdown (ListItem)
  const nameDropdown = form.addListItem().setTitle(CONFIG.formHeaders.name).setRequired(true);
  nameDropdown.setChoiceValues(["Loading..."]);

  const numDropdown = form.addListItem()
  .setTitle(CONFIG.formHeaders.times)
  .setChoiceValues(['1', '2', '3', '4', '5']) // Set the dropdown options
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
  const recipientEmail = CONFIG.ids.adminEmails.join(",");

  // Send email
  MailApp.sendEmail(recipientEmail, emailSubject, emailBody);

  const file = DriveApp.getFileById(form.getId());
  const targetFolder = DriveApp.getFolderById(CONFIG.ids.formsFolder);
  file.moveTo(targetFolder);
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

function setupAvailability(sheetName, year, month) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
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
  const roles = CONFIG.roles;

  // Add each role with empty cells under each Sunday
  roles.forEach(function (role) {
    const roleRow = [role];
    serviceDates.forEach(function () {
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
  const emptyData = roles.map(role => [role, ...Array(5).fill("")]);

  // Add empty data under the "Availability" heading for each role
  emptyData.forEach(function (dataRow) {
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

  const today = new Date();

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

  const newTabName = `${planMonthName} Availability`;
  const deleteTabName = `${oldMonthName} Availability`;
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

  if (metadataSheet) {
    const oldFormId = metadataSheet.getRange("B1").getValue(); // Get the old form ID from metadata
    if (oldFormId) {
      try {
        const oldForm = FormApp.openById(oldFormId); // Open the form using the ID
        oldForm.removeDestination(); // Remove the link to the spreadsheet
        console.log("De-linked form with ID: " + oldFormId);
      } catch (e) {
        console.log("Could not de-link or find the old form: " + e.message);
      }
    }
  }

  const sheets = ss.getSheets();
  sheets.forEach(function(sheet) {
    if (sheet.getName().startsWith("Form Responses")) {
      const toDelete = sheet.getName();
      ss.deleteSheet(sheet);
      console.log("Deleted old Form Responses tab: " + toDelete);
    }
  })
  createNewFormForMonth(planMonth, planYear, planMonthName);
  console.log(`Created new form for ${planMonthName}`);
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

  console.log('--- STARTING updateAvailability ---');

  const today = new Date();
  const planDate = new Date(today);
  planDate.setMonth(today.getMonth() + 1);
  const planMonthName = planDate.toLocaleString('default', { month: 'long' });
  const sheetName = planMonthName + " Availability";

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

  const serviceDateKeys = dateHeaders.map(h => {
    return String(h).trim().substring(0, 5);
  });

  console.log('Standardized Date Keys: ' + serviceDateKeys.join(', '));

  // Initialize the availability object
  const availability = {};
  let roleOrder = CONFIG.roles;

  // Standardize roleOrder to uppercase for case-insensitive matching
  roleOrder = roleOrder.map(role => role.toUpperCase());

  // Process each row in the Ministry Members sheet
  for (let i = 1; i < databaseData.length; i++) {
    const row = databaseData[i];
    let name = row[0] ? row[0].trim() : "";
    
    if (!name) continue; 
    
    const roles = row[1]
      ? row[1].toString().split(",").map(role => {
          return role.trim().toUpperCase();
        })
      : [];
    const timesWilling = row[2] ? row[2].toString().trim() : "";
    const rawUnavailableDates = row[3] ? row[3].toString() : "";
    
    const unavailableDates = rawUnavailableDates
      ? rawUnavailableDates.split(",").map(dateStr => {
      const trimmedDateStr = dateStr.trim();
      const parsedDate = new Date(trimmedDateStr);
      
      if (!isNaN(parsedDate.getTime())) {
        const mm = ('0' + (parsedDate.getMonth() + 1)).slice(-2);
        const dd = ('0' + parsedDate.getDate()).slice(-2);
        return mm + '/' + dd;
      } else {
        // Fallback: just take the first 5 characters (e.g., "MM/dd") if parse fails
        return trimmedDateStr.substring(0, 5);
      }
    })
  : [];

    if (!name || !roles.length) continue;

    // Format the name as "Firstname L."
    const nameParts = name.split(" ");
    if (nameParts.length > 1) {
      name = nameParts[0] + " " + nameParts[1].charAt(0).toUpperCase() + ".";
    }

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
      CONFIG.sheetHeaders.comments
    ];
    dbSheet.appendRow(headers);
    dbSheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");
    
    // Add Dummy Data
    const dummyRow = ["John Doe", "WL, ACOUSTIC", "4", "", "Excited to serve!"];
    dbSheet.appendRow(dummyRow);
    console.log("Created Ministry Members sheet with dummy data.");
  }

  // 2. Create Form Metadata sheet if it doesn't exist
  let metaSheet = ss.getSheetByName(CONFIG.sheetNames.formMetadata);
  if (!metaSheet) {
    metaSheet = ss.insertSheet(CONFIG.sheetNames.formMetadata);
    metaSheet.appendRow(["Form Name", "Form ID"]);
    console.log("Created Form Metadata sheet.");
  }
  
  console.log("Initialization complete. You can now run monthlySetup().");
}