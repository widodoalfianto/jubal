/**
 * 
 * Google Apps Script for automating church music ministry scheduling.
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

var databaseFileId = "1cTlyG3m3i3OYZU7X2LYPJ03TUMKkFVakSIbhhVVH1mE"; // Database file ID
var ministrySheetName = "Ministry Members"; // Sheet name for ministry members

var formsFolderId = '1RMITTfCVaYzc0RBBtF0caE8RI5yeq57d'
var formNameHeader = 'Select your name';
var formTimesHeader = 'How many times are you willing to serve this month?';
var formDatesHeader = 'Which days are you NOT available? If re-submitting, please re-submit this section also';
var formCommentsHeader = 'Comments(optional)';

var headerRowIndex = 13;
var dateRowIndex = 1;

var sheetsNameHeader = 'Name';
var sheetsRolesHeader = 'Roles';
var sheetsTimesHeader = 'Times Willing to Serve';
var sheetsDatesHeader = 'Unavailable Dates';
var sheetsCommentsHeader = 'Comments';

function onFormSubmit(e) {
  updateDatabase(e);
}

function updateDatabase(e) {
  try {
    var databaseSS = SpreadsheetApp.openById(databaseFileId);
    var databaseSheet = databaseSS.getSheetByName(ministrySheetName);
    var databaseData = databaseSheet.getDataRange().getValues();

    Logger.log('--- STARTING UPDATE DATABASE ---');
    Logger.log('Event Data: ' + JSON.stringify(e));

    // Extract form responses using namedValues
    var responses = e.namedValues;
    var name = responses[formNameHeader];
    var timesWilling = responses[formTimesHeader];

    var unavailableDates = [];
    var rawDates = responses[formDatesHeader];

    if (rawDates && rawDates.length > 0) {
      Logger.log('Raw Unavailable Dates from Form: ' + rawDates[0]);

      // 1. Split the single string response (rawDates[0]) into an array of individual date strings.
      unavailableDates = rawDates[0].split(',').map(function(dateStr, index) {

        // Logger.log('Processing item ' + index + ': ' + dateStr.trim());

        // 2. Use regex to remove ' - ' followed by any string to the end of the date item.
        // Also trim any leading/trailing spaces from the split process.
        var cleanStr = dateStr.replace(/\s-\s.*$/, '').trim();

        // Logger.log('  -> Cleaned String: ' + cleanStr);

        // 3. Convert to Date object to ensure format is MM/dd (and handle potential variations)
        var date = new Date(cleanStr);
        if (!isNaN(date.getTime())) {
          var mm = ('0' + (date.getMonth() + 1)).slice(-2);
          var dd = ('0' + date.getDate()).slice(-2);
          var finalDate = mm + '/' + dd;
          // Logger.log('  -> Final MM/dd Date: ' + finalDate);
          return finalDate;
        }

        // Fallback for unparseable strings (should be rare)
        Logger.log('  -> Fallback returned (unparseable): ' + cleanStr);
        return cleanStr;
      });
    }

    var unavailableDatesString = unavailableDates.join(','); // Prepare for storage
    var comments = responses[formCommentsHeader];

    // Log final extracted values before database write
    Logger.log('Name: ' + name);
    Logger.log('Times Willing to Serve: ' + timesWilling);
    Logger.log('Parsed Unavailable Dates: ' + unavailableDatesString);
    Logger.log('Comments: ' + comments);
    
    // --- Database Update Logic ---

    var found = false;
    for (var i = 1; i < databaseData.length; i++) {
      if (databaseData[i][0] == name) {
        // Update the corresponding row
        databaseSheet.getRange(i + 1, 3).setValue(timesWilling);
        databaseSheet.getRange(i + 1, 4).setValue(unavailableDatesString); // Use the joined string
        databaseSheet.getRange(i + 1, 5).setValue(comments);
        found = true;
        Logger.log('Updated existing row ' + (i + 1) + ' for ' + name);
        break;
      }
    }

    if (!found) {
      // If no match is found, append a new row
      var lastRow = databaseSheet.getLastRow() + 1;
      databaseSheet.getRange(lastRow, 1).setValue(name);
      databaseSheet.getRange(lastRow, 3).setValue(timesWilling);
      databaseSheet.getRange(lastRow, 4).setValue(unavailableDatesString);
      databaseSheet.getRange(lastRow, 5).setValue(comments);
      Logger.log('Added new row ' + lastRow + ' for ' + name);
    }
  } catch (error) {
    Logger.log('!!! ERROR in updateDatabase: ' + error.message);
  }
  updateAvailability();
  Logger.log('Updated availability and finished execution.');
}

function getServiceDates(year, month) {
  var serviceDates = [];
  
  // Get the first day of the month
  var firstDay = new Date(year, month, 1);
  
  // Find the first Friday of the month
  var firstFriday = new Date(firstDay);
  while (firstFriday.getDay() !== 5) { // 5 represents Friday
    firstFriday.setDate(firstFriday.getDate() + 1);
  }
  serviceDates.push(Utilities.formatDate(firstFriday, Session.getScriptTimeZone(), 'MM/dd') + ' - Corporate Prayer');
  
  // Iterate through the days of the month to find all Sundays
  var currentDate = new Date(firstDay);
  while (currentDate.getMonth() === month) {
    if (currentDate.getDay() === 0) { // 0 represents Sunday
      serviceDates.push(Utilities.formatDate(currentDate, Session.getScriptTimeZone(), 'MM/dd'));
    }
    currentDate.setDate(currentDate.getDate() + 1);
  }
  return serviceDates;
}

function testForm() {
  var today = new Date();

  var planDate = new Date(today);
  planDate.setMonth(today.getMonth() + 1);

  var oldDate = new Date(today);
  oldDate.setMonth(today.getMonth() - 1);

  var todayMonthName = today.toLocaleString('default', { month: 'long' });
  var todayMonth = today.getMonth();
  var todayYear = today.getFullYear();

  var planMonthName = planDate.toLocaleString('default', { month: 'long' });
  var planMonth = planDate.getMonth();
  var planYear = planDate.getFullYear();

  var oldMonthName = oldDate.toLocaleString('default', { month: 'long' });
  var oldMonth = oldDate.getMonth();
  var oldYear = oldDate.getFullYear();

  createNewFormForMonth(planMonth, planYear, planMonthName);
}

//TODO: Include metadata processing in this function for modularity
function createNewFormForMonth(month, year, monthName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var metadataSheet = ss.getSheetByName("Form Metadata") || ss.insertSheet("Form Metadata");

  // Create a new form for the upcoming month
  var formTitle = "Music Ministry Availability - " + monthName;
  var form = FormApp.create(formTitle);

  // Name Dropdown (ListItem)
  var nameDropdown = form.addListItem().setTitle(formNameHeader).setRequired(true);
  nameDropdown.setChoiceValues(["Loading..."]);

  // const rolesMC = form.addCheckboxItem();
  // rolesMC.setTitle("Select Your Roles").setChoices([
  //   rolesMC.createChoice('WL'),
  //   rolesMC.createChoice('Singer'),
  //   rolesMC.createChoice('Acoustic'),
  //   rolesMC.createChoice('Keyboard'),
  //   rolesMC.createChoice('EG'),
  //   rolesMC.createChoice('Bass'),
  //   rolesMC.createChoice('Drums')
  // ]);

  // Number of Times Willing to Serve
  // form.addTextItem().setTitle("Number of Times Willing to Serve").setRequired(true);
  var numDropdown = form.addListItem()
  .setTitle(formTimesHeader)
  .setChoiceValues(['1', '2', '3', '4', '5']) // Set the dropdown options
  .setRequired(true); // Make the question required
  
  // Add next month's form metadata
  var lastRow = metadataSheet.getLastRow();
  metadataSheet.getRange(lastRow + 1, 1).setValue(monthName + " Form");
  metadataSheet.getRange(lastRow + 1, 2).setValue(form.getId());

  // Update the dropdown with real names
  updateFormDropdown();

  // Add the service dates to the form for unavailable dates selection
  var serviceDates = getServiceDates(year, month);
  var dateChoices = serviceDates;

  const availMC = form.addCheckboxItem();
  availMC.setTitle(formDatesHeader)
    .setChoices(dateChoices.map(date => availMC.createChoice(date)));

  // Optional comments section
  form.addTextItem().setTitle(formCommentsHeader).setRequired(false);

  // Link the form responses to a new sheet in the current spreadsheet
  form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());  // Link the form to the new response sheet
  Logger.log("Linked form responses to new sheet");

  // Get the links for the edit and responder URLs
  var editUrl = form.getEditUrl(); // Edit link for the form owner
  var responderUrl = form.getPublishedUrl(); // Responder link for the participants

  // Send email notification about the new form
  var emailSubject = "New Music Ministry Availability Form Created";
  var emailBody = "A new Music Ministry Availability Form has been created for the month of " + monthName + ".\n\n" +
                  "You can access and fill out the form using the following link:\n" + responderUrl + "\n\n" +
                  "If you need to edit the form, use the following link:\n" + editUrl + "\n\n" +
                  "Please submit your availability as soon as possible.";
  var recipientEmail = "widodoalfianto94@gmail.com"; // Replace with your email address

  // Send email
  MailApp.sendEmail(recipientEmail, emailSubject, emailBody);

  var file = DriveApp.getFileById(form.getId());
  var targetFolder = DriveApp.getFolderById(formsFolderId);
  file.moveTo(targetFolder);
}

function updateFormDropdown() {
  var ss = SpreadsheetApp.openById(databaseFileId);
  var databaseSheet = ss.getSheetByName(ministrySheetName);
  var metadataSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Metadata");

  if (!metadataSheet) {
    Logger.log("Form Metadata sheet missing.");
    return;
  }

  var formId = metadataSheet.getRange("B2").getValue();
  if (!formId) {
    Logger.log("No Form ID found.");
    return;
  }

  // Retrieve the list of names from the "Ministry Members" sheet
  var names = databaseSheet.getRange("A2:A" + databaseSheet.getLastRow()).getValues();
  names = names.flat().filter(String); // Flatten the array and remove any empty strings

  // Open the form using the Form ID
  var form = FormApp.openById(formId);

  // Locate the dropdown question by its title
  var items = form.getItems(FormApp.ItemType.LIST);
  var dropdownTitle = formNameHeader; // Adjust this to match your question title
  var dropdownItem = null;

  for (var i = 0; i < items.length; i++) {
    if (items[i].getTitle() === dropdownTitle) {
      dropdownItem = items[i].asListItem();
      break;
    }
  }

  if (dropdownItem) {
    // Update the dropdown choices
    dropdownItem.setChoiceValues(names);
    Logger.log("Dropdown updated with names from the sheet.");
  } else {
    Logger.log("Dropdown question not found.");
  }
}

function setupAvailability(sheetName, year, month) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName); // If the sheet doesn't exist, create it
  } else {
    sheet.clear(); // Clear the existing sheet if it exists
  }

  // Get next month's Sundays dynamically
  var serviceDates = getServiceDates(year, month);

  var headerRow = ["Schedule"].concat(serviceDates);
  sheet.appendRow(headerRow); // Adding the header row to the sheet

  // Select the header row range and make it bold
  var headerRange = sheet.getRange(1, 1, 1, headerRow.length);
  headerRange.setFontWeight("bold"); // Make the header text bold

  // Define the roles (without any members for now)
  var roles = ["WL", "SINGER", "ACOUSTIC", "KEYBOARD", "EG", "BASS", "DRUMS"];

  // Add each role with empty cells under each Sunday
  roles.forEach(function (role) {
    var roleRow = [role];
    serviceDates.forEach(function () {
      roleRow.push(""); // Adding empty cells for each Sunday
    });
    sheet.appendRow(roleRow); // Add the row for the role
  });

  // Apply bold formatting to all the rows with roles
  var lastRow = sheet.getLastRow();
  sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).setFontWeight("bold"); // Make role rows bold

  // Add 3 empty rows of space before the availability section
  var insertionRow = lastRow + 1;
  sheet.insertRowsAfter(insertionRow, 3);

  // Add "Availability" above the role section
  sheet.getRange(insertionRow + 3, 1).setValue("Availability").setFontWeight("bold");
  var availabilityRange = sheet.getRange(sheet.getLastRow(), 1, 1, sheet.getLastColumn());
  availabilityRange.setFontWeight("bold"); // Make the "Availability" text bold

  // Auto-resize the columns to fit the content
  sheet.autoResizeColumns(1, sheet.getLastColumn());

  // Set up empty data below the "Availability" section for each role
  var emptyData = [
    ["WL", "", "", "", ""],  // Example for WL role
    ["SINGER", "", "", "", ""],  // Example for SINGER role
    ["ACOUSTIC", "", "", "", ""],  // Example for ACOUSTIC role
    ["KEYBOARD", "", "", "", ""],  // Example for KEYBOARD role
    ["EG", "", "", "", ""],  // Example for ELECTRIC/SYNTH role
    ["BASS", "", "", "", ""],  // Example for BASS role
    ["DRUMS", "", "", "", ""],  // Example for DRUMS/CAJON role
  ];

  // Add empty data under the "Availability" heading for each role
  emptyData.forEach(function (dataRow) {
    sheet.appendRow(dataRow); // Add the empty data row for the role
  });
}

function clearByHeader(header) {
    // Open the spreadsheet by its ID
  var ss = SpreadsheetApp.openById(databaseFileId);
  
  // Access the "Ministry Members" sheet
  var sheet = ss.getSheetByName(ministrySheetName);
  
  if (!sheet) {
    Logger.log("Sheet not found: " + ministrySheetName);
    return;
  }
  
  // Get the data range of the sheet
  var dataRange = sheet.getDataRange();
  
  // Get the values in the first row to find the "Not Available Dates" column
  var headers = dataRange.getValues()[0];
  
  // Find the index of the provided header column
  var colIndex = headers.indexOf(header) + 1; // +1 to convert to 1-based index
  
  if (colIndex === 0) {
    Logger.log(header + ' column not found.');
    return;
  }
  
  // Determine the range to clear: from row 2 to the last row in the identified column
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    Logger.log("No data to clear.");
    return;
  }
  
  // Clear the contents of the column, starting from row 2
  sheet.getRange(2, colIndex, lastRow - 1).clearContent();
  
  Logger.log(header + ' column cleared.');
}

function monthlySetup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var metadataSheet = ss.getSheetByName("Form Metadata") || ss.insertSheet("Form Metadata");

  var today = new Date();

  var planDate = new Date(today);
  planDate.setMonth(today.getMonth() + 1);

  var oldDate = new Date(today);
  oldDate.setMonth(today.getMonth() - 1);

  var todayMonthName = today.toLocaleString('default', { month: 'long' });
  var todayMonth = today.getMonth();
  var todayYear = today.getFullYear();

  var planMonthName = planDate.toLocaleString('default', { month: 'long' });
  var planMonth = planDate.getMonth();
  var planYear = planDate.getFullYear();

  var oldMonthName = oldDate.toLocaleString('default', { month: 'long' });
  var oldMonth = oldDate.getMonth();
  var oldYear = oldDate.getFullYear();

  var newTabName = `${planMonthName} Availability`;
  var deleteTabName = `${oldMonthName} Availability`;
  setupAvailability(newTabName, planYear, planMonth);

  clearByHeader(sheetsTimesHeader);
  clearByHeader(sheetsDatesHeader);
  clearByHeader(sheetsCommentsHeader);

  if (!ss.getSheetByName(newTabName)) {
    ss.insertSheet(newTabName);
    Logger.log("Created new tab: " + newTabName);
  }

  var oldSheet = ss.getSheetByName(deleteTabName);
  if (oldSheet) {
    ss.deleteSheet(oldSheet);
    Logger.log("Deleted old tab: " + deleteTabName);
  }

  // Store the form ID in the "Form Metadata" sheet in the specified structure
  // Clear any old form metadata if we have more than 2 entries
  var lastRow = metadataSheet.getLastRow();
  if (lastRow > 1) {
    var currentMonthFormLabel = metadataSheet.getRange(2, 1).getValue();  // Get the current month's form label
    var currentMonthFormId = metadataSheet.getRange(2, 2).getValue();  // Get the current month's form ID

    // Move the current month's form metadata to row 1
    metadataSheet.getRange(1, 1).setValue(currentMonthFormLabel);  // Move label to row 1
    metadataSheet.getRange(1, 2).setValue(currentMonthFormId);  // Move form ID to row 1
    metadataSheet.deleteRow(2);
  }

  if (metadataSheet) {
    var oldFormId = metadataSheet.getRange("B1").getValue(); // Get the old form ID from metadata
    if (oldFormId) {
      try {
        var oldForm = FormApp.openById(oldFormId); // Open the form using the ID
        oldForm.removeDestination(); // Remove the link to the spreadsheet
        Logger.log("De-linked form with ID: " + oldFormId);
      } catch (e) {
        Logger.log("Could not de-link or find the old form: " + e.message);
      }
    }
  }

  var sheets = ss.getSheets();
  sheets.forEach(function(sheet) {
    if (sheet.getName().startsWith("Form Responses")) {
      toDelete = sheet.getName();
      ss.deleteSheet(sheet);
      Logger.log("Deleted old Form Responses tab: " + toDelete);
    }
  })
  createNewFormForMonth(planMonth, planYear, planMonthName);
  Logger.log(`Created new form for ${planMonthName}`);
}

function findFormResponseSheet() {
  // Open the active spreadsheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Get all sheets in the spreadsheet
  var sheets = ss.getSheets();
  
  // Iterate through each sheet to find the form response sheet
  for (var i = 0; i < sheets.length; i++) {
    var sheet = sheets[i];
    var sheetName = sheet.getName();
    
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
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  Logger.log('--- STARTING updateAvailability ---');

  var today = new Date();
  var planDate = new Date(today);
  planDate.setMonth(today.getMonth() + 1);
  var planMonthName = planDate.toLocaleString('default', { month: 'long' });
  var sheetName = planMonthName + " Availability";

  var matrixSheet = ss.getSheetByName(sheetName);
  var databaseSheet = ss.getSheetByName(ministrySheetName);

  if (!matrixSheet || !databaseSheet) {
    Logger.log("Error: One or more required sheets are missing.");
    return;
  }

  var databaseData = databaseSheet.getDataRange().getValues();

  if (!databaseData.length) {
    Logger.log("No data found in the Ministry Members sheet.");
    return;
  }

  // Get Date Headers from the sheet (Row 1, starting from column 2)
  var lastCol = matrixSheet.getLastColumn();
  if (lastCol <= 1) {
    Logger.log("Error: Availability matrix has no date columns.");
    return;
  }
  var headerRowValues = matrixSheet.getRange(1, 2, 1, lastCol - 1).getValues();
  var dateHeaders = headerRowValues[0];
  Logger.log('Matrix Date Headers (Raw): ' + dateHeaders.join(', '));

  var lastCol = matrixSheet.getLastColumn();

  var dateHeaders = matrixSheet
    .getRange(dateRowIndex, 2, 1, lastCol - 1)
    .getDisplayValues()[0];

  var serviceDateKeys = dateHeaders.map(function(h) {
    return String(h).trim().substring(0, 5);
  });

  Logger.log('Standardized Date Keys: ' + serviceDateKeys.join(', '));

  // Initialize the availability object
  var availability = {};
  var roleOrder = ["WL", "SINGER", "ACOUSTIC", "KEYBOARD", "EG", "BASS", "DRUMS"];

  // Standardize roleOrder to uppercase for case-insensitive matching
  roleOrder = roleOrder.map(function(role) { return role.toUpperCase(); });

  // Process each row in the Ministry Members sheet
  for (var i = 1; i < databaseData.length; i++) {
    var row = databaseData[i];
    var name = row[0] ? row[0].trim() : "";
    
    if (!name) continue; 
    
    var roles = row[1]
      ? row[1].toString().split(",").map(function(role) {
          return role.trim().toUpperCase();
        })
      : [];
    var timesWilling = row[2] ? row[2].toString().trim() : "";
    var rawUnavailableDates = row[3] ? row[3].toString() : "";
    
    var unavailableDates = rawUnavailableDates
      ? rawUnavailableDates.split(",").map(function(dateStr) {
      var trimmedDateStr = dateStr.trim();
      var parsedDate = new Date(trimmedDateStr);
      
      if (!isNaN(parsedDate.getTime())) {
        var mm = ('0' + (parsedDate.getMonth() + 1)).slice(-2);
        var dd = ('0' + parsedDate.getDate()).slice(-2);
        return mm + '/' + dd;
      } else {
        // Fallback: just take the first 5 characters (e.g., "MM/dd") if parse fails
        return trimmedDateStr.substring(0, 5);
      }
    })
  : [];

    if (!name || !roles.length) continue;

    // Format the name as "Firstname L."
    var nameParts = name.split(" ");
    if (nameParts.length > 1) {
      name = nameParts[0] + " " + nameParts[1].charAt(0).toUpperCase() + ".";
    }

    // If "Times Willing to Serve" is blank, mark unavailable for all dates
    var isUnavailableAllMonth = timesWilling === "";

    roles.forEach(function(role) {
      if (!availability[role]) availability[role] = {};
      
      // *** Use serviceDateKeys for reliable matching ***
      serviceDateKeys.forEach(function(dateKey) {
        
        // dateKey is the clean "MM/dd" string (e.g., "12/14")
        var date = dateKey; 
        
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
  var numRoles = roleOrder.length;
  // Use serviceDateKeys.length to define the range width
  var clearRange = matrixSheet.getRange(headerRowIndex, 2, numRoles, serviceDateKeys.length); 
  clearRange.clearContent();

  // Update the availability matrix in the sheet
  var roleRowIndex = headerRowIndex;
  roleOrder.forEach(function(role) {
    var roleData = availability[role];
    if (roleData) {
      var namesRow = serviceDateKeys.map(function(dateKey) {
        // dateKey is the clean "MM/dd" string, used as the lookup key
        return roleData[dateKey] ? roleData[dateKey].join("\n") : "";
      });
      
      // Set values in the sheet
      var range = matrixSheet.getRange(roleRowIndex, 2, 1, namesRow.length);
      range.setValues([namesRow]);
      range.setWrap(false); // Disable text wrapping for the range
      roleRowIndex++;
    }
  });

  matrixSheet.autoResizeColumns(1, matrixSheet.getLastColumn() - 1);
  matrixSheet.autoResizeRows(headerRowIndex, roleRowIndex - headerRowIndex + 1);
  Logger.log("Availability matrix updated in sheet: " + sheetName);
  Logger.log('--- FINISHED updateAvailability ---');
}