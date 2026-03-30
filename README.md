# 🎵 Jubal

This repository contains the Google Apps Script and Google Sheets logic used for **monthly ministry scheduling**.

The script is currently configured to:
- **Create a new availability form and schedule sheet** on the **8th day of every month**
- Triggered automatically by a **time-based Apps Script trigger**

## 🏗️ Project Structure

- **`Jubal.js`**: The core logic controller. Handles form submissions, database updates, and sheet generation.
- **`Config.js`**: Centralized configuration file for IDs, role definitions, and email settings.
- **`Testing.js`**: End-to-end test suite to verify the full lifecycle of the application (excluded from production).

## 💻 Technologies Used

- **Google Apps Script**: Server-less JavaScript platform for automation.
- **Google Sheets API**: For database management and schedule visualization.
- **Google Forms API**: For collecting volunteer availability.
- **Modern JavaScript (ES6+)**: Utilizes `const`/`let`, arrow functions, and template literals for clean, maintainable code.

---

## 🛠️ Adoption & Setup

To use this workflow for your own ministry, follow these steps:

### 1. Installation
1. Create a new **Google Sheet**.
2. Open **Extensions > Apps Script**.
3. Copy the contents of `Jubal.js` and `Config.js` into the script editor (ensure you create the separate files).

### 2. Configuration (`Config.js`)
Update the `CONFIG` object with your specific details:
- **`ids.formsFolder`**: The ID of a Google Drive folder where monthly forms will be stored.
- **`ids.adminEmails`**: A list of email addresses that should receive notifications.
- **`roles`**: Update the array with the specific roles in your ministry (e.g., "WL", "Keys", "Media").

### 3. Initialization
1. In the Apps Script editor, select the `initializeProject` function from the dropdown menu.
2. Click **Run**. This will create the necessary sheets, including:
   - `Ministry Members`
   - `Form Metadata`
   - `Settings`
   - `Events`
   - `Monthly Events`
3. **Important**: Delete the dummy row in "Ministry Members" and add your actual team members.

### 4. Triggers
Set up the automation triggers in the Apps Script dashboard (Clock icon on the left):
1. **Monthly Form Creation**:
   - Function: `monthlySetup` | Event Source: `Time-driven` | Type: `Month timer` (e.g., 8th of the month).
2. **Database Update**:
   - Function: `onFormSubmit` | Event Source: `From spreadsheet` | Event Type: `On form submit`.

---

## � Improvements Roadmap

We welcome contributions! If you're interested in helping improve this project, here are some ideas currently on the roadmap:

1. ✅ **Email alerts when Apps Script encounters errors**
2. ✨ **Enhance the current availability sheet with:**
   - Dropdowns for member names
   - Warnings when the number of assignments exceeds willingness
3. 📧 **Email notifications to ministers**
   - Including their personal schedule and the full monthly schedule
4. 🧹 **Code clean-up and refactoring**
5. 🔁 **Integrate a CI/CD pipeline** for version control and deployment

---

## 📝 Notes

If you'd like to collaborate on the Google Sheets file directly, please email:

📬 **[widodoalfianto94@gmail.com](mailto:widodoalfianto94@gmail.com)** for edit access.

---

## 🙌 Contributions

Pull requests are welcome! Feel free to fork the repo, make changes, and open a PR.

---

## Events

The default schedule is now just **all Sundays** for the target month.

Use the `Events` sheet to add optional recurring events that differ by church. This is the preferred configuration sheet for:
- monthly recurring events, such as first-Friday `Corporate Prayer`
- yearly recurring events, such as `Christmas`
- movable yearly events, such as `Easter`

Recommended columns:
- **Enabled**
- **Rule ID**
- **Label**
- **Recurrence**: `monthly` or `yearly`
- **Rule Type**: `every_weekday`, `nth_weekday`, `fixed_date`, or `easter_offset`
- **Month**: `all`, a month number, or a month name
- **Weekday**
- **Ordinal**
- **Day Of Month**
- **Offset Days**
- **Include In Form**
- **Include In Schedule**
- **Sort Order**
- **Type**
- **Notes**

Examples:
- Sundays every month:
  - `TRUE | sunday_service |  | monthly | every_weekday | all | Sunday | every |  | 0 | TRUE | TRUE | 20 | service |`
- First Friday Corporate Prayer:
  - `TRUE | corporate_prayer | Corporate Prayer | monthly | nth_weekday | all | Friday | 1 |  | 0 | TRUE | TRUE | 10 | prayer |`
- Easter:
  - `TRUE | easter | Easter | yearly | easter_offset | all |  |  |  | 0 | TRUE | TRUE | 30 | special |`
- Christmas:
  - `TRUE | christmas | Christmas | yearly | fixed_date | 12 |  |  | 25 | 0 | TRUE | TRUE | 40 | special |`

Legacy note:
- Existing `Recurring Events` sheets are still read for backward compatibility.

## Monthly Events & Overrides

You can override or add special events for any month using the `Monthly Events` sheet (create it in the spreadsheet). Columns:
- **Year**: numeric year (e.g. `2026`)
- **Month**: numeric month (`1`-`12`) or month name (`March`)
- **Date**: a date for the event (supports `MM/DD`, `YYYY-MM-DD`, `Mar 29`, etc.)
- **Action**: `ADD` or `REMOVE`
- **Label**: short label (e.g. `Easter`, `Christmas`, `Corporate Prayer`)
- **Rule ID**: optional stable identifier for matching an existing recurring event
- **Type**: optional free-text field

Behavior:
- The script starts from:
  - Sundays by default
  - plus any enabled recurring rules from `Events`
- Then `Monthly Events` can remove or add one-off events for that specific month.

Admin workflows:
- To move a recurring event (e.g. move Corporate Prayer to the 2nd Friday), add:
  - one `REMOVE` row for the original date
  - one `ADD` row for the new date
- To add special events (Easter, Christmas), add rows in `Monthly Events` for the relevant months.
- After editing the Availability sheet header manually, run `syncCurrentFormWithAvailability()` from the Apps Script editor to update the live form's date choices.

This makes the Availability sheet header the authoritative source for the form choices and scheduling matrix, so manual edits are supported.
