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
   - `Recurring`
   - `Events`
3. **Important**: Delete the dummy row in "Ministry Members" and add your actual team members.
4. `initializeProject()` now also tries to switch role entry to checkboxes when it is safe.
   It adds role checkbox columns starting in column `G` and makes the `Roles` column auto-generate from those checkboxes.
5. If `initializeProject()` skips that migration because columns `G+` already contain data, you can review the sheet and run `migrateMemberRolesToCheckboxes()` manually.

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

## Recurring

The default schedule is now just **all Sundays** for the target month.

Use the `Recurring` sheet to define repeatable patterns for your church. This is the preferred sheet for:
- weekly recurring events, such as Sunday services
- monthly recurring events, such as first-Friday `Corporate Prayer`
- yearly recurring events, such as `Christmas`
- movable yearly events, such as `Easter`

Most churches only need these columns:
- **Enabled**
- **Event**
- **Frequency**: `Weekly`, `Monthly`, `Yearly`, or `Easter`
- **Weekday**
- **Week Of Month**
- **Month**
- **Day**
- **Include In Form**
- **Include In Schedule**
- **Notes**

Examples:
- Sundays every week:
  - `TRUE |  | Weekly | Sunday | every | all |  | TRUE | TRUE | Leave Event blank to show plain Sunday dates`
- First Friday Corporate Prayer:
  - `TRUE | Corporate Prayer | Monthly | Friday | 1 | all |  | TRUE | TRUE |`
- Easter:
  - `TRUE | Easter | Easter |  |  | all |  | TRUE | TRUE |`
- Christmas:
  - `TRUE | Christmas | Yearly |  |  | 12 | 25 | TRUE | TRUE |`

Legacy note:
- Existing `Events` sheets from earlier versions may contain recurring rules. Running `initializeProject()` will rename them to `Recurring` when it is safe to do so.
- Existing `Recurring Events` sheets are still read for backward compatibility.

## Events

Use the `Events` sheet for one-off additions or removals in a specific month. This is the sheet you edit when “this month is different from normal.”

Recommended columns:
- **Enabled**
- **Date**: use a real date cell in `YYYY-MM-DD` format
- **Event**: the event name shown on the form and schedule
- **Action**: `ADD` or `REMOVE`
- **Recurring Event**: optional. Use the same event name from `Recurring` when moving or cancelling a recurring event
- **Include In Form**
- **Include In Schedule**
- **Notes**

Behavior:
- The script starts from:
  - Sundays by default
  - plus any enabled recurring rules from `Recurring`
- Then `Events` can remove or add one-off events for that specific month.
- Past one-off events can be moved to `Events Archive` automatically during `monthlySetup()` based on the `events_archive_frequency` and `events_archive_month` settings.
- Built-in example rows stay in `Events` as templates and are not archived.

Admin workflows:
- To move a recurring event (e.g. move Corporate Prayer to the 2nd Friday), add:
  - one `REMOVE` row for the original date
  - one `ADD` row for the new date
- To add special events (Good Friday, retreat nights, Christmas Eve), add rows in `Events` for the relevant dates.
- After editing the Availability sheet header manually, run `syncCurrentFormWithAvailability()` from the Apps Script editor to update the live form's date choices.

Useful settings:
- `events_archive_frequency`: `Off`, `Monthly`, `Quarterly`, or `Yearly`
- `events_archive_month`: month to run yearly archiving, such as `January`

This makes the Availability sheet header the authoritative source for the form choices and scheduling matrix, so manual edits are supported.
