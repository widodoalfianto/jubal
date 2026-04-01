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
- **`ids.formsFolder`**: Optional starter folder ID for new installs. Leave blank if each church will set this in `Settings`.
- **`ids.adminEmails`**: Optional starter admin emails for new installs. Leave blank if each church will manage admins only through the `Admins` sheet.
- **`roles`**: Starter roles used to seed the `Roles` sheet on first setup.

### 3. Initialization
1. In the Apps Script editor, select the `initializeProject` function from the dropdown menu.
2. Click **Run**. This will create the necessary sheets, including:
   - `Ministry Members`
   - `Form Metadata`
   - `Settings`
   - `Admins`
   - `Roles`
   - `Recurring`
   - `Events`
3. Add your actual team members to `Ministry Members` before running the first monthly setup.
4. `initializeProject()` now also tries to switch role entry to checkboxes when it is safe.
   It adds role checkbox columns starting in column `G` and makes the `Roles` column auto-generate from those checkboxes.
5. If `initializeProject()` skips that migration because columns `G+` already contain data, you can review the sheet and run `migrateMemberRolesToCheckboxes()` manually.
6. Use the `Admins` sheet to add or remove notification recipients with checkboxes and email validation.
7. Use the `Roles` sheet to add or disable ministry roles. The system uses that sheet to build the role checkboxes in `Ministry Members`.
8. `Ministry Members` is ordered for daily admin use: `Name`, `Unavailable Dates`, `Times Willing To Serve`, `Comments`, `Roles`, then role checkboxes, with `Canonical Name` kept as the last column.

### 4. Triggers
If you deploy through the GitHub Actions pipeline, the automation triggers are synced automatically after each deploy by running `syncAutomationTriggers()`.

If you are installing or repairing a spreadsheet manually, run `syncAutomationTriggers()` once from the Apps Script editor. It will ensure these trigger types exist:
1. **Daily Automation Scheduler**:
   - Function: `runDailyAutomation` | Event Source: `Time-driven` | Type: `Day timer`
   - This runs daily and calls both:
     - `sendAdminPlanningReminderIfDue()`
     - `monthlySetup()`
   - Both functions still use their own in-code day guards, so the scheduler is safe to run every day.
2. **Database Update**:
   - Function: `onFormSubmit` | Event Source: `From spreadsheet` | Event Type: `On form submit`
3. **Manual Override**:
   - If you need to run the monthly setup immediately from the editor, use `runMonthlySetupNow()`.

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

## 🙌 Contributions

Pull requests are welcome! Feel free to fork the repo, make changes, and open a PR.

---

## Recurring

The default schedule is now just **all Sundays** for the target month.

Use the `Recurring` sheet to define repeatable patterns for your church. This is the preferred sheet for:
- weekly recurring events, such as Sunday services
- monthly recurring events, such as first-Friday `Corporate Prayer`
- simple repeating patterns that happen the same way most months

Most churches only need these columns:
- **Enabled**
- **Event**
- **Frequency**: `Weekly` or `Monthly`
- **Weekday**
- **Week Of Month**
- **Include In Form**
- **Include In Schedule**
- **Notes**

Examples:
- Sundays every week:
  - `TRUE |  | Weekly | Sunday | every | TRUE | TRUE | Leave Event blank to show plain Sunday dates`
- First Friday Corporate Prayer:
  - `TRUE | Corporate Prayer | Monthly | Friday | 1 | TRUE | TRUE |`

Use the `Events` sheet, not `Recurring`, for dated specials like:
- Easter
- Christmas
- Good Friday
- Christmas Eve
- retreats or one-off ministry nights

Legacy note:
- Existing `Events` sheets from earlier versions may contain recurring rules. Running `initializeProject()` will rename them to `Recurring` when it is safe to do so.
- Existing `Recurring Events` sheets are still read for backward compatibility.

## Events

Use the `Events` sheet for one-off additions or removals in a specific month. This is the sheet you edit when “this month is different from normal.”

Best admin workflow:
- use the `Add Special Event` menu option from the spreadsheet menu for one-time additions
- pick the date with the dialog's date field
- let the script add the row to `Events` for you
- if you need to cancel or move a recurring date, edit the `Events` sheet directly and use `REMOVE`

Recommended columns:
- **Enabled**
- **Date**: click into the cell and use the calendar date picker
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
- Easter and Christmas should be entered here as dated event rows when you want them included.
- Past one-off events can be moved to `Events Archive` automatically during `monthlySetup()` based on the `events_archive_frequency` setting.
- Built-in example rows stay in `Events` as templates and are not archived.

Admin workflows:
- To move a recurring event (e.g. move Corporate Prayer to the 2nd Friday), add:
  - one `REMOVE` row for the original date
  - one `ADD` row for the new date
- To add special events (Good Friday, retreat nights, Christmas Eve), add rows in `Events` for the relevant dates.
- After changing `Events` after next month has already been generated, use the `Scheduling` menu in the spreadsheet:
  - `Apply Event Changes to Next Month` rebuilds next month's sheet from `Recurring` and `Events`
  - `Refresh Form Dates` updates the live form choices from the sheet
  - `Refresh Availability Sheet` rebuilds the availability list from `Ministry Members`

## Admins

Use the `Admins` sheet to control who receives reminders and alerts.

Columns:
- **Enabled**: check this when the person should receive notifications
- **Email**: validated email address
- **Notes**: optional admin-facing note

This is the preferred self-service place for managing notification recipients. Admins can simply add a new row, enter the email address, and check **Enabled**.

## Roles

Use the `Roles` sheet to control which ministry roles are active.

Columns:
- **Enabled**: check this when the role should be active
- **Role**: role name shown in schedules and member role checkboxes
- **Notes**: optional admin-facing note

To add a new role, add a new row and check **Enabled**. The next setup/update cycle will sync that role into `Ministry Members`.
The disabled example row stays at the top and is highlighted so admins can quickly see the pattern to follow.

Useful settings:
- `time_zone`: time zone used for event generation and reminder emails
- `admin_reminder_enabled`: `TRUE` or `FALSE`
- `admin_reminder_day`: day of month to remind admins to review next month before the form is created
- `form_creation_day`: the day of month when the daily `monthlySetup` trigger should actually create the next month form and availability sheet
- `times_choices`: the choices shown in the form question for how many times someone is willing to serve
- `events_archive_frequency`: `Off`, `Monthly`, `Quarterly`, or `Yearly`
- `forms_folder_id`: only change this if you want future monthly forms stored in a different Drive folder

Availability tabs are now named only by month, such as `April` or `May`.

The `Events` sheet keeps its example rows highlighted at the top so admins always have a simple pattern to copy when adding one-time events or cancellations.

Archive timing:
- `Monthly`: archives old one-time events on the 1st day of each month
- `Quarterly`: archives on January 1, April 1, July 1, and October 1
- `Yearly`: archives on January 1

This makes the Availability sheet header the authoritative source for the form choices and scheduling matrix, so manual edits are supported.
