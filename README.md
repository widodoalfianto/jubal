# 📅 ministryScheduler

This repository contains the Google Apps Script and Google Sheets logic used for **monthly ministry scheduling**.

The script is currently configured to:
- **Create a new availability form and schedule sheet** on the **8th day of every month**
- Triggered automatically by a **time-based Apps Script trigger**

## 🏗️ Project Structure

- **`MinistryScheduler.js`**: The core logic controller. Handles form submissions, database updates, and sheet generation.
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
3. Copy the contents of `MinistryScheduler.js` and `Config.js` into the script editor (ensure you create the separate files).

### 2. Configuration (`Config.js`)
Update the `CONFIG` object with your specific details:
- **`ids.formsFolder`**: The ID of a Google Drive folder where monthly forms will be stored.
- **`ids.adminEmails`**: A list of email addresses that should receive notifications.
- **`roles`**: Update the array with the specific roles in your ministry (e.g., "WL", "Keys", "Media").

### 3. Initialization
1. In the Apps Script editor, select the `initializeProject` function from the dropdown menu.
2. Click **Run**. This will create the necessary sheets ("Ministry Members", "Form Metadata") and populate them with dummy data.
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