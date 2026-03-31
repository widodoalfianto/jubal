/**
 * Email and reminder helpers.
 */

function formatHumanDateTime(value, timeZone) {
  if (!value) return '';
  const parsed = Object.prototype.toString.call(value) === '[object Date]' ? value : new Date(value);
  if (isNaN(parsed.getTime())) return String(value);
  return Utilities.formatDate(parsed, timeZone || safeGetScriptTimeZone(), "MMMM d, yyyy 'at' h:mm a z");
}

function formatDayOfMonthHuman(day) {
  const parsed = toIntegerOrDefault(day, 0);
  if (parsed <= 0) return '';
  const mod100 = parsed % 100;
  if (mod100 >= 11 && mod100 <= 13) return `${parsed}th`;
  const mod10 = parsed % 10;
  if (mod10 === 1) return `${parsed}st`;
  if (mod10 === 2) return `${parsed}nd`;
  if (mod10 === 3) return `${parsed}rd`;
  return `${parsed}th`;
}

function getFriendlyReconciliationAction(actionRequired, name) {
  const normalized = String(actionRequired || '').trim();
  if (!normalized) return 'Please review this item.';
  if (normalized.indexOf('New member added') === 0) {
    return `${name || 'This person'} was added as a new row because the submission did not match anyone already listed in Ministry Members. Please check whether this is a brand new person or a duplicate/spelling variation.`;
  }
  return normalized;
}

function buildReconciliationAlertPayload(rowsToAlert, sheet, settings) {
  const runtimeSettings = settings || loadRuntimeSettings();
  const sheetLink = getSheetRangeUrl(sheet, 1, 'A', 'G');
  const membersLink = getSheetUrlByName(CONFIG.sheetNames.ministryMembers);
  const items = rowsToAlert.map((item, index) => {
    const row = item.row;
    const submittedAt = formatHumanDateTime(row[0], runtimeSettings.timeZone);
    const name = String(row[2] || '').trim() || 'Unknown name';
    const parseErrors = String(row[4] || '').trim();
    const action = getFriendlyReconciliationAction(row[5], name);
    const rowLink = getSheetRangeUrl(sheet, item.idx, 'A', 'G');

    const lines = [
      `${index + 1}. ${name}`,
      `Submitted: ${submittedAt}`,
      `What needs review: ${action}`
    ];
    if (parseErrors) lines.push(`Dates we could not understand: ${parseErrors}`);
    lines.push(`Open this item: ${rowLink}`);
    return lines.join('\n');
  });

  const itemLabel = rowsToAlert.length === 1 ? 'item' : 'items';
  return {
    subject: `${runtimeSettings.churchName}: ${rowsToAlert.length} reconciliation ${itemLabel} need${rowsToAlert.length === 1 ? 's' : ''} review`,
    body: [
      `Church: ${runtimeSettings.churchName}`,
      '',
      `Please review the following ${itemLabel} in Jubal:`,
      '',
      items.join('\n\n'),
      '',
      `Reconciliation sheet: ${sheetLink}`,
      `Ministry Members sheet: ${membersLink}`
    ].join('\n')
  };
}

function addReconciliationEntry(e, submittedName, matchedCanonical, parseErrors, actionRequired) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(CONFIG.sheetNames.reconciliation);
    if (!sheet) {
      sheet = ss.insertSheet(CONFIG.sheetNames.reconciliation);
      sheet.appendRow(['timestamp', 'formId', 'submittedName', 'matchedCanonical', 'parseErrors', 'actionRequired', 'alerted']);
      applySheetTheme(sheet);
    }

    const ts = new Date().toISOString();
    const meta = ss.getSheetByName(CONFIG.sheetNames.formMetadata);
    let formId = '';
    if (meta) formId = (meta.getRange('B2').getValue() || meta.getRange('B1').getValue() || '').toString();

    const errorsStr = Array.isArray(parseErrors) ? parseErrors.join('; ') : (parseErrors || '');
    const rowIndex = sheet.getLastRow() + 1;
    sheet.appendRow([ts, formId, submittedName || '', matchedCanonical || '', errorsStr, actionRequired || '', '']);
    fitSheetToContent(sheet);
    applyTableBordersToDataRange(sheet);

    sendReconciliationAlert(sheet, rowIndex);
  } catch (err) {
    console.error('addReconciliationEntry failed: ' + err.message);
  }
}

function sendReconciliationAlert(sheet, rowIndex) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!sheet) sheet = ss.getSheetByName(CONFIG.sheetNames.reconciliation);
    if (!sheet) return;

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return;

    const startRow = rowIndex ? rowIndex : 2;
    const range = sheet.getRange(startRow, 1, lastRow - startRow + 1, sheet.getLastColumn());
    const values = range.getValues();

    const rowsToAlert = [];
    for (let i = 0; i < values.length; i++) {
      const row = values[i];
      const alerted = (row[6] || '').toString().trim();
      if (!alerted) {
        rowsToAlert.push({ idx: startRow + i, row });
      }
    }

    if (!rowsToAlert.length) return;

    const email = buildReconciliationAlertPayload(rowsToAlert, sheet, loadRuntimeSettings());
    const sent = sendEmailToAdmins(email.subject, email.body);
    if (!sent) return;

    rowsToAlert.forEach(r => {
      try {
        sheet.getRange(r.idx, 7).setValue(new Date().toISOString());
      } catch (inner) {
        console.error('Failed to mark alerted for row ' + r.idx + ': ' + inner.message);
      }
    });
  } catch (err) {
    console.error('sendReconciliationAlert failed: ' + err.message);
  }
}

function sendNewFormCreatedEmail(monthName, responderUrl, editUrl, settings) {
  const runtimeSettings = settings || loadRuntimeSettings();
  const emailSubject = `${runtimeSettings.churchName}: New availability form created for ${monthName}`;
  const emailBody = `Church: ${runtimeSettings.churchName}\n\n` +
                  'A new Music Ministry Availability Form has been created for the month of ' + monthName + '.\n\n' +
                  'You can access and fill out the form using the following link:\n' + responderUrl + '\n\n' +
                  'If you need to edit the form, use the following link:\n' + editUrl + '\n\n' +
                  'Please submit your availability as soon as possible.';

  sendEmailToAdmins(emailSubject, emailBody, runtimeSettings);
}

function getAdminReminderPropertyKey(referenceDate) {
  const date = referenceDate ? new Date(referenceDate) : new Date();
  const planDate = new Date(date);
  planDate.setMonth(planDate.getMonth() + 1);
  return `adminReminderSent:${planDate.getFullYear()}-${('0' + (planDate.getMonth() + 1)).slice(-2)}`;
}

function getMonthlySetupPropertyKey(referenceDate) {
  const date = referenceDate ? new Date(referenceDate) : new Date();
  const planDate = new Date(date);
  planDate.setMonth(planDate.getMonth() + 1);
  return `monthlySetupCompleted:${planDate.getFullYear()}-${('0' + (planDate.getMonth() + 1)).slice(-2)}`;
}

function shouldRunMonthlySetupToday(referenceDate, settings) {
  const runtimeSettings = settings || loadRuntimeSettings();
  const today = referenceDate ? new Date(referenceDate) : new Date();
  return today.getDate() === runtimeSettings.formCreationDay;
}

function buildAdminPlanningReminder(referenceDate, settings) {
  const runtimeSettings = settings || loadRuntimeSettings();
  const today = referenceDate ? new Date(referenceDate) : new Date();
  const planDate = new Date(today);
  planDate.setMonth(planDate.getMonth() + 1);

  const planMonthName = Utilities.formatDate(planDate, runtimeSettings.timeZone, 'MMMM yyyy');
  const monthlySetupDay = formatDayOfMonthHuman(runtimeSettings.formCreationDay);
  const subject = `${runtimeSettings.churchName}: ${planMonthName} Events Reminder`;
  const bodyLines = [
    `Church: ${runtimeSettings.churchName}`,
    '',
    `Please review the schedule setup for ${planMonthName}.`,
    '',
    'Before the monthly form is created, please check these areas:',
    `- Recurring schedule: ${getSheetUrlByName(CONFIG.sheetNames.recurring)}`,
    `- One-time events and changes: ${getSheetUrlByName(CONFIG.sheetNames.events)}`,
    `- Ministry roles: ${getSheetUrlByName(CONFIG.sheetNames.rolesConfig)}`,
    `- Ministry members and role updates: ${getSheetUrlByName(CONFIG.sheetNames.ministryMembers)}`,
    `- Admin contacts and notifications: ${getSheetUrlByName(CONFIG.sheetNames.admins)}`,
    '',
    'Recommended checklist:',
    '- Confirm your normal recurring events are correct.',
    '- Add any special events, cancellations, or moved dates for next month.',
    '- Add or disable roles in the Roles sheet if your ministry roster changed.',
    '- Add any new members or update role checkboxes as needed.',
    '- Add or remove admin recipients in the Admins sheet if notification recipients need to change.',
    '',
    `Monthly setup day: the ${monthlySetupDay} of each month.`,
    'Once that looks right, the next month form and availability sheet can be generated by the monthly setup trigger.',
    `Reminder date: ${formatHumanDateTime(today, runtimeSettings.timeZone)}`
  ];

  return {
    subject: subject,
    body: bodyLines.join('\n')
  };
}

function sendAdminPlanningReminderIfDue() {
  const runtimeSettings = loadRuntimeSettings();
  if (!runtimeSettings.adminReminderEnabled) return { status: 'disabled' };
  if (!getAdminRecipientList(runtimeSettings).length) return { status: 'no_recipients' };

  const today = new Date();
  if (today.getDate() !== runtimeSettings.adminReminderDay) {
    return { status: 'not_due_today', day: today.getDate(), reminderDay: runtimeSettings.adminReminderDay };
  }

  const propertyKey = getAdminReminderPropertyKey(today);
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty(propertyKey)) {
    return { status: 'already_sent', key: propertyKey };
  }

  const reminder = buildAdminPlanningReminder(today, runtimeSettings);
  const sent = sendEmailToAdmins(reminder.subject, reminder.body, runtimeSettings);
  if (!sent) return { status: 'no_recipients' };

  props.setProperty(propertyKey, new Date().toISOString());
  return { status: 'sent', key: propertyKey };
}

function sendAdminPlanningReminderNow() {
  const runtimeSettings = loadRuntimeSettings();
  const reminder = buildAdminPlanningReminder(new Date(), runtimeSettings);
  const sent = sendEmailToAdmins(reminder.subject, reminder.body, runtimeSettings);
  return { status: sent ? 'sent' : 'no_recipients' };
}
