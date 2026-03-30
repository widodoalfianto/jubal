/**
 * Configuration for Jubal.
 * Update these values to customize for your organization.
 */
const CONFIG = {
  ids: {
    formsFolder: "1RMITTfCVaYzc0RBBtF0caE8RI5yeq57d",
    adminEmails: ["widodoalfianto94@gmail.com"]
  },
  sheetNames: {
    ministryMembers: "Ministry Members",
    formMetadata: "Form Metadata",
    reconciliation: "Reconciliation",
    settings: "Settings",
    admins: "Admins",
    rolesConfig: "Roles",
    recurring: "Recurring",
    events: "Events",
    eventsArchive: "Events Archive",
    recurringEvents: "Recurring Events",
    monthlyEvents: "Monthly Events"
  },
  formHeaders: {
    name: 'Select your name',
    times: 'How many times are you willing to serve this month?',
    dates: 'Which days are you NOT available? If re-submitting, please re-submit this section also',
    comments: 'Comments(optional)'
  },
  sheetHeaders: {
    name: 'Name',
    roles: 'Roles',
    times: 'Times Willing to Serve',
    dates: 'Unavailable Dates',
    comments: 'Comments',
    canonicalName: 'Canonical Name'
  },
  roles: ["WL", "SINGER", "ACOUSTIC", "KEYBOARD", "EG", "BASS", "DRUMS"],
  defaults: {
    churchName: "Music Ministry",
    timeZone: "America/Los_Angeles",
    formCreationDay: 8,
    timesChoices: ["1", "2", "3", "4", "5"],
    availabilitySheetSuffix: "Availability",
    adminReminderEnabled: true,
    adminReminderDay: 5,
    eventsArchiveFrequency: "Yearly"
  },
  themes: {
    ministryMembers: { header: '#cfe2f3', tab: '#3c78d8', text: '#000000' },
    formMetadata: { header: '#cfe2f3', tab: '#6d9eeb', text: '#000000' },
    settings: { header: '#d9ead3', tab: '#6aa84f', text: '#000000' },
    admins: { header: '#fff2cc', tab: '#f1c232', text: '#000000' },
    rolesConfig: { header: '#fce5cd', tab: '#e69138', text: '#000000' },
    recurring: { header: '#d9d2e9', tab: '#674ea7', text: '#000000' },
    events: { header: '#d9d9d9', tab: '#666666', text: '#000000' },
    eventsArchive: { header: '#eeeeee', tab: '#999999', text: '#000000' },
    reconciliation: { header: '#f4cccc', tab: '#cc0000', text: '#000000' }
  },
  layout: {
    headerRowIndex: 13,
    dateRowIndex: 1
  }
};
