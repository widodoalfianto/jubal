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
    eventsArchiveFrequency: "Yearly",
    eventsArchiveMonth: "January"
  },
  layout: {
    headerRowIndex: 13,
    dateRowIndex: 1
  }
};
