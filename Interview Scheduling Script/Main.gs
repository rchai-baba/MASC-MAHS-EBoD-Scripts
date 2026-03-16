// ╔══════════════════════════════════════════════════════════════════════════════╗
// ║  MAIN — Config, Menu, Triggers, Timestamp                                ║
// ╚══════════════════════════════════════════════════════════════════════════════╝

// ─── Calendly Import Config ───
const CONFIG = {
  CALENDLY_API_TOKEN: 'eyJraWQiOiIxY2UxZTEzNjE3ZGNmNzY2YjNjZWJjY2Y4ZGM1YmFmYThhNjVlNjg0MDIzZjdjMzJiZTgzNDliMjM4MDEzNWI0IiwidHlwIjoiUEFUIiwiYWxnIjoiRVMyNTYifQ.eyJpc3MiOiJodHRwczovL2F1dGguY2FsZW5kbHkuY29tIiwiaWF0IjoxNzczMzQwNDk3LCJqdGkiOiIwMjhkNDkyMy05NjgzLTRkZTUtYjk3NC0wMTQxMzAyN2VlNjMiLCJ1c2VyX3V1aWQiOiI5MjlkNDRlMi05NzgwLTRhYTQtYWE2Mi0wMmFkN2FjYmZlNTEiLCJzY29wZSI6InNjaGVkdWxlZF9ldmVudHM6cmVhZCB1c2VyczpyZWFkIn0.idqRzOjdLA4JTrxGfUFNROAnkGVqkCNmfnvQkZzRwCbLo4bcW6QZ0bPioUzQ-PfPfo2lXdUqdJrbTgvIypB5sQ',
  SCHEDULE_SHEET_NAME: 'auto',
  CSV_SHEET_NAME: 'calendly',
  DAYS_AHEAD: 80,
  COL_DATE:       1,
  COL_TIME:       2,
  COL_EBOD:       3,
  COL_BOD1:       4,
  COL_BOD2:       5,
  COL_NAME:       6,
  COL_EMAIL:      7,
  COL_PHONE:      8,
  COL_SCORES:     9,
  COL_ZOOM:      10,
  DATA_START_ROW: 2,
  EVENT_TYPE_FILTER: '',
  LAST_UPDATED_CELL: 'L1',

  // ─── Phone Number Lookup ───
  PHONE_SOURCE_SHEET_NAME:     'phone',
  PHONE_SOURCE_COL_EMAIL:      1,
  PHONE_SOURCE_COL_PHONE:      2,
  PHONE_SOURCE_DATA_START_ROW: 2,
  PHONE_TARGET_COL_EMAIL:      7,   // G in schedule sheet
  PHONE_TARGET_COL_PHONE:      8,   // H in schedule sheet
};


// ═══════════════════════════════════════
//  MENU
// ═══════════════════════════════════════

function onOpen() {
  SpreadsheetApp.getUi().createMenu('BOD Tools')
    .addItem('Import from Calendly API', 'importFromCalendlyAPI')
    .addItem('Import from pasted CSV', 'importFromCSV')
    .addSeparator()
    .addItem('Fill phone numbers', 'fillPhoneNumbers')
    .addSeparator()
    .addItem('Run full sync (Calendly + Phones)', 'runFullSync')
    .addToUi();
}


// ═══════════════════════════════════════
//  FULL SYNC — Calendly → Phones → Timestamp
//  This is what the hourly trigger calls.
// ═══════════════════════════════════════

function runFullSync() {
  Logger.log('========== runFullSync START ==========');
  importFromCalendlyAPI();
  fillPhoneNumbers();
  stampLastUpdated_();
  Logger.log('========== runFullSync COMPLETE ==========');
}


// ═══════════════════════════════════════
//  LAST UPDATED TIMESTAMP
// ═══════════════════════════════════════

function stampLastUpdated_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SCHEDULE_SHEET_NAME);
  if (!sheet) return;
  const now = new Date();
  const ts = Utilities.formatDate(now, 'America/New_York', 'M/d/yyyy h:mm:ss a') + ' ET';
  sheet.getRange(CONFIG.LAST_UPDATED_CELL).setValue('Last updated: ' + ts);
  Logger.log('Stamped "' + ts + '" → ' + CONFIG.SCHEDULE_SHEET_NAME + '!' + CONFIG.LAST_UPDATED_CELL);
}


// ═══════════════════════════════════════
//  TRIGGERS
//  Run createHourlyTrigger() once from the script editor.
// ═══════════════════════════════════════

function createHourlyTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'runFullSync') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('runFullSync').timeBased().everyHours(1).create();
  Logger.log('Hourly trigger created for runFullSync');
}

function removeHourlyTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'runFullSync') ScriptApp.deleteTrigger(t);
  });
  Logger.log('Trigger removed');
}


// ═══════════════════════════════════════
//  SHARED UTILITIES
// ═══════════════════════════════════════

function get12Hour_(h) { return h === 0 ? 12 : h > 12 ? h - 12 : h; }
function padZero_(n) { return n < 10 ? '0' + n : '' + n; }

function parseCalendlyDateTime_(str) {
  const m = str.match(/(\d{4})-(\d{2})-(\d{2})\s+(\d{1,2}):(\d{2})\s*(am|pm)/i);
  if (!m) { Logger.log('    PARSE FAIL: "' + str + '"'); return null; }
  let h = parseInt(m[4], 10);
  const ampm = m[6].toLowerCase();
  if (ampm === 'pm' && h !== 12) h += 12;
  if (ampm === 'am' && h === 12) h = 0;
  return new Date(parseInt(m[1]), parseInt(m[2]) - 1, parseInt(m[3]), h, parseInt(m[5]), 0);
}

function convertToEastern_(utcDate) {
  const s = Utilities.formatDate(utcDate, 'America/New_York', 'yyyy-MM-dd HH:mm');
  const p = s.split(/[-\s:]/);
  return new Date(parseInt(p[0]), parseInt(p[1]) - 1, parseInt(p[2]), parseInt(p[3]), parseInt(p[4]), 0);
}
