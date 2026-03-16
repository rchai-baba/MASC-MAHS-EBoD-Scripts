// ╔══════════════════════════════════════════════════════════════════════════════╗
// ║  PHONE NUMBER LOOKUP — Applications → Interview Schedule                  ║
// ║                                                                            ║
// ║  Matches emails between the Applications sheet and the schedule sheet,     ║
// ║  then fills in the phone numbers.                                          ║
// ╚══════════════════════════════════════════════════════════════════════════════╝

const PHONE_CONFIG = {
  // ─── Source: Applications sheet ───
  SOURCE_SHEET_NAME:    'phone',
  SOURCE_COL_EMAIL:     1,   // E — email column
  SOURCE_COL_PHONE:     2,   // F — phone column
  SOURCE_DATA_START_ROW: 2,  // first data row (after header)

  // ─── Target: Interview schedule sheet ───
  TARGET_SHEET_NAME:    'auto',
  TARGET_COL_EMAIL:     7,   // G — email column
  TARGET_COL_PHONE:     8,   // H — phone column
  TARGET_DATA_START_ROW: 2,  // first data row (after header)
};


/**
 * Reads emails + phone numbers from the Applications sheet,
 * then matches by email to fill phone numbers in the schedule sheet.
 */
function fillPhoneNumbers() {
  Logger.log('========== fillPhoneNumbers START ==========');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('getActiveSpreadsheet() is null. Open this from Extensions > Apps Script inside the Google Sheet.');

  const allTabs = ss.getSheets().map(function(s) { return '"' + s.getName() + '"'; });
  Logger.log('All tabs: [' + allTabs.join(', ') + ']');

  // ── Step 1: Read source (Applications) sheet ──
  const srcSheet = ss.getSheetByName(PHONE_CONFIG.SOURCE_SHEET_NAME);
  Logger.log('Looking for source "' + PHONE_CONFIG.SOURCE_SHEET_NAME + '" → ' + (srcSheet ? 'FOUND' : 'NOT FOUND'));
  if (!srcSheet) throw new Error('"' + PHONE_CONFIG.SOURCE_SHEET_NAME + '" not found. Tabs: ' + allTabs.join(', '));

  const srcLastRow = srcSheet.getLastRow();
  const srcNumRows = srcLastRow - PHONE_CONFIG.SOURCE_DATA_START_ROW + 1;
  Logger.log('Source sheet: ' + srcLastRow + ' rows total, ' + srcNumRows + ' data rows');

  if (srcNumRows <= 0) {
    Logger.log('ERROR: No data in source sheet.');
    return;
  }

  const maxSrcCol = Math.max(PHONE_CONFIG.SOURCE_COL_EMAIL, PHONE_CONFIG.SOURCE_COL_PHONE);
  const srcData = srcSheet.getRange(PHONE_CONFIG.SOURCE_DATA_START_ROW, 1, srcNumRows, maxSrcCol).getValues();

  // Log first 3 source rows for debugging
  Logger.log('');
  Logger.log('FIRST 3 SOURCE ROWS:');
  for (let i = 0; i < Math.min(3, srcData.length); i++) {
    const email = srcData[i][PHONE_CONFIG.SOURCE_COL_EMAIL - 1];
    const phone = srcData[i][PHONE_CONFIG.SOURCE_COL_PHONE - 1];
    Logger.log('  Row ' + (PHONE_CONFIG.SOURCE_DATA_START_ROW + i) + ': email="' + email + '" phone="' + phone + '"');
  }

  // Build email → phone lookup (lowercase email keys)
  let phoneLookup = {};
  let srcCount = 0;
  let emptyCount = 0;

  for (let i = 0; i < srcData.length; i++) {
    const email = String(srcData[i][PHONE_CONFIG.SOURCE_COL_EMAIL - 1] || '').trim().toLowerCase();
    const phone = String(srcData[i][PHONE_CONFIG.SOURCE_COL_PHONE - 1] || '').trim();

    if (email && phone) {
      phoneLookup[email] = phone;
      srcCount++;
    } else {
      emptyCount++;
    }
  }

  Logger.log('');
  Logger.log('Phone lookup built: ' + srcCount + ' entries (' + emptyCount + ' rows had missing email or phone)');

  const sampleEmails = Object.keys(phoneLookup).slice(0, 5);
  for (const e of sampleEmails) {
    Logger.log('  ' + e + ' → ' + phoneLookup[e]);
  }

  // ── Step 2: Read target (schedule) sheet ──
  const tgtSheet = ss.getSheetByName(PHONE_CONFIG.TARGET_SHEET_NAME);
  Logger.log('');
  Logger.log('Looking for target "' + PHONE_CONFIG.TARGET_SHEET_NAME + '" → ' + (tgtSheet ? 'FOUND' : 'NOT FOUND'));
  if (!tgtSheet) throw new Error('"' + PHONE_CONFIG.TARGET_SHEET_NAME + '" not found. Tabs: ' + allTabs.join(', '));

  const tgtLastRow = tgtSheet.getLastRow();
  const tgtNumRows = tgtLastRow - PHONE_CONFIG.TARGET_DATA_START_ROW + 1;
  Logger.log('Target sheet: ' + tgtLastRow + ' rows total, ' + tgtNumRows + ' data rows');

  if (tgtNumRows <= 0) {
    Logger.log('No data rows in target sheet.');
    return;
  }

  const maxTgtCol = Math.max(PHONE_CONFIG.TARGET_COL_EMAIL, PHONE_CONFIG.TARGET_COL_PHONE);
  const tgtData = tgtSheet.getRange(PHONE_CONFIG.TARGET_DATA_START_ROW, 1, tgtNumRows, maxTgtCol).getValues();

  // ── Step 3: Match emails and fill phone numbers ──
  Logger.log('');
  Logger.log('MATCHING EMAILS:');
  let matchCount = 0, noEmail = 0, noPhone = 0, alreadyFilled = 0;

  for (let i = 0; i < tgtData.length; i++) {
    const email = String(tgtData[i][PHONE_CONFIG.TARGET_COL_EMAIL - 1] || '').trim().toLowerCase();
    const existingPhone = String(tgtData[i][PHONE_CONFIG.TARGET_COL_PHONE - 1] || '').trim();
    const actualRow = PHONE_CONFIG.TARGET_DATA_START_ROW + i;

    if (!email) {
      noEmail++;
      continue;
    }

    // Skip if phone already filled
    if (existingPhone) {
      alreadyFilled++;
      Logger.log('  — Row ' + actualRow + ': "' + email + '" already has phone "' + existingPhone + '"');
      continue;
    }

    const phone = phoneLookup[email];
    if (!phone) {
      noPhone++;
      Logger.log('  ✗ Row ' + actualRow + ': "' + email + '" — not found in ' + PHONE_CONFIG.SOURCE_SHEET_NAME);
      continue;
    }

    tgtSheet.getRange(actualRow, PHONE_CONFIG.TARGET_COL_PHONE).setValue(phone);
    matchCount++;
    Logger.log('  ✓ Row ' + actualRow + ': "' + email + '" → ' + phone);
  }

  SpreadsheetApp.flush();

  Logger.log('');
  Logger.log('══════════════════════════════');
  Logger.log('  SUMMARY');
  Logger.log('  Filled:                  ' + matchCount);
  Logger.log('  Already had phone:       ' + alreadyFilled);
  Logger.log('  Empty email in schedule: ' + noEmail);
  Logger.log('  Email not in Apps sheet: ' + noPhone);
  Logger.log('══════════════════════════════');
  Logger.log('========== fillPhoneNumbers DONE ==========');
}
