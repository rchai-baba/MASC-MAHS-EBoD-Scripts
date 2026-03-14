// ╔══════════════════════════════════════════════════════════════════════════════╗
// ║  CALENDLY → EBOD/BOD INTERVIEW SHEET — Google Apps Script                 ║
// ╚══════════════════════════════════════════════════════════════════════════════╝

const CONFIG = {
  CALENDLY_API_TOKEN: 'YOUR_API_KEY_HERE!!!'
  SCHEDULE_SHEET_NAME: 'auto',
  CSV_SHEET_NAME: 'CalendlyCSV',
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
};

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Calendly Import')
    .addItem('Import from Calendly API', 'importFromCalendlyAPI')
    .addSeparator()
    .addItem('Import from pasted CSV', 'importFromCSV')
    .addToUi();
}


// ═══════════════════════════════════════
//  MODE 1: IMPORT FROM CALENDLY API
// ═══════════════════════════════════════

function importFromCalendlyAPI() {
  Logger.log('========== importFromCalendlyAPI START ==========');

  if (CONFIG.CALENDLY_API_TOKEN === 'YOUR_CALENDLY_PERSONAL_ACCESS_TOKEN') {
    throw new Error('Set your CALENDLY_API_TOKEN in CONFIG. Get one at https://calendly.com/integrations/api_webhooks');
  }

  Logger.log('Step 1: Fetching current user...');
  const me = calendlyGet_('https://api.calendly.com/users/me');
  const userUri = me.resource.uri;
  const orgUri = me.resource.current_organization;
  Logger.log('  User: ' + userUri);
  Logger.log('  Org: ' + orgUri);

  const now = new Date();
  const minTime = now.toISOString();
  const futureDate = new Date(now.getTime() + CONFIG.DAYS_AHEAD * 86400000);
  const maxTime = futureDate.toISOString();
  Logger.log('Step 2: Date range: ' + minTime + ' → ' + maxTime);

  Logger.log('Step 3: Fetching events...');
  let allEvents = [];
  let nextPageToken = null;
  let pageNum = 0;

  do {
    pageNum++;
    let url = 'https://api.calendly.com/scheduled_events'
      + '?organization=' + encodeURIComponent(orgUri)
      + '&min_start_time=' + encodeURIComponent(minTime)
      + '&max_start_time=' + encodeURIComponent(maxTime)
      + '&status=active&count=100';
    if (nextPageToken) url += '&page_token=' + encodeURIComponent(nextPageToken);

    const resp = JSON.parse(calendlyGet_(url, true));
    const pageEvents = resp.collection || [];
    allEvents = allEvents.concat(pageEvents);
    Logger.log('  Page ' + pageNum + ': ' + pageEvents.length + ' events (total: ' + allEvents.length + ')');
    nextPageToken = (resp.pagination && resp.pagination.next_page_token) || null;
  } while (nextPageToken);

  Logger.log('Step 4: Fetching invitees...');
  let calendlyRows = [];
  for (let e = 0; e < allEvents.length; e++) {
    const event = allEvents[e];
    if (CONFIG.EVENT_TYPE_FILTER && event.name !== CONFIG.EVENT_TYPE_FILTER) continue;

    const eventUuid = event.uri.split('/').pop();
    const invResp = JSON.parse(calendlyGet_('https://api.calendly.com/scheduled_events/' + eventUuid + '/invitees', true));

    for (const inv of (invResp.collection || [])) {
      if (inv.status === 'canceled') continue;
      let zoomLink = (event.location && event.location.join_url) ? event.location.join_url : '';
      calendlyRows.push({ name: inv.name || '', email: inv.email || '', startTime: event.start_time, zoomLink: zoomLink });
      Logger.log('  ' + inv.name + ' | ' + event.start_time);
    }
  }

  Logger.log('Total invitees: ' + calendlyRows.length);
  Logger.log('Step 5: Converting times...');

  const parsedEvents = calendlyRows.map(function(row) {
    const utcDate = new Date(row.startTime);
    const etDate = convertToEastern_(utcDate);
    const p = { month: etDate.getMonth() + 1, day: etDate.getDate(), hour12: get12Hour_(etDate.getHours()), minute: etDate.getMinutes(), name: row.name, email: row.email, zoomLink: row.zoomLink };
    Logger.log('  ' + row.name + ': ' + p.month + '/' + p.day + ' ' + p.hour12 + ':' + padZero_(p.minute));
    return p;
  });

  const count = writeEventsToSheet_(parsedEvents);
  Logger.log('========== DONE — Matched ' + count + '/' + parsedEvents.length + ' ==========');
}

function calendlyGet_(url, rawResponse) {
  const resp = UrlFetchApp.fetch(url, { method: 'get', headers: { 'Authorization': 'Bearer ' + CONFIG.CALENDLY_API_TOKEN, 'Content-Type': 'application/json' }, muteHttpExceptions: true });
  if (resp.getResponseCode() !== 200) throw new Error('Calendly API ' + resp.getResponseCode() + ': ' + resp.getContentText());
  return rawResponse ? resp.getContentText() : JSON.parse(resp.getContentText());
}


// ═══════════════════════════════════════
//  MODE 2: IMPORT FROM PASTED CSV
// ═══════════════════════════════════════

function importFromCSV() {
  Logger.log('========== importFromCSV START ==========');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Logger.log('Spreadsheet: ' + (ss ? ss.getName() : 'NULL!'));
  if (!ss) throw new Error('getActiveSpreadsheet() is null. Open this script from Extensions > Apps Script inside the Google Sheet.');

  // List all tabs
  const allTabs = ss.getSheets().map(function(s) { return '"' + s.getName() + '"'; });
  Logger.log('All tabs: [' + allTabs.join(', ') + ']');

  const csvSheet = ss.getSheetByName(CONFIG.CSV_SHEET_NAME);
  Logger.log('Looking for "' + CONFIG.CSV_SHEET_NAME + '" → ' + (csvSheet ? 'FOUND' : 'NOT FOUND'));
  if (!csvSheet) throw new Error('Tab "' + CONFIG.CSV_SHEET_NAME + '" not found. Tabs: ' + allTabs.join(', '));

  const data = csvSheet.getDataRange().getValues();
  Logger.log('CSV data: ' + data.length + ' rows, ' + (data[0] ? data[0].length : 0) + ' columns');
  if (data.length < 2) throw new Error('CSV sheet is empty.');

  // Log raw first data row for debugging
  Logger.log('');
  Logger.log('RAW HEADER ROW (row 1):');
  for (let c = 0; c < Math.min(data[0].length, 15); c++) {
    Logger.log('  Col ' + c + ': "' + data[0][c] + '"');
  }

  Logger.log('');
  Logger.log('RAW FIRST DATA ROW (row 2):');
  for (let c = 0; c < Math.min(data[1].length, 15); c++) {
    Logger.log('  Col ' + c + ': "' + data[1][c] + '" (type: ' + typeof data[1][c] + ', isDate: ' + (data[1][c] instanceof Date) + ')');
  }

  const headers = data[0].map(function(h) { return String(h).trim(); });
  Logger.log('');
  Logger.log('Parsed headers: [' + headers.join(' | ') + ']');

  const colIdx = {
    name:      headers.indexOf('Invitee Name'),
    email:     headers.indexOf('Invitee Email'),
    startTime: headers.indexOf('Start Date & Time'),
    location:  headers.indexOf('Location'),
    canceled:  headers.indexOf('Canceled'),
    eventType: headers.indexOf('Event Type Name'),
  };

  Logger.log('Column indices → name:' + colIdx.name + ' email:' + colIdx.email +
    ' startTime:' + colIdx.startTime + ' location:' + colIdx.location +
    ' canceled:' + colIdx.canceled + ' eventType:' + colIdx.eventType);

  if (colIdx.name === -1 || colIdx.email === -1 || colIdx.startTime === -1) {
    throw new Error('Missing columns! Need "Invitee Name", "Invitee Email", "Start Date & Time". Got: ' + headers.join(', '));
  }

  // Parse rows
  let parsedEvents = [];
  let skippedCanceled = 0, skippedFilter = 0, skippedParse = 0;

  Logger.log('');
  Logger.log('PARSING CSV ROWS:');

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const canceled = String(row[colIdx.canceled] || '').toLowerCase().trim();
    if (canceled === 'true') {
      skippedCanceled++;
      Logger.log('  Row ' + (i+1) + ': SKIP (canceled) "' + row[colIdx.name] + '"');
      continue;
    }

    if (CONFIG.EVENT_TYPE_FILTER && colIdx.eventType !== -1) {
      const et = String(row[colIdx.eventType] || '').trim();
      if (et && et !== CONFIG.EVENT_TYPE_FILTER) { skippedFilter++; continue; }
    }

    const name = String(row[colIdx.name] || '').trim();
    const email = String(row[colIdx.email] || '').trim();
    const startRaw = row[colIdx.startTime];

    Logger.log('  Row ' + (i+1) + ': name="' + name + '" startRaw="' + startRaw + '" type=' + typeof startRaw + ' isDate=' + (startRaw instanceof Date));

    let parsedDate;
    if (startRaw instanceof Date) {
      parsedDate = startRaw;
    } else {
      parsedDate = parseCalendlyDateTime_(String(startRaw).trim());
    }

    if (!parsedDate || !name) {
      skippedParse++;
      Logger.log('    → SKIP: parsedDate=' + parsedDate + ', name="' + name + '"');
      continue;
    }

    let zoomLink = '';
    if (colIdx.location !== -1) {
      const locStr = String(row[colIdx.location] || '');
      const m = locStr.match(/https?:\/\/[^\s,]+/);
      if (m) zoomLink = m[0];
    }

    const h24 = parsedDate.getHours();
    const h12 = get12Hour_(h24);
    const min = parsedDate.getMinutes();
    const mo = parsedDate.getMonth() + 1;
    const dy = parsedDate.getDate();
    const key = mo + '/' + dy + '|' + h12 + ':' + padZero_(min);

    Logger.log('    → OK: ' + mo + '/' + dy + ' h24=' + h24 + ' h12=' + h12 + ':' + padZero_(min) + ' key="' + key + '"');

    parsedEvents.push({ month: mo, day: dy, hour12: h12, minute: min, name: name, email: email, zoomLink: zoomLink });
  }

  Logger.log('');
  Logger.log('CSV SUMMARY: total=' + (data.length - 1) + ' parsed=' + parsedEvents.length +
    ' canceled=' + skippedCanceled + ' filtered=' + skippedFilter + ' parseError=' + skippedParse);

  Logger.log('');
  const count = writeEventsToSheet_(parsedEvents);
  Logger.log('========== DONE — Matched ' + count + '/' + parsedEvents.length + ' ==========');
}


// ═══════════════════════════════════════
//  WRITE EVENTS TO SCHEDULE SHEET
// ═══════════════════════════════════════

function writeEventsToSheet_(events) {
  Logger.log('--- writeEventsToSheet_ ---');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SCHEDULE_SHEET_NAME);
  if (!sheet) {
    const tabs = ss.getSheets().map(function(s) { return s.getName(); });
    throw new Error('Sheet "' + CONFIG.SCHEDULE_SHEET_NAME + '" not found! Tabs: ' + tabs.join(', '));
  }

  const lastRow = sheet.getLastRow();
  const numRows = lastRow - CONFIG.DATA_START_ROW + 1;
  Logger.log('Schedule sheet: ' + lastRow + ' rows total, ' + numRows + ' data rows');

  if (numRows <= 0) {
    Logger.log('ERROR: No data rows!');
    return 0;
  }

  const dataRange = sheet.getRange(CONFIG.DATA_START_ROW, 1, numRows, CONFIG.COL_ZOOM);
  const sheetData = dataRange.getValues();

  // Log first 5 raw schedule rows
  Logger.log('');
  Logger.log('FIRST 5 RAW SCHEDULE ROWS:');
  for (let i = 0; i < Math.min(5, sheetData.length); i++) {
    const dateVal = sheetData[i][CONFIG.COL_DATE - 1];
    const timeVal = sheetData[i][CONFIG.COL_TIME - 1];
    const nameVal = sheetData[i][CONFIG.COL_NAME - 1];
    Logger.log('  Row ' + (CONFIG.DATA_START_ROW + i) + ': date="' + dateVal + '" (' + typeof dateVal + ')' +
      ' | time="' + timeVal + '" (' + typeof timeVal + ', isDate=' + (timeVal instanceof Date) + ')' +
      ' | name="' + nameVal + '"');
  }

  // Build slot map
  Logger.log('');
  Logger.log('BUILDING SLOT MAP:');
  let slotMap = {};
  let slotCount = 0;
  let skipCount = 0;

  for (let i = 0; i < sheetData.length; i++) {
    const dateCell = sheetData[i][CONFIG.COL_DATE - 1];
    const timeCell = sheetData[i][CONFIG.COL_TIME - 1];
    const rowNum = CONFIG.DATA_START_ROW + i;

    if (!dateCell || !timeCell) { skipCount++; continue; }

    // Parse date
    const dateStr = String(dateCell);
    const dateMatch = dateStr.match(/\((\d{1,2})\/(\d{1,2})\)/);
    if (!dateMatch) {
      if (i < 5) Logger.log('  Row ' + rowNum + ': "' + dateStr + '" — no (m/d) match, SKIP');
      skipCount++;
      continue;
    }
    const sMonth = parseInt(dateMatch[1], 10);
    const sDay = parseInt(dateMatch[2], 10);

    // Parse time
    let sHour, sMin;
    if (timeCell instanceof Date) {
      sHour = timeCell.getHours();
      sMin = timeCell.getMinutes();
    } else if (typeof timeCell === 'number') {
      // Might be a serial number (fraction of day)
      // 0.5 = 12:00, 0.0625 = 1:30, etc.
      const totalMinutes = Math.round(timeCell * 24 * 60);
      sHour = Math.floor(totalMinutes / 60);
      sMin = totalMinutes % 60;
      if (i < 5) Logger.log('  Row ' + rowNum + ': time is NUMBER ' + timeCell + ' → ' + sHour + ':' + padZero_(sMin));
    } else {
      const tStr = String(timeCell);
      const tMatch = tStr.match(/(\d{1,2}):(\d{2})/);
      if (tMatch) {
        sHour = parseInt(tMatch[1], 10);
        sMin = parseInt(tMatch[2], 10);
      } else {
        if (i < 5) Logger.log('  Row ' + rowNum + ': time "' + tStr + '" — cannot parse, SKIP');
        skipCount++;
        continue;
      }
    }

    const key = sMonth + '/' + sDay + '|' + sHour + ':' + padZero_(sMin);
    if (!slotMap[key]) slotMap[key] = [];
    slotMap[key].push(i);
    slotCount++;

    if (slotCount <= 10) {
      Logger.log('  Row ' + rowNum + ': "' + dateStr + '" time=' + sHour + ':' + padZero_(sMin) + ' → key="' + key + '"');
    }
  }

  const allKeys = Object.keys(slotMap);
  Logger.log('Slot map: ' + allKeys.length + ' unique keys, ' + slotCount + ' slots, ' + skipCount + ' skipped');
  Logger.log('Sample keys: [' + allKeys.slice(0, 20).join(', ') + ']');

  // Match
  Logger.log('');
  Logger.log('MATCHING:');
  let matchCount = 0, noSlot = 0;

  for (const evt of events) {
    const key = evt.month + '/' + evt.day + '|' + evt.hour12 + ':' + padZero_(evt.minute);
    const indices = slotMap[key];

    if (!indices || indices.length === 0) {
      noSlot++;
      Logger.log('  ✗ "' + evt.name + '" key="' + key + '" — NO SLOT');
      continue;
    }

    const rowIdx = indices[0];
    const actualRow = CONFIG.DATA_START_ROW + rowIdx;

    sheet.getRange(actualRow, CONFIG.COL_NAME).setValue(evt.name);
    sheet.getRange(actualRow, CONFIG.COL_EMAIL).setValue(evt.email);
    if (evt.zoomLink) sheet.getRange(actualRow, CONFIG.COL_ZOOM).setValue(evt.zoomLink);

    matchCount++;
    Logger.log('  ✓ "' + evt.name + '" → Row ' + actualRow + ' key="' + key + '"');
  }

  SpreadsheetApp.flush();

  Logger.log('');
  Logger.log('RESULTS: matched=' + matchCount + ' noSlot=' + noSlot + ' total=' + events.length);
  return matchCount;
}


// ═══════════════════════════════════════
//  UTILITIES
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


// ═══════════════════════════════════════
//  AUTO-RUN TRIGGERS
// ═══════════════════════════════════════

function createHourlyTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'importFromCalendlyAPI') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('importFromCalendlyAPI').timeBased().everyHours(1).create();
  Logger.log('Hourly trigger created');
}

function removeHourlyTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'importFromCalendlyAPI') ScriptApp.deleteTrigger(t);
  });
  Logger.log('Trigger removed');
}