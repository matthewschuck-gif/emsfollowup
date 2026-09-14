/**
 * EMS Incident Follow-Up — Apps Script Web App
 *
 * SETUP (one time):
 * 1. Open the EMS Incident Report response spreadsheet
 * 2. Extensions → Apps Script → paste this entire file
 * 3. Click Deploy → New Deployment
 *    - Type: Web App
 *    - Execute as: Me
 *    - Who has access: Anyone
 * 4. Click Deploy → copy the Web App URL
 * 5. Paste that URL into the APPS_SCRIPT_URL variable
 *    in index.html
 *
 * UPDATING AN EXISTING DEPLOYMENT:
 * Editing this code does NOT change what a live Web App URL serves.
 * Go to Deploy → Manage deployments → pencil icon on the existing
 * deployment → Version: "New version" → Deploy. Otherwise the URL
 * keeps running the old code.
 */

var SHEET_ID            = '1VUDvmkqG6F7-1PQWWc3_FWQxGz6bj_tBRdi4UuAWqP0';
var TAB_NAME             = 'Staff Follow-Up';
var STREAMLINED_TAB_NAME = 'Staff-Entered Reports';

// Handle POST from dashboard
function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    if (data.action === 'createStreamlined') {
      var ts = createStreamlined(data);
      return ContentService
        .createTextOutput(JSON.stringify({ success: true, timestamp: ts }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    writeFollowUp(data);
    return ContentService
      .createTextOutput(JSON.stringify({ success: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Handle GET — returns { followUps: {...}, streamlined: [...] } so any
// staff member's dashboard can sync to what's actually on the sheets.
//
// Apps Script does not send an Access-Control-Allow-Origin header, so a
// plain fetch() from a page hosted on a different domain gets its response
// silently blocked by the browser (the request still runs — that's why
// direct/incognito visits to this URL work fine — only reading it from JS
// on another origin fails). If a ?callback= param is present, respond
// JSONP-style instead: a <script> tag load isn't subject to CORS at all.
function doGet(e) {
  var callback = e && e.parameter && e.parameter.callback;
  try {
    var result = {
      followUps:   readAllFollowUps(),
      streamlined: readStreamlinedReports()
    };
    return respond_(result, callback);
  } catch (err) {
    return respond_({ error: err.toString() }, callback);
  }
}

function respond_(obj, callback) {
  if (callback) {
    return ContentService
      .createTextOutput(callback + '(' + JSON.stringify(obj) + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function readAllFollowUps() {
  var ss    = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(TAB_NAME);
  var result = {};

  if (!sheet) return result; // no follow-ups saved yet

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return result;

  // Column order written by writeFollowUp():
  // 1 Submission Timestamp, 2 Student 800#, 3 Student Name, 4 Grade,
  // 5 divider, 6 Assigned To, 7 Follow-Up Adult, 8 Meeting Date, 9 Notes,
  // 10 Logged in PS, 11 SRO Contacted, 12 Parent Contacted,
  // 13 Contact Method, 14 Teachers Notified, 15 Status, 16 Closed At,
  // 17 Last Updated
  var rows = sheet.getRange(2, 1, lastRow - 1, 17).getValues();

  rows.forEach(function(row) {
    var ts = String(row[0] || '').trim();
    if (!ts) return;
    result[ts] = {
      timestamp:    ts,
      studentId:    row[1]  || '',
      studentName:  row[2]  || '',
      grade:        row[3]  || '',
      assignedTo:   row[5]  || '',
      adult:        row[6]  || '',
      meetDate:     row[7]  || '',
      notes:        row[8]  || '',
      ps:           row[9]  || '',
      sro:          row[10] || '',
      parent:       row[11] || '',
      contact:      row[12] || '',
      teachers:     row[13] || '',
      status:       row[14] || 'open',
      closedAt:     row[15] || '',
      lastUpdated:  row[16] || ''
    };
  });

  return result;
}

function writeFollowUp(data) {
  var ss    = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(TAB_NAME);

  // Auto-create the tab if it doesn't exist yet
  if (!sheet) {
    sheet = ss.insertSheet(TAB_NAME);
    var headers = [
      'Submission Timestamp', 'Student 800#', 'Student Name', 'Grade',
      '── STAFF FOLLOW-UP ──',
      'Assigned To', 'Follow-Up Adult', 'Meeting Date', 'Notes',
      'Logged in PS', 'SRO Contacted', 'Parent Contacted',
      'Contact Method', 'Teachers Notified',
      'Status', 'Closed At', 'Last Updated'
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length)
      .setBackground('#490e6f').setFontColor('#ffe100')
      .setFontWeight('bold').setFontSize(11);
    sheet.setFrozenRows(1);
  }

  var ts = data.timestamp || '';

  // Search for existing row with this timestamp
  var lastRow   = sheet.getLastRow();
  var existRow  = -1;

  if (lastRow > 1) {
    var tsCol = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (var i = 0; i < tsCol.length; i++) {
      if (String(tsCol[i][0]).trim() === ts.trim()) {
        existRow = i + 2; // 1-indexed, offset for header
        break;
      }
    }
  }

  var now = new Date().toLocaleString('en-US', {
    month: 'short', day: 'numeric', year: 'numeric',
    hour: 'numeric', minute: '2-digit', hour12: true
  });

  var rowData = [
    ts,
    data.studentId   || '',
    data.studentName || '',
    data.grade       || '',
    '',                        // divider
    data.assignedTo  || '',
    data.adult       || '',
    data.meetDate    || '',
    data.notes       || '',
    data.ps          || '',
    data.sro         || '',
    data.parent      || '',
    data.contact     || '',
    data.teachers    || '',
    data.status      || 'open',
    data.closedAt    || '',
    now
  ];

  if (existRow > 0) {
    // Update existing row
    sheet.getRange(existRow, 1, 1, rowData.length).setValues([rowData]);
  } else {
    // Append new row
    sheet.appendRow(rowData);
    existRow = sheet.getLastRow();
  }

  // Color-code status cell (column 15)
  var statusCell = sheet.getRange(existRow, 15);
  var status = data.status || 'open';
  if (status === 'open') {
    statusCell.setBackground('#fee2e2').setFontColor('#b91c1c');
  } else if (status === 'progress') {
    statusCell.setBackground('#fef3c7').setFontColor('#d97706');
  } else if (status === 'closed') {
    statusCell.setBackground('#dcfce7').setFontColor('#16a34a');
  }
}

// ── STREAMLINED / IN-PERSON REPORTS ────────────────────────────────────
// Used when a staff member is meeting with a student directly and no
// incident form was ever submitted. One save writes the incident-style
// details to their own sheet (so the real Form Responses sheet is never
// touched) AND writes the follow-up portion into the same Staff Follow-Up
// sheet used for everything else, keyed by the same timestamp — so once
// created, the case is tracked identically to a form-submitted incident.

function createStreamlined(data) {
  var ss    = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(STREAMLINED_TAB_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(STREAMLINED_TAB_NAME);
    var headers = [
      'Timestamp', 'First Name', 'Last Name', 'Grade', 'Homeroom Teacher',
      '800 Number (Student ID)', 'Where did it happen?', 'When did it happen?',
      'Who was involved?', 'Who witnessed it?', 'Is there any evidence?',
      'What happened?', "What do you think needs to happen to resolve this situation?",
      'Entered By'
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length)
      .setBackground('#490e6f').setFontColor('#ffe100')
      .setFontWeight('bold').setFontSize(11);
    sheet.setFrozenRows(1);
  }

  var ts = data.timestamp || new Date().toISOString();

  var rowData = [
    ts,
    data.firstName    || '',
    data.lastName     || '',
    data.grade        || '',
    data.homeroom     || '',
    data.studentId    || '',
    data.where        || '',
    data.when         || '',
    data.who          || '',
    data.witnesses    || '',
    data.evidence     || '',
    data.whatHappened || '',
    data.resolution   || '',
    data.enteredBy    || ''
  ];
  sheet.appendRow(rowData);

  writeFollowUp({
    timestamp:   ts,
    studentId:   data.studentId || '',
    studentName: ((data.firstName || '') + ' ' + (data.lastName || '')).trim(),
    grade:       data.grade || '',
    status:      data.status || 'open',
    assignedTo:  data.assignedTo || '',
    adult:       data.adult || '',
    meetDate:    data.meetDate || '',
    notes:       data.notes || '',
    ps:          data.ps || '',
    sro:         data.sro || '',
    parent:      data.parent || '',
    contact:     data.contact || '',
    teachers:    data.teachers || '',
    closedAt:    data.closedAt || ''
  });

  return ts;
}

function readStreamlinedReports() {
  var ss    = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(STREAMLINED_TAB_NAME);
  var results = [];

  if (!sheet) return results;

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return results;

  var rows = sheet.getRange(2, 1, lastRow - 1, 14).getValues();

  rows.forEach(function(row) {
    var ts = String(row[0] || '').trim();
    if (!ts) return;
    results.push({
      'Timestamp':                ts,
      'First Name':               row[1]  || '',
      'Last Name':                row[2]  || '',
      'Grade':                    row[3]  || '',
      'Homeroom Teacher':         row[4]  || '',
      '800 Number (Student ID)':  row[5]  || '',
      'Where did it happen?':     row[6]  || '',
      'When did it happen?':      row[7]  || '',
      'Who was involved?':        row[8]  || '',
      'Who witnessed it?':        row[9]  || '',
      'Is there any evidence?':   row[10] || '',
      'What happened?':           row[11] || '',
      "What do you think needs to happen to resolve this situation?": row[12] || '',
      'Entered By':                row[13] || ''
    });
  });

  return results;
}
