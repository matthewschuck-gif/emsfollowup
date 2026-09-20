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

// Staff emails to notify whenever a new in-person ("+") report is saved.
// Same team/format as ALERT_EMAIL in form-trigger.gs (the script bound to
// the incident report Form) -- edit this comma-separated list directly to
// add or remove addresses, and keep both lists in sync.
var ALERT_EMAIL = 'Christopher_Montagna@easdpa.org, Kevin_Kuhn@easdpa.org, J_Lugar@easdpa.org, Kristin_Mincarelli@easdpa.org, A_Mowbray@easdpa.org, Matthew_schuck@easdpa.org, K_Wagner@easdpa.org, Kelly_Rigg@easdpa.org, Sara_Judge@easdpa.org, Lindsey_Carr@easdpa.org, Nicole_Zimmerman@easdpa.org, Donna_Schlinkman@easdpa.org, Kristin_Mincarelli@easdpa.org,matthew_schuck@easdpa.org';

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
      'Entered By',
      // ── BULLYING/CYBERBULLYING BRANCH (PSBA 249-AR-1) ──
      'Believed to be Bullying?', 'Complainant/Reporter Name', 'Complainant Home Address',
      'Complainant Phone Number', 'School Building', 'Date of Alleged Incident(s)',
      "Alleged Offender(s)", 'If Directed at Someone Else, Identify Them',
      'Statements Made (threats/requests/demands)', 'Type of Bullying',
      'Cyber: Platform/App/Method Used', 'Cyber: Device Used', 'Cyber: Username(s)/Account(s) Involved',
      'Cyber: During or Outside School Hours', 'Cyber: Date/Time of Online Activity', 'Cyber: Evidence Saved',
      'In-Person: Specific Location', 'In-Person: During or Outside School Hours',
      'In-Person: Physical, Verbal, or Both', 'In-Person: Injuries/Physical Contact',
      'Certified True and Complete'
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
    data.enteredBy    || '',
    data.bullyFlag         || '',
    data.bullyReporterName    || '',
    data.bullyReporterAddress || '',
    data.bullyReporterPhone   || '',
    data.bullySchool          || '',
    data.bullyIncidentDate    || '',
    data.bullyOffenders       || '',
    data.bullyTargetOther     || '',
    data.bullyStatements      || '',
    data.bullyType             || '',
    data.cyberPlatform    || '',
    data.cyberDevice      || '',
    data.cyberAccounts    || '',
    data.cyberHours       || '',
    data.cyberWhen        || '',
    data.cyberEvidence    || '',
    data.inpersonLocation || '',
    data.inpersonHours    || '',
    data.inpersonNature   || '',
    data.inpersonInjury   || '',
    data.bullyCertify ? 'Yes' : ''
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

  notifyStreamlinedReport_(data, ts);

  return ts;
}

// Emails ALERT_EMAIL whenever a staff member saves a new in-person report.
// Mirrors the look of onIncidentSubmit()'s alert in the incident-form
// creator script. A failed send never blocks the report from saving.
function notifyStreamlinedReport_(data, ts) {
  if (!ALERT_EMAIL) return;
  try {
    var firstName = data.firstName || '—';
    var lastName  = data.lastName  || '—';
    var grade     = data.grade     || '—';
    var isBullying = data.bullyFlag === 'Yes';

    var submittedAt = Utilities.formatDate(new Date(ts), Session.getScriptTimeZone(), 'MMMM d, yyyy \'at\' h:mm a');

    var subject = (isBullying ? '⚠️ Bullying/Cyberbullying — ' : '🧑‍🤝‍🧑 In-Person Report — ')
      + firstName + ' ' + lastName + ' (' + grade + ' grade)';

    var body =
      'A staff member entered an in-person report at ' + submittedAt + '.\n\n' +
      '══════════════════════════════════════\n' +
      'STUDENT INFORMATION\n' +
      '══════════════════════════════════════\n' +
      '800 Number:    ' + (data.studentId || '—') + '\n' +
      'Name:          ' + firstName + ' ' + lastName + '\n' +
      'Grade:         ' + grade + '\n' +
      'Homeroom:      ' + (data.homeroom || '—') + '\n' +
      'Entered by:    ' + (data.enteredBy || '—') + '\n\n' +
      '══════════════════════════════════════\n' +
      'INCIDENT DETAILS\n' +
      '══════════════════════════════════════\n' +
      'Where:         ' + (data.where || '—') + '\n' +
      'When:          ' + (data.when  || '—') + '\n\n' +
      'What happened:\n' + (data.whatHappened || '—') + '\n\n' +
      (isBullying ?
        '══════════════════════════════════════\n' +
        'BULLYING/CYBERBULLYING REPORT (249-AR-1)\n' +
        '══════════════════════════════════════\n' +
        'Type:               ' + (data.bullyType || '—') + '\n' +
        'Reporter:            ' + (data.bullyReporterName || '—') + '\n' +
        'Alleged offender(s): ' + (data.bullyOffenders || '—') + '\n' +
        'School building:     ' + (data.bullySchool || '—') + '\n' +
        'Date of incident:    ' + (data.bullyIncidentDate || '—') + '\n\n'
        : '') +
      '══════════════════════════════════════\n' +
      'STAFF FOLLOW-UP\n' +
      '══════════════════════════════════════\n' +
      'Status: ' + (data.status || 'open') + '\n\n' +
      'This report is logged in the EMS Incident Follow-Up sheet.\n';

    var htmlBody =
      '<div style="font-family:Arial,sans-serif;max-width:600px;color:#1a1a1a;">' +
      '<div style="background:#490e6f;padding:16px 20px;border-radius:8px 8px 0 0;">' +
        '<h2 style="color:#ffe100;margin:0;font-size:18px;">' + (isBullying ? '⚠️ Bullying/Cyberbullying Report' : '🧑‍🤝‍🧑 In-Person Report') + '</h2>' +
        '<p style="color:rgba(255,255,255,.7);margin:4px 0 0;font-size:13px;">Ephrata Middle School · ' + submittedAt + '</p>' +
      '</div>' +
      '<div style="background:#f3edf8;padding:14px 20px;border-left:4px solid #490e6f;">' +
        '<p style="margin:0;font-size:15px;font-weight:700;">' + firstName + ' ' + lastName + '</p>' +
        '<p style="margin:2px 0 0;font-size:13px;color:#52525b;">' + grade + ' grade &nbsp;·&nbsp; 800#: ' + (data.studentId || '—') + ' &nbsp;·&nbsp; Entered by ' + (data.enteredBy || '—') + '</p>' +
      '</div>' +
      '<div style="padding:16px 20px;background:#fff;border:1px solid #e4e4e7;">' +
        '<table style="width:100%;border-collapse:collapse;font-size:14px;">' +
          '<tr><td style="padding:4px 0;color:#52525b;width:80px;vertical-align:top;">Where</td><td style="padding:4px 0;font-weight:600;">' + (data.where || '—') + '</td></tr>' +
          '<tr><td style="padding:4px 0;color:#52525b;vertical-align:top;">When</td><td style="padding:4px 0;">' + (data.when || '—') + '</td></tr>' +
        '</table>' +
        '<div style="margin-top:12px;">' +
          '<p style="font-size:12px;font-weight:700;text-transform:uppercase;color:#490e6f;margin:0 0 4px;">What Happened</p>' +
          '<p style="margin:0;font-size:14px;background:#fafafa;padding:8px 10px;border-radius:6px;border:1px solid #e4e4e7;">' + (data.whatHappened || '—').replace(/\n/g,'<br>') + '</p>' +
        '</div>' +
      '</div>' +
      (isBullying ?
      '<div style="padding:14px 20px;background:#fee2e2;border:1px solid #e4e4e7;border-top:none;">' +
        '<p style="font-size:12px;font-weight:700;text-transform:uppercase;color:#b91c1c;margin:0 0 6px;">⚠ Flagged as Possible Bullying — ' + (data.bullyType || 'Type not specified') + '</p>' +
        '<p style="font-size:13px;margin:0;color:#52525b;">Alleged offender(s): ' + (data.bullyOffenders || '—') + ' &nbsp;·&nbsp; School: ' + (data.bullySchool || '—') + '</p>' +
      '</div>' : '') +
      '<div style="padding:14px 20px;background:#fef3c7;border:1px solid #e4e4e7;border-top:none;border-radius:0 0 8px 8px;">' +
        '<p style="font-size:12px;font-weight:700;text-transform:uppercase;color:#d97706;margin:0 0 6px;">⚡ Staff Follow-Up</p>' +
        '<p style="font-size:13px;margin:0;color:#52525b;">Status: ' + (data.status || 'open') + '. Open the dashboard for full details.</p>' +
      '</div>' +
      '</div>';

    MailApp.sendEmail({
      to: ALERT_EMAIL,
      subject: subject,
      body: body,
      htmlBody: htmlBody,
      name: 'EMS Incident Report System',
      replyTo: 'matthew_schuck@easdpa.org'
    });
  } catch (err) {
    // Don't let an email failure block the report from saving.
  }
}

function readStreamlinedReports() {
  var ss    = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(STREAMLINED_TAB_NAME);
  var results = [];

  if (!sheet) return results;

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return results;

  var rows = sheet.getRange(2, 1, lastRow - 1, 35).getValues();

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
      'Entered By':                row[13] || '',
      // ── BULLYING/CYBERBULLYING BRANCH ──
      'Believed to be Bullying?':              row[14] || '',
      'Complainant/Reporter Name':             row[15] || '',
      'Complainant Home Address':              row[16] || '',
      'Complainant Phone Number':              row[17] || '',
      'School Building':                       row[18] || '',
      'Date of Alleged Incident(s)':           row[19] || '',
      'Alleged Offender(s)':                   row[20] || '',
      'If Directed at Someone Else, Identify Them': row[21] || '',
      'Statements Made (threats/requests/demands)': row[22] || '',
      'Type of Bullying':                      row[23] || '',
      'Cyber: Platform/App/Method Used':       row[24] || '',
      'Cyber: Device Used':                    row[25] || '',
      'Cyber: Username(s)/Account(s) Involved': row[26] || '',
      'Cyber: During or Outside School Hours': row[27] || '',
      'Cyber: Date/Time of Online Activity':   row[28] || '',
      'Cyber: Evidence Saved':                 row[29] || '',
      'In-Person: Specific Location':          row[30] || '',
      'In-Person: During or Outside School Hours': row[31] || '',
      'In-Person: Physical, Verbal, or Both':  row[32] || '',
      'In-Person: Injuries/Physical Contact':  row[33] || '',
      'Certified True and Complete':           row[34] || ''
    });
  });

  return results;
}
