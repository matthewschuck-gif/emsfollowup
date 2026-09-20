/**
 * EMS Student Incident Report — FORM-BOUND SCRIPT
 *
 * WHERE THIS GOES: open the actual Google Form students use (not the
 * response Sheet) → Extensions → Apps Script. This project is bound to
 * the Form and its onFormSubmit trigger, and is the ONLY script that
 * needs to live there.
 *
 * WHAT IT DOES: emails ALERT_EMAIL every time a student submits the
 * incident report, and (one-time) can add the "Understanding Bullying"
 * branch to the form itself.
 */

// ── EMAIL RECIPIENTS ──────────────────────────────────────────────
// To add more recipients, separate with commas — no spaces
var ALERT_EMAIL = 'Christopher_Montagna@easdpa.org, Kevin_Kuhn@easdpa.org, J_Lugar@easdpa.org, Kristin_Mincarelli@easdpa.org, A_Mowbray@easdpa.org, Matthew_schuck@easdpa.org, K_Wagner@easdpa.org, Kelly_Rigg@easdpa.org, Sara_Judge@easdpa.org, Lindsey_Carr@easdpa.org, Nicole_Zimmerman@easdpa.org, Donna_Schlinkman@easdpa.org, Kristin_Mincarelli@easdpa.org,matthew_schuck@easdpa.org';
// ─────────────────────────────────────────────────────────────────

// The exact text of the branching question added by
// addBullyingBranchToExistingForm() below — used to detect the answer
// in onIncidentSubmit() and by the staff dashboard (index.html).
var BULLYING_QUESTION_TITLE = 'Based on this definition, do you believe what you are reporting is bullying or cyberbullying?';

/**
 * Fires on every form submission.
 * DO NOT rename this function — the trigger references it by name.
 */
function onIncidentSubmit(e) {
  try {
    var r   = e.response;
    var ts  = r.getTimestamp();
    var ans = r.getItemResponses();

    var data = {};
    ans.forEach(function(item) {
      data[item.getItem().getTitle()] = item.getResponse() || '—';
    });

    var studentId    = data['800 Number (Student ID)']   || '—';
    var firstName    = data['First Name']                 || '—';
    var lastName     = data['Last Name']                  || '—';
    var grade        = data['Grade']                      || '—';
    var homeroom     = data['Homeroom Teacher']            || '—';
    var dateStr      = data["Today's Date"]               || '—';
    var where        = data['Where did it happen?']       || '—';
    var when         = data['When did it happen?']        || '—';
    var involved     = data['Who was involved?']          || '—';
    var witnesses    = data['Who witnessed it?']          || '—';
    var evidence     = data['Is there any evidence?']     || '—';
    var whatHappened = data['What happened?']             || '—';
    var nextSteps    = data['What do you think needs to happen to resolve this situation?'] || '—';

    // ── Bullying/cyberbullying branch (only present if the student saw
    // and answered it — see addBullyingBranchToExistingForm() below) ──
    var isBullying          = data[BULLYING_QUESTION_TITLE] === 'Yes';
    var bullyPattern         = data['Has this happened more than once, or is it part of an ongoing pattern?'] || '—';
    var bullyToldStop        = data['Have you personally told the other student(s) to stop?'] || '—';
    var bullyToldStopDetail  = data['If yes, what happened when you told them to stop?'] || '';
    var bullyOthers          = data['Has anyone else (another student, teacher, or staff member) already talked to the other student(s) about this or told them to stop?'] || '—';
    var bullyOneSided        = data['Is this one-sided, or do you also respond to, start, or take part in comments or actions toward the other student(s)?'] || '—';
    var bullyOneSidedDetail  = data['If you also take part, please briefly describe what you have said or done.'] || '';

    var submittedAt = Utilities.formatDate(ts, Session.getScriptTimeZone(), 'MMMM d, yyyy \'at\' h:mm a');

    var subject = (isBullying ? '⚠️ Bullying/Cyberbullying Reported — ' : '⚠️ Incident Report Submitted — ')
      + firstName + ' ' + lastName + ' (' + grade + ' grade)';

    var body =
      'An incident report was submitted at ' + submittedAt + '.\n\n' +
      '══════════════════════════════════════\n' +
      'STUDENT INFORMATION\n' +
      '══════════════════════════════════════\n' +
      '800 Number:    ' + studentId   + '\n' +
      'Name:          ' + firstName + ' ' + lastName + '\n' +
      'Grade:         ' + grade       + '\n' +
      'Homeroom:      ' + homeroom    + '\n' +
      'Date entered:  ' + dateStr     + '\n\n' +
      '══════════════════════════════════════\n' +
      'INCIDENT DETAILS\n' +
      '══════════════════════════════════════\n' +
      'Where:         ' + where       + '\n' +
      'When:          ' + when        + '\n\n' +
      'Who was involved:\n' + involved + '\n\n' +
      'Witnesses:\n' + witnesses + '\n\n' +
      'Evidence:\n' + evidence + '\n\n' +
      'What happened:\n' + whatHappened + '\n\n' +
      (isBullying ?
        '══════════════════════════════════════\n' +
        'BULLYING/CYBERBULLYING (student self-reported)\n' +
        '══════════════════════════════════════\n' +
        'Ongoing pattern?            ' + bullyPattern + '\n' +
        'Reporter told them to stop? ' + bullyToldStop + (bullyToldStopDetail ? ' — ' + bullyToldStopDetail : '') + '\n' +
        'Someone else addressed it?  ' + bullyOthers + '\n' +
        'One-sided or mutual?        ' + bullyOneSided + (bullyOneSidedDetail ? ' — ' + bullyOneSidedDetail : '') + '\n\n'
        : '') +
      '══════════════════════════════════════\n' +
      'STUDENT\'S PERSPECTIVE\n' +
      '══════════════════════════════════════\n' +
      'What should happen:\n' + nextSteps + '\n\n' +
      '══════════════════════════════════════\n' +
      'STAFF FOLLOW-UP REQUIRED\n' +
      '══════════════════════════════════════\n' +
      '  • Follow-up meeting (adult name + date)\n' +
      '  • Notes\n' +
      '  • Logged in PowerSchool (Yes/No)\n' +
      '  • SRO Contacted (Yes/No)\n' +
      '  • Parent/Guardian contacted (Phone/Email/In Person/None)\n' +
      '  • Teachers communicated with (Yes/No)\n\n';

    var htmlBody =
      '<div style="font-family:Arial,sans-serif;max-width:600px;color:#1a1a1a;">' +
      '<div style="background:#490e6f;padding:16px 20px;border-radius:8px 8px 0 0;">' +
        '<h2 style="color:#ffe100;margin:0;font-size:18px;">⚠️ Incident Report Submitted</h2>' +
        '<p style="color:rgba(255,255,255,.7);margin:4px 0 0;font-size:13px;">Ephrata Middle School · ' + submittedAt + '</p>' +
      '</div>' +
      '<div style="background:#f3edf8;padding:14px 20px;border-left:4px solid #490e6f;">' +
        '<p style="margin:0;font-size:15px;font-weight:700;">' + firstName + ' ' + lastName + '</p>' +
        '<p style="margin:2px 0 0;font-size:13px;color:#52525b;">' + grade + ' grade &nbsp;·&nbsp; 800#: ' + studentId + ' &nbsp;·&nbsp; Homeroom: ' + homeroom + '</p>' +
      '</div>' +
      '<div style="padding:16px 20px;background:#fff;border:1px solid #e4e4e7;">' +
        '<h3 style="font-size:12px;text-transform:uppercase;letter-spacing:.5px;color:#490e6f;margin:0 0 10px;">Incident Details</h3>' +
        '<table style="width:100%;border-collapse:collapse;font-size:14px;">' +
          '<tr><td style="padding:4px 0;color:#52525b;width:80px;vertical-align:top;">Where</td><td style="padding:4px 0;font-weight:600;">' + where + '</td></tr>' +
          '<tr><td style="padding:4px 0;color:#52525b;vertical-align:top;">When</td><td style="padding:4px 0;">' + when + '</td></tr>' +
        '</table>' +
        '<div style="margin-top:12px;">' +
          '<p style="font-size:12px;font-weight:700;text-transform:uppercase;color:#490e6f;margin:0 0 4px;">Who Was Involved</p>' +
          '<p style="margin:0;font-size:14px;background:#fafafa;padding:8px 10px;border-radius:6px;border:1px solid #e4e4e7;">' + involved.replace(/\n/g,'<br>') + '</p>' +
        '</div>' +
        '<div style="margin-top:10px;">' +
          '<p style="font-size:12px;font-weight:700;text-transform:uppercase;color:#490e6f;margin:0 0 4px;">Witnesses</p>' +
          '<p style="margin:0;font-size:14px;background:#fafafa;padding:8px 10px;border-radius:6px;border:1px solid #e4e4e7;">' + witnesses.replace(/\n/g,'<br>') + '</p>' +
        '</div>' +
        (evidence && evidence !== '—' ?
        '<div style="margin-top:10px;">' +
          '<p style="font-size:12px;font-weight:700;text-transform:uppercase;color:#490e6f;margin:0 0 4px;">Evidence</p>' +
          '<p style="margin:0;font-size:14px;background:#fafafa;padding:8px 10px;border-radius:6px;border:1px solid #e4e4e7;">' + evidence.replace(/\n/g,'<br>') + '</p>' +
        '</div>' : '') +
        '<div style="margin-top:10px;">' +
          '<p style="font-size:12px;font-weight:700;text-transform:uppercase;color:#490e6f;margin:0 0 4px;">What Happened</p>' +
          '<p style="margin:0;font-size:14px;background:#fafafa;padding:8px 10px;border-radius:6px;border:1px solid #e4e4e7;">' + whatHappened.replace(/\n/g,'<br>') + '</p>' +
        '</div>' +
      '</div>' +
      (isBullying ?
      '<div style="padding:14px 20px;background:#fee2e2;border:1px solid #e4e4e7;border-top:none;">' +
        '<p style="font-size:12px;font-weight:700;text-transform:uppercase;color:#b91c1c;margin:0 0 6px;">⚠ Bullying/Cyberbullying (student self-reported)</p>' +
        '<table style="width:100%;border-collapse:collapse;font-size:13px;color:#52525b;">' +
          '<tr><td style="padding:2px 0;width:170px;vertical-align:top;">Ongoing pattern?</td><td style="padding:2px 0;">' + bullyPattern + '</td></tr>' +
          '<tr><td style="padding:2px 0;vertical-align:top;">Told them to stop?</td><td style="padding:2px 0;">' + bullyToldStop + (bullyToldStopDetail ? ' — ' + bullyToldStopDetail : '') + '</td></tr>' +
          '<tr><td style="padding:2px 0;vertical-align:top;">Someone else addressed it?</td><td style="padding:2px 0;">' + bullyOthers + '</td></tr>' +
          '<tr><td style="padding:2px 0;vertical-align:top;">One-sided or mutual?</td><td style="padding:2px 0;">' + bullyOneSided + (bullyOneSidedDetail ? ' — ' + bullyOneSidedDetail : '') + '</td></tr>' +
        '</table>' +
      '</div>' : '') +
      (nextSteps && nextSteps !== '—' ?
      '<div style="padding:14px 20px;background:#fff8dc;border:1px solid #e4e4e7;border-top:none;">' +
        '<p style="font-size:12px;font-weight:700;text-transform:uppercase;color:#c9a800;margin:0 0 4px;">Student\'s Perspective</p>' +
        '<p style="margin:0;font-size:14px;">' + nextSteps.replace(/\n/g,'<br>') + '</p>' +
      '</div>' : '') +
      '<div style="padding:14px 20px;background:#fef3c7;border:1px solid #e4e4e7;border-top:none;">' +
        '<p style="font-size:12px;font-weight:700;text-transform:uppercase;color:#d97706;margin:0 0 6px;">⚡ Staff Follow-Up Required</p>' +
        '<p style="font-size:13px;margin:0;color:#52525b;">Complete the staff section: follow-up meeting, PS log, SRO contact, parent contact, teacher communication.</p>' +
      '</div>' +
      '<div style="padding:12px 20px;background:#490e6f;border-radius:0 0 8px 8px;text-align:center;">' +
  '<a href="https://bit.ly/emsfollowup" style="color:#ffe100;font-size:13px;font-weight:700;text-decoration:none;">Open Follow-Up Dashboard → bit.ly/emsfollowup</a>' +
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

  } catch(err) {
    Logger.log('Email send error: ' + err.toString());
  }
}


// ══════════════════════════════════════════════════════════════════════
// ONE-TIME SETUP — run addBullyingBranchToExistingForm() a single time
// from this same script project (it's already bound to the form, so no
// Form ID needed) to insert the "Understanding Bullying" branch into the
// live form. Re-running it a second time would duplicate the questions,
// so only run it once; if you ever need to change the wording afterward,
// edit the questions directly in the Form editor instead of re-running this.
// ══════════════════════════════════════════════════════════════════════

var BULLYING_DEFINITION =
  'Ephrata Middle School defines bullying/cyberbullying (Board Policy 249) as an intentional ' +
  'electronic, written, verbal, or physical act or series of acts directed at another student ' +
  '(or students) that:\n\n' +
  '• is SEVERE, PERSISTENT, or PERVASIVE — meaning it keeps happening again and again over time, ' +
  'or is widespread and hard to get away from. This is usually a PATTERN, not just a single ' +
  'one-time disagreement or argument, and\n' +
  "• has the effect of hurting the other student's ability to feel safe or to learn, creating a " +
  'threatening or hostile environment, or seriously disrupting school.\n\n' +
  'There is usually a power difference between the student(s) doing it and the student it is ' +
  'happening to, and it can happen in person, in writing, or electronically/online ' +
  '(cyberbullying) — during school, on a bus, at a school activity, or off-campus if it affects ' +
  'the school environment.\n\n' +
  'Please answer honestly based on this definition — there is no wrong answer, and a staff ' +
  'member will follow up with you either way.';

// Builds every item for the bullying branch (the "Understanding Bullying"
// page, its Yes/No question, the "Bullying/Cyberbullying Details" page,
// and a fresh "Next Steps" page break) and wires the Yes/No navigation.
// Returns handles the caller uses to set final item order with moveItem().
function buildBullyingBranchItems_(form, nextStepsTitle, nextStepsHelp) {
  var pageBullyDetails = form.addPageBreakItem()
    .setTitle('Bullying/Cyberbullying Details')
    .setHelpText('Thank you for being honest. These questions help staff understand the full situation.');

  var pageNextSteps = form.addPageBreakItem()
    .setTitle(nextStepsTitle)
    .setHelpText(nextStepsHelp);

  var pageUnderstanding = form.addPageBreakItem()
    .setTitle('Understanding Bullying')
    .setHelpText(BULLYING_DEFINITION);

  var bullyQuestion = form.addMultipleChoiceItem();
  bullyQuestion.setTitle(BULLYING_QUESTION_TITLE)
    .setChoices([
      bullyQuestion.createChoice('Yes', pageBullyDetails),
      bullyQuestion.createChoice('No', pageNextSteps)
    ])
    .setRequired(true);

  var qPattern = form.addMultipleChoiceItem()
    .setTitle('Has this happened more than once, or is it part of an ongoing pattern?')
    .setChoiceValues([
      'Yes — this has happened more than once / it is ongoing',
      'No — this was a one-time incident',
      "I'm not sure"
    ])
    .setRequired(true);

  var qToldToStop = form.addMultipleChoiceItem()
    .setTitle('Have you personally told the other student(s) to stop?')
    .setChoiceValues(['Yes', 'No'])
    .setRequired(true);

  var qToldToStopDetail = form.addParagraphTextItem()
    .setTitle('If yes, what happened when you told them to stop?')
    .setRequired(false);

  var qOthersAddressed = form.addMultipleChoiceItem()
    .setTitle('Has anyone else (another student, teacher, or staff member) already talked to the other student(s) about this or told them to stop?')
    .setChoiceValues(['Yes', 'No', "I don't know"])
    .setRequired(true);

  var qOneSided = form.addMultipleChoiceItem()
    .setTitle('Is this one-sided, or do you also respond to, start, or take part in comments or actions toward the other student(s)?')
    .setChoiceValues([
      'This is one-sided — I do not respond or take part',
      'We both take part / respond to each other',
      "I'm not sure"
    ])
    .setRequired(true);

  var qOneSidedDetail = form.addParagraphTextItem()
    .setTitle('If you also take part, please briefly describe what you have said or done.')
    .setRequired(false);

  return {
    pageUnderstanding: pageUnderstanding,
    bullyQuestion: bullyQuestion,
    pageBullyDetails: pageBullyDetails,
    qPattern: qPattern,
    qToldToStop: qToldToStop,
    qToldToStopDetail: qToldToStopDetail,
    qOthersAddressed: qOthersAddressed,
    qOneSided: qOneSided,
    qOneSidedDetail: qOneSidedDetail,
    pageNextSteps: pageNextSteps
  };
}

function findItemByTitle_(form, title) {
  var items = form.getItems();
  for (var i = 0; i < items.length; i++) {
    if (items[i].getTitle() === title) return items[i];
  }
  return null;
}

/**
 * Run this ONCE, from this same script project (Extensions → Apps Script
 * on the actual form — it's already bound, so FormApp.getActiveForm()
 * finds it automatically, no ID to paste).
 *
 * Inserts the "Understanding Bullying" branch into the live form in
 * place, right after "What happened?" and before "Next Steps", so
 * existing responses and the existing response sheet are untouched —
 * new columns just appear the next time someone answers the new questions.
 */
function addBullyingBranchToExistingForm() {
  var form = FormApp.getActiveForm();

  // Idempotency guard: if a prior run already added the branch, don't add
  // a second copy. If you deliberately want to re-add it (e.g. after fully
  // deleting every trace of it from the form), delete this check.
  if (findItemByTitle_(form, BULLYING_QUESTION_TITLE)) {
    Logger.log('Bullying branch already present on this form -- nothing to do. ' +
      'If you want to rebuild it, first delete every bullying-related item from the form.');
    return;
  }

  var whatHappenedItem  = findItemByTitle_(form, 'What happened?');
  var nextStepsQuestion = findItemByTitle_(form, 'What do you think needs to happen to resolve this situation?');
  // "Next Steps" may currently be the original SectionHeaderItem, or (if a
  // previous run got partway through) already a leftover PageBreakItem --
  // check the type before assuming which one it is.
  var nextStepsItem = findItemByTitle_(form, 'Next Steps');

  if (!whatHappenedItem || !nextStepsQuestion) {
    throw new Error(
      'Could not find "What happened?" and/or "What do you think needs to happen..." on this ' +
      'form. Confirm this script is bound to the live incident form and that its question ' +
      'titles have not changed, or adjust the titles findItemByTitle_() searches for.'
    );
  }

  var nextStepsHelp = 'Your perspective matters. Help us understand what you think should happen.';
  if (nextStepsItem && nextStepsItem.getType() === FormApp.ItemType.SECTION_HEADER) {
    nextStepsHelp = nextStepsItem.asSectionHeaderItem().getHelpText();
  }

  var bully = buildBullyingBranchItems_(form, 'Next Steps', nextStepsHelp);

  // The new pageNextSteps page break carries the same title/help text, so
  // whatever used to occupy that spot (the old cosmetic header, or a stray
  // page break from an earlier partial run) is now redundant.
  if (nextStepsItem) form.deleteItem(nextStepsItem);

  var afterWhatHappenedIndex = whatHappenedItem.getIndex() + 1;

  // moveItem(Item, Integer) rejects the specific subtypes addPageBreakItem()/
  // addMultipleChoiceItem() return (a quirk of Apps Script's overload
  // matching), so use the (fromIndex, toIndex) integer overload instead,
  // re-reading each item's current index right before moving it.
  [
    bully.pageUnderstanding, bully.bullyQuestion,
    bully.pageBullyDetails, bully.qPattern, bully.qToldToStop, bully.qToldToStopDetail,
    bully.qOthersAddressed, bully.qOneSided, bully.qOneSidedDetail,
    bully.pageNextSteps
  ].forEach(function(item, i) {
    form.moveItem(item.getIndex(), afterWhatHappenedIndex + i);
  });

  Logger.log('Bullying branch added successfully.');
  Logger.log('Open the form to review: ' + form.getEditUrl());
}
