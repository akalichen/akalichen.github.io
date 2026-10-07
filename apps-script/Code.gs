/**
 * Electrobiotech Lab — CX43 microscope usage log + training backend
 * ------------------------------------------------------------------
 * Container-bound Google Apps Script. Lives inside the Google Sheet that
 * stores the records (Extensions > Apps Script). Deployed as a Web App that
 * both microscope_log.html (usage log) and training.html (quiz passes) POST to.
 *
 * On every submission it:
 *   1. appends one row to the "Log" sheet (usage) or "Training" sheet (quiz pass),
 *      creating headers if needed,
 *   2. exports the whole spreadsheet as .xlsx,
 *   3. emails the .xlsx (cumulative) + a summary of the new entry to NOTIFY_EMAIL,
 *   4. optionally sends the user a short confirmation.
 */

var CONFIG = {
  NOTIFY_EMAIL: 'cheng.li@oregonstate.edu',
  SEND_USER_CONFIRMATION: true,
  SHEET_NAME: 'Log',
  TRAINING_SHEET_NAME: 'Training',
  SUBJECT_PREFIX: '[CX43 log]',
  TRAINING_SUBJECT_PREFIX: '[CX43 training]',
  TIMEZONE: 'America/Los_Angeles'
};

// Column order in the sheet. Keys must match the form field names.
var FIELDS = [
  ['timestamp',          'Submitted (server time)'],
  ['name',               'Name'],
  ['email',              'Email'],
  ['affiliation',        'Lab / PI / affiliation'],
  ['role',               'Role'],
  ['trained',            'Trained'],
  ['date',               'Date of use'],
  ['start_time',         'Start'],
  ['end_time',           'End'],
  ['project',            'Project / purpose'],
  ['objectives',         'Objectives'],
  ['mode',               'Observation mode'],
  ['oil',                'Immersion oil'],
  ['oil_cleaned',        'Oil cleaned'],
  ['camera',             'EP50 camera'],
  ['camera_storage',     'Images saved to'],
  ['sample_type',        'Sample type'],
  ['biosafety',          'Biosafety / hazard'],
  ['sample_description', 'Sample description'],
  ['samples_removed',    'Samples removed & stage wiped'],
  ['condition',          'Condition on arrival'],
  ['notes',              'Issues / comments'],
  ['user_agent',         'Browser'],
  ['page',               'Submitted from']
];

// Columns for the Training sheet (quiz passes from training.html).
var TRAINING_FIELDS = [
  ['timestamp',   'Submitted (server time)'],
  ['name',        'Name'],
  ['email',       'Email'],
  ['affiliation', 'Lab / PI / affiliation'],
  ['course',      'Course'],
  ['score',       'Score'],
  ['total',       'Out of'],
  ['passed',      'Passed'],
  ['attempts',    'Attempts'],
  ['questions',   'Question IDs'],
  ['user_agent',  'Browser'],
  ['page',        'Submitted from']
];

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var payload = JSON.parse(e.postData.contents || '{}');

    // Spam / sanity checks
    if (payload.website) return respond({ ok: false, error: 'Rejected' });
    if (payload.form === 'training') return handleTraining_(payload);
    if (!payload.name || !payload.date) return respond({ ok: false, error: 'Missing required fields' });

    var sheet = getSheet_();
    var now = new Date();
    var row = FIELDS.map(function (f) {
      var k = f[0];
      if (k === 'timestamp') return now;
      var v = payload[k];
      if (v === undefined || v === null) return '';
      return String(v).slice(0, 1000);
    });
    sheet.appendRow(row);
    var rowNumber = sheet.getLastRow() - 1; // entry number (excluding header)

    sendNotification_(payload, rowNumber, now);
    if (CONFIG.SEND_USER_CONFIRMATION && isEmail_(payload.email)) sendUserConfirmation_(payload, rowNumber);

    return respond({ ok: true, row: rowNumber });
  } catch (err) {
    return respond({ ok: false, error: String(err && err.message || err) });
  } finally {
    lock.releaseLock();
  }
}

// Allows a quick browser check that the deployment is live.
function doGet() {
  return respond({ ok: true, service: 'CX43 usage log', entries: Math.max(getSheet_().getLastRow() - 1, 0) });
}

function respond(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function getSheet_(name, fields) {
  name = name || CONFIG.SHEET_NAME; fields = fields || FIELDS;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(name) || ss.insertSheet(name);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(fields.map(function (f) { return f[1]; }));
    sheet.getRange(1, 1, 1, fields.length).setFontWeight('bold').setBackground('#FAF0DC');
    sheet.setFrozenRows(1);
    sheet.getRange('A:A').setNumberFormat('yyyy-mm-dd hh:mm');
  }
  return sheet;
}

// ---------------------------------------------------------------------------
// Training quiz passes
// ---------------------------------------------------------------------------
function handleTraining_(p) {
  if (!p.name || !isEmail_(p.email) || String(p.passed) !== 'Yes') return respond({ ok: false, error: 'Missing required fields' });
  var sheet = getSheet_(CONFIG.TRAINING_SHEET_NAME, TRAINING_FIELDS);
  var now = new Date();
  var row = TRAINING_FIELDS.map(function (f) {
    if (f[0] === 'timestamp') return now;
    var v = p[f[0]];
    return (v === undefined || v === null) ? '' : String(v).slice(0, 1000);
  });
  sheet.appendRow(row);
  var rowNumber = sheet.getLastRow() - 1;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var when = Utilities.formatDate(now, CONFIG.TIMEZONE, 'yyyy-MM-dd HH:mm');
  var html =
    '<div style="font-family:Arial,sans-serif;font-size:14px;color:#1f2937">' +
    '<h2 style="color:#15803d;margin:0 0 6px">CX43 training passed — ' + escapeHtml_(p.name) + '</h2>' +
    '<p style="margin:0 0 12px;color:#555">' + when + ' · record #' + rowNumber + ' · total trained users: ' + rowNumber + '</p>' +
    '<table style="border-collapse:collapse">' +
    '<tr><td style="padding:4px 10px 4px 0;color:#555">Email</td><td>' + escapeHtml_(String(p.email)) + '</td></tr>' +
    '<tr><td style="padding:4px 10px 4px 0;color:#555">Affiliation</td><td>' + escapeHtml_(String(p.affiliation || '')) + '</td></tr>' +
    '<tr><td style="padding:4px 10px 4px 0;color:#555">Score</td><td>' + escapeHtml_(String(p.score)) + ' / ' + escapeHtml_(String(p.total)) + '</td></tr>' +
    '<tr><td style="padding:4px 10px 4px 0;color:#555">Attempts</td><td>' + escapeHtml_(String(p.attempts || 1)) + '</td></tr>' +
    '<tr><td style="padding:4px 10px 4px 0;color:#555">Questions</td><td style="color:#777;font-size:12px">' + escapeHtml_(String(p.questions || '')) + '</td></tr>' +
    '</table>' +
    '<p style="margin-top:16px">The complete spreadsheet (usage log + training records) is attached. Live sheet: <a href="' + ss.getUrl() + '">' + escapeHtml_(ss.getName()) + '</a></p></div>';
  var attachments = [];
  try { attachments.push(exportXlsx_()); } catch (err) { html += '<p style="color:#b91c1c">Attachment failed: ' + escapeHtml_(String(err)) + '</p>'; }
  MailApp.sendEmail({
    to: CONFIG.NOTIFY_EMAIL,
    subject: CONFIG.TRAINING_SUBJECT_PREFIX + ' PASS — ' + p.name + ' (' + p.score + '/' + p.total + ')',
    htmlBody: html, attachments: attachments, name: 'Electrobiotech Lab usage log'
  });
  if (CONFIG.SEND_USER_CONFIRMATION) {
    MailApp.sendEmail({
      to: p.email,
      subject: 'CX43 microscope training — passed',
      htmlBody: '<p>Hi ' + escapeHtml_(p.name) + ',</p><p>You passed the CX43 microscope training quiz (' + escapeHtml_(String(p.score)) + '/' + escapeHtml_(String(p.total)) + ') on ' + when + '. Your completion is on record.</p>' +
        '<p>Please fill in the usage log for every session: https://www.electrobiotechlab.com/microscope_log.html</p><p>— Electrobiotech Lab, Oregon State University</p>',
      name: 'Electrobiotech Lab usage log'
    });
  }
  return respond({ ok: true, row: rowNumber });
}

function exportXlsx_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var url = 'https://docs.google.com/spreadsheets/d/' + ss.getId() + '/export?format=xlsx';
  var resp = UrlFetchApp.fetch(url, {
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true
  });
  if (resp.getResponseCode() !== 200) throw new Error('xlsx export failed: HTTP ' + resp.getResponseCode());
  var stamp = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy-MM-dd_HHmm');
  return resp.getBlob().setName('CX43_usage_log_' + stamp + '.xlsx');
}

function sendNotification_(p, rowNumber, now) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var when = Utilities.formatDate(now, CONFIG.TIMEZONE, 'yyyy-MM-dd HH:mm');
  var rows = FIELDS.filter(function (f) { return f[0] !== 'timestamp' && f[0] !== 'user_agent' && f[0] !== 'page'; })
    .map(function (f) {
      var v = p[f[0]] ? escapeHtml_(String(p[f[0]])) : '<span style="color:#999">—</span>';
      return '<tr><td style="padding:4px 10px 4px 0;color:#555;white-space:nowrap">' + f[1] + '</td><td style="padding:4px 0">' + v + '</td></tr>';
    }).join('');

  var flags = [];
  if (p.oil === 'Yes' && p.oil_cleaned !== 'Yes, with lens paper') flags.push('Immersion oil used and not yet confirmed cleaned.');
  if (p.condition && p.condition !== 'Clean and working') flags.push('Condition on arrival: ' + escapeHtml_(p.condition));
  if (p.notes) flags.push('User left a comment.');

  var html =
    '<div style="font-family:Arial,sans-serif;font-size:14px;color:#1f2937">' +
    '<h2 style="color:#DC4405;margin:0 0 6px">CX43 usage log — entry #' + rowNumber + '</h2>' +
    '<p style="margin:0 0 12px;color:#555">Submitted ' + when + ' · total entries: ' + rowNumber + '</p>' +
    (flags.length ? '<p style="background:#FAF0DC;border-left:4px solid #D3832B;padding:8px 12px">' + flags.join('<br>') + '</p>' : '') +
    '<table style="border-collapse:collapse">' + rows + '</table>' +
    '<p style="margin-top:16px">The complete, cumulative log is attached as an Excel file. Live sheet: <a href="' + ss.getUrl() + '">' + escapeHtml_(ss.getName()) + '</a></p>' +
    '</div>';

  var subject = CONFIG.SUBJECT_PREFIX + ' #' + rowNumber + ' — ' + p.name + ' · ' + p.date + (p.start_time ? ' ' + p.start_time : '');

  var attachments = [];
  try { attachments.push(exportXlsx_()); } catch (err) { html += '<p style="color:#b91c1c">Attachment failed: ' + escapeHtml_(String(err)) + '</p>'; }

  MailApp.sendEmail({
    to: CONFIG.NOTIFY_EMAIL,
    subject: subject,
    htmlBody: html,
    attachments: attachments,
    name: 'Electrobiotech Lab usage log'
  });
}

function sendUserConfirmation_(p, rowNumber) {
  MailApp.sendEmail({
    to: p.email,
    subject: 'CX43 microscope session logged (#' + rowNumber + ')',
    htmlBody:
      '<p>Hi ' + escapeHtml_(p.name) + ',</p>' +
      '<p>Your CX43 microscope session on <b>' + escapeHtml_(p.date) + '</b>' + (p.start_time ? ' starting ' + escapeHtml_(p.start_time) : '') +
      ' has been logged (entry #' + rowNumber + ').</p>' +
      '<p>Reminder: clean immersion oil from the objectives with lens paper, switch off the illuminator and camera, cover the microscope, and remove your samples.</p>' +
      '<p>— Electrobiotech Lab, Oregon State University</p>',
    name: 'Electrobiotech Lab usage log'
  });
}

function isEmail_(s) { return typeof s === 'string' && /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(s); }

function escapeHtml_(s) {
  return s.replace(/[&<>"']/g, function (c) {
    return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
  });
}

/**
 * Run this ONCE from the editor (Run > testSubmission) to trigger the
 * authorization prompt and confirm the email + attachment arrive.
 * Delete the test row from the sheet afterwards if you like.
 */
function testTrainingSubmission() {
  var fake = { postData: { contents: JSON.stringify({
    form: 'training', course: 'CX43 microscope training', name: 'Test User', email: CONFIG.NOTIFY_EMAIL,
    affiliation: 'Li Lab', score: 7, total: 7, passed: 'Yes', attempts: 1, questions: 'oil-band; bft-focus; na-def', user_agent: 'Apps Script test', page: 'editor'
  }) } };
  Logger.log(doPost(fake).getContent());
}

function testSubmission() {
  var fake = {
    postData: {
      contents: JSON.stringify({
        name: 'Test User', email: CONFIG.NOTIFY_EMAIL, affiliation: 'Li Lab', role: 'Faculty', trained: 'Yes',
        date: Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy-MM-dd'), start_time: '09:00', end_time: '09:30',
        project: 'Deployment test', objectives: '10x; 40x', mode: 'Phase contrast', oil: 'No', camera: 'Yes',
        camera_storage: 'USB drive', sample_type: 'Biological — live culture / enrichment', biosafety: 'BSL-1',
        sample_description: 'Test entry — safe to delete', samples_removed: 'Yes', condition: 'Clean and working',
        notes: '', user_agent: 'Apps Script test', page: 'editor'
      })
    }
  };
  Logger.log(doPost(fake).getContent());
}
