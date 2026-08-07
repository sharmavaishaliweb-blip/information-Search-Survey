/**
 * Online Information Search Behaviour Experimental Platform
 * Google Apps Script backend — Sheets + Drive
 *
 * QUICK SETUP (about 5 minutes):
 * 1. Open https://script.google.com → New project
 * 2. Delete any stub code and paste THIS entire file
 * 3. Save the project (name it e.g. "Search Survey Backend")
 * 4. Select setupAll in the function dropdown → Run
 *    Approve Google permissions when prompted
 *    (Creates the Google Sheet + Drive folder automatically)
 * 5. View → Execution log / Logger to copy the Sheet URL
 *    Or run getSetupInfo and check the log
 * 6. Deploy → New deployment → Type: Web app
 *      - Description: survey-v1
 *      - Execute as: Me
 *      - Who has access: Anyone
 * 7. Copy the Web App URL
 * 8. Paste it into index.html → APPS_SCRIPT_URL
 *
 * Each survey Submit will:
 *  - Append one row to the Responses sheet
 *  - Upload the screen recording to the Drive folder
 *  - Save the Drive file URL in that same row
 */

var SHEET_NAME = 'Responses';
var FOLDER_NAME = 'Information Search Survey Recordings';

var HEADERS = [
  'Timestamp',
  'Full Name',
  'Age',
  'Gender',
  'Education',
  'Online Purchase Frequency',
  'Selected Ad',
  'FB Q1 Familiar',
  'FB Q2 Know Well',
  'FB Q3 Know More',
  'FB Q4 Provided Facts',
  'FB Q5 Practical Info',
  'Sem Unimportant–Important',
  'Sem Irrelevant–Relevant',
  'Sem Means Nothing–Means A Lot',
  'Sem Worthless–Valuable',
  'Sem Not Needed–Needed',
  'PI Q7 Consider Purchase',
  'PI Q8 Likelihood High',
  'PI Q9 Willingness High',
  'PI Q10 Probability High',
  'Google Searches',
  'Chat Messages',
  'Site Clicks',
  'Official Links Clicked',
  'Browsed URLs',
  'Metrics Summary JSON',
  'Metrics Full JSON',
  'SN1 Google', 'SN2 Google', 'TE1 Google', 'TE2 Google', 'CT1 Google', 'CT2 Google', 'QD1 Google', 'QD2 Google',
  'SN1 Direct', 'SN2 Direct', 'TE1 Direct', 'TE2 Direct', 'CT1 Direct', 'CT2 Direct', 'QD1 Direct', 'QD2 Direct',
  'SN1 AI', 'SN2 AI', 'TE1 AI', 'TE2 AI', 'CT1 AI', 'CT2 AI', 'QD1 AI', 'QD2 AI',
  'Drive Video URL',
  'Drive File ID'
];

/**
 * Run once after pasting this script.
 * Creates spreadsheet + Drive folder and stores their IDs in Script Properties.
 */
function setupAll() {
  var props = PropertiesService.getScriptProperties();

  var ss = SpreadsheetApp.create('Online Information Search Behaviour — Responses');
  var sheet = ss.getSheets()[0];
  sheet.setName(SHEET_NAME);
  sheet.clear();
  sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
  sheet.autoResizeColumns(1, Math.min(HEADERS.length, 15));

  var folder = DriveApp.createFolder(FOLDER_NAME);

  props.setProperty('SPREADSHEET_ID', ss.getId());
  props.setProperty('DRIVE_FOLDER_ID', folder.getId());

  var info = {
    ok: true,
    spreadsheetId: ss.getId(),
    spreadsheetUrl: ss.getUrl(),
    driveFolderId: folder.getId(),
    driveFolderUrl: folder.getUrl()
  };

  Logger.log(JSON.stringify(info, null, 2));
  return info;
}

/** Safe to re-run — prints current Sheet/Drive links. */
function getSetupInfo() {
  var props = PropertiesService.getScriptProperties();
  var ssId = props.getProperty('SPREADSHEET_ID');
  var folderId = props.getProperty('DRIVE_FOLDER_ID');
  var info = {
    spreadsheetId: ssId || null,
    spreadsheetUrl: ssId ? ('https://docs.google.com/spreadsheets/d/' + ssId + '/edit') : null,
    driveFolderId: folderId || null,
    driveFolderUrl: folderId ? ('https://drive.google.com/drive/folders/' + folderId) : null
  };
  Logger.log(JSON.stringify(info, null, 2));
  return info;
}

function ensureSheetHeaders_(sheet) {
  var lastCol = sheet.getLastColumn();
  if (lastCol < 1) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
    return;
  }
  var existing = sheet.getRange(1, 1, 1, Math.max(lastCol, HEADERS.length)).getValues()[0];
  var needsReset = false;
  for (var i = 0; i < HEADERS.length; i++) {
    if (String(existing[i] || '') !== HEADERS[i]) {
      needsReset = true;
      break;
    }
  }
  if (needsReset && sheet.getLastRow() <= 1) {
    sheet.clear();
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
  }
}

function getSpreadsheet_() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('SPREADSHEET_ID');
  if (!id) {
    throw new Error('Spreadsheet not configured. Run setupAll() once in the Apps Script editor.');
  }
  return SpreadsheetApp.openById(id);
}

function getDriveFolder_() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('DRIVE_FOLDER_ID');
  if (!id) {
    var folder = DriveApp.createFolder(FOLDER_NAME);
    props.setProperty('DRIVE_FOLDER_ID', folder.getId());
    return folder;
  }
  return DriveApp.getFolderById(id);
}

function doGet() {
  var info = {};
  try { info = getSetupInfo(); } catch (e) {}
  return jsonResponse_({
    ok: true,
    message: 'Information Search Survey endpoint is live. Use POST to submit participant data + video.',
    setup: info
  });
}

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return jsonResponse_({ ok: false, error: 'Empty request body' });
    }

    var payload = JSON.parse(e.postData.contents);
    var answers = payload.answers || {};
    var videoBase64 = payload.videoBase64 || '';
    var mimeType = payload.mimeType || 'video/webm';
    var safeName = String((answers.fullName || 'participant')).replace(/[^\w\-]+/g, '_').slice(0, 40);
    var fileName = payload.fileName || ('survey-' + safeName + '-' + Date.now() + '.webm');

    var driveUrl = '';
    var driveFileId = '';

    if (videoBase64) {
      var cleaned = String(videoBase64).replace(/^data:[^;]+;base64,/, '');
      // Apps Script practical limit ~35–45MB decoded; fail clearly if huge
      if (cleaned.length > 60 * 1024 * 1024) {
        return jsonResponse_({
          ok: false,
          error: 'Recording is too large for Apps Script upload. Ask participant to share a shorter screen region or re-record briefly.'
        });
      }
      var bytes = Utilities.base64Decode(cleaned);
      var blob = Utilities.newBlob(bytes, mimeType, fileName);
      var folder = getDriveFolder_();
      var file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      driveFileId = file.getId();
      driveUrl = file.getUrl();
    }

    var summary = (answers.searchMetrics && answers.searchMetrics.summary) ? answers.searchMetrics.summary : {};
    var g = summary.google || {};
    var d = summary.direct || {};
    var a = summary.ai || {};

    var row = [
      new Date().toISOString(),
      answers.fullName || '',
      answers.age || '',
      answers.gender || '',
      answers.education || '',
      answers.frequency || '',
      answers.selectedAd || '',
      answers.fb_q1 || '',
      answers.fb_q2 || '',
      answers.fb_q3 || '',
      answers.fb_q4 || '',
      answers.fb_q5 || '',
      answers.sem_1 || '',
      answers.sem_2 || '',
      answers.sem_3 || '',
      answers.sem_4 || '',
      answers.sem_5 || '',
      answers.pi_q7 || '',
      answers.pi_q8 || '',
      answers.pi_q9 || '',
      answers.pi_q10 || '',
      stringifyList_(answers.googleSearches),
      stringifyList_(answers.chatMessages),
      stringifyList_(answers.siteClicks),
      stringifyList_(answers.officialClicks),
      stringifyList_(answers.browsedUrls),
      summary && Object.keys(summary).length ? JSON.stringify(summary) : '',
      answers.searchMetrics ? JSON.stringify(answers.searchMetrics) : '',
      nullToBlank_(g.SN1), nullToBlank_(g.SN2), nullToBlank_(g.TE1_sec), nullToBlank_(g.TE2_sec), nullToBlank_(g.CT1), nullToBlank_(g.CT2), nullToBlank_(g.QD1), nullToBlank_(g.QD2),
      nullToBlank_(d.SN1), nullToBlank_(d.SN2), nullToBlank_(d.TE1_sec), nullToBlank_(d.TE2_sec), nullToBlank_(d.CT1), nullToBlank_(d.CT2), nullToBlank_(d.QD1), nullToBlank_(d.QD2),
      nullToBlank_(a.SN1), nullToBlank_(a.SN2), nullToBlank_(a.TE1_sec), nullToBlank_(a.TE2_sec), nullToBlank_(a.CT1), nullToBlank_(a.CT2), nullToBlank_(a.QD1), nullToBlank_(a.QD2),
      driveUrl,
      driveFileId
    ];

    var ss = getSpreadsheet_();
    var sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
    }
    ensureSheetHeaders_(sheet);
    sheet.appendRow(row);

    return jsonResponse_({
      ok: true,
      driveUrl: driveUrl,
      driveFileId: driveFileId,
      spreadsheetUrl: ss.getUrl()
    });
  } catch (err) {
    return jsonResponse_({ ok: false, error: String(err) });
  }
}

function nullToBlank_(v) {
  return (v === null || typeof v === 'undefined') ? '' : v;
}

function jsonResponse_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function stringifyList_(value) {
  if (!value) return '';
  if (Object.prototype.toString.call(value) !== '[object Array]') {
    return String(value);
  }
  return value.map(function (item) {
    if (item == null) return '';
    if (typeof item === 'string' || typeof item === 'number' || typeof item === 'boolean') {
      return String(item);
    }
    return JSON.stringify(item);
  }).join(' | ');
}
