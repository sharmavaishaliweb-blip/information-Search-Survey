/**
 * Information Search Survey — Google Apps Script backend
 *
 * SETUP (use the Gmail that should own the Sheet + Drive files):
 * 1. Go to https://sheets.google.com and create a blank spreadsheet
 *    named "Information Search Survey Responses" (any name is fine).
 * 2. Extensions → Apps Script. Delete any stub code and paste THIS entire file.
 * 3. In the script editor, set SPREADSHEET_ID below (from the Sheet URL).
 *    Optional: set DRIVE_FOLDER_ID to save recordings into a specific folder.
 * 4. Select setupSheet from the function dropdown → Run once.
 *    Approve Google permissions when prompted.
 * 5. Deploy → New deployment → Type: Web app
 *      - Execute as: Me
 *      - Who has access: Anyone
 * 6. Copy the Web App URL into index.html → APPS_SCRIPT_URL
 */

// ===== CONFIG — fill these in =====
var SPREADSHEET_ID = 'PASTE_YOUR_SPREADSHEET_ID_HERE';
var DRIVE_FOLDER_ID = ''; // optional; leave blank to save in My Drive root
var SHEET_NAME = 'Responses';

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
  'Official Links Clicked',
  'Drive Video URL',
  'Drive File ID'
];

function setupSheet() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }
  sheet.clear();
  sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
}

function doGet() {
  return ContentService
    .createTextOutput(JSON.stringify({
      ok: true,
      message: 'Information Search Survey endpoint is live. Use POST to submit.'
    }))
    .setMimeType(ContentService.MimeType.JSON);
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
    var fileName = payload.fileName || ('survey-session-' + Date.now() + '.webm');

    var driveUrl = '';
    var driveFileId = '';

    if (videoBase64) {
      var cleaned = videoBase64.replace(/^data:[^;]+;base64,/, '');
      var bytes = Utilities.base64Decode(cleaned);
      var blob = Utilities.newBlob(bytes, mimeType, fileName);

      var file;
      if (DRIVE_FOLDER_ID) {
        var folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
        file = folder.createFile(blob);
      } else {
        file = DriveApp.createFile(blob);
      }

      // Anyone with the link can view (researchers can tighten later)
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      driveFileId = file.getId();
      driveUrl = file.getUrl();
    }

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
      (answers.googleSearches || []).join(' | '),
      (answers.chatMessages || []).join(' | '),
      (answers.officialClicks || []).join(' | '),
      driveUrl,
      driveFileId
    ];

    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) {
      setupSheet();
      sheet = ss.getSheetByName(SHEET_NAME);
    }
    sheet.appendRow(row);

    return jsonResponse_({
      ok: true,
      driveUrl: driveUrl,
      driveFileId: driveFileId
    });
  } catch (err) {
    return jsonResponse_({ ok: false, error: String(err) });
  }
}

function jsonResponse_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
