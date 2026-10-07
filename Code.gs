/**
 * UHS Kaparpura — Online Admission backend
 * Deploy this as a Google Apps Script Web App.
 * It writes every submission to a Google Sheet, and stores the
 * uploaded photo + signature in a Google Drive folder.
 *
 * SETUP (one-time):
 * 1. Create a blank Google Sheet. Copy its ID from the URL:
 *      https://docs.google.com/spreadsheets/d/  <-- SHEET_ID -->  /edit
 * 2. Create a Google Drive folder (e.g. "UHSK Admission Uploads").
 *    Copy its ID from the URL:
 *      https://drive.google.com/drive/folders/  <-- FOLDER_ID -->
 * 3. Paste both IDs below.
 * 4. In the Apps Script editor: Extensions > Apps Script (from the Sheet),
 *    delete any starter code, paste this whole file in.
 * 5. Click Deploy > New deployment > type: Web app.
 *      - Execute as: Me
 *      - Who has access: Anyone
 * 6. Copy the deployment URL (ends in /exec) into CONFIG.WEB_APP_URL
 *    inside index.html.
 * 7. Re-deploy (Manage deployments > Edit > New version) any time you
 *    change this script.
 */

// ====================== CONFIGURATION ======================
var SHEET_ID = '1g462lSD6PBngZicD7sTVPYlmibrKO4xDq-ndG3QBngQ';
var SHEET_NAME = 'data';
var DRIVE_FOLDER_ID = '1izJr-oWqhexXsZ6b8sBdOp700_ouZ8vw';

// Optional lightweight shared secret so random visitors can't scrape
// applicant data (Aadhar / bank details) just by guessing the web app URL.
// Set the same value in CONFIG.API_KEY inside index.html.
// Leave both blank ('') to disable this check.
var API_KEY = '';

// ---- Admin dashboard login ----
// The admin password is NOT written in this file (so it's safe even if you
// commit Code.gs to a public GitHub repo). Instead it's stored privately
// inside this Apps Script project:
//   Apps Script editor -> ⚙️ Project Settings -> Script Properties
//   -> Add script property -> name: ADMIN_PASSWORD, value: (choose a password)
// The person opening admin.html simply types that password each time they
// want to log in — nothing about it is ever saved in any file.
var ADMIN_PASSWORD_PROPERTY = 'ADMIN_PASSWORD';

var TIMEZONE = 'Asia/Kolkata';

// Shown in the admin dashboard header so you can confirm the NEW code is deployed.
var BACKEND_VERSION = 'v4-2026-10-07';

// Column order in the sheet (row 1 = headers, auto-created on first run).
var HEADERS = [
  'applicationId', 'submittedAt', 'admClass', 'stream',
  'pen', 'apaar', 'aadhar', 'eshiksha',
  'studentName', 'dob', 'gender', 'category', 'religion',
  'motherName', 'fatherName', 'aadharOwner', 'guardianAadhar', 'mobile',
  'address', 'pincode',
  'distance', 'cwsn', 'income', 'bloodGroup', 'height', 'weight',
  'prevUdise',
  'bankAccount', 'ifsc', 'bankName', 'accHolder', 'accRelation',
  'photoUrl', 'signatureUrl'
];
// =============================================================

// Converts a dob value into 'dd/mm/yyyy' text.
// Accepts: a native <input type="date"> value ('yyyy-mm-dd'), an already
// dd/mm/yyyy string (left as-is), or a real Date object (e.g. if a sheet
// cell auto-converted an older entry) — always normalizes to dd/mm/yyyy so
// storage and every API response stay consistent.
function formatDateDMY_(value) {
  if (value === '' || value === null || value === undefined) return '';
  if (Object.prototype.toString.call(value) === '[object Date]') {
    return Utilities.formatDate(value, TIMEZONE, 'dd/MM/yyyy');
  }
  var str = String(value).trim();
  var iso = str.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (iso) return iso[3] + '/' + iso[2] + '/' + iso[1];
  var dmy = str.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (dmy) return str; // already dd/mm/yyyy
  return str;
}

function getSheet_() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getFolder_() {
  return DriveApp.getFolderById(DRIVE_FOLDER_ID);
}

function genAppId_() {
  var stamp = Utilities.formatDate(new Date(), TIMEZONE, 'yyMMdd');
  var rand = Math.floor(1000 + Math.random() * 9000);
  return 'UHSK' + stamp + rand;
}

// Decodes a base64 data-URL image and saves it into the Drive folder.
// Returns a viewable link, or '' if there was no image.
function saveImage_(dataUrl, filenamePrefix) {
  if (!dataUrl) return '';
  var match = String(dataUrl).match(/^data:(image\/[a-zA-Z0-9.+-]+);base64,(.+)$/);
  if (!match) return '';
  var mime = match[1];
  var base64 = match[2];
  var ext = mime.split('/')[1] === 'jpeg' ? 'jpg' : mime.split('/')[1];
  var bytes = Utilities.base64Decode(base64);
  var blob = Utilities.newBlob(bytes, mime, filenamePrefix + '.' + ext);
  var file = getFolder_().createFile(blob);
  // Anyone with the (unguessable) link can view — needed so the printed
  // application PDF can display the photo/signature as an <img>.
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return 'https://drive.google.com/thumbnail?id=' + file.getId() + '&sz=w1000';
}

function checkKey_(providedKey) {
  if (!API_KEY) return true; // check disabled
  return providedKey === API_KEY;
}

// Checks the admin password against the one stored in Script Properties.
// Returns true/false. If no ADMIN_PASSWORD property has been set yet,
// this always fails closed (denies access) rather than opening the
// dashboard to everyone by accident.
function checkAdminPassword_(providedPassword) {
  var real = PropertiesService.getScriptProperties().getProperty(ADMIN_PASSWORD_PROPERTY);
  if (!real) return false;
  return String(providedPassword || '') === real;
}

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}



// ---------------------- POST: router ----------------------
function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);

    if (body.action === 'submit') return handleSubmit_(body);
    if (body.action === 'adminList') return handleAdminList_(body);
    if (body.action === 'adminDelete') return handleAdminDelete_(body);
    if (body.action === 'adminDeleteMultiple') return handleAdminDeleteMultiple_(body);
    if (body.action === 'adminUpdate') return handleAdminUpdate_(body);

    return jsonOut_({ success: false, error: 'Unknown action' });
  } catch (err) {
    return jsonOut_({ success: false, error: err.message });
  }
}

// ---- Public: submit a new application (guarded by the light API key) ----
function handleSubmit_(body) {
  if (!checkKey_(body.apiKey)) {
    return jsonOut_({ success: false, error: 'Unauthorized' });
  }

  var appId = genAppId_();
  var photoUrl = saveImage_(body.photo, appId + '_photo');
  var signatureUrl = saveImage_(body.signature, appId + '_signature');

  var dobFormatted = formatDateDMY_(body.dob);

  var row = HEADERS.map(function (key) {
    if (key === 'applicationId') return appId;
    if (key === 'submittedAt') return Utilities.formatDate(new Date(), TIMEZONE, 'dd-MM-yyyy HH:mm');
    if (key === 'dob') return dobFormatted;
    if (key === 'photoUrl') return photoUrl;
    if (key === 'signatureUrl') return signatureUrl;
    return body[key] !== undefined ? body[key] : '';
  });

  // LockService avoids two simultaneous submissions clashing on the same row.
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sheet = getSheet_();
    sheet.appendRow(row);
    // Force the DOB cell to stay as literal text ('@' format) so Google
    // Sheets doesn't silently auto-convert the dd/mm/yyyy string into its
    // own Date/serial value (which would show/export in a different format).
    var dobColIdx = HEADERS.indexOf('dob') + 1;
    var lastRow = sheet.getLastRow();
    sheet.getRange(lastRow, dobColIdx).setNumberFormat('@').setValue(dobFormatted);
  } finally {
    lock.releaseLock();
  }

  return jsonOut_({ success: true, applicationId: appId, photoUrl: photoUrl, signatureUrl: signatureUrl });
}

// ---- Admin: list every application (guarded by the admin password) ----
function handleAdminList_(body) {
  if (!checkAdminPassword_(body.password)) {
    return jsonOut_({ success: false, error: 'Unauthorized' });
  }

  var data = getSheet_().getDataRange().getValues();
  var headers = data[0];
  var records = [];
  for (var i = 1; i < data.length; i++) {
    var record = {};
    headers.forEach(function (h, idx) {
      record[h] = (h === 'dob') ? formatDateDMY_(data[i][idx]) : data[i][idx];
    });
    record._row = i + 1; // 1-based sheet row number, needed for delete
    records.push(record);
  }
  records.reverse(); // newest first

  return jsonOut_({ success: true, records: records, version: BACKEND_VERSION, sheet: SHEET_NAME });
}

// ---- Admin: delete one application by its sheet row number ----
function handleAdminDelete_(body) {
  if (!checkAdminPassword_(body.password)) {
    return jsonOut_({ success: false, error: 'Unauthorized' });
  }
  var appId = String(body.applicationId || '').trim();
  if (!appId) {
    return jsonOut_({ success: false, error: 'applicationId missing' });
  }

  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sheet = getSheet_();
    var col = HEADERS.indexOf('applicationId') + 1;

    // Find EVERY row carrying this applicationId (by id, never by a stale row number).
    var last = sheet.getLastRow();
    var rows = [];
    if (last >= 2) {
      var vals = sheet.getRange(2, col, last - 1, 1).getValues();
      for (var i = 0; i < vals.length; i++) {
        if (String(vals[i][0]).trim() === appId) rows.push(i + 2);
      }
    }
    if (rows.length === 0) {
      return jsonOut_({ success: false, error: 'Application sheet "' + SHEET_NAME + '" mein nahi mili (id: ' + appId + '). Refresh karke dekhein.' });
    }

    rows.sort(function (a, b) { return b - a; }); // bottom first
    rows.forEach(function (r) { sheet.deleteRow(r); });
    SpreadsheetApp.flush();

    var still = findRowByAppId_(sheet, appId);
    if (still) {
      return jsonOut_({ success: false, error: 'Delete verify failed: row ' + still + ' abhi bhi maujood hai (sheet "' + SHEET_NAME + '"). Sheet mein is row par formula/protection to nahi?' });
    }
    return jsonOut_({ success: true, removedRows: rows, backend: BACKEND_VERSION });
  } finally {
    lock.releaseLock();
  }
}

// Finds the sheet row number (1-based) for an applicationId. Returns 0 if not found.
function findRowByAppId_(sheet, appId) {
  var last = sheet.getLastRow();
  if (last < 2) return 0;
  var col = HEADERS.indexOf('applicationId') + 1;
  var vals = sheet.getRange(2, col, last - 1, 1).getValues();
  for (var i = 0; i < vals.length; i++) {
    if (String(vals[i][0]) === appId) return i + 2;
  }
  return 0;
}

// ---- Admin: delete MANY applications at once (by applicationId) ----
// Looks rows up by applicationId (not by row number) and deletes from the
// bottom up, so row shifts can never delete the wrong application.
function handleAdminDeleteMultiple_(body) {
  if (!checkAdminPassword_(body.password)) {
    return jsonOut_({ success: false, error: 'Unauthorized' });
  }
  var ids = body.applicationIds;
  if (!ids || Object.prototype.toString.call(ids) !== '[object Array]' || ids.length === 0) {
    return jsonOut_({ success: false, error: 'No applications selected' });
  }
  var wanted = {};
  ids.forEach(function (id) { wanted[String(id)] = true; });

  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  var deleted = 0;
  try {
    var sheet = getSheet_();
    var last = sheet.getLastRow();
    if (last >= 2) {
      var col = HEADERS.indexOf('applicationId') + 1;
      var vals = sheet.getRange(2, col, last - 1, 1).getValues();
      var rowsToDelete = [];
      for (var i = 0; i < vals.length; i++) {
        if (wanted[String(vals[i][0])]) rowsToDelete.push(i + 2);
      }
      rowsToDelete.sort(function (a, b) { return b - a; }); // bottom first
      rowsToDelete.forEach(function (r) { sheet.deleteRow(r); });
      SpreadsheetApp.flush();
      deleted = rowsToDelete.length;

      // verify: none of the requested ids may remain in the sheet
      var after = sheet.getLastRow();
      var remaining = 0;
      if (after >= 2) {
        sheet.getRange(2, col, after - 1, 1).getValues().forEach(function (v) {
          if (wanted[String(v[0])]) remaining++;
        });
      }
      if (remaining > 0) {
        return jsonOut_({ success: false, error: 'Delete verify failed: ' + remaining + ' application abhi bhi sheet mein maujood hain.' });
      }
    }
  } finally {
    lock.releaseLock();
  }
  return jsonOut_({ success: true, deleted: deleted, notFound: ids.length - deleted });
}

// ---- Admin: edit the details of one application ----
// Only these columns may be changed; id, timestamp and image links are fixed.
var NON_EDITABLE_FIELDS = ['applicationId', 'submittedAt', 'photoUrl', 'signatureUrl'];

// Moves the Drive file referenced by a thumbnail URL (…?id=FILE_ID&…) to trash.
// Failures are ignored so a missing/already-deleted file never blocks an edit.
function trashDriveFileByUrl_(url) {
  try {
    var m = String(url || '').match(/[?&]id=([a-zA-Z0-9_-]+)/);
    if (m) DriveApp.getFileById(m[1]).setTrashed(true);
  } catch (err) { /* ignore */ }
}

function handleAdminUpdate_(body) {
  if (!checkAdminPassword_(body.password)) {
    return jsonOut_({ success: false, error: 'Unauthorized' });
  }
  var appId = String(body.applicationId || '').trim();
  var updates = body.updates;
  if (!appId || !updates || typeof updates !== 'object') {
    return jsonOut_({ success: false, error: 'Invalid request' });
  }

  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sheet = getSheet_();
    var rowNum = findRowByAppId_(sheet, appId);
    if (!rowNum) {
      return jsonOut_({ success: false, error: 'Application not found. Please refresh and try again.' });
    }
    var changed = 0;
    HEADERS.forEach(function (key, idx) {
      if (NON_EDITABLE_FIELDS.indexOf(key) !== -1) return;
      if (!Object.prototype.hasOwnProperty.call(updates, key)) return;
      var val = updates[key];
      val = (key === 'dob') ? formatDateDMY_(val)
                            : String(val === null || val === undefined ? '' : val).trim();
      // '@' = plain text, so Sheets never strips leading zeros (PEN, bank a/c, etc.)
      sheet.getRange(rowNum, idx + 1).setNumberFormat('@').setValue(val);
      changed++;
    });
    // Optional: replace photo and/or signature. This action already requires the
    // admin password (checked above), so only the admin can change images.
    var imgResult = {};
    [['photo', 'photoUrl', '_photo'], ['signature', 'signatureUrl', '_signature']].forEach(function (p) {
      var dataUrl = body[p[0]];
      if (!dataUrl) return;
      var colIdx = HEADERS.indexOf(p[1]) + 1;
      var oldUrl = String(sheet.getRange(rowNum, colIdx).getValue() || '');
      var newUrl = saveImage_(dataUrl, appId + p[2]);
      if (!newUrl) return; // not a valid image data URL — leave the old one
      sheet.getRange(rowNum, colIdx).setValue(newUrl);
      trashDriveFileByUrl_(oldUrl); // old image is no longer needed (and was link-shared)
      imgResult[p[1]] = newUrl;
      changed++;
    });
    return jsonOut_({ success: true, changed: changed, images: imgResult });
  } finally {
    lock.releaseLock();
  }
}

// ---------------------- GET: public self-lookup / health check ----------------------
function doGet(e) {
  try {
    var action = e.parameter.action;

    if (action === 'search') {
      if (!checkKey_(e.parameter.apiKey)) {
        return jsonOut_({ found: false, error: 'Unauthorized' });
      }
      var cls = e.parameter.class || '';
      var type = e.parameter.type || 'appId';
      var value = String(e.parameter.value || '').trim();

      var data = getSheet_().getDataRange().getValues();
      var headers = data[0];
      var classIdx = headers.indexOf('admClass');
      var appIdIdx = headers.indexOf('applicationId');
      var penIdx = headers.indexOf('pen');

      for (var i = data.length - 1; i >= 1; i--) { // newest first
        var row = data[i];
        var classMatch = String(row[classIdx]) === cls;
        var keyMatch = type === 'appId'
          ? String(row[appIdIdx]) === value
          : String(row[penIdx]) === value;
        if (classMatch && keyMatch) {
          var record = {};
          headers.forEach(function (h, idx) {
            record[h] = (h === 'dob') ? formatDateDMY_(row[idx]) : row[idx];
          });
          return jsonOut_({ found: true, record: record });
        }
      }
      return jsonOut_({ found: false });
    }

    return jsonOut_({ ok: true, message: 'UHS Kaparpura admission API is running.' });
  } catch (err) {
    return jsonOut_({ error: err.message });
  }
}
