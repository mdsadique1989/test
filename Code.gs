// Code.gs - FORM S02/UDISE  (one sheet row per student)
//
// Data goes to a tab called "Students" (created automatically).
// Your old wide sheet is NOT touched - keep it as a backup.

var SHEET_NAME = 'Students';
var TIMEZONE = 'Asia/Kolkata';

var HEADERS = [
  'Timestamp', 'Submission ID', 'Academic Year', 'UDISE State', 'District', 'Block',
  'UDISE Code', 'School Name', 'School Contact Number',
  'S.No', 'Name', 'Gender', 'DOB', 'CWSN', 'Class', 'Section', 'Admission Date',
  'Mother Name', 'Father Name', 'Guardian Name', 'Mobile', 'Alt Mobile',
  'Aadhaar', 'Aadhaar Name', 'Reason'
];
var COL = {};
HEADERS.forEach(function (h, i) { COL[h] = i; });

// ---------- web app entry points ----------

function doPost(e) {
  var lock = LockService.getScriptLock();
  // Only one save at a time, so two people can never slip in the same Aadhaar together.
  if (!lock.tryLock(30000)) {
    return respond_({ status: 'error', message: 'Server is busy. Please try again in a moment.' });
  }
  try {
    var data = JSON.parse(e.postData.contents);
    return respond_(saveSubmission_(data));
  } catch (err) {
    return respond_({ status: 'error', message: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function doGet() {
  return respond_({ status: 'ok' });
}

function respond_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ---------- helpers ----------

function clean_(v) { return String(v == null ? '' : v).replace(/\s+/g, ' ').trim(); }
function upper_(v) { return clean_(v).toUpperCase(); }
function digits_(v) { return String(v == null ? '' : v).replace(/\D/g, ''); }

// 2026-2027 / 2026/27 / 2026 - 27  ->  2026-27
function normYear_(v) {
  var s = clean_(v);
  var m = s.match(/^(\d{4})\s*[-\/\u2013]\s*(\d{2}|\d{4})$/);
  return m ? m[1] + '-' + m[2].slice(-2) : s;
}

function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// Builds one sheet row (array in HEADERS order). Values are cleaned here.
function buildRow_(ts, subId, school, s, index) {
  return [
    ts, subId,
    normYear_(school.academicYear), upper_(school.udiseState), upper_(school.district), upper_(school.block),
    upper_(school.udiseCode), upper_(school.schoolName), digits_(school.schoolContact),
    String(s.sno || index + 1),
    upper_(s.name), clean_(s.gender), clean_(s.dob), clean_(s.cwsn), upper_(s.class), upper_(s.section),
    clean_(s.admissionDate),
    upper_(s.motherName), upper_(s.fatherName), upper_(s.guardianName),
    digits_(s.mobile), digits_(s.altMobile),
    digits_(s.aadhaar), upper_(s.aadhaarName), clean_(s.reason)
  ];
}

// Writes rows as plain text so Aadhaar / mobile / dates are never converted to numbers or dates.
function writeRows_(sheet, rows) {
  if (!rows.length) return;
  var range = sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, HEADERS.length);
  range.setNumberFormat('@');
  range.setValues(rows);
}

// ---------- main save logic ----------

function saveSubmission_(data) {
  var school = {
    academicYear: data.academicYear,
    udiseState: data.udiseState || data.UdiseState,   // accept old key spelling too
    district: data.district,
    block: data.block,
    udiseCode: data.udiseCode || data.UdiseCode,
    schoolName: data.schoolName,
    schoolContact: data.schoolContact
  };

  if (!clean_(school.block) || !clean_(school.udiseCode) || !clean_(school.schoolName)) {
    return { status: 'error', message: 'Block, UDISE Code and School Name are required.' };
  }

  // Skip empty student slots.
  var students = (data.students || []).filter(function (s) {
    return s && (clean_(s.name) || digits_(s.aadhaar) || digits_(s.mobile));
  });
  if (!students.length) return { status: 'error', message: 'No student details received.' };

  // Validate (the form checks too, but never trust the browser alone).
  var seen = {};
  for (var i = 0; i < students.length; i++) {
    var s = students[i];
    var label = 'Student ' + (s.sno || i + 1);
    var a = digits_(s.aadhaar), m = digits_(s.mobile), alt = digits_(s.altMobile);
    if (!clean_(s.name)) return { status: 'error', message: label + ': Name is required.' };
    if (!/^\d{12}$/.test(a)) return { status: 'error', message: label + ': Aadhaar must be exactly 12 digits.' };
    if (!/^\d{10}$/.test(m)) return { status: 'error', message: label + ': Mobile number is mandatory (10 digits).' };
    if (alt && !/^\d{10}$/.test(alt)) return { status: 'error', message: label + ': Alternate number must be 10 digits.' };
    if (seen[a]) return { status: 'error', message: 'Same Aadhaar entered twice in this form (' + a + ').' };
    seen[a] = true;
  }

  var sheet = getSheet_();
  var n = sheet.getLastRow() - 1;                       // data rows (excluding header)
  var subId = clean_(data.submissionId) || Utilities.getUuid();

  // Read only the two columns we need.
  var idVals = n > 0 ? sheet.getRange(2, COL['Submission ID'] + 1, n, 1).getValues() : [];
  var aVals = n > 0 ? sheet.getRange(2, COL['Aadhaar'] + 1, n, 1).getValues() : [];

  var rowsByAadhaar = {};   // aadhaar -> [{row, id}]
  var ownRows = [];         // rows already saved under this submission ID
  var ownAadhaar = {};
  for (var r = 0; r < n; r++) {
    var ea = digits_(aVals[r][0]);
    var eid = String(idVals[r][0]);
    if (ea) (rowsByAadhaar[ea] = rowsByAadhaar[ea] || []).push({ row: r + 2, id: eid });
    if (eid === subId) { ownRows.push(r + 2); ownAadhaar[ea] = true; }
  }

  // Same submission ID saved before?
  //  - shares an Aadhaar with it  -> user is correcting a saved form: replace the old rows.
  //  - shares none                -> user typed a NEW set of students without clearing: keep old rows.
  var correction = false;
  if (ownRows.length) {
    correction = students.some(function (st) { return ownAadhaar[digits_(st.aadhaar)]; });
    if (!correction) subId = Utilities.getUuid();
  }

  // Duplicate check against everything already in the sheet.
  var dups = [];
  students.forEach(function (st, idx) {
    var a = digits_(st.aadhaar);
    var clash = (rowsByAadhaar[a] || []).some(function (x) { return x.id !== subId; });
    if (clash) dups.push({ sno: st.sno || idx + 1, aadhaar: a });
  });
  if (dups.length) {
    return {
      status: 'duplicate',
      duplicates: dups,
      message: 'Already registered. This AADHAAR number exists in the records:\n' +
        dups.map(function (d) { return '  Student ' + d.sno + ' - ' + d.aadhaar; }).join('\n') +
        '\n\nNothing was saved. Please remove or correct it and try again.'
    };
  }

  // Replace old rows on correction (delete bottom-up so row numbers stay valid).
  if (correction) {
    ownRows.sort(function (x, y) { return y - x; }).forEach(function (row) { sheet.deleteRow(row); });
  }

  var ts = Utilities.formatDate(new Date(), TIMEZONE, 'dd/MM/yyyy HH:mm:ss');
  var rows = students.map(function (st, idx) { return buildRow_(ts, subId, school, st, idx); });
  writeRows_(sheet, rows);

  return { status: 'success', submissionId: subId, saved: rows.length };
}

// ---------- ONE-TIME: convert your old wide sheet into the vertical "Students" sheet ----------
// Run once from the Apps Script editor (select migrateOldSheet > Run).
// Safe to run twice: rows already migrated are skipped. The old sheet is left untouched.

function migrateOldSheet() {
  var OLD_SHEET_NAME = 'Sheet1';   // <-- change to your old tab's name if different

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var old = ss.getSheetByName(OLD_SHEET_NAME);
  if (!old) throw new Error('Old sheet "' + OLD_SHEET_NAME + '" not found. Edit OLD_SHEET_NAME.');
  if (OLD_SHEET_NAME === SHEET_NAME) throw new Error('OLD_SHEET_NAME must differ from ' + SHEET_NAME);

  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var vals = old.getDataRange().getValues();
    var head = vals[0];
    var c = function (name) {
      var i = head.indexOf(name);
      if (i < 0) throw new Error('Column not found in old sheet: ' + name);
      return i;
    };

    var target = getSheet_();
    var n = target.getLastRow() - 1;
    var done = {};
    if (n > 0) {
      target.getRange(2, COL['Submission ID'] + 1, n, 1).getValues()
        .forEach(function (x) { done[String(x[0])] = true; });
    }

    var out = [], seenAadhaar = {}, dupLog = [];
    for (var r = 1; r < vals.length; r++) {
      var v = vals[r];
      var subId = 'OLD-' + (r + 1);              // r+1 = row number in old sheet
      if (done[subId]) continue;

      var tsRaw = v[c('Timestamp')];
      var ts = tsRaw instanceof Date ? Utilities.formatDate(tsRaw, TIMEZONE, 'dd/MM/yyyy HH:mm:ss') : String(tsRaw);
      var school = {
        academicYear: v[c('Academic Year')], udiseState: v[c('UDISE State')], district: v[c('District')],
        block: v[c('Block')], udiseCode: v[c('UDISE Code')], schoolName: v[c('School Name')],
        schoolContact: v[c('School Contact Number')]
      };

      for (var k = 1; k <= 3; k++) {
        var p = 'S' + k + ' ';
        var s = {
          sno: k, name: v[c(p + 'Name')], gender: v[c(p + 'Gender')], dob: v[c(p + 'DOB')], cwsn: v[c(p + 'CWSN')],
          class: v[c(p + 'Class')], section: v[c(p + 'Section')], admissionDate: v[c(p + 'Admission Date')],
          motherName: v[c(p + 'Mother Name')], fatherName: v[c(p + 'Father Name')], guardianName: v[c(p + 'Guardian Name')],
          mobile: v[c(p + 'Mobile')], altMobile: v[c(p + 'Alt Mobile')], aadhaar: v[c(p + 'Aadhaar')],
          aadhaarName: v[c(p + 'Aadhaar Name')], reason: v[c(p + 'Reason')]
        };
        if (!(clean_(s.name) || digits_(s.aadhaar) || digits_(s.mobile))) continue;   // empty slot
        var a = digits_(s.aadhaar);
        if (a && seenAadhaar[a]) dupLog.push('Aadhaar ' + a + ' appears again in old row ' + (r + 1) + ' (S' + k + ')');
        seenAadhaar[a] = true;
        out.push(buildRow_(ts, subId, school, s, k - 1));
      }
    }

    writeRows_(target, out);
    Logger.log('Migrated ' + out.length + ' student rows. Duplicate Aadhaar in old data: ' + dupLog.length);
    dupLog.forEach(function (d) { Logger.log(d); });
  } finally {
    lock.releaseLock();
  }
}
