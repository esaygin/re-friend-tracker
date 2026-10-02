/**
 * Backend for Pero's 40th website.
 * Saves RSVPs to this Google Sheet and uploaded photos to a Google Drive folder.
 * Setup steps are in PARTY-SETUP.md.
 */

// Google Drive folder for the photo album. Leave empty and setup() creates one for you.
var PHOTO_FOLDER_ID = '';

// Optional: an email address to notify on each new RSVP. Leave empty for no emails.
var NOTIFY_EMAIL = '';

var RSVP_HEADERS = ['Updated', 'Family', 'Names', 'Email', 'Phone', 'Attending', 'Adults', 'Children',
  'Children ages', 'Hotel', 'Arrival', 'Departure', 'Dietary needs', 'Message for Pero'];
var PHOTO_HEADERS = ['Received', 'From', 'File', 'Note', 'Link'];

/** Run once from the Apps Script editor to create the tabs and the photo folder. */
function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  sheet_(ss, 'RSVPs', RSVP_HEADERS);
  sheet_(ss, 'Photos', PHOTO_HEADERS);
  sheet_(ss, 'Summary', []);
  var props = PropertiesService.getScriptProperties();
  if (!PHOTO_FOLDER_ID && !props.getProperty('PHOTO_FOLDER_ID')) {
    var folder = DriveApp.createFolder("Pero's 40th - photo album");
    props.setProperty('PHOTO_FOLDER_ID', folder.getId());
    Logger.log('Created photo folder: ' + folder.getUrl());
  }
  updateSummary_();
}

function doGet() {
  return json_({ ok: true, message: "Pero's 40th backend is running." });
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    var data = JSON.parse(e.postData.contents);
    if (data.website) return json_({ ok: true }); // spam bot filled the hidden field
    lock.waitLock(20000);
    if (data.type === 'rsvp') return json_(saveRsvp_(data));
    if (data.type === 'photo') return json_(savePhoto_(data));
    return json_({ ok: false, error: 'unknown request' });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  } finally {
    try { lock.releaseLock(); } catch (ignored) {}
  }
}

function saveRsvp_(d) {
  var email = clean_(d.email).toLowerCase();
  if (!email || !clean_(d.family) || (d.attending !== 'Yes' && d.attending !== 'No')) {
    return { ok: false, error: 'missing fields' };
  }
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = sheet_(ss, 'RSVPs', RSVP_HEADERS);
  var ages = (d.kidAges || []).map(function (a) { return a === '0' || a === 0 ? '<1' : String(a); }).join(', ');
  var row = [new Date(), safe_(d.family), safe_(d.names), safe_(email), safe_(d.phone), d.attending,
    num_(d.adults), num_(d.children), ages, safe_(d.hotel), safe_(d.arrival), safe_(d.departure),
    safe_(d.dietary), safe_(d.message)];

  // One row per family: a second reply with the same email replaces the first.
  var last = sh.getLastRow();
  var existing = -1;
  if (last > 1) {
    var emails = sh.getRange(2, 4, last - 1, 1).getValues();
    for (var i = 0; i < emails.length; i++) {
      if (String(emails[i][0]).replace(/^'/, '').toLowerCase() === email) { existing = i + 2; break; }
    }
  }
  if (existing > 0) sh.getRange(existing, 1, 1, row.length).setValues([row]);
  else sh.appendRow(row);
  updateSummary_();

  if (NOTIFY_EMAIL) {
    var subject = "Pero's 40th RSVP: " + clean_(d.family) + ' - ' + (d.attending === 'Yes' ? 'coming' : 'not coming');
    var body = clean_(d.names) + ' (' + email + ')\n' +
      (d.attending === 'Yes'
        ? num_(d.adults) + ' adults, ' + num_(d.children) + ' children' + (ages ? ' (ages ' + ages + ')' : '') +
          '\nHotel: ' + clean_(d.hotel) + '\nDietary: ' + clean_(d.dietary)
        : 'Not coming') +
      (d.message ? '\n\nMessage: ' + clean_(d.message) : '');
    MailApp.sendEmail(NOTIFY_EMAIL, subject, body);
  }
  return { ok: true, updated: existing > 0 };
}

function savePhoto_(d) {
  if (!d.data) return { ok: false, error: 'no file' };
  var mime = /^image\//.test(d.mimeType || '') ? d.mimeType : 'image/jpeg';
  var from = clean_(d.from).slice(0, 60) || 'Guest';
  var name = (from + ' - ' + clean_(d.filename || 'photo.jpg')).replace(/[\\\/:*?"<>|]/g, '_').slice(0, 150);
  var blob = Utilities.newBlob(Utilities.base64Decode(d.data), mime, name);
  var file = photoFolder_().createFile(blob);
  if (d.note) file.setDescription(clean_(d.note).slice(0, 2000));
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  sheet_(ss, 'Photos', PHOTO_HEADERS).appendRow([new Date(), safe_(from), safe_(name), safe_(d.note), file.getUrl()]);
  return { ok: true };
}

/** Rebuilds the Summary tab with totals for planning dinner, rooms and the kids' corner. */
function updateSummary_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var rows = sheet_(ss, 'RSVPs', RSVP_HEADERS).getDataRange().getValues().slice(1);
  var t = { replied: rows.length, coming: 0, declined: 0, adults: 0, kids: 0, hotelYes: 0, hotelMaybe: 0,
    hotelPeople: 0, ages: { 'Under 3': 0, '3 to 5': 0, '6 to 9': 0, '10 to 13': 0, '14 to 17': 0 } };
  rows.forEach(function (r) {
    if (r[5] === 'Yes') {
      t.coming++; t.adults += +r[6] || 0; t.kids += +r[7] || 0;
      if (r[9] === 'Yes') { t.hotelYes++; t.hotelPeople += (+r[6] || 0) + (+r[7] || 0); }
      if (r[9] === 'Maybe') t.hotelMaybe++;
      String(r[8]).split(',').forEach(function (a) {
        a = a.trim(); if (!a) return;
        var n = a === '<1' ? 0 : +a;
        var k = n < 3 ? 'Under 3' : n < 6 ? '3 to 5' : n < 10 ? '6 to 9' : n < 14 ? '10 to 13' : '14 to 17';
        t.ages[k]++;
      });
    } else if (r[5] === 'No') t.declined++;
  });
  var out = [
    ["Pero's 40th - summary", ''],
    ['Last updated', new Date()],
    ['', ''],
    ['Families replied', t.replied],
    ['Families coming', t.coming],
    ['Families not coming', t.declined],
    ['Adults coming', t.adults],
    ['Children coming', t.kids],
    ['Guests at dinner', t.adults + t.kids],
    ['', ''],
    ['Families staying at the hotel', t.hotelYes],
    ['People staying at the hotel', t.hotelPeople],
    ['Families still deciding', t.hotelMaybe],
    ['', ''],
    ['Children by age', '']
  ];
  Object.keys(t.ages).forEach(function (k) { out.push(['  ' + k, t.ages[k]]); });
  var sh = sheet_(ss, 'Summary', []);
  sh.clearContents();
  sh.getRange(1, 1, out.length, 2).setValues(out);
  sh.getRange(1, 1).setFontWeight('bold').setFontSize(14);
  sh.setColumnWidth(1, 260);
}

function photoFolder_() {
  var id = PHOTO_FOLDER_ID || PropertiesService.getScriptProperties().getProperty('PHOTO_FOLDER_ID');
  if (!id) throw new Error('Photo folder not set up. Run setup() first.');
  return DriveApp.getFolderById(id);
}

function sheet_(ss, name, headers) {
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    if (headers.length) {
      sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
      sh.setFrozenRows(1);
    }
  }
  return sh;
}

function clean_(v) { return String(v == null ? '' : v).trim().slice(0, 5000); }
// Stops text starting with = + - @ from being treated as a spreadsheet formula.
function safe_(v) { var s = clean_(v); return /^[=+\-@]/.test(s) ? "'" + s : s; }
function num_(v) { var n = parseInt(v, 10); return isNaN(n) || n < 0 ? 0 : Math.min(n, 20); }
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
