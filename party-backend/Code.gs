/**
 * Backend for Peros 40th website.
 * Saves RSVPs to this Google Sheet and uploaded photos to a Google Drive folder.
 * Setup steps are in PARTY-SETUP.md.
 */

// Google Drive folder for the photo album. Leave empty and setup() creates one for you.
var PHOTO_FOLDER_ID = '';

// Optional: an email address to notify on each new RSVP. Leave empty for no emails.
var NOTIFY_EMAIL = '';

var RSVP_HEADERS = ['Updated', 'Name', 'Coming', 'Adults', 'Kids', 'Total people', 'Kids ages', 'Hotel nights', 'Mobile'];
var PHOTO_HEADERS = ['Received', 'From', 'File', 'Note', 'Link'];

/** Run once from the Apps Script editor to create the tabs and the photo folder. */
function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  sheet_(ss, 'RSVPs', RSVP_HEADERS);
  sheet_(ss, 'Photos', PHOTO_HEADERS);
  sheet_(ss, 'Summary', []);
  var props = PropertiesService.getScriptProperties();
  if (!PHOTO_FOLDER_ID && !props.getProperty('PHOTO_FOLDER_ID')) {
    var folder = DriveApp.createFolder("Peros 40th - photo album");
    props.setProperty('PHOTO_FOLDER_ID', folder.getId());
    Logger.log('Created photo folder: ' + folder.getUrl());
  }
  updateSummary_();
}

function doGet() {
  return json_({ ok: true, message: "Peros 40th backend is running." });
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
  var name = clean_(d.name);
  if (!name || (d.attending !== 'Yes' && d.attending !== 'No')) return { ok: false, error: 'missing fields' };
  var yes = d.attending === 'Yes';
  var adults = yes ? num_(d.adults) : 0;
  var kids = yes ? num_(d.kids) : 0;
  var nights = yes ? Math.min(num_(d.nights), 2) : 0;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = sheet_(ss, 'RSVPs', RSVP_HEADERS);
  var row = [new Date(), safe_(name), d.attending, adults, kids, adults + kids,
    yes && kids ? safe_(d.kidsAges) : '', nights, safe_(d.phone)];

  // One row per family: replying again with the same name replaces the earlier answer.
  var key = nameKey_(name);
  var last = sh.getLastRow();
  var existing = -1;
  if (last > 1) {
    var names = sh.getRange(2, 2, last - 1, 1).getValues();
    for (var i = 0; i < names.length; i++) {
      if (nameKey_(names[i][0]) === key) { existing = i + 2; break; }
    }
  }
  if (existing > 0) sh.getRange(existing, 1, 1, row.length).setValues([row]);
  else sh.appendRow(row);
  updateSummary_();

  if (NOTIFY_EMAIL) {
    var subject = 'Peros 40th RSVP: ' + name + ' - ' + (yes ? 'coming' : 'not coming');
    var body = yes
      ? adults + ' adults, ' + kids + ' kids' + (row[6] ? ' (ages ' + clean_(d.kidsAges) + ')' : '') +
        '\nHotel: ' + (nights ? nights + ' night(s)' : 'not staying')
      : 'Not coming';
    if (d.phone) body += '\nMobile: ' + clean_(d.phone);
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
  var t = { replied: rows.length, coming: 0, declined: 0, adults: 0, kids: 0,
    one: 0, onePeople: 0, two: 0, twoPeople: 0, dinnerOnly: 0 };
  rows.forEach(function (r) {
    if (r[2] === 'Yes') {
      var people = (+r[3] || 0) + (+r[4] || 0);
      t.coming++; t.adults += +r[3] || 0; t.kids += +r[4] || 0;
      if (+r[7] === 1) { t.one++; t.onePeople += people; }
      else if (+r[7] === 2) { t.two++; t.twoPeople += people; }
      else t.dinnerOnly++;
    } else if (r[2] === 'No') t.declined++;
  });
  var out = [
    ['Peros 40th - summary', ''],
    ['Last updated', new Date()],
    ['', ''],
    ['Families replied', t.replied],
    ['Families coming', t.coming],
    ['Families not coming', t.declined],
    ['', ''],
    ['Adults coming', t.adults],
    ['Kids coming', t.kids],
    ['Total people', t.adults + t.kids],
    ['', ''],
    ['Families staying 1 night', t.one],
    ['  People staying 1 night', t.onePeople],
    ['Families staying 2 nights', t.two],
    ['  People staying 2 nights', t.twoPeople],
    ['Families coming for dinner only', t.dinnerOnly]
  ];
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

function nameKey_(v) { return String(v).replace(/^'/, '').toLowerCase().replace(/[^a-z0-9\u00c0-\u024f]+/g, ' ').trim(); }
function clean_(v) { return String(v == null ? '' : v).trim().slice(0, 5000); }
// Stops text starting with = + - @ from being treated as a spreadsheet formula.
function safe_(v) { var s = clean_(v); return /^[=+\-@]/.test(s) ? "'" + s : s; }
function num_(v) { var n = parseInt(v, 10); return isNaN(n) || n < 0 ? 0 : Math.min(n, 20); }
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
