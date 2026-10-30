/**
 * திருமலை குடும்பம் — பாதயாத்திரை registration backend.
 *
 * Receives one submission from index.html, allocates the next ID number on a
 * strict first-come-first-served basis, files the photo and signatures in
 * Drive, and appends a row to the sheet.
 *
 * ── SETUP ──────────────────────────────────────────────────────────────────
 *  1. Create a new Google Sheet (this is where registrations land).
 *  2. Extensions ▸ Apps Script. Delete the stub, paste this file, save.
 *  3. Deploy ▸ New deployment ▸ type "Web app".
 *       Execute as:        Me
 *       Who has access:    Anyone                 <-- required; see note below
 *  4. Authorise when prompted (it needs Sheets + Drive access).
 *  5. Copy the /exec URL it gives you and paste it into SUBMIT_URL in
 *     index.html.
 *
 *  Note on "Anyone": the browser calls this without a login, so the endpoint
 *  must be open. Anyone holding the URL can append rows, so treat the URL as
 *  semi-private — don't publish index.html with it filled in. It grants no
 *  read access to the sheet.
 *
 *  Re-deploying after an edit: Deploy ▸ Manage deployments ▸ pencil ▸
 *  Version "New version". Keeping the same deployment keeps the same URL.
 */

const SHEET_NAME = 'Registrations';
const DRIVE_FOLDER = 'Pathayathirai Uploads';

// The group's yatra year — the fixed part of the ID, as printed on the card
// template (TK14___). This is NOT the per-pilgrim year dropdown.
const EVENT_YEAR = '14';
const SERIAL_DIGITS = 3;

const HEADERS = [
  'Submitted at', 'ID number', 'Serial', 'வருஷம்', 'பெயர்', 'தந்தை பெயர்',
  'பிறந்த தேதி', 'ஆதார் எண்', 'முகவரி', 'கைபேசி எண்', 'இரத்த வகை',
  'பொறுப்பாளர் பெயர்', 'பொறுப்பாளர் கைபேசி', 'இடம்', 'தேதி',
  'Photo', 'Signature', 'Incharge signature'
];

function doPost(e) {
  // One allocation at a time. Without this, two submissions landing together
  // would read the same last row and be handed the same ID number.
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
  } catch (err) {
    return json_({ ok: false, error: 'Server busy, please submit again.' });
  }

  try {
    if (!e || !e.postData || !e.postData.contents) {
      return json_({ ok: false, error: 'Empty request body.' });
    }
    const data = JSON.parse(e.postData.contents);

    const sheet = getSheet_();
    // Row 1 is the header, so the current last row IS the next serial number.
    const serial = sheet.getLastRow();
    const idNumber = 'TK' + EVENT_YEAR + padLeft_(serial, SERIAL_DIGITS);

    const folder = getFolder_();
    const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd-HHmmss');
    const base = idNumber + '_' + stamp;

    const photoUrl = saveDataUrl_(folder, data.photo, base + '_photo');
    const sigUrl = saveDataUrl_(folder, data.signature, base + '_signature');
    const inSigUrl = saveDataUrl_(folder, data.inchargeSignature, base + '_incharge_sig');

    sheet.appendRow([
      new Date(),
      idNumber,
      serial,
      data.year || '',
      data.name || '',
      data.fatherName || '',
      data.dob || '',
      asText_(data.aadhar),
      data.address || '',
      asText_(data.mobile),
      data.bloodGroup || '',
      data.inchargeName || '',
      asText_(data.inchargeMobile),
      data.place || '',
      data.declDate || '',
      photoUrl,
      sigUrl,
      inSigUrl
    ]);
    SpreadsheetApp.flush();

    return json_({ ok: true, idNumber: idNumber, serial: serial });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  } finally {
    lock.releaseLock();
  }
}

/** Lets you confirm the deployment is live by opening the /exec URL. */
function doGet() {
  return json_({ ok: true, service: 'pathayathirai-registration' });
}

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
  }
  if (sh.getLastRow() === 0) {
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
  }
  return sh;
}

function getFolder_() {
  const it = DriveApp.getFoldersByName(DRIVE_FOLDER);
  return it.hasNext() ? it.next() : DriveApp.createFolder(DRIVE_FOLDER);
}

/** Decodes a data: URL into Drive and returns the file's link. */
function saveDataUrl_(folder, dataUrl, filename) {
  if (!dataUrl) return '';
  const m = /^data:([^;]+);base64,(.+)$/.exec(dataUrl);
  if (!m) return '';
  const mime = m[1];
  const ext = mime.split('/')[1].replace('jpeg', 'jpg');
  const blob = Utilities.newBlob(Utilities.base64Decode(m[2]), mime, filename + '.' + ext);
  return folder.createFile(blob).getUrl();
}

/** Keeps Aadhaar/mobile as text so Sheets doesn't render them as 1.23457E+11. */
function asText_(v) {
  return v ? "'" + v : '';
}

function padLeft_(n, width) {
  let s = String(n);
  while (s.length < width) s = '0' + s;
  return s;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
