// "Where do you want to go?" — ATL edition
// Serves the app and reads/writes the "places" tab of this spreadsheet.

const SHEET_NAME = 'places';
// The page design lives in this Drive folder. Claude drops a new "app.html" there for each update;
// the newest one is served. If the folder is empty or unreadable, the Index file in this project is used.
const APP_FOLDER_ID = '1ojbqE5397ZpKKyc5_kaw98HLmFLxW2bq';
const APP_FILE_NAME = 'app.html';

function appHtml_() {
  try {
    const files = DriveApp.getFolderById(APP_FOLDER_ID).getFilesByName(APP_FILE_NAME);
    let best = null;
    while (files.hasNext()) {
      const f = files.next();
      if (f.isTrashed()) continue;
      if (!best || f.getDateCreated() > best.getDateCreated()) best = f;
    }
    if (best) return HtmlService.createHtmlOutput(best.getBlob().getDataAsString('UTF-8'));
  } catch (e) {
    console.warn('Falling back to Index: ' + e);
  }
  return HtmlService.createHtmlOutputFromFile('Index');
}

// Run this once from the editor (select it, press Run) so Google asks for Drive permission.
function authorize() {
  return DriveApp.getFolderById(APP_FOLDER_ID).getName();
}

// Home-screen / tab icon (public image in the app files folder).
const ICON_URL = 'https://lh3.googleusercontent.com/d/1YaRKHUkSINU0QWcVK5qhg2zwXDqBSpUz';

function doGet() {
  const out = appHtml_()
    .setTitle('Where do you want to go?')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .addMetaTag('apple-mobile-web-app-capable', 'yes');
  try { out.addMetaTag('apple-mobile-web-app-title', 'Where to go?'); } catch (e) {}
  // Lets the home-screen launcher page (GitHub Pages) show the app inside itself.
  out.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  try { out.setFaviconUrl(ICON_URL); }
  catch (e) { try { out.setFaviconUrl(ICON_URL + '#icon.png'); } catch (e2) { console.warn('Icon not set: ' + e2); } }
  return out;
}

function sheet_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sh) throw new Error('No tab named "' + SHEET_NAME + '" in this spreadsheet.');
  return sh;
}

function headers_(sh) {
  return sh.getRange(1, 1, 1, sh.getLastColumn()).getDisplayValues()[0];
}

// All rows as objects keyed by the header row. Everything comes back as text.
function getPlaces() {
  const sh = sheet_();
  const values = sh.getDataRange().getDisplayValues();
  const head = values.shift();
  return values
    .filter(function (r) { return r[0]; })
    .map(function (r) {
      const o = {};
      head.forEach(function (k, i) { o[k] = r[i]; });
      return o;
    });
}

// Create or update one row. Only the keys present in `rec` change; other cells keep their value.
function savePlace(rec) {
  if (!rec || !rec.id) throw new Error('Missing id');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = sheet_();
    const head = headers_(sh);
    const last = sh.getLastRow();
    const ids = last > 1 ? sh.getRange(2, 1, last - 1, 1).getDisplayValues().map(function (r) { return r[0]; }) : [];
    const idx = ids.indexOf(String(rec.id));
    const rowNum = idx >= 0 ? idx + 2 : last + 1;
    const current = idx >= 0 ? sh.getRange(rowNum, 1, 1, head.length).getDisplayValues()[0] : head.map(function () { return ''; });
    const row = head.map(function (k, i) {
      if (!(k in rec)) return current[i];
      const v = rec[k];
      return v === null || v === undefined ? '' : String(v);
    });
    const range = sh.getRange(rowNum, 1, 1, head.length);
    range.setNumberFormat('@');
    range.setValues([row]);
    return true;
  } finally {
    lock.releaseLock();
  }
}

function deletePlace(id) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = sheet_();
    const last = sh.getLastRow();
    if (last < 2) return false;
    const ids = sh.getRange(2, 1, last - 1, 1).getDisplayValues().map(function (r) { return r[0]; });
    const idx = ids.indexOf(String(id));
    if (idx < 0) return false;
    sh.deleteRow(idx + 2);
    return true;
  } finally {
    lock.releaseLock();
  }
}
