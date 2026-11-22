/**
 * PORTAL TAUTAN WEBSITE
 * Data tersimpan di Google Sheet (sheet "Links"), logo di Google Drive.
 *
 * CARA PAKAI:
 * 1. Buka script.google.com > Proyek baru (atau Extensions > Apps Script dari Google Sheet).
 * 2. Tempel kode ini di Code.gs, buat file HTML bernama "Index" lalu tempel isi Index.html.
 * 3. Ganti ADMIN_PASSWORD di bawah.
 * 4. Deploy > New deployment > Web app > Execute as: Me, Who has access: Anyone.
 */

const ADMIN_PASSWORD = 'admin123';          // GANTI PASSWORD INI
const SHEET_NAME = 'Links';
const LOGO_FOLDER_NAME = 'Logo Portal Tautan';
const HEADERS = ['ID', 'Judul', 'Link', 'Logo', 'Dibuat'];

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Portal Tautan')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* ---------- Helper ---------- */
function getSheet_() {
  const props = PropertiesService.getScriptProperties();
  let ss = null;
  const id = props.getProperty('SPREADSHEET_ID');
  if (id) {
    try { ss = SpreadsheetApp.openById(id); } catch (e) { ss = null; }
  }
  if (!ss) {
    ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) ss = SpreadsheetApp.create('Database Portal Tautan');
    props.setProperty('SPREADSHEET_ID', ss.getId());
  }
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

function getLogoFolder_() {
  const it = DriveApp.getFoldersByName(LOGO_FOLDER_NAME);
  return it.hasNext() ? it.next() : DriveApp.createFolder(LOGO_FOLDER_NAME);
}

function checkToken_(token) {
  if (!token || !CacheService.getScriptCache().get('adm_' + token)) {
    throw new Error('Sesi admin berakhir, silakan login ulang.');
  }
}

function normalizeUrl_(u) {
  u = String(u || '').trim();
  if (!u) throw new Error('Alamat link wajib diisi.');
  if (!/^https?:\/\//i.test(u)) u = 'https://' + u;
  return u;
}

function saveLogo_(logoData) {
  const bytes = Utilities.base64Decode(logoData.base64);
  const blob = Utilities.newBlob(bytes, logoData.mime || 'image/png', logoData.name || 'logo.png');
  const file = getLogoFolder_().createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return 'https://lh3.googleusercontent.com/d/' + file.getId();
}

/* ---------- Publik ---------- */
function getLinks() {
  const sh = getSheet_();
  const last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, 4).getValues()
    .filter(r => r[0])
    .map(r => ({ id: String(r[0]), judul: r[1], link: r[2], logo: r[3] }));
}

/* ---------- Admin ---------- */
function adminLogin(password) {
  if (String(password) !== ADMIN_PASSWORD) return { ok: false, message: 'Password salah.' };
  const token = Utilities.getUuid();
  CacheService.getScriptCache().put('adm_' + token, '1', 60 * 60 * 3); // 3 jam
  return { ok: true, token: token };
}

function adminSave(token, data) {
  checkToken_(token);
  const judul = String(data.judul || '').trim();
  if (!judul) throw new Error('Nama judul wajib diisi.');
  const link = normalizeUrl_(data.link);

  const sh = getSheet_();
  let logo = data.logoUrl ? String(data.logoUrl).trim() : '';
  if (data.logoData && data.logoData.base64) logo = saveLogo_(data.logoData);

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    if (data.id) {
      const last = sh.getLastRow();
      const ids = last > 1 ? sh.getRange(2, 1, last - 1, 1).getValues().flat().map(String) : [];
      const idx = ids.indexOf(String(data.id));
      if (idx === -1) throw new Error('Data tidak ditemukan.');
      const row = idx + 2;
      const oldLogo = sh.getRange(row, 4).getValue();
      sh.getRange(row, 2, 1, 3).setValues([[judul, link, logo || (data.keepLogo ? oldLogo : '')]]);
    } else {
      sh.appendRow([Utilities.getUuid(), judul, link, logo, new Date()]);
    }
  } finally {
    lock.releaseLock();
  }
  return getLinks();
}

function adminDelete(token, id) {
  checkToken_(token);
  const sh = getSheet_();
  const last = sh.getLastRow();
  if (last < 2) return [];
  const ids = sh.getRange(2, 1, last - 1, 1).getValues().flat().map(String);
  const idx = ids.indexOf(String(id));
  if (idx !== -1) sh.deleteRow(idx + 2);
  return getLinks();
}
