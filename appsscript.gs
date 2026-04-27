// ═══════════════════════════════════════════════════════════════
//  LinkSpace — Google Apps Script Backend
//  Password tersimpan di spreadsheet (sheet "Config")
//  Jalankan setupSheets() SEKALI sebelum deploy
// ═══════════════════════════════════════════════════════════════

// ── GANTI DENGAN ID SPREADSHEET ANDA ──
const SPREADSHEET_ID = 'GANTI_DENGAN_ID_SPREADSHEET_ANDA';

const SHEET_CONFIG   = 'Config';
const SHEET_PROFILE  = 'Profile';
const SHEET_LINKS    = 'Links';
const SHEET_FLOATERS = 'Floaters';

// Password default saat pertama kali setup
const DEFAULT_PASSWORD = 'admin123';

// ═══════════════════════════════════════════════════════════════
//  ENTRY POINT
// ═══════════════════════════════════════════════════════════════
function doPost(e) {
  try {
    const body   = JSON.parse(e.postData.contents);
    const action = body.action;
    let result;

    switch (action) {
      case 'getAll':          result = getAll();                     break;
      case 'login':           result = login(body);                  break;
      case 'changePassword':  result = changePassword(body);         break;
      case 'saveProfile':     result = saveProfile(body);            break;
      case 'addLink':         result = addLink(body);                break;
      case 'updateLink':      result = updateLink(body);             break;
      case 'deleteLink':      result = deleteLink(body);             break;
      case 'addFloater':      result = addFloater(body);             break;
      case 'deleteFloater':   result = deleteFloater(body);          break;
      default:
        result = { ok: false, error: 'Unknown action: ' + action };
    }

    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Untuk test langsung di browser (GET)
function doGet(e) {
  const result = getAll();
  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ═══════════════════════════════════════════════════════════════
//  SETUP — Jalankan fungsi ini SEKALI dari editor Apps Script
//  Menu: Run > setupSheets
// ═══════════════════════════════════════════════════════════════
function setupSheets() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  // ── Config (password) ──
  let cfg = ss.getSheetByName(SHEET_CONFIG);
  if (!cfg) cfg = ss.insertSheet(SHEET_CONFIG);
  cfg.clearContents();
  cfg.getRange(1,1,2,2).setValues([
    ['key',      'value'],
    ['password', DEFAULT_PASSWORD],
  ]);
  cfg.getRange('A:B').setColumnWidth(0, 180);

  // ── Profile ──
  let prof = ss.getSheetByName(SHEET_PROFILE);
  if (!prof) prof = ss.insertSheet(SHEET_PROFILE);
  prof.clearContents();
  prof.getRange(1,1,2,7).setValues([
    ['name','tagline','avatar_url','color1','color2','updated_at','_reserved'],
    ['','','','#2979d4','#1a4a9e',new Date().toISOString(),''],
  ]);

  // ── Links ──
  let lnk = ss.getSheetByName(SHEET_LINKS);
  if (!lnk) lnk = ss.insertSheet(SHEET_LINKS);
  lnk.clearContents();
  lnk.getRange(1,1,1,5).setValues([['title','url','icon','order','created_at']]);

  // ── Floaters ──
  let flt = ss.getSheetByName(SHEET_FLOATERS);
  if (!flt) flt = ss.insertSheet(SHEET_FLOATERS);
  flt.clearContents();
  flt.getRange(1,1,1,5).setValues([['url','size','left','top','created_at']]);

  Logger.log('✅ Setup selesai! Sheet Config, Profile, Links, Floaters dibuat.');
  Logger.log('Password default: ' + DEFAULT_PASSWORD);
}

// ═══════════════════════════════════════════════════════════════
//  HELPERS — Config sheet
// ═══════════════════════════════════════════════════════════════
function getConfig(key) {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_CONFIG);
  const data  = sheet.getDataRange().getValues();
  // Row 0 = header, data mulai row 1
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === key) return String(data[i][1]);
  }
  return null;
}

function setConfig(key, value) {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_CONFIG);
  const data  = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === key) {
      sheet.getRange(i + 1, 2).setValue(value);
      return;
    }
  }
  // Key belum ada, tambahkan baris baru
  sheet.appendRow([key, value]);
}

// ═══════════════════════════════════════════════════════════════
//  AUTH
// ═══════════════════════════════════════════════════════════════
function login(body) {
  const storedPwd = getConfig('password') || DEFAULT_PASSWORD;
  if (body.password !== storedPwd) {
    throw new Error('Password salah');
  }
  // Buat token sederhana berbasis timestamp
  // Token ini tidak divalidasi server (stateless), hanya dipakai
  // di client sebagai tanda sudah login dalam sesi ini.
  const token = 'ls_' + Date.now() + '_' + Math.random().toString(36).slice(2);
  return { ok: true, token: token };
}

function changePassword(body) {
  const storedPwd = getConfig('password') || DEFAULT_PASSWORD;
  if (body.oldPassword !== storedPwd) {
    throw new Error('Password lama salah');
  }
  if (!body.newPassword || body.newPassword.length < 8) {
    throw new Error('Password baru minimal 8 karakter');
  }
  setConfig('password', body.newPassword);
  return { ok: true };
}

// ═══════════════════════════════════════════════════════════════
//  GET ALL
// ═══════════════════════════════════════════════════════════════
function getAll() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  // Profile
  const profSheet = ss.getSheetByName(SHEET_PROFILE);
  const profRow   = profSheet.getRange(2,1,1,6).getValues()[0];
  const profile = {
    name:       String(profRow[0] || ''),
    tagline:    String(profRow[1] || ''),
    avatar_url: String(profRow[2] || ''),
    color1:     String(profRow[3] || '#2979d4'),
    color2:     String(profRow[4] || '#1a4a9e'),
  };

  // Links
  const lnkSheet = ss.getSheetByName(SHEET_LINKS);
  const lnkData  = lnkSheet.getDataRange().getValues();
  const links = [];
  for (let i = 1; i < lnkData.length; i++) {
    const r = lnkData[i];
    if (!r[0] && !r[1]) continue;
    links.push({
      title: String(r[0] || ''),
      url:   String(r[1] || ''),
      icon:  String(r[2] || ''),
      order: r[3] || 99,
      _row:  i + 1,
    });
  }

  // Floaters
  const fltSheet = ss.getSheetByName(SHEET_FLOATERS);
  const fltData  = fltSheet.getDataRange().getValues();
  const floaters = [];
  for (let i = 1; i < fltData.length; i++) {
    const r = fltData[i];
    if (!r[0]) continue;
    floaters.push({
      url:  String(r[0] || ''),
      size: Number(r[1] || 120),
      left: Number(r[2] || 10),
      top:  Number(r[3] || 20),
      _row: i + 1,
    });
  }

  return { ok: true, profile, links, floaters };
}

// ═══════════════════════════════════════════════════════════════
//  PROFILE
// ═══════════════════════════════════════════════════════════════
function saveProfile(body) {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_PROFILE);
  sheet.getRange(2,1,1,6).setValues([[
    body.name       || '',
    body.tagline    || '',
    body.avatar_url || '',
    body.color1     || '#2979d4',
    body.color2     || '#1a4a9e',
    new Date().toISOString(),
  ]]);
  return { ok: true };
}

// ═══════════════════════════════════════════════════════════════
//  LINKS
// ═══════════════════════════════════════════════════════════════
function addLink(body) {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_LINKS);
  sheet.appendRow([
    body.title || '',
    body.url   || '',
    body.icon  || '',
    Number(body.order) || 99,
    new Date().toISOString(),
  ]);
  return { ok: true };
}

function updateLink(body) {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_LINKS);
  const row   = parseInt(body.rowIndex);
  if (!row || row < 2) throw new Error('Row tidak valid');
  sheet.getRange(row,1,1,5).setValues([[
    body.title || '',
    body.url   || '',
    body.icon  || '',
    Number(body.order) || 99,
    new Date().toISOString(),
  ]]);
  return { ok: true };
}

function deleteLink(body) {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_LINKS);
  const row   = parseInt(body.rowIndex);
  if (!row || row < 2) throw new Error('Row tidak valid');
  sheet.deleteRow(row);
  return { ok: true };
}

// ═══════════════════════════════════════════════════════════════
//  FLOATERS
// ═══════════════════════════════════════════════════════════════
function addFloater(body) {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_FLOATERS);
  sheet.appendRow([
    body.url  || '',
    Number(body.size) || 120,
    Number(body.left) || 10,
    Number(body.top)  || 20,
    new Date().toISOString(),
  ]);
  return { ok: true };
}

function deleteFloater(body) {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_FLOATERS);
  const row   = parseInt(body.rowIndex);
  if (!row || row < 2) throw new Error('Row tidak valid');
  sheet.deleteRow(row);
  return { ok: true };
}
