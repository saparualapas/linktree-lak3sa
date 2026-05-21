// ═══════════════════════════════════════════════════════════════
//  Lak3saLink — Google Apps Script Backend
//  Password tersimpan di spreadsheet (sheet "Config")
//  Jalankan setupSheets() SEKALI sebelum deploy
// ═══════════════════════════════════════════════════════════════

// ── GANTI DENGAN ID SPREADSHEET ANDA ──
const SPREADSHEET_ID = 'GANTI_DENGAN_ID_SPREADSHEET_ANDA';

const SHEET_CONFIG   = 'Config';
const SHEET_PROFILE  = 'Profile';
const SHEET_LINKS    = 'Links';
const SHEET_FLOATERS = 'Floaters';
const SHEET_PHOTOS   = 'Photos';

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
      case 'addPhoto':        result = addPhoto(body);               break;
      case 'updatePhoto':     result = updatePhoto(body);            break;
      case 'deletePhoto':     result = deletePhoto(body);            break;
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

  // ── Photos ──
  let pht = ss.getSheetByName(SHEET_PHOTOS);
  if (!pht) pht = ss.insertSheet(SHEET_PHOTOS);
  pht.clearContents();
  // data_b64 berisi string base64 data URL foto (bisa panjang, set kolom lebar)
  pht.getRange(1,1,1,4).setValues([['title','order','data_b64','created_at']]);

  Logger.log('✅ Setup selesai! Sheet Config, Profile, Links, Floaters, Photos dibuat.');
  Logger.log('Password default: ' + DEFAULT_PASSWORD);
}

// ═══════════════════════════════════════════════════════════════
//  HELPERS — Config sheet
// ═══════════════════════════════════════════════════════════════
function getConfig(key) {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_CONFIG);
  const data  = sheet.getDataRange().getValues();
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

  // Photos
  const phtSheet = ss.getSheetByName(SHEET_PHOTOS);
  const photos   = [];
  if (phtSheet) {
    const phtData = phtSheet.getDataRange().getValues();
    for (let i = 1; i < phtData.length; i++) {
      const r = phtData[i];
      if (!r[0]) continue;
      photos.push({
        title:    String(r[0] || ''),
        order:    r[1] || 99,
        data_b64: String(r[2] || ''),
        _row:     i + 1,
      });
    }
  }

  return { ok: true, profile, links, floaters, photos };
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

// ═══════════════════════════════════════════════════════════════
//  PHOTOS
//  Foto disimpan sebagai data URL base64 langsung di spreadsheet.
//  Batasi ukuran file di sisi client (maks ~1.5 MB sebelum encode).
//  Base64 dari 1.5 MB ≈ 2 MB string — masih dalam batas sel Sheets.
// ═══════════════════════════════════════════════════════════════
function addPhoto(body) {
  if (!body.title) throw new Error('Judul foto wajib diisi');
  if (!body.data_b64) throw new Error('Data foto tidak ada');

  // Validasi ukuran: base64 string > ~2.1 MB kemungkinan terlalu besar
  if (body.data_b64.length > 2200000) {
    throw new Error('Foto terlalu besar. Maksimal 1.5 MB.');
  }

  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet   = ss.getSheetByName(SHEET_PHOTOS);

  // Buat sheet Photos jika belum ada (untuk instalasi lama)
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_PHOTOS);
    sheet.getRange(1,1,1,4).setValues([['title','order','data_b64','created_at']]);
  }

  sheet.appendRow([
    body.title || '',
    Number(body.order) || 99,
    body.data_b64,
    new Date().toISOString(),
  ]);
  return { ok: true };
}

function updatePhoto(body) {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_PHOTOS);
  if (!sheet) throw new Error('Sheet Photos tidak ditemukan');
  const row = parseInt(body.rowIndex);
  if (!row || row < 2) throw new Error('Row tidak valid');
  // Hanya update judul dan urutan (tidak ubah data foto)
  sheet.getRange(row, 1).setValue(body.title || '');
  sheet.getRange(row, 2).setValue(Number(body.order) || 99);
  return { ok: true };
}

function deletePhoto(body) {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_PHOTOS);
  if (!sheet) throw new Error('Sheet Photos tidak ditemukan');
  const row = parseInt(body.rowIndex);
  if (!row || row < 2) throw new Error('Row tidak valid');
  sheet.deleteRow(row);
  return { ok: true };
}
