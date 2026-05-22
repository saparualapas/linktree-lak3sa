// ═══════════════════════════════════════════════════════════════
//  Lak3saLink — Google Apps Script Backend
//  Jalankan setupSheets() SEKALI sebelum deploy pertama kali
// ═══════════════════════════════════════════════════════════════

const SPREADSHEET_ID = 'GANTI_DENGAN_ID_SPREADSHEET_ANDA';
const SHEET_CONFIG   = 'Config';
const SHEET_PROFILE  = 'Profile';
const SHEET_LINKS    = 'Links';
const SHEET_FLOATERS = 'Floaters';
const SHEET_PHOTOS   = 'Photos';
const DEFAULT_PWD    = 'admin123';
const MAX_PHOTO_B64  = 690000;   // ~500 KB file → ~690 KB base64
const MAX_AVATAR_B64 = 415000;   // ~300 KB file → ~415 KB base64
const MAX_FLT_B64    = 4150000;  // ~3 MB file → ~4.1 MB base64

// ═══════════════════════════════════════════════════════════════
//  ENTRY POINTS
// ═══════════════════════════════════════════════════════════════
function doPost(e) {
  try {
    const body   = JSON.parse(e.postData.contents);
    const action = String(body.action || '');
    const ALLOWED = ['getAll','login','changePassword','saveProfile',
                     'addLink','updateLink','deleteLink',
                     'addFloater','updateFloater','deleteFloater',
                     'addPhoto','updatePhoto','deletePhoto'];
    if (!ALLOWED.includes(action)) {
      return json({ ok:false, error:'Action tidak dikenali' });
    }
    const handlers = {
      getAll, login, changePassword, saveProfile,
      addLink, updateLink, deleteLink,
      addFloater, updateFloater, deleteFloater,
      addPhoto, updatePhoto, deletePhoto,
    };
    return json(handlers[action](body));
  } catch (err) {
    Logger.log('doPost error: ' + err.toString());
    return json({ ok:false, error: err.message || err.toString() });
  }
}

function doGet(e) {
  return json(getAll());
}

function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ═══════════════════════════════════════════════════════════════
//  SETUP — Jalankan sekali dari Apps Script Editor
// ═══════════════════════════════════════════════════════════════
function setupSheets() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  function ensureSheet(name, headers, defaultRow) {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    sh.clearContents();
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    if (defaultRow) sh.getRange(2, 1, 1, defaultRow.length).setValues([defaultRow]);
    return sh;
  }

  ensureSheet(SHEET_CONFIG,   ['key','value'],
    [['password', DEFAULT_PWD]]);
  // Config hanya 1 data row, clear lagi dan isi 2 baris (header + password)
  const cfg = ss.getSheetByName(SHEET_CONFIG);
  cfg.clearContents();
  cfg.getRange(1,1,2,2).setValues([['key','value'],['password',DEFAULT_PWD]]);

  ensureSheet(SHEET_PROFILE,
    ['name','tagline','avatar_url','avatar_b64','color1','color2','updated_at'],
    ['','','','','#2979d4','#1a4a9e', new Date().toISOString()]);

  ensureSheet(SHEET_LINKS,    ['title','url','icon','order','created_at']);
  ensureSheet(SHEET_FLOATERS, ['url','data_b64','size','left','top','created_at']);
  ensureSheet(SHEET_PHOTOS,   ['title','order','data_b64','created_at']);

  Logger.log('✅ Setup selesai. Password default: ' + DEFAULT_PWD);
}

// ═══════════════════════════════════════════════════════════════
//  HELPERS
// ═══════════════════════════════════════════════════════════════
function ss() { return SpreadsheetApp.openById(SPREADSHEET_ID); }
function sh(name) {
  const sheet = ss().getSheetByName(name);
  if (!sheet) throw new Error('Sheet "' + name + '" tidak ditemukan. Jalankan setupSheets().');
  return sheet;
}

function getConfig(key) {
  const data = sh(SHEET_CONFIG).getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === key) return String(data[i][1]);
  }
  return null;
}

function setConfig(key, value) {
  const sheet = sh(SHEET_CONFIG);
  const data  = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === key) { sheet.getRange(i+1, 2).setValue(value); return; }
  }
  sheet.appendRow([key, value]);
}

function validRow(row) {
  const n = parseInt(row);
  if (!n || n < 2 || isNaN(n)) throw new Error('Row index tidak valid');
  return n;
}

function sanitize(s, max) {
  return String(s || '').slice(0, max || 500);
}

// ═══════════════════════════════════════════════════════════════
//  AUTH
// ═══════════════════════════════════════════════════════════════
function login(body) {
  const pwd = getConfig('password') || DEFAULT_PWD;
  if (!body.password || body.password !== pwd) throw new Error('Password salah');
  const token = 'lk_' + Date.now().toString(36) + '_' + Math.random().toString(36).slice(2,8);
  return { ok:true, token };
}

function changePassword(body) {
  const pwd = getConfig('password') || DEFAULT_PWD;
  if (body.oldPassword !== pwd) throw new Error('Password lama salah');
  const np = sanitize(body.newPassword, 64);
  if (np.length < 8) throw new Error('Password baru minimal 8 karakter');
  setConfig('password', np);
  return { ok:true };
}

// ═══════════════════════════════════════════════════════════════
//  GET ALL
// ═══════════════════════════════════════════════════════════════
function getAll() {
  // Profile — kolom: name,tagline,avatar_url,avatar_b64,color1,color2,updated_at
  const pr = sh(SHEET_PROFILE).getRange(2,1,1,7).getValues()[0];
  const profile = {
    name:        String(pr[0]||''),
    tagline:     String(pr[1]||''),
    avatar_url:  String(pr[2]||''),
    avatar_b64:  String(pr[3]||''),
    color1:      String(pr[4]||'#2979d4'),
    color2:      String(pr[5]||'#1a4a9e'),
  };

  // Links
  const ld = sh(SHEET_LINKS).getDataRange().getValues();
  const links = [];
  for (let i=1; i<ld.length; i++) {
    const r=ld[i]; if(!r[0]&&!r[1]) continue;
    links.push({ title:String(r[0]||''), url:String(r[1]||''), icon:String(r[2]||''), order:r[3]||99, _row:i+1 });
  }

  // Floaters — kolom: url,data_b64,size,left,top,created_at
  const fd = sh(SHEET_FLOATERS).getDataRange().getValues();
  const floaters = [];
  for (let i=1; i<fd.length; i++) {
    const r=fd[i]; if(!r[0]&&!r[1]) continue;
    floaters.push({ url:String(r[0]||''), data_b64:String(r[1]||''), size:Number(r[2]||120), left:Number(r[3]||10), top:Number(r[4]||20), _row:i+1 });
  }

  // Photos
  const photos = [];
  const phSheet = ss().getSheetByName(SHEET_PHOTOS);
  if (phSheet) {
    const pd = phSheet.getDataRange().getValues();
    for (let i=1; i<pd.length; i++) {
      const r=pd[i]; if(!r[0]) continue;
      photos.push({ title:String(r[0]||''), order:r[1]||99, data_b64:String(r[2]||''), _row:i+1 });
    }
  }

  return { ok:true, profile, links, floaters, photos };
}

// ═══════════════════════════════════════════════════════════════
//  PROFILE
// ═══════════════════════════════════════════════════════════════
function saveProfile(body) {
  // Validasi avatar_b64 jika ada
  const avB64 = body.avatar_b64 || '';
  if (avB64) {
    if (avB64.length > MAX_AVATAR_B64) throw new Error('Foto profil terlalu besar. Maks 300 KB.');
    if (!avB64.startsWith('data:image/')) throw new Error('Format foto profil tidak valid.');
  }
  sh(SHEET_PROFILE).getRange(2,1,1,7).setValues([[
    sanitize(body.name, 60),
    sanitize(body.tagline, 120),
    sanitize(body.avatar_url, 300),
    avB64,
    sanitize(body.color1, 10) || '#2979d4',
    sanitize(body.color2, 10) || '#1a4a9e',
    new Date().toISOString(),
  ]]);
  return { ok:true };
}

// ═══════════════════════════════════════════════════════════════
//  LINKS
// ═══════════════════════════════════════════════════════════════
function addLink(body) {
  if (!body.title || !body.url) throw new Error('Judul dan URL wajib diisi');
  sh(SHEET_LINKS).appendRow([
    sanitize(body.title, 80), sanitize(body.url, 500),
    sanitize(body.icon, 50), Number(body.order)||99, new Date().toISOString(),
  ]);
  return { ok:true };
}

function updateLink(body) {
  const row = validRow(body.rowIndex);
  if (!body.url) throw new Error('URL wajib diisi');
  sh(SHEET_LINKS).getRange(row,1,1,5).setValues([[
    sanitize(body.title,80), sanitize(body.url,500),
    sanitize(body.icon,50), Number(body.order)||99, new Date().toISOString(),
  ]]);
  return { ok:true };
}

function deleteLink(body) {
  sh(SHEET_LINKS).deleteRow(validRow(body.rowIndex));
  return { ok:true };
}

// ═══════════════════════════════════════════════════════════════
//  FLOATERS
// ═══════════════════════════════════════════════════════════════
function addFloater(body) {
  const url = sanitize(body.url, 300);
  const b64 = body.data_b64 || '';
  if (!url && !b64) throw new Error('URL atau file gambar wajib diisi');
  if (b64) {
    if (b64.length > MAX_FLT_B64) throw new Error('Gambar BG terlalu besar. Maks 3 MB.');
    if (!b64.startsWith('data:image/')) throw new Error('Format gambar tidak valid.');
  }
  sh(SHEET_FLOATERS).appendRow([
    url, b64,
    Number(body.size)||120, Number(body.left)||10, Number(body.top)||20,
    new Date().toISOString(),
  ]);
  return { ok:true };
}

function updateFloater(body) {
  const row = validRow(body.rowIndex);
  const url = sanitize(body.url, 300);
  const b64 = body.data_b64 || '';
  if (!url && !b64) throw new Error('URL wajib diisi');
  sh(SHEET_FLOATERS).getRange(row,1,1,6).setValues([[
    url, b64,
    Number(body.size)||120, Number(body.left)||0, Number(body.top)||0,
    new Date().toISOString(),
  ]]);
  return { ok:true };
}

function deleteFloater(body) {
  sh(SHEET_FLOATERS).deleteRow(validRow(body.rowIndex));
  return { ok:true };
}

// ═══════════════════════════════════════════════════════════════
//  PHOTOS
//  Foto disimpan Base64 di Sheets. Batas 300 KB file ≈ ~410 KB b64.
//  Batas sel Sheets 50.000 karakter — untuk keamanan, validasi di sini.
// ═══════════════════════════════════════════════════════════════
function getOrCreatePhotosSheet() {
  let ph = ss().getSheetByName(SHEET_PHOTOS);
  if (!ph) {
    ph = ss().insertSheet(SHEET_PHOTOS);
    ph.getRange(1,1,1,4).setValues([['title','order','data_b64','created_at']]);
  }
  return ph;
}

function addPhoto(body) {
  if (!body.title) throw new Error('Judul foto wajib diisi');
  if (!body.data_b64) throw new Error('Data foto tidak ada');
  // Validasi panjang base64 (300 KB file → ~410.000 char base64)
  if (body.data_b64.length > MAX_PHOTO_B64) {
    throw new Error('Foto terlalu besar. Maksimal 500 KB.');
  }
  // Validasi format data URL
  if (!body.data_b64.startsWith('data:image/')) {
    throw new Error('Format foto tidak valid.');
  }
  getOrCreatePhotosSheet().appendRow([
    sanitize(body.title, 80), Number(body.order)||99,
    body.data_b64, new Date().toISOString(),
  ]);
  return { ok:true };
}

function updatePhoto(body) {
  const row = validRow(body.rowIndex);
  // Hanya update judul & urutan, data foto tidak diubah
  const ph = getOrCreatePhotosSheet();
  ph.getRange(row, 1).setValue(sanitize(body.title, 80));
  ph.getRange(row, 2).setValue(Number(body.order)||99);
  return { ok:true };
}

function deletePhoto(body) {
  getOrCreatePhotosSheet().deleteRow(validRow(body.rowIndex));
  return { ok:true };
}
