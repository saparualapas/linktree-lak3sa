// ═══════════════════════════════════════════════════════════════
//  LinkSpace — Google Apps Script Backend
//  Salin seluruh kode ini ke Google Apps Script, lalu Deploy
//  sebagai Web App dengan akses "Anyone"
// ═══════════════════════════════════════════════════════════════

// ─── GANTI ID SPREADSHEET ANDA DI SINI ───
const SPREADSHEET_ID = 'GANTI_DENGAN_ID_SPREADSHEET_ANDA';

// Nama sheet (jangan diubah kecuali Anda juga mengubah di kode)
const SHEET_PROFILE  = 'Profile';
const SHEET_LINKS    = 'Links';
const SHEET_FLOATERS = 'Floaters';

// ═══════════════════════════════════════════════════════════════
//  ENTRY POINT — POST
// ═══════════════════════════════════════════════════════════════
function doPost(e) {
  const headers = {
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'POST, GET, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type',
    'Content-Type': 'application/json',
  };

  try {
    const body   = JSON.parse(e.postData.contents);
    const action = body.action;
    let result;

    switch (action) {
      case 'getAll':        result = getAll();                  break;
      case 'saveProfile':   result = saveProfile(body);         break;
      case 'addLink':       result = addLink(body);             break;
      case 'updateLink':    result = updateLink(body);          break;
      case 'deleteLink':    result = deleteLink(body);          break;
      case 'addFloater':    result = addFloater(body);          break;
      case 'deleteFloater': result = deleteFloater(body);       break;
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

// Untuk handle preflight OPTIONS (CORS)
function doGet(e) {
  // Bisa dipakai untuk test: buka URL Apps Script langsung di browser
  const result = getAll();
  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ═══════════════════════════════════════════════════════════════
//  SETUP SHEETS — Jalankan fungsi ini SEKALI untuk membuat sheet
// ═══════════════════════════════════════════════════════════════
function setupSheets() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  // ── Profile sheet ──
  let profileSheet = ss.getSheetByName(SHEET_PROFILE);
  if (!profileSheet) {
    profileSheet = ss.insertSheet(SHEET_PROFILE);
  }
  // Header row
  profileSheet.getRange(1, 1, 1, 6).setValues([
    ['name', 'tagline', 'avatar_url', 'color1', 'color2', 'updated_at']
  ]);
  // Baris data profil (hanya 1 baris, row 2)
  const existingProfile = profileSheet.getRange(2, 1, 1, 1).getValue();
  if (!existingProfile) {
    profileSheet.getRange(2, 1, 1, 6).setValues([
      ['', '', '', '#2979d4', '#1a4a9e', new Date().toISOString()]
    ]);
  }

  // ── Links sheet ──
  let linksSheet = ss.getSheetByName(SHEET_LINKS);
  if (!linksSheet) {
    linksSheet = ss.insertSheet(SHEET_LINKS);
  }
  linksSheet.getRange(1, 1, 1, 5).setValues([
    ['title', 'url', 'icon', 'order', 'created_at']
  ]);

  // ── Floaters sheet ──
  let floatersSheet = ss.getSheetByName(SHEET_FLOATERS);
  if (!floatersSheet) {
    floatersSheet = ss.insertSheet(SHEET_FLOATERS);
  }
  floatersSheet.getRange(1, 1, 1, 5).setValues([
    ['url', 'size', 'left', 'top', 'created_at']
  ]);

  Logger.log('✅ Setup selesai! Sheet Profile, Links, dan Floaters telah dibuat.');
}

// ═══════════════════════════════════════════════════════════════
//  GET ALL DATA
// ═══════════════════════════════════════════════════════════════
function getAll() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  // Profile
  const profileSheet = ss.getSheetByName(SHEET_PROFILE);
  const profData     = profileSheet.getRange(2, 1, 1, 6).getValues()[0];
  const profile = {
    name:       profData[0] || '',
    tagline:    profData[1] || '',
    avatar_url: profData[2] || '',
    color1:     profData[3] || '#2979d4',
    color2:     profData[4] || '#1a4a9e',
  };

  // Links
  const linksSheet = ss.getSheetByName(SHEET_LINKS);
  const linksData  = linksSheet.getDataRange().getValues();
  const links = [];
  for (let i = 1; i < linksData.length; i++) {
    const row = linksData[i];
    if (!row[0] && !row[1]) continue; // Skip empty rows
    links.push({
      title:  row[0] || '',
      url:    row[1] || '',
      icon:   row[2] || '',
      order:  row[3] || 99,
      _row:   i + 1, // Actual sheet row number (1-indexed)
    });
  }

  // Floaters
  const floatersSheet = ss.getSheetByName(SHEET_FLOATERS);
  const floatersData  = floatersSheet.getDataRange().getValues();
  const floaters = [];
  for (let i = 1; i < floatersData.length; i++) {
    const row = floatersData[i];
    if (!row[0]) continue;
    floaters.push({
      url:  row[0] || '',
      size: row[1] || 120,
      left: row[2] || 10,
      top:  row[3] || 20,
      _row: i + 1,
    });
  }

  return { ok: true, profile, links, floaters };
}

// ═══════════════════════════════════════════════════════════════
//  PROFILE
// ═══════════════════════════════════════════════════════════════
function saveProfile(body) {
  const ss          = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet       = ss.getSheetByName(SHEET_PROFILE);
  // Selalu tulis ke baris 2 (1 baris data profil)
  sheet.getRange(2, 1, 1, 6).setValues([[
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
    body.order || 99,
    new Date().toISOString(),
  ]);
  return { ok: true };
}

function updateLink(body) {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_LINKS);
  const row   = parseInt(body.rowIndex);
  if (!row || row < 2) throw new Error('Invalid row index');
  sheet.getRange(row, 1, 1, 5).setValues([[
    body.title || '',
    body.url   || '',
    body.icon  || '',
    body.order || 99,
    new Date().toISOString(),
  ]]);
  return { ok: true };
}

function deleteLink(body) {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_LINKS);
  const row   = parseInt(body.rowIndex);
  if (!row || row < 2) throw new Error('Invalid row index');
  sheet.deleteRow(row);
  return { ok: true };
}

// ═══════════════════════════════════════════════════════════════
//  FLOATERS (Gambar Background)
// ═══════════════════════════════════════════════════════════════
function addFloater(body) {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_FLOATERS);
  sheet.appendRow([
    body.url  || '',
    body.size || 120,
    body.left || 10,
    body.top  || 20,
    new Date().toISOString(),
  ]);
  return { ok: true };
}

function deleteFloater(body) {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_FLOATERS);
  const row   = parseInt(body.rowIndex);
  if (!row || row < 2) throw new Error('Invalid row index');
  sheet.deleteRow(row);
  return { ok: true };
}
