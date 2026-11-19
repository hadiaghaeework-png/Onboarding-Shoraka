/*******************************************************
 * ابزارهای مشترک سمت سرور
 *******************************************************/

function ss_() {
  const id = CONFIG.SPREADSHEET_ID ||
    PropertiesService.getScriptProperties().getProperty('ACADEMY_SPREADSHEET_ID');
  if (!id) throw new Error('ابتدا تابع setupAcademy را یک‌بار اجرا کنید.');
  return SpreadsheetApp.openById(id);
}

function sheet_(name) {
  const sh = ss_().getSheetByName(name);
  if (!sh) throw new Error('شیت «' + name + '» وجود ندارد. setupAcademy را اجرا کنید.');
  return sh;
}

function ensureSheet_(name, headers) {
  const ss = ss_();
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) sh.appendRow(headers);
  else {
    // اگر ستون جدیدی به تعریف اضافه شده باشد، سرستون‌ها را کامل کن
    const cur = sh.getRange(1, 1, 1, Math.max(headers.length, sh.getLastColumn())).getValues()[0];
    headers.forEach((h, i) => { if (cur[i] !== h) sh.getRange(1, i + 1).setValue(h); });
  }
  return sh;
}

function rows_(name) {
  const sh = sheet_(name);
  if (sh.getLastRow() < 2) return [];
  const values = sh.getDataRange().getValues();
  const headers = values[0];
  const out = [];
  for (let i = 1; i < values.length; i++) {
    const r = values[i];
    if (r.every(c => c === '' || c === null)) continue;
    const o = { __row: i + 1 };
    headers.forEach((h, j) => { o[h] = r[j] instanceof Date ? r[j].toISOString() : r[j]; });
    out.push(o);
  }
  return out;
}

function newId_(prefix) {
  return prefix + '-' + Utilities.getUuid().slice(0, 8) + Date.now().toString(36);
}

function truthy_(v) {
  if (v === true) return true;
  const s = String(v).toLowerCase().trim();
  return s === 'true' || s === 'yes' || s === '1';
}

function splitCsv_(v) {
  return String(v || '').split(',').map(s => s.trim()).filter(Boolean);
}

function escapeQuery_(s) {
  return String(s).replace(/\\/g, '\\\\').replace(/'/g, "\\'");
}

function fileIcon_(mime) {
  mime = String(mime || '');
  if (mime === 'application/vnd.google-apps.folder') return 'FOLDER';
  if (mime === 'application/pdf') return 'PDF';
  if (mime.indexOf('video/') === 0) return 'VIDEO';
  if (mime.indexOf('image/') === 0) return 'IMAGE';
  if (mime.indexOf('audio/') === 0) return 'AUDIO';
  if (mime.indexOf('spreadsheet') !== -1 || mime.indexOf('excel') !== -1 || mime.indexOf('csv') !== -1) return 'SHEET';
  if (mime.indexOf('document') !== -1 || mime.indexOf('word') !== -1) return 'DOC';
  if (mime.indexOf('presentation') !== -1 || mime.indexOf('powerpoint') !== -1 || mime.indexOf('slides') !== -1) return 'SLIDES';
  if (mime.indexOf('zip') !== -1 || mime.indexOf('compressed') !== -1) return 'ARCHIVE';
  if (mime.indexOf('text/') === 0) return 'TEXT';
  return 'FILE';
}

function formatSize_(bytes) {
  if (!bytes) return '';
  const n = Number(bytes);
  if (isNaN(n)) return '';
  if (n >= 1073741824) return (n / 1073741824).toFixed(1) + ' GB';
  if (n >= 1048576) return (n / 1048576).toFixed(1) + ' MB';
  return Math.max(1, Math.round(n / 1024)) + ' KB';
}

function formatDateFa_(d) {
  try { return new Date(d).toLocaleDateString('fa-IR'); } catch (e) { return ''; }
}

function nowIso_() { return new Date().toISOString(); }

function ok_(data) { return Object.assign({ ok: true }, data || {}); }

function cacheGet_(key) {
  try { const v = CacheService.getScriptCache().get(key); return v ? JSON.parse(v) : null; }
  catch (e) { return null; }
}
function cachePut_(key, value, seconds) {
  try { CacheService.getScriptCache().put(key, JSON.stringify(value), seconds || 600); }
  catch (e) { /* اندازه بیش از حد — کش اختیاری است */ }
}
function cacheRemove_(key) { try { CacheService.getScriptCache().remove(key); } catch (e) {} }

function logActivity_(email, action, detail) {
  try {
    sheet_('Activity').appendRow([newId_('ACT'), nowIso_(), email, action, detail || '']);
  } catch (e) { /* ثبت فعالیت نباید عملیات اصلی را متوقف کند */ }
}
