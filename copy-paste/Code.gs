// ===== Config.gs =====
/* پیکربندی نسخه در حال توسعه. احراز هویت و مجوز از شیت Users خوانده می‌شود؛
   استقرار آزمایشی با اجرای مالک و محدودیت دامنه، تنها پس از تأیید هویت Workspace. */

const CONFIG = {
  ROOT_FOLDER_ID: 'PUT_YOUR_DRIVE_FOLDER_ID_HERE',
  PAGE_SIZE: 12,
  SPREADSHEET_ID: '',                 // خالی = همان فایل ساخته‌شده توسط setupAcademy در Script Properties
  ORG_NAME: 'بیمه‌بازار',
  ORG_NAME_EN: 'BimeBazar',
  LOG_LOGINS: false                   // ثبت ورود در Activity (می‌تواند پرحجم شود)
};

const TEAM_CONFIG = [
  { key: 'account',   name: 'Account',   aliases: ['Account', 'اکانت'],                 color: 'blue',   icon: '◈' },
  { key: 'mlm',       name: 'MLM',       aliases: ['MLM'],                              color: 'orange', icon: '◇' },
  { key: 'bnpl',      name: 'BNPL',      aliases: ['BNPL'],                             color: 'green',  icon: '▣' },
  { key: 'telesales', name: 'Tele Sales', aliases: ['Tele Sales', 'Telesales', 'تله‌سلز'], color: 'purple', icon: '◉' }
];

// مدیر اولیه فقط با راه‌اندازی از ویرایشگر در شیت Users ثبت می‌شود.


const SHEET_DEFS = {
  Users:        ['Email', 'Name', 'Role', 'Teams', 'MentorEmail', 'CanManage', 'Active', 'CreatedAt'],
  Credentials:  ['Username', 'Salt', 'PasswordHash', 'Iterations', 'Version'],
  Paths:        ['PathId', 'TeamKey', 'Title', 'Description', 'Order', 'Active', 'CreatedAt'],
  PathItems:    ['ItemId', 'PathId', 'Order', 'Title', 'DriveItemId', 'ItemType', 'Required', 'QuizId', 'EstimatedMin', 'Active'],
  Progress:     ['Email', 'PathItemId', 'Status', 'CompletedAt', 'Score'],
  Quizzes:      ['QuizId', 'RefType', 'RefId', 'Title', 'Description', 'PassScore', 'TimeLimitMin', 'MaxAttempts', 'ShuffleQuestions', 'ShuffleOptions', 'ShowAnswers', 'Active', 'CreatedAt', 'TeamKey'],
  Questions:    ['QuestionId', 'QuizId', 'Order', 'Type', 'Text', 'OptionsJSON', 'CorrectJSON', 'Points'],
  Attempts:     ['AttemptId', 'Timestamp', 'Email', 'QuizId', 'Score', 'Passed', 'Status', 'AnswersJSON', 'Feedback', 'TimeSpentSec'],
  Announcements:['Id', 'Title', 'Body', 'Audience', 'Active', 'CreatedAt'],
  Activity:     ['Id', 'Timestamp', 'Email', 'Action', 'Detail'],
  MentorNotes:  ['Id', 'Timestamp', 'MentorEmail', 'MenteeEmail', 'Note'],
  Notifications:['Id', 'Email', 'Title', 'Body', 'CreatedAt', 'PathId'],
  NotificationReads:['Email', 'NotificationId', 'ReadAt']
};

const ROLE_LABELS = {
  admin: 'مدیر سامانه',
  team_manager: 'مدیر تیم',
  mentor: 'منتور',
  learner: 'کارشناس'
};

const QUESTION_TYPES = ['mcq', 'multi', 'tf', 'essay', 'scenario'];


// ===== Util.gs =====
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


// ===== Auth.gs =====
/*******************************************************
 * هویت و کنترل دسترسی (RBAC مبتنی بر نقش و تیم)
 * همه توابع سرویس‌دهنده باید از این لایه عبور کنند.
 *******************************************************/

function normalizeRole_(role) {
  const r = String(role || '').toLowerCase().trim();
  return ['admin', 'mentor', 'learner'].indexOf(r) !== -1 ? r : 'learner';
}

function roleLabel_(role) {
  return ROLE_LABELS[role] || ROLE_LABELS.learner;
}

// این متغیر فقط در اجرای apiRequest از نشست معتبر پر می‌شود.
let CURRENT_SESSION_USERNAME_ = '';
let CURRENT_USER_CONTEXT_ = null;

function getUserContext() {
  const username = CURRENT_SESSION_USERNAME_;
  if (!username) throw new Error('نشست ورود معتبر نیست. دوباره وارد شوید.');
  if (CURRENT_USER_CONTEXT_) return CURRENT_USER_CONTEXT_;
  const u = rows_('Users').find(x => String(x.Email).toLowerCase().trim() === username);
  if (!u || !truthy_(u.Active)) throw new Error('حساب شما غیرفعال است یا وجود ندارد.');
  const storedRole = String(u.Role || '').trim().toLowerCase();
  if (['admin', 'mentor', 'learner'].indexOf(storedRole) === -1) throw new Error('نقش حساب نامعتبر است.');
  const role = normalizeRole_(storedRole), teams = splitCsv_(u.Teams);
  if (role === 'learner' && (teams.length !== 1 || !TEAM_CONFIG.some(t => t.key === teams[0]))) throw new Error('تیم این کارشناس باید دقیقاً یک تیم معتبر باشد.');
  if (role === 'mentor' && (!teams.length || teams.some(t => !TEAM_CONFIG.some(x => x.key === t)))) throw new Error('تیم‌های منتور نامعتبر است.');
  return CURRENT_USER_CONTEXT_ = {email: username, name: String(u.Name || username), role: role, roleLabel: roleLabel_(role), teams: teams, canManage: role === 'admin', mentorEmail: String(u.MentorEmail || '').toLowerCase(), demo: false};
}

// PBKDF2-HMAC-SHA256 with a unique per-user salt. Never store raw passwords in Sheets.
function derivePassword_(password, salt, iterations) {
  const key = Utilities.newBlob(String(password), 'text/plain').getBytes();
  const saltBytes = Utilities.base64Decode(salt);
  let u = Utilities.computeHmacSha256Signature(saltBytes.concat([0, 0, 0, 1]), key);
  const result = u.slice();
  for (let i = 1; i < iterations; i++) {
    u = Utilities.computeHmacSha256Signature(u, key);
    for (let j = 0; j < result.length; j++) result[j] = (result[j] ^ u[j]);
  }
  return Utilities.base64Encode(result);
}
function newSecret_() {
  return (Utilities.getUuid() + Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, '').toLowerCase();
}
function normalizeUsername_(name) { return String(name || '').toLowerCase().trim(); }
function savePassword_(username, password) {
  username = normalizeUsername_(username);
  if (!/^[a-z0-9._@-]{3,100}$/.test(username)) throw new Error('نام کاربری باید ۳ تا ۱۰۰ حرف انگلیسی، عدد یا نشانه مجاز باشد.');
  if (typeof password !== 'string' || password.length < 12 || password.length > 128) throw new Error('رمز باید بین ۱۲ تا ۱۲۸ نویسه باشد.');
  const salt = Utilities.base64Encode(Utilities.newBlob(newSecret_()).getBytes());
  const iterations = 20000;
  const row = [username, salt, derivePassword_(password, salt, iterations), iterations, newSecret_()];
  const sh = sheet_('Credentials'), values = sh.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) if (normalizeUsername_(values[i][0]) === username) { sh.getRange(i + 1, 1, 1, row.length).setValues([row]); return; }
  sh.appendRow(row);
}
function constantEqual_(a, b) {
  a = String(a); b = String(b); let n = a.length ^ b.length;
  for (let i = 0; i < Math.max(a.length, b.length); i++) n |= (a.charCodeAt(i) || 0) ^ (b.charCodeAt(i) || 0);
  return n === 0;
}

// Security-critical cache failures must fail closed; never use the optional UI cache wrappers here.
function securityCacheGet_(key) { const raw = CacheService.getScriptCache().get(key); return raw ? JSON.parse(raw) : null; }
function securityCachePut_(key, value, ttl) { CacheService.getScriptCache().put(key, JSON.stringify(value), ttl); }
function securityCacheRemove_(key) { CacheService.getScriptCache().remove(key); }

/** Public login entry; all other privileged operations require apiRequest. */
function login(username, password) {
  username = normalizeUsername_(username);
  if (!/^[a-z0-9._@-]{3,100}$/.test(username) || typeof password !== 'string' || password.length > 128) throw new Error('نام کاربری یا رمز عبور نادرست است.');
  const failKey = 'login_fail_' + Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, username));
  const lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    const failures = securityCacheGet_(failKey) || { count: 0 };
    if (failures.count >= 5) throw new Error('تعداد تلاش ناموفق بیش از حد است. ۱۵ دقیقه بعد دوباره تلاش کنید.');
    failures.count++; securityCachePut_(failKey, failures, 900);
  } finally { lock.releaseLock(); }
  const cred = rows_('Credentials').find(x => normalizeUsername_(x.Username) === username);
  const user = rows_('Users').find(x => normalizeUsername_(x.Email) === username);
  const dummy = Utilities.base64Encode(Utilities.newBlob('fixed-dummy-salt-for-login').getBytes());
  const actual = derivePassword_(password, cred ? String(cred.Salt) : dummy, cred ? Number(cred.Iterations) : 20000);
  if (!cred || !user || !truthy_(user.Active) || !constantEqual_(actual, String(cred.PasswordHash))) throw new Error('نام کاربری یا رمز عبور نادرست است.');
  // لاگ ورود پیش از اعلام موفقیت به‌صورت همزمان در Sheet ذخیره می‌شود.
  const auditLock = LockService.getScriptLock(); auditLock.waitLock(30000);
  try { sheet_('Activity').appendRow([newId_('ACT'), nowIso_(), username, 'login', 'ورود موفق']); }
  finally { auditLock.releaseLock(); }
  securityCacheRemove_(failKey);
  cachePut_('presence:' + username, {at: nowIso_()}, 180);
  const token = newSecret_();
  securityCachePut_('session_' + token, {username: username, version: String(cred.Version), expires: Date.now() + 6*60*60*1000}, 21600);
  return ok_({token: token});
}

function logout(token) { if (typeof token === 'string' && /^[a-f0-9]{96}$/.test(token)) securityCacheRemove_('session_' + token); return ok_({}); }

function apiRequest(token, operation, args) {
  if (typeof token !== 'string' || !/^[a-f0-9]{96}$/.test(token)) throw new Error('نشست ورود معتبر نیست. دوباره وارد شوید.');
  const session = securityCacheGet_('session_' + token);
  if (!session || Date.now() >= session.expires) throw new Error('نشست ورود منقضی شده است. دوباره وارد شوید.');
  const cred = rows_('Credentials').find(c => normalizeUsername_(c.Username) === session.username);
  if (!cred || String(cred.Version) !== session.version) throw new Error('نشست ورود منقضی شده است. دوباره وارد شوید.');
  const endpoints = {
    getBootstrap, listMyPaths, getPath, markItemComplete, getLibraryPage, getItemDetail, markDriveItemComplete,
    getQuizForAttempt, submitQuiz, getMyTranscript, listAnnouncements, listNotifications, markNotificationRead,
    adminListUsers, adminUpsertUser, adminSetUserActive, adminTeamsStatus, adminListPaths, adminSavePath, adminDeletePath,
    adminPathItems, adminSavePathItem, adminDeletePathItem, adminDriveMeta, adminListQuizzes, adminGetQuiz, adminSaveQuiz,
    adminDeleteQuiz, adminSaveQuestion, adminDeleteQuestion, adminListAttempts, adminGetAttemptDetail, adminReviewAttempt,
    adminListAnnouncements, adminSaveAnnouncement, adminDeleteAnnouncement, reportSummary, reportTeamMembers, reportMember,
    reportIncomplete, reportExportCsv, addMentorNote, presencePing
  };
  if (!Object.prototype.hasOwnProperty.call(endpoints, operation) || !Array.isArray(args) || args.length > 10) throw new Error('درخواست نامعتبر است.');
  CURRENT_SESSION_USERNAME_ = session.username;
  try { getUserContext(); return endpoints[operation].apply(null, args); }
  finally { CURRENT_SESSION_USERNAME_ = ''; CURRENT_USER_CONTEXT_ = null; }
}

function requireRole_(user, roles) {
  if (roles.indexOf(user.role) === -1) {
    throw new Error('شما مجوز انجام این عملیات را ندارید.');
  }
}

/** تیم‌های مجاز این کاربر (admin همه تیم‌های پیکربندی‌شده را می‌بیند) */
function allowedTeams_(user) {
  if (user.role === 'admin') return TEAM_CONFIG.map(t => t.key);
  return (user.teams || []).filter(k => TEAM_CONFIG.some(t => t.key === k));
}

function assertTeamAccess_(user, teamKey) {
  if (allowedTeams_(user).indexOf(teamKey) === -1) {
    throw new Error('به این تیم دسترسی ندارید.');
  }
}

/** آیا کاربر می‌تواند محتوا/آزمون تیم را مدیریت کند؟ (ادمین یا مدیر صاحب مجوز) */
function canManageTeam_(user, teamKey) {
  if (user.role === 'admin') return true;
  if (user.role === 'team_manager' && user.canManage && allowedTeams_(user).indexOf(teamKey) !== -1) return true;
  return false;
}

function teamConfigByKey_(key) {
  return TEAM_CONFIG.find(t => t.key === key) || null;
}

/** آیا targetId (پوشه/فایل Drive) داخل teamRootId قرار دارد؟ حداکثر ۳۰ سطح بالاروی */
function isWithinTeam_(targetId, teamRootId) {
  if (String(targetId) === String(teamRootId)) return true;
  let current = targetId;
  for (let i = 0; i < 30; i++) {
    let meta;
    try {
      meta = Drive.Files.get(current, { fields: 'id,parents', supportsAllDrives: true });
    } catch (e) { return false; }
    const parents = meta.parents || [];
    if (parents.indexOf(teamRootId) !== -1) return true;
    if (!parents.length) return false;
    current = parents[0];
    if (String(current) === String(teamRootId)) return true;
  }
  return false;
}

/** نشان حضور تقریبی است و ذخیره دائمی فعالیت محسوب نمی‌شود. */
function presencePing() {
  const user = getUserContext();
  cachePut_('presence:' + user.email, {at: nowIso_()}, 180);
  return ok_({});
}


// ===== Drive.gs =====
/*******************************************************
 * اتصال Google Drive — کتابخانه محتوا
 * فهرست‌برداری صفحه‌ای (بدون بارگذاری کل Drive) + کنترل دسترسی سطح پوشه
 *******************************************************/

function listChildren_(folderId, foldersOnly) {
  const q = ["'" + folderId + "' in parents", 'trashed = false'];
  if (foldersOnly) q.push("mimeType = 'application/vnd.google-apps.folder'");
  const files = [];
  let pageToken = null;
  do {
    const res = Drive.Files.list({
      q: q.join(' and '),
      pageSize: 200,
      pageToken: pageToken || undefined,
      fields: 'nextPageToken,files(id,name,mimeType)',
      supportsAllDrives: true,
      includeItemsFromAllDrives: true
    });
    (res.files || []).forEach(f => files.push(f));
    pageToken = res.nextPageToken;
  } while (pageToken && files.length < 1000);
  return files;
}

/** نگاشت کلید تیم → پوشه Drive (کَش ۱۰ دقیقه‌ای) */
function findTeamFolders_(user) {
  const cacheKey = 'team_folders_v1';
  let folders = cacheGet_(cacheKey);
  if (!folders) {
    try {
      folders = listChildren_(CONFIG.ROOT_FOLDER_ID, true);
    } catch (e) {
      throw new Error('دسترسی به پوشه اصلی Drive ممکن نشد. ROOT_FOLDER_ID و اشتراک‌گذاری پوشه را بررسی کنید. جزئیات: ' + e.message);
    }
    cachePut_(cacheKey, folders, 600);
  }
  return TEAM_CONFIG
    .filter(team => allowedTeams_(user).indexOf(team.key) !== -1)
    .map(team => {
      const folder = folders.find(f =>
        team.aliases.some(alias => String(f.name).toLowerCase() === String(alias).toLowerCase())
      );
      return folder ? {
        id: folder.id, key: team.key, name: team.name,
        color: team.color, icon: team.icon, mimeType: folder.mimeType
      } : null;
    })
    .filter(Boolean);
}

function getTeamFolderOrThrow_(user, teamKey) {
  assertTeamAccess_(user, teamKey);
  const team = findTeamFolders_(user).find(t => t.key === teamKey);
  if (!team) throw new Error('پوشه تیم «' + (teamConfigByKey_(teamKey) || {}).name + '» در Drive پیدا نشد.');
  return team;
}

function getDriveName_(id) {
  const cacheKey = 'fname_' + id;
  const cached = cacheGet_(cacheKey);
  if (cached) return cached;
  const meta = Drive.Files.get(id, { fields: 'name,mimeType', supportsAllDrives: true });
  cachePut_(cacheKey, meta.name, 600);
  return meta.name;
}

/**
 * صفحه‌ای از محتوای یک پوشه برمی‌گرداند.
 * دسترسی: targetId باید داخل پوشه تیم مجاز کاربر باشد.
 */
function getLibraryPage(folderId, pageToken, searchText, typeFilter, teamKey) {
  const user = getUserContext();
  const team = getTeamFolderOrThrow_(user, teamKey);
  const targetId = folderId || team.id;
  if (!isWithinTeam_(targetId, team.id)) {
    throw new Error('دسترسی به این پوشه مجاز نیست.');
  }

  const q = ["'" + targetId + "' in parents", 'trashed = false'];
  if (searchText) q.push("name contains '" + escapeQuery_(searchText) + "'");
  if (typeFilter === 'folder') q.push("mimeType = 'application/vnd.google-apps.folder'");
  if (typeFilter === 'file') q.push("mimeType != 'application/vnd.google-apps.folder'");

  // شمارش سریع نتایج صفحه نه کل نتایج (محدودیت Drive API برای count وجود ندارد)
  const response = Drive.Files.list({
    q: q.join(' and '),
    pageSize: CONFIG.PAGE_SIZE,
    pageToken: pageToken || undefined,
    orderBy: 'folder,name',
    fields: 'nextPageToken,files(id,name,mimeType,modifiedTime,size,webViewLink,thumbnailLink,iconLink,parents,description)',
    supportsAllDrives: true,
    includeItemsFromAllDrives: true
  });

  const myProgress = getMyProgressMap_(user.email);
  const quizMap = getQuizRefMap_();

  const items = (response.files || []).map(f => {
    const isFolder = f.mimeType === 'application/vnd.google-apps.folder';
    const quizId = quizMap[f.id] || '';
    return {
      id: f.id,
      name: f.name,
      mimeType: f.mimeType,
      kind: isFolder ? 'folder' : 'file',
      modified: f.modifiedTime || '',
      size: formatSize_(f.size),
      url: f.webViewLink || ('https://drive.google.com/open?id=' + f.id),
      previewUrl: isFolder ? '' : 'https://drive.google.com/file/d/' + f.id + '/preview',
      thumbnail: f.thumbnailLink || '',
      icon: fileIcon_(f.mimeType),
      quiz: quizId ? { available: true, quizId: quizId } : { available: false },
      status: myProgress['drive:' + f.id] ? 'completed' : 'not_started'
    };
  });

  return ok_({
    items: items,
    nextPageToken: response.nextPageToken || '',
    hasMore: !!response.nextPageToken,
    pageSize: CONFIG.PAGE_SIZE,
    folderId: targetId,
    folderName: getDriveName_(targetId),
    teamId: team.id
  });
}

/** جزئیات یک آیتم + آزمون متصل + وضعیت کاربر */
function getItemDetail(teamKey, itemId) {
  const user = getUserContext();
  const team = getTeamFolderOrThrow_(user, teamKey);
  if (!isWithinTeam_(itemId, team.id)) throw new Error('دسترسی به این آیتم مجاز نیست.');

  const meta = Drive.Files.get(itemId, {
    fields: 'id,name,mimeType,modifiedTime,size,webViewLink,thumbnailLink,description',
    supportsAllDrives: true
  });
  const quizMap = getQuizRefMap_();
  const quizId = quizMap[itemId] || '';
  const progress = getMyProgressMap_(user.email);
  return ok_({
    item: {
      id: meta.id, name: meta.name, mimeType: meta.mimeType,
      kind: meta.mimeType === 'application/vnd.google-apps.folder' ? 'folder' : 'file',
      modified: meta.modifiedTime || '', size: formatSize_(meta.size),
      url: meta.webViewLink || '', thumbnail: meta.thumbnailLink || '',
      description: meta.description || '', icon: fileIcon_(meta.mimeType),
      status: progress['drive:' + meta.id] ? 'completed' : 'not_started'
    },
    quiz: quizId ? getQuizMetaById_(quizId) : { available: false }
  });
}

/** ثبت تکمیل مشاهده یک فایل Drive (جدا از مسیر یادگیری) */
function markDriveItemComplete(teamKey, itemId) {
  const user = getUserContext();
  const team = getTeamFolderOrThrow_(user, teamKey);
  if (!isWithinTeam_(itemId, team.id)) throw new Error('دسترسی به این آیتم مجاز نیست.');
  markProgress_(user.email, 'drive:' + itemId, 100);
  logActivity_(user.email, 'item_complete', itemId);
  return ok_({});
}


// ===== Paths.gs =====
/*******************************************************
 * مسیرهای یادگیری + پیشرفت
 *******************************************************/

function getMyProgressMap_(email) {
  const map = {};
  rows_('Progress').forEach(p => {
    if (String(p.Email).toLowerCase() === email) {
      map[String(p.PathItemId)] = { status: String(p.Status), completedAt: p.CompletedAt, score: Number(p.Score) || 0 };
    }
  });
  return map;
}

function markProgress_(email, itemKey, score) {
  const sh = sheet_('Progress');
  const values = sh.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]).toLowerCase() === email && String(values[i][1]) === String(itemKey)) {
      sh.getRange(i + 1, 3).setValue('completed');
      sh.getRange(i + 1, 4).setValue(nowIso_());
      sh.getRange(i + 1, 5).setValue(score || 100);
      return;
    }
  }
  sh.appendRow([email, itemKey, 'completed', nowIso_(), score || 100]);
}

function pathProgress_(items, progressMap) {
  const total = items.length;
  const done = items.filter(it => progressMap['item:' + it.ItemId]).length;
  return { total: total, done: done, percent: total ? Math.round(done / total * 100) : 0 };
}

/** آیا کاربر در آزمونِ متصل به این آیتم قبول شده؟ */
function hasPassedQuiz_(email, quizId) {
  if (!quizId) return true; // بدون آزمون → همین که بگذراند کافی نیست؛ توسط تیک «مشاهده کردم» ثبت می‌شود
  return rows_('Attempts').some(a =>
    String(a.Email).toLowerCase() === email && String(a.QuizId) === String(quizId) && truthy_(a.Passed));
}

function listPathsForTeams_(teamKeys) {
  return rows_('Paths')
    .filter(p => truthy_(p.Active) && teamKeys.indexOf(String(p.TeamKey)) !== -1)
    .sort((a, b) => Number(a.Order || 0) - Number(b.Order || 0));
}

function itemsOfPath_(pathId) {
  return rows_('PathItems')
    .filter(i => String(i.PathId) === String(pathId) && truthy_(i.Active))
    .sort((a, b) => Number(a.Order || 0) - Number(b.Order || 0));
}

/** مسیرهای قابل مشاهده کاربر به همراه پیشرفت */
function listMyPaths(teamKey) {
  const user = getUserContext();
  const allowed = allowedTeams_(user);
  const keys = teamKey ? allowed.filter(k => k === teamKey) : allowed;
  if (teamKey) assertTeamAccess_(user, teamKey);

  const progress = getMyProgressMap_(user.email);
  const allItems = rows_('PathItems').filter(i => truthy_(i.Active));
  const itemsByPath = {};
  allItems.forEach(i => { (itemsByPath[String(i.PathId)] || (itemsByPath[String(i.PathId)] = [])).push(i); });
  const paths = listPathsForTeams_(keys).map(p => {
    const items = itemsByPath[String(p.PathId)] || [];
    const prog = pathProgress_(items, progress);
    const team = teamConfigByKey_(p.TeamKey);
    return {
      pathId: String(p.PathId), teamKey: String(p.TeamKey), teamName: team ? team.name : p.TeamKey,
      teamColor: team ? team.color : 'blue', teamIcon: team ? team.icon : '◈',
      title: String(p.Title), description: String(p.Description || ''), order: Number(p.Order || 0),
      total: prog.total, done: prog.done, percent: prog.percent
    };
  });
  return ok_({ paths: paths });
}

/** جزئیات یک مسیر + وضعیت هر قدم برای کاربر جاری */
function getPath(pathId) {
  const user = getUserContext();
  const pathRow = rows_('Paths').find(p => String(p.PathId) === String(pathId));
  if (!pathRow || !truthy_(pathRow.Active)) throw new Error('مسیر یادگیری پیدا نشد.');
  assertTeamAccess_(user, String(pathRow.TeamKey));

  const progress = getMyProgressMap_(user.email);
  const itemRows = rows_('PathItems').filter(i => String(i.PathId) === String(pathId) && truthy_(i.Active))
    .sort((a, b) => Number(a.Order || 0) - Number(b.Order || 0));
  const ids = new Set(itemRows.map(i => String(i.QuizId || '')).filter(Boolean));
  const passed = new Set(rows_('Attempts').filter(a => String(a.Email).toLowerCase() === user.email && truthy_(a.Passed)).map(a => String(a.QuizId)));
  const quizzes = {};
  rows_('Quizzes').filter(q => ids.has(String(q.QuizId)) && truthy_(q.Active)).forEach(q => { quizzes[String(q.QuizId)] = q; });
  const counts = {};
  if (ids.size) rows_('Questions').forEach(q => { const id = String(q.QuizId); if (ids.has(id)) counts[id] = (counts[id] || 0) + 1; });

  const resolved = itemRows.map(item => {
    const id = String(item.ItemId);
    const key = 'item:' + id;
    const done = !!progress[key];
    const required = truthy_(item.Required);
    const quizId = String(item.QuizId || '');
    const quizPassed = !quizId || passed.has(quizId);
    const q = quizzes[quizId];
    let status;
    if (done) status = 'completed';
    else if (quizId && !quizPassed) status = 'quiz_pending';
    else status = 'available';
    return {
      itemId: id, order: Number(item.Order || 0), title: String(item.Title),
      driveItemId: String(item.DriveItemId || ''), itemType: String(item.ItemType || 'drive_file'),
      required: required, quizId: quizId,
      quiz: quizId ? (q ? { available: true, quizId: quizId, title: String(q.Title || ''), description: String(q.Description || ''), passScore: Number(q.PassScore) || 70, questionCount: counts[quizId] || 0, timeLimitMin: Number(q.TimeLimitMin) || 0, maxAttempts: Number(q.MaxAttempts) || 0 } : { available: false }) : null,
      estimatedMin: Number(item.EstimatedMin || 0),
      status: status, completedAt: progress[key] ? progress[key].completedAt : ''
    };
  });

  const team = teamConfigByKey_(String(pathRow.TeamKey));
  const prog = pathProgress_(itemRows, progress);
  return ok_({
    path: {
      pathId: String(pathRow.PathId), teamKey: String(pathRow.TeamKey),
      teamName: team ? team.name : String(pathRow.TeamKey),
      title: String(pathRow.Title), description: String(pathRow.Description || ''),
      order: Number(pathRow.Order || 0), total: prog.total, done: prog.done, percent: prog.percent
    },
    items: resolved
  });
}

/** ثبت تکمیل یک قدم از مسیر — اگر آزمون دارد، قبولی لازم است */
function markItemComplete(pathItemId) {
  const user = getUserContext();
  const item = rows_('PathItems').find(i => String(i.ItemId) === String(pathItemId));
  if (!item || !truthy_(item.Active)) throw new Error('آیتم مسیر پیدا نشد.');
  const pathRow = rows_('Paths').find(p => String(p.PathId) === String(item.PathId));
  if (!pathRow || !truthy_(pathRow.Active)) throw new Error('مسیر این آیتم پیدا نشد.');
  assertTeamAccess_(user, String(pathRow.TeamKey));
  const step = getPath(String(pathRow.PathId)).items.find(i => i.itemId === String(pathItemId));
  if (!step) throw new Error('مرحله در این مسیر فعال نیست.');

  const quizId = String(item.QuizId || '');
  if (quizId && !hasPassedQuiz_(user.email, quizId)) {
    throw new Error('برای تکمیل این مرحله ابتدا باید در آزمون متصل قبول شوید.');
  }
  markProgress_(user.email, 'item:' + pathItemId, 100);
  notifyPathCompletion_(user.email, String(pathRow.PathId));
  logActivity_(user.email, 'path_item_complete', String(item.Title));
  return ok_({});
}


// ===== Quizzes.gs =====
/*******************************************************
 * موتور آزمون — ساخت، ارسال، نمره‌دهی، بازبینی
 *******************************************************/

function getQuizRefMap_() {
  // RefType 'drive' → آیتمی از Drive که آزمون به آن متصل است
  const map = {};
  rows_('Quizzes').forEach(q => {
    if (truthy_(q.Active) && /drive|item/i.test(String(q.RefType)) && q.RefId) {
      map[String(q.RefId)] = String(q.QuizId);
    }
  });
  return map;
}

function getQuizRow_(quizId) {
  return rows_('Quizzes').find(q => String(q.QuizId) === String(quizId)) || null;
}

function getQuizMetaById_(quizId) {
  const q = getQuizRow_(quizId);
  if (!q || !truthy_(q.Active)) return { available: false };
  const questions = rows_('Questions').filter(x => String(x.QuizId) === String(quizId));
  return {
    available: true,
    quizId: String(q.QuizId),
    title: String(q.Title),
    description: String(q.Description || ''),
    passScore: Number(q.PassScore) || 70,
    questionCount: questions.length,
    timeLimitMin: Number(q.TimeLimitMin) || 0,
    maxAttempts: Number(q.MaxAttempts) || 0
  };
}

function questionsOf_(quizId) {
  return rows_('Questions')
    .filter(q => String(q.QuizId) === String(quizId))
    .sort((a, b) => Number(a.Order || 0) - Number(b.Order || 0));
}

function parseJsonSafe_(s, fallback) {
  try { const v = JSON.parse(s); return v == null ? fallback : v; } catch (e) { return fallback; }
}

function countMyAttempts_(email, quizId) {
  return rows_('Attempts').filter(a =>
    String(a.Email).toLowerCase() === email && String(a.QuizId) === String(quizId)
      && String(a.Status) !== 'in_progress').length;
}

function shuffle_(arr) {
  const a = arr.slice();
  for (let i = a.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    const tmp = a[i]; a[i] = a[j]; a[j] = tmp;
  }
  return a;
}

/**
 * آزمون را برای شروع برمی‌گرداند (بدون پاسخ صحیح) + ساخت توکن زمان اگر آزمون محدود زمانی است.
 */
// آزمون تنها از مسیر فعال و قابل دسترس یا از یک آیتم Drive مجاز قابل اجراست.
function assertQuizAccess_(user, quiz) {
  if (quiz.TeamKey) assertTeamAccess_(user, String(quiz.TeamKey));
  const paths = rows_('Paths');
  const linked = rows_('PathItems').filter(i => String(i.QuizId || '') === String(quiz.QuizId) && truthy_(i.Active));
  for (const item of linked) {
    const path = paths.find(p => String(p.PathId) === String(item.PathId) && truthy_(p.Active));
    if (!path || allowedTeams_(user).indexOf(String(path.TeamKey)) === -1) continue;
    return;
  }
  if (/^(drive|item)$/i.test(String(quiz.RefType || '')) && quiz.RefId) {
    const teams = findTeamFolders_(user);
    if (teams.some(t => isWithinTeam_(String(quiz.RefId), t.id))) {
      // اتصال به مسیر نباید برای دور زدن پیش‌نیاز همان مسیر استفاده شود.
      if (!linked.length) return;
    }
  }
  throw new Error('به این آزمون دسترسی ندارید.');
}

function getQuizForAttempt(quizId) {
  const user = getUserContext();
  const q = getQuizRow_(quizId);
  if (!q || !truthy_(q.Active)) throw new Error('این آزمون فعال نیست یا پیدا نشد.');
  assertQuizAccess_(user, q);

  const maxAttempts = Number(q.MaxAttempts) || 0;
  const used = countMyAttempts_(user.email, quizId);
  if (maxAttempts > 0 && used >= maxAttempts) {
    throw new Error('تعداد تلاش‌های مجاز شما برای این آزمون به پایان رسیده است.');
  }

  const timeLimitMin = Number(q.TimeLimitMin) || 0;
  // One-use token for every attempt. Cache entries can expire early, so expiry fails closed.
  const attemptToken = Utilities.getUuid();
  const ttl = timeLimitMin > 0 ? Math.min(timeLimitMin * 60 + 60, 21600) : 21600;
  cachePut_('quiz_tok_' + attemptToken, { email: user.email, quizId: String(quizId), deadline: timeLimitMin > 0 ? Date.now() + timeLimitMin * 60000 + 10000 : Date.now() + ttl * 1000 }, ttl);

  let questions = questionsOf_(quizId).map(x => ({
    questionId: String(x.QuestionId),
    order: Number(x.Order || 0),
    type: String(x.Type),
    text: String(x.Text),
    options: parseJsonSafe_(x.OptionsJSON, []),
    points: Number(x.Points) || 1
  }));
  if (truthy_(q.ShuffleQuestions)) questions = shuffle_(questions);

  // تصادفی‌سازی گزینه‌ها: نگاشت ایندکس اصلی حفظ می‌شود
  const shuffleOptions = truthy_(q.ShuffleOptions);
  questions = questions.map(question => {
    if (!shuffleOptions || question.options.length < 2) return question;
    const tagged = question.options.map((text, index) => ({ index: index, text: text }));
    return Object.assign({}, question, { options: shuffle_(tagged) });
  });

  return ok_({
    quiz: {
      quizId: String(q.QuizId),
      title: String(q.Title),
      description: String(q.Description || ''),
      passScore: Number(q.PassScore) || 70,
      timeLimitMin: timeLimitMin,
      maxAttempts: maxAttempts,
      attemptsUsed: used,
      totalPoints: questions.reduce((s, x) => s + x.points, 0),
      shuffleOptions: shuffleOptions
    },
    questions: questions,
    attemptToken: attemptToken
  });
}

/**
 * ثبت و نمره‌دهی آزمون.
 * answers: [{ questionId, value }]
 *   mcq|scenario|tf → value: index (number)
 *   multi → value: array of indexes
 *   essay → value: string
 */
function submitQuiz(quizId, answers, timeSpentSec, attemptToken) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try { return submitQuiz_(quizId, answers, timeSpentSec, attemptToken); }
  finally { lock.releaseLock(); }
}

function submitQuiz_(quizId, answers, timeSpentSec, attemptToken) {
  const user = getUserContext();
  const q = getQuizRow_(quizId);
  if (!q || !truthy_(q.Active)) throw new Error('این آزمون فعال نیست یا پیدا نشد.');
  assertQuizAccess_(user, q);

  const tok = attemptToken ? cacheGet_('quiz_tok_' + attemptToken) : null;
  if (!tok || tok.email !== user.email || String(tok.quizId) !== String(quizId) || Date.now() > tok.deadline) {
    throw new Error('زمان آزمون به پایان رسیده یا نشست آزمون نامعتبر است.');
  }

  const maxAttempts = Number(q.MaxAttempts) || 0;
  const used = countMyAttempts_(user.email, quizId);
  if (maxAttempts > 0 && used >= maxAttempts) {
    throw new Error('تعداد تلاش‌های مجاز شما برای این آزمون به پایان رسیده است.');
  }

  const questions = questionsOf_(quizId);
  if (!questions.length) throw new Error('این آزمون سؤالی ندارد.');

  const totalPoints = questions.reduce((s, x) => s + (Number(x.Points) || 1), 0);
  const autoPointsAll = questions.filter(x => String(x.Type) !== 'essay')
    .reduce((s, x) => s + (Number(x.Points) || 1), 0);
  const hasEssay = questions.some(x => String(x.Type) === 'essay');

  const ansMap = {};
  (answers || []).forEach(a => { if (a && a.questionId) ansMap[String(a.questionId)] = a.value; });

  let earnedAuto = 0;
  const review = [];
  questions.forEach(question => {
    const qid = String(question.QuestionId);
    const type = String(question.Type);
    const points = Number(question.Points) || 1;
    const correct = parseJsonSafe_(question.CorrectJSON, []);
    const given = ansMap[qid];
    const item = { questionId: qid, type: type, points: points, correct: correct, given: (given === undefined ? null : given), isCorrect: null, awarded: 0 };
    if (type === 'essay') {
      item.isCorrect = null; // نیازمند بازبینی
    } else if (type === 'multi') {
      const g = Array.isArray(given) ? given.map(Number).sort() : [];
      const c = (Array.isArray(correct) ? correct : [correct]).map(Number).sort();
      item.isCorrect = g.length === c.length && g.every((v, i) => v === c[i]);
    } else {
      const g = (given === null || given === undefined || given === '') ? -1 : Number(given);
      const c = Array.isArray(correct) ? Number(correct[0]) : Number(correct);
      item.isCorrect = g === c;
    }
    if (item.isCorrect === true) { earnedAuto += points; item.awarded = points; }
    review.push(item);
  });

  const provisionalScore = totalPoints ? Math.round((earnedAuto / totalPoints) * 100) : 0;
  const passScore = Number(q.PassScore) || 70;
  const status = hasEssay ? 'pending_review' : 'auto';
  const passed = !hasEssay && provisionalScore >= passScore;

  const attemptId = newId_('ATT');
  cacheRemove_('quiz_tok_' + attemptToken);
  sheet_('Attempts').appendRow([
    attemptId, nowIso_(), user.email, String(quizId),
    provisionalScore, passed, status,
    JSON.stringify(answers || []), '', Number(timeSpentSec) || 0
  ]);

  // قبولی → تکمیل خودکار آیتم‌های مسیر که به این آزمون متصل‌اند
  if (!hasEssay && passed) {
    rows_('PathItems').filter(it => String(it.QuizId || '') === String(quizId) && truthy_(it.Active)).forEach(it => {
      const path = rows_('Paths').find(p => String(p.PathId) === String(it.PathId) && truthy_(p.Active));
      if (!path || allowedTeams_(user).indexOf(String(path.TeamKey)) === -1) return;
      const step = getPath(String(path.PathId)).items.find(i => i.itemId === String(it.ItemId));
      if (step && step.status !== 'locked') { markProgress_(user.email, 'item:' + String(it.ItemId), provisionalScore); notifyPathCompletion_(user.email, String(path.PathId)); }
    });
  }
  logActivity_(user.email, 'quiz_submit', String(q.Title) + ' — ' + provisionalScore + '٪');

  const showAnswers = String(q.ShowAnswers || 'after_pass');
  const reveal = showAnswers === 'always' || (showAnswers === 'after_pass' && passed);

  return ok_({
    attemptId: attemptId,
    score: provisionalScore,
    passed: passed,
    status: status,
    correctCount: review.filter(r => r.isCorrect === true).length,
    wrongCount: review.filter(r => r.isCorrect === false).length,
    pendingCount: review.filter(r => r.isCorrect === null).length,
    totalQuestions: questions.length,
    passScore: passScore,
    attemptsUsed: used + 1,
    maxAttempts: maxAttempts,
    review: reveal ? review.map(r => ({
      questionId: r.questionId, given: r.given, correct: r.correct,
      isCorrect: r.isCorrect, awarded: r.awarded, points: r.points
    })) : null
  });
}

/** کارنامه کاربر جاری */
function getMyTranscript() {
  const user = getUserContext();
  const quizzes = rows_('Quizzes');
  const quizMap = {};
  quizzes.forEach(q => { quizMap[String(q.QuizId)] = q; });
  const mine = rows_('Attempts')
    .filter(a => String(a.Email).toLowerCase() === user.email && String(a.Status) !== 'in_progress')
    .sort((a, b) => new Date(b.Timestamp) - new Date(a.Timestamp));
  const attempts = mine.map(a => {
    const q = quizMap[String(a.QuizId)] || {};
    return {
      attemptId: String(a.AttemptId), quizId: String(a.QuizId),
      quizTitle: String(q.Title || 'آزمون حذف‌شده'),
      score: Number(a.Score) || 0, passed: truthy_(a.Passed),
      status: String(a.Status), timestamp: a.Timestamp,
      timeSpentSec: Number(a.TimeSpentSec) || 0, feedback: String(a.Feedback || '')
    };
  });
  const best = {};
  attempts.forEach(a => {
    if (!best[a.quizId] || a.score > best[a.quizId].score) best[a.quizId] = a;
  });
  const summary = {
    totalAttempts: attempts.length,
    passedQuizzes: Object.keys(best).filter(k => best[k].passed).length,
    uniqueQuizzes: Object.keys(best).length,
    avgScore: attempts.length ? Math.round(attempts.reduce((s, a) => s + a.score, 0) / attempts.length) : 0
  };
  return ok_({ attempts: attempts, summary: summary });
}


// ===== Reports.gs =====
/*******************************************************
 * گزارش‌ها، خروجی CSV (قابل باز شدن در Excel) و یادداشت منتور
 *******************************************************/

function activeUsers_() {
  return rows_('Users').filter(u => truthy_(u.Active)).map(u => ({
    email: String(u.Email).toLowerCase(),
    name: String(u.Name || u.Email),
    role: normalizeRole_(u.Role),
    teams: splitCsv_(u.Teams),
    mentorEmail: String(u.MentorEmail || '').toLowerCase(),
    canManage: truthy_(u.CanManage)
  }));
}

function pathItemsByTeam_(teamKey) {
  const pathIds = rows_('Paths')
    .filter(p => String(p.TeamKey) === teamKey && truthy_(p.Active))
    .map(p => String(p.PathId));
  return rows_('PathItems').filter(i => pathIds.indexOf(String(i.PathId)) !== -1 && truthy_(i.Active));
}

function userCompletionPercent_(userEmail, teams) {
  let total = 0, done = 0;
  const progress = getMyProgressMap_(userEmail);
  teams.forEach(teamKey => {
    const items = pathItemsByTeam_(teamKey);
    total += items.length;
    done += items.filter(i => progress['item:' + String(i.ItemId)]).length;
  });
  return { total: total, done: done, percent: total ? Math.round(done / total * 100) : 0 };
}

// Read each sheet once per report instead of once per member.
function completionSnapshot_() {
  const pathTeams = {};
  rows_('Paths').filter(p => truthy_(p.Active)).forEach(p => { pathTeams[String(p.PathId)] = String(p.TeamKey); });
  const items = {};
  rows_('PathItems').filter(i => truthy_(i.Active)).forEach(i => {
    const team = pathTeams[String(i.PathId)];
    if (team) (items[team] || (items[team] = [])).push('item:' + String(i.ItemId));
  });
  const done = {};
  rows_('Progress').forEach(p => {
    const email = String(p.Email).toLowerCase();
    (done[email] || (done[email] = {}))[String(p.PathItemId)] = true;
  });
  return (email, teams) => {
    let total = 0, count = 0;
    teams.forEach(t => (items[t] || []).forEach(id => { total++; if ((done[email] || {})[id]) count++; }));
    return {total: total, done: count, percent: total ? Math.round(count * 100 / total) : 0};
  };
}

/** گزارش کلی — ادمین: همه تیم‌ها (یا یک تیم)، مدیر تیم: تیم‌های خودش */
function reportSummary(teamKey) {
  const user = getUserContext();
  const completion = completionSnapshot_();
  requireRole_(user, ['admin', 'team_manager']);
  const allowed = allowedTeams_(user);
  let keys = allowed;
  if (teamKey) { assertTeamAccess_(user, teamKey); keys = [teamKey]; }

  const users = activeUsers_();
  const attempts = rows_('Attempts').filter(a => String(a.Status) !== 'in_progress');

  const teams = keys.map(key => {
    const team = teamConfigByKey_(key);
    const members = users.filter(u => u.teams.indexOf(key) !== -1);
    const memberEmails = members.map(m => m.email);
    const teamAttempts = attempts.filter(a => memberEmails.indexOf(String(a.Email).toLowerCase()) !== -1);
    const scored = teamAttempts.filter(a => a.Score !== '' && a.Score != null);
    const avgScore = scored.length ? Math.round(scored.reduce((s, a) => s + Number(a.Score || 0), 0) / scored.length) : 0;
    const pendingReview = teamAttempts.filter(a => String(a.Status) === 'pending_review').length;

    let sumPercent = 0, completedMembers = 0;
    members.forEach(m => {
      const c = completion(m.email, [key]);
      sumPercent += c.percent;
      if (c.percent >= 100 && c.total > 0) completedMembers++;
    });

    return {
      key: key, name: team ? team.name : key, icon: team ? team.icon : '◈', color: team ? team.color : 'blue',
      members: members.length,
      avgCompletion: members.length ? Math.round(sumPercent / members.length) : 0,
      completedMembers: completedMembers,
      attempts: teamAttempts.length,
      avgScore: avgScore,
      passedAttempts: teamAttempts.filter(a => truthy_(a.Passed)).length,
      pendingReview: pendingReview
    };
  });

  const totals = {
    members: teams.reduce((s, t) => s + t.members, 0),
    attempts: teams.reduce((s, t) => s + t.attempts, 0),
    pendingReview: teams.reduce((s, t) => s + t.pendingReview, 0),
    avgScore: teams.length && teams.some(t => t.attempts)
      ? Math.round(teams.reduce((s, t) => s + t.avgScore * t.attempts, 0) / Math.max(1, teams.reduce((s, t) => s + t.attempts, 0)))
      : 0
  };
  return ok_({ teams: teams, totals: totals });
}

/** ریزداده اعضای یک تیم — ادمین/مدیر همان تیم/منتور (فقط تیمن‌های تحت پوشش خودش) */
function reportTeamMembers(teamKey) {
  const user = getUserContext();
  const completion = completionSnapshot_();
  assertTeamAccess_(user, teamKey);
  const menteeMode = user.role === 'mentor';
  if (!menteeMode) requireRole_(user, ['admin', 'team_manager']);

  const users = activeUsers_().filter(u => u.teams.indexOf(teamKey) !== -1);
  const attempts = rows_('Attempts').filter(a => String(a.Status) !== 'in_progress');

  const members = users
    .filter(u => !menteeMode || u.mentorEmail === user.email)
    .map(u => {
      const mine = attempts.filter(a => String(a.Email).toLowerCase() === u.email);
      const scored = mine.filter(a => a.Score !== '' && a.Score != null);
      const c = completion(u.email, [teamKey]);
      let last = '';
      mine.forEach(a => { const t = new Date(a.Timestamp).getTime(); if (!last || t > new Date(last).getTime()) last = a.Timestamp; });
      return {
        email: u.email, name: u.name, role: u.role,
        completion: c.percent, itemsDone: c.done, itemsTotal: c.total,
        attempts: mine.length,
        avgScore: scored.length ? Math.round(scored.reduce((s, a) => s + Number(a.Score || 0), 0) / scored.length) : 0,
        passed: mine.filter(a => truthy_(a.Passed)).length,
        pendingReview: mine.filter(a => String(a.Status) === 'pending_review').length,
        lastActivity: last
      };
    })
    .sort((a, b) => b.completion - a.completion);

  return ok_({ members: members });
}

/** کارنامه کامل یک نفر — برای ادمین/مدیرِ تیمش/منتورِ او/خودِ شخص */
function reportMember(email) {
  const user = getUserContext();
  const target = String(email || '').toLowerCase().trim();
  const isSelf = target === user.email;
  const users = activeUsers_();
  const targetUser = users.find(u => u.email === target)
    || { email: target, name: target, role: 'learner', teams: [], mentorEmail: '' };

  if (!isSelf) {
    if (user.role === 'admin') { /* ok */ }
    else if (user.role === 'team_manager' && targetUser.teams.some(t => allowedTeams_(user).indexOf(t) !== -1)) { /* ok */ }
    else if (user.role === 'mentor' && targetUser.mentorEmail === user.email) { /* ok */ }
    else throw new Error('شما مجوز مشاهده کارنامه این کاربر را ندارید.');
  }

  const progressMap = getMyProgressMap_(target);
  const quizMap = {};
  rows_('Quizzes').forEach(q => { quizMap[String(q.QuizId)] = q; });

  const paths = rows_('Paths').filter(p => truthy_(p.Active) && targetUser.teams.indexOf(String(p.TeamKey)) !== -1)
    .map(p => {
      const items = itemsOfPath_(p.PathId);
      const done = items.filter(i => progressMap['item:' + String(i.ItemId)]).length;
      return {
        pathId: String(p.PathId), title: String(p.Title), teamKey: String(p.TeamKey),
        total: items.length, done: done, percent: items.length ? Math.round(done / items.length * 100) : 0
      };
    });

  const attempts = rows_('Attempts')
    .filter(a => String(a.Email).toLowerCase() === target && String(a.Status) !== 'in_progress')
    .sort((a, b) => new Date(b.Timestamp) - new Date(a.Timestamp))
    .map(a => ({
      attemptId: String(a.AttemptId), quizId: String(a.QuizId),
      quizTitle: String((quizMap[String(a.QuizId)] || {}).Title || 'آزمون حذف‌شده'),
      score: Number(a.Score) || 0, passed: truthy_(a.Passed), status: String(a.Status),
      timestamp: a.Timestamp, timeSpentSec: Number(a.TimeSpentSec) || 0, feedback: String(a.Feedback || '')
    }));

  const notes = rows_('MentorNotes')
    .filter(n => String(n.MenteeEmail).toLowerCase() === target)
    .sort((a, b) => new Date(b.Timestamp) - new Date(a.Timestamp))
    .map(n => ({ id: String(n.Id), mentorEmail: String(n.MentorEmail), timestamp: n.Timestamp, note: String(n.Note) }));

  return ok_({
    member: { email: target, name: targetUser.name, role: targetUser.role, teams: targetUser.teams, mentorEmail: targetUser.mentorEmail },
    paths: paths, attempts: attempts, notes: notes
  });
}

/** آموزش‌های تکمیل‌نشده در تیم‌های مجاز */
function reportIncomplete(teamKey) {
  const user = getUserContext();
  const completion = completionSnapshot_();
  requireRole_(user, ['admin', 'team_manager']);
  const allowed = allowedTeams_(user);
  let keys = allowed;
  if (teamKey) { assertTeamAccess_(user, teamKey); keys = [teamKey]; }

  const users = activeUsers_();
  const rowsOut = [];
  keys.forEach(key => {
    const members = users.filter(u => u.teams.indexOf(key) !== -1);
    members.forEach(m => {
      const c = completion(m.email, [key]);
      if (c.total > 0 && c.percent < 100) {
        rowsOut.push({
          teamKey: key, teamName: (teamConfigByKey_(key) || {}).name || key,
          email: m.email, name: m.name,
          completion: c.percent, remaining: c.total - c.done
        });
      }
    });
  });
  rowsOut.sort((a, b) => a.completion - b.completion);
  return ok_({ rows: rowsOut });
}

/** خروجی CSV (با BOM برای Excel فارسی) — base64 برمی‌گرداند */
function reportExportCsv(kind, teamKey) {
  const user = getUserContext();
  const completion = completionSnapshot_();
  requireRole_(user, ['admin', 'team_manager']);
  if (teamKey) assertTeamAccess_(user, teamKey);

  let header = [], lines = [], filename = 'report.csv';

  if (kind === 'team_summary') {
    const data = reportSummary(teamKey || '').teams;
    header = ['تیم', 'تعداد اعضا', 'میانگین تکمیل (%)', 'اعضای تکمیل‌شده', 'تعداد تلاش آزمون', 'میانگین نمره', 'تلاش‌های قبول', 'در انتظار بازبینی'];
    lines = data.map(t => [t.name, t.members, t.avgCompletion, t.completedMembers, t.attempts, t.avgScore, t.passedAttempts, t.pendingReview]);
    filename = 'bimebazar-team-summary.csv';
  } else if (kind === 'user_progress' || kind === 'incomplete') {
    const keys = teamKey ? [teamKey] : allowedTeams_(user);
    const users = activeUsers_();
    header = ['تیم', 'نام', 'ایمیل', 'نقش', 'تکمیل (%)', 'آیتم انجام‌شده', 'کل آیتم‌ها'];
    keys.forEach(key => {
      users.filter(u => u.teams.indexOf(key) !== -1).forEach(u => {
        const c = completion(u.email, [key]);
        if (kind === 'incomplete' && !(c.total > 0 && c.percent < 100)) return;
        lines.push([(teamConfigByKey_(key) || {}).name || key, u.name, u.email, roleLabel_(u.role), c.percent, c.done, c.total]);
      });
    });
    filename = kind === 'incomplete' ? 'bimebazar-incomplete.csv' : 'bimebazar-user-progress.csv';
  } else if (kind === 'attempts') {
    const allowed = allowedTeams_(user);
    const keys = teamKey ? [teamKey] : allowed;
    const users = activeUsers_().filter(u => u.teams.some(t => keys.indexOf(t) !== -1));
    const emails = users.map(u => u.email);
    const quizMap = {};
    rows_('Quizzes').forEach(q => { quizMap[String(q.QuizId)] = q; });
    header = ['تاریخ', 'نام', 'ایمیل', 'آزمون', 'نمره', 'وضعیت', 'نتیجه', 'زمان (ثانیه)'];
    rows_('Attempts').filter(a => emails.indexOf(String(a.Email).toLowerCase()) !== -1).forEach(a => {
      const u = users.find(x => x.email === String(a.Email).toLowerCase());
      lines.push([
        Utilities.formatDate(new Date(a.Timestamp), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm'),
        u ? u.name : '', a.Email, (quizMap[String(a.QuizId)] || {}).Title || '',
        a.Score, a.Status, truthy_(a.Passed) ? 'قبول' : 'رد', a.TimeSpentSec || 0
      ]);
    });
    filename = 'bimebazar-attempts.csv';
  } else {
    throw new Error('نوع خروجی نامعتبر است.');
  }

  const esc = v => {
    let s = String(v == null ? '' : v);
    if (/^[\s]*[=+@-]/.test(s)) s = "'" + s;
    return /[",\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
  };
  const csv = '﻿' + [header].concat(lines).map(r => r.map(esc).join(',')).join('\r\n');
  return ok_({ filename: filename, csvBase64: Utilities.base64Encode(csv, Utilities.Charset.UTF_8) });
}

/** ثبت یادداشت منتور روی کارنامه یک نفر تحت پوشش */
function addMentorNote(menteeEmail, note) {
  const user = getUserContext();
  const target = String(menteeEmail || '').toLowerCase().trim();
  if (!note || !String(note).trim()) throw new Error('متن بازخورد خالی است.');
  const targetUser = activeUsers_().find(u => u.email === target);
  const isAdminOrManager = user.role === 'admin' ||
    (user.role === 'team_manager' && targetUser && targetUser.teams.some(t => allowedTeams_(user).indexOf(t) !== -1));
  const isMentor = targetUser && targetUser.mentorEmail === user.email;
  if (!isAdminOrManager && !isMentor) throw new Error('شما منتور این کاربر نیستید.');
  sheet_('MentorNotes').appendRow([newId_('NOTE'), nowIso_(), user.email, target, String(note).trim()]);
  logActivity_(user.email, 'mentor_note', target);
  return ok_({});
}


// ===== Admin.gs =====
/*******************************************************
 * پنل مدیریت — کاربران، محتوا، آزمون‌ها، بازبینی، اعلانات
 * همه توابع مجوز را سرور-ساید بررسی می‌کنند.
 *******************************************************/

/* ---------- کاربران ---------- */

function adminListUsers() {
  const user = getUserContext();
  requireRole_(user, ['admin']);
  const userRows = rows_('Users');
  const lastLogin = {};
  rows_('Activity').filter(a => String(a.Action) === 'login').forEach(a => { const key = String(a.Email).toLowerCase(); if (!lastLogin[key] || new Date(a.Timestamp) > new Date(lastLogin[key])) lastLogin[key] = a.Timestamp; });
  const presence = {};
  try { const cache = CacheService.getScriptCache(); const keys = userRows.map(u => 'presence:' + String(u.Email).toLowerCase()); for (let i = 0; i < keys.length; i += 100) Object.assign(presence, cache.getAll(keys.slice(i, i + 100))); } catch (_) {}
  const users = userRows.map(u => ({
    email: String(u.Email), name: String(u.Name || ''), role: normalizeRole_(u.Role),
    teams: splitCsv_(u.Teams), mentorEmail: String(u.MentorEmail || ''),
    canManage: truthy_(u.CanManage), active: truthy_(u.Active), createdAt: u.CreatedAt,
    lastLogin: lastLogin[String(u.Email).toLowerCase()] || '', online: !!presence['presence:' + String(u.Email).toLowerCase()]
  }));
  return ok_({ users: users, teams: TEAM_CONFIG.map(t => ({ key: t.key, name: t.name })) });
}

function adminUpsertUser(payload) {
  const user = getUserContext();
  requireRole_(user, ['admin']);
  if (!payload || !payload.email) throw new Error('ایمیل کاربر الزامی است.');
  const email = String(payload.email).toLowerCase().trim();
  if (!/^[a-z0-9._@-]{3,100}$/.test(email)) throw new Error('نام کاربری نامعتبر است.');
  if (payload.password && (typeof payload.password !== 'string' || payload.password.length < 12 || payload.password.length > 128)) throw new Error('رمز باید ۱۲ تا ۱۲۸ نویسه باشد.');
  if (['admin', 'mentor', 'learner'].indexOf(String(payload.role)) === -1) throw new Error('نقش باید مدیر، منتور یا کارشناس باشد.');
  const role = normalizeRole_(payload.role);
  if (email === user.email && role !== 'admin') throw new Error('نمی‌توانید نقش خودتان را تغییر دهید.');
  const teams = (payload.teams || []).filter(k => TEAM_CONFIG.some(t => t.key === k));
  if (role === 'mentor' && (!teams.length || teams.length !== (payload.teams || []).length)) throw new Error('برای منتور حداقل یک تیم معتبر انتخاب کنید.');
  if (role === 'learner' && (teams.length !== 1 || (payload.teams || []).length !== 1)) throw new Error('برای هر کارشناس دقیقاً یک تیم معتبر انتخاب کنید.');

  const mentorName = String(payload.mentorEmail || '').toLowerCase().trim();
  if (mentorName) {
    const mentor = rows_('Users').find(u => String(u.Email).toLowerCase().trim() === mentorName && String(u.Role).toLowerCase() === 'mentor' && truthy_(u.Active));
    if (role !== 'learner' || !mentor || !splitCsv_(mentor.Teams).some(t => teams.indexOf(t) !== -1)) throw new Error('منتور فعال و هم‌تیم انتخاب کنید.');
  }
  const sh = sheet_('Users');
  const values = sh.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]).toLowerCase().trim() === email) {
      sh.getRange(i + 1, 2).setValue(String(payload.name || ''));
      sh.getRange(i + 1, 3).setValue(role);
      sh.getRange(i + 1, 4).setValue(teams.join(','));
      sh.getRange(i + 1, 5).setValue(String(payload.mentorEmail || '').toLowerCase().trim());
      sh.getRange(i + 1, 6).setValue(!!payload.canManage);
      if (payload.active !== undefined) sh.getRange(i + 1, 7).setValue(payload.active !== false);
      if (payload.password) savePassword_(email, payload.password);
      logActivity_(user.email, 'user_update', email);
      return ok_({ saved: email });
    }
  }
  if (!payload.password) throw new Error('برای کاربر جدید رمز عبور الزامی است.');
  savePassword_(email, payload.password);
  sh.appendRow([email, String(payload.name || ''), role, teams.join(','),
    String(payload.mentorEmail || '').toLowerCase().trim(), !!payload.canManage, true, nowIso_()]);
  logActivity_(user.email, 'user_create', email);
  return ok_({ saved: email, created: true });
}

function adminSetUserActive(email, active) {
  const user = getUserContext();
  requireRole_(user, ['admin']);
  const target = String(email || '').toLowerCase().trim();
  if (target === user.email && !active) throw new Error('نمی‌توانید حساب خودتان را غیرفعال کنید.');
  const sh = sheet_('Users');
  const values = sh.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]).toLowerCase().trim() === target) {
      sh.getRange(i + 1, 7).setValue(!!active);
      logActivity_(user.email, active ? 'user_activate' : 'user_deactivate', target);
      return ok_({});
    }
  }
  throw new Error('کاربر پیدا نشد.');
}

/** وضعیت پوشه هر تیم در Drive (برای تصمیم‌گیری ادمین) */
function adminTeamsStatus() {
  const user = getUserContext();
  requireRole_(user, ['admin']);
  let rootName = '';
  let folders = [];
  let error = '';
  try {
    rootName = Drive.Files.get(CONFIG.ROOT_FOLDER_ID, { fields: 'name', supportsAllDrives: true }).name;
    folders = listChildren_(CONFIG.ROOT_FOLDER_ID, true);
  } catch (e) { error = e.message; }
  return ok_({
    rootName: rootName, error: error,
    teams: TEAM_CONFIG.map(t => {
      const f = folders.find(x => t.aliases.some(a => String(x.name).toLowerCase() === String(a).toLowerCase()));
      return { key: t.key, name: t.name, found: !!f, folderId: f ? f.id : '' };
    })
  });
}

/* ---------- محتوا (مسیرها و آیتم‌ها) ---------- */

function manageScope_(user, teamKey) {
  if (!canManageTeam_(user, teamKey)) throw new Error('برای مدیریت محتوای این تیم مجوز ندارید.');
}

function adminListPaths(teamKey) {
  const user = getUserContext();
  const allowed = allowedTeams_(user).filter(k => canManageTeam_(user, k));
  const keys = teamKey ? allowed.filter(k => k === teamKey) : allowed;
  const items = rows_('PathItems');
  const paths = rows_('Paths')
    .filter(p => keys.indexOf(String(p.TeamKey)) !== -1)
    .sort((a, b) => String(a.TeamKey).localeCompare(String(b.TeamKey)) || Number(a.Order || 0) - Number(b.Order || 0))
    .map(p => ({
      pathId: String(p.PathId), teamKey: String(p.TeamKey), title: String(p.Title),
      description: String(p.Description || ''), order: Number(p.Order || 0), active: truthy_(p.Active),
      itemCount: items.filter(i => String(i.PathId) === String(p.PathId) && truthy_(i.Active)).length
    }));
  const quizzes = adminListQuizzes().quizzes;
  return ok_({ paths: paths, teams: TEAM_CONFIG.filter(t => allowed.indexOf(t.key) !== -1), quizzes: quizzes });
}

function adminSavePath(payload) {
  const user = getUserContext();
  const teamKey = String(payload.teamKey || '');
  manageScope_(user, teamKey);
  if (!payload.title || !String(payload.title).trim()) throw new Error('عنوان مسیر الزامی است.');

  const sh = sheet_('Paths');
  const id = String(payload.pathId || '').trim();
  if (id) {
    const values = sh.getDataRange().getValues();
    for (let i = 1; i < values.length; i++) {
      if (String(values[i][0]) === id) {
        manageScope_(user, String(values[i][1]));
        sh.getRange(i + 1, 3).setValue(String(payload.title).trim());
        sh.getRange(i + 1, 4).setValue(String(payload.description || '').trim());
        sh.getRange(i + 1, 5).setValue(Number(payload.order) || 0);
        sh.getRange(i + 1, 6).setValue(payload.active !== false);
        return ok_({ pathId: id });
      }
    }
    throw new Error('مسیر پیدا نشد.');
  }
  const pathId = newId_('PATH');
  sh.appendRow([pathId, teamKey, String(payload.title).trim(), String(payload.description || '').trim(), Number(payload.order) || 0, payload.active !== false, nowIso_()]);
  logActivity_(user.email, 'path_create', payload.title);
  return ok_({ pathId: pathId, created: true });
}

function adminDeletePath(pathId) {
  const user = getUserContext();
  const sh = sheet_('Paths');
  const values = sh.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(pathId)) {
      manageScope_(user, String(values[i][1]));
      sh.getRange(i + 1, 6).setValue(false);
      logActivity_(user.email, 'path_deactivate', String(pathId));
      return ok_({});
    }
  }
  throw new Error('مسیر پیدا نشد.');
}

function adminSavePathItem(payload) {
  const user = getUserContext();
  const pathId = String(payload.pathId || '');
  const pathRow = rows_('Paths').find(p => String(p.PathId) === pathId);
  if (!pathRow) throw new Error('مسیر پیدا نشد.');
  manageScope_(user, String(pathRow.TeamKey));
  if (!payload.title || !String(payload.title).trim()) throw new Error('عنوان آیتم الزامی است.');

  const driveId = parseDriveId_(payload.driveItemId);
  const quizId = String(payload.quizId || '').trim();
  if (quizId) {
    const quiz = getQuizRow_(quizId);
    if (!quiz || !truthy_(quiz.Active) || (quiz.TeamKey && String(quiz.TeamKey) !== String(pathRow.TeamKey))) throw new Error('آزمون فعال این تیم نیست.');
    const linkedPaths = rows_('PathItems').filter(i => String(i.QuizId || '') === quizId).map(i => String(i.PathId));
    if (rows_('Paths').some(p => linkedPaths.indexOf(String(p.PathId)) !== -1 && String(p.TeamKey) !== String(pathRow.TeamKey))) throw new Error('این آزمون به تیم دیگری متصل است.');
  }
  if (driveId) {
    const team = getTeamFolderOrThrow_(user, String(pathRow.TeamKey));
    if (!isWithinTeam_(driveId, team.id)) throw new Error('فایل باید داخل پوشه همین تیم باشد.');
  }

  const sh = sheet_('PathItems');
  const id = String(payload.itemId || '').trim();
  const row = [
    id || newId_('ITEM'), pathId, Number(payload.order) || 0,
    String(payload.title).trim(), driveId,
    String(payload.itemType || (driveId ? 'drive_file' : 'link')),
    payload.required !== false, quizId, Number(payload.estimatedMin) || 0,
    payload.active !== false
  ];
  if (id) {
    const values = sh.getDataRange().getValues();
    for (let i = 1; i < values.length; i++) {
      if (String(values[i][0]) === id) {
        sh.getRange(i + 1, 1, 1, row.length).setValues([row]);
        logActivity_(user.email, 'path_item_update', row[3]);
        return ok_({ itemId: id });
      }
    }
    throw new Error('آیتم پیدا نشد.');
  }
  sh.appendRow(row);
  logActivity_(user.email, 'path_item_create', row[3]);
  return ok_({ itemId: row[0], created: true });
}

function adminDeletePathItem(itemId) {
  const user = getUserContext();
  const sh = sheet_('PathItems');
  const values = sh.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(itemId)) {
      const pathRow = rows_('Paths').find(p => String(p.PathId) === String(values[i][1]));
      if (pathRow) manageScope_(user, String(pathRow.TeamKey));
      sh.getRange(i + 1, 10).setValue(false);
      return ok_({});
    }
  }
  throw new Error('آیتم پیدا نشد.');
}

function adminPathItems(pathId) {
  const user = getUserContext();
  const pathRow = rows_('Paths').find(p => String(p.PathId) === String(pathId));
  if (!pathRow) throw new Error('مسیر پیدا نشد.');
  manageScope_(user, String(pathRow.TeamKey));
  const items = rows_('PathItems').filter(i => String(i.PathId) === String(pathId)).sort((a,b) => Number(a.Order || 0) - Number(b.Order || 0)).map(i => ({
    itemId: String(i.ItemId), pathId: String(i.PathId), order: Number(i.Order || 0),
    title: String(i.Title), driveItemId: String(i.DriveItemId || ''), itemType: String(i.ItemType),
    required: truthy_(i.Required), quizId: String(i.QuizId || ''),
    estimatedMin: Number(i.EstimatedMin || 0), active: truthy_(i.Active)
  }));
  return ok_({ items: items, path: { pathId: String(pathRow.PathId), title: String(pathRow.Title), teamKey: String(pathRow.TeamKey) } });
}

/** استخراج شناسه Drive از ورودی (آیدی خام یا لینک) */
function parseDriveId_(input) {
  const s = String(input || '').trim();
  if (!s) return '';
  const m1 = s.match(/\/d\/([a-zA-Z0-9_-]{10,})/);
  if (m1) return m1[1];
  const m2 = s.match(/[?&]id=([a-zA-Z0-9_-]{10,})/);
  if (m2) return m2[1];
  if (/^[a-zA-Z0-9_-]{10,}$/.test(s)) return s;
  return s; // اگر الگو نخورد همان را نگه دار؛ نام فایلی که تایپ شده باشد به درد نمی‌خورد ولی خطا هم نمی‌دهیم
}

/** متادیتای فایل Drive برای پیش‌نمایش/اتصال (بدون دانلود محتوا) */
function adminDriveMeta(driveIdOrUrl) {
  const user = getUserContext();
  if (user.role === 'learner' || user.role === 'mentor') throw new Error('مجوز ندارید.');
  const id = parseDriveId_(driveIdOrUrl);
  if (!id) throw new Error('شناسه یا لینک Drive نامعتبر است.');
  if (user.role !== 'admin' && !findTeamFolders_(user).some(t => canManageTeam_(user, t.key) && isWithinTeam_(id, t.id))) throw new Error('دسترسی به این فایل مجاز نیست.');
  try {
    const meta = Drive.Files.get(id, {
      fields: 'id,name,mimeType,size,modifiedTime,webViewLink',
      supportsAllDrives: true
    });
    return ok_({ id: meta.id, name: meta.name, mimeType: meta.mimeType, size: formatSize_(meta.size), modified: meta.modifiedTime || '', url: meta.webViewLink || '', icon: fileIcon_(meta.mimeType) });
  } catch (e) {
    throw new Error('فایل در Drive پیدا نشد یا شما به آن دسترسی ندارید: ' + e.message);
  }
}

/* ---------- آزمون‌ها ---------- */

/** فهرست آزمون‌های قابل مدیریت برای این کاربر */
function adminListQuizzes() {
  const user = getUserContext();
  requireRole_(user, ['admin', 'team_manager']);
  const allowed = allowedTeams_(user);

  // نگاشت آزمون ← تیم (از طریق آیتم‌های مسیر متصل به آزمون)
  const pathTeam = {};
  rows_('Paths').forEach(p => { pathTeam[String(p.PathId)] = String(p.TeamKey); });
  const linkedQuizTeams = {};
  rows_('PathItems').forEach(i => { const id = String(i.QuizId || ''), team = pathTeam[String(i.PathId)]; if (id && team) (linkedQuizTeams[id] || (linkedQuizTeams[id] = new Set())).add(team); });
  const allowedQuizzes = new Set();
  if (user.role === 'admin') {
    rows_('Quizzes').forEach(q => allowedQuizzes.add(String(q.QuizId)));
  } else {
    rows_('PathItems').forEach(i => {
      const team = pathTeam[String(i.PathId)];
      if (team && canManageTeam_(user, team) && i.QuizId) allowedQuizzes.add(String(i.QuizId));
    });
  }

  const counts = {};
  rows_('Questions').forEach(q => { counts[String(q.QuizId)] = (counts[String(q.QuizId)] || 0) + 1; });
  const quizzes = rows_('Quizzes')
    .filter(q => allowedQuizzes.has(String(q.QuizId)))
    .sort((a, b) => new Date(b.CreatedAt || 0) - new Date(a.CreatedAt || 0))
    .map(q => ({
      quizId: String(q.QuizId), teamKey: String(q.TeamKey || ((linkedQuizTeams[String(q.QuizId)] && linkedQuizTeams[String(q.QuizId)].size === 1) ? Array.from(linkedQuizTeams[String(q.QuizId)])[0] : '')), refType: String(q.RefType || 'standalone'), refId: String(q.RefId || ''),
      title: String(q.Title), description: String(q.Description || ''),
      passScore: Number(q.PassScore) || 70, timeLimitMin: Number(q.TimeLimitMin) || 0,
      maxAttempts: Number(q.MaxAttempts) || 0, shuffleQuestions: truthy_(q.ShuffleQuestions),
      shuffleOptions: truthy_(q.ShuffleOptions), showAnswers: String(q.ShowAnswers || 'after_pass'),
      active: truthy_(q.Active), questionCount: counts[String(q.QuizId)] || 0, createdAt: q.CreatedAt
    }));
  return ok_({ quizzes: quizzes });
}

function adminGetQuiz(quizId) {
  const user = getUserContext();
  requireRole_(user, ['admin', 'team_manager']);
  const q = getQuizRow_(quizId);
  if (!q) throw new Error('آزمون پیدا نشد.');
  assertQuizManageScope_(user, q);
  const questions = questionsOf_(quizId).map(x => ({
    questionId: String(x.QuestionId), order: Number(x.Order || 0), type: String(x.Type),
    text: String(x.Text), options: parseJsonSafe_(x.OptionsJSON, []),
    correct: parseJsonSafe_(x.CorrectJSON, []), points: Number(x.Points) || 1
  }));
  return ok_({
    quiz: {
      quizId: String(q.QuizId), teamKey: String(q.TeamKey || ''), refType: String(q.RefType || 'standalone'), refId: String(q.RefId || ''),
      title: String(q.Title), description: String(q.Description || ''),
      passScore: Number(q.PassScore) || 70, timeLimitMin: Number(q.TimeLimitMin) || 0,
      maxAttempts: Number(q.MaxAttempts) || 0, shuffleQuestions: truthy_(q.ShuffleQuestions),
      shuffleOptions: truthy_(q.ShuffleOptions), showAnswers: String(q.ShowAnswers || 'after_pass'),
      active: truthy_(q.Active)
    },
    questions: questions
  });
}

function assertQuizManageScope_(user, quizRow) {
  if (user.role === 'admin') return;
  const refType = String(quizRow.RefType || 'standalone');
  const refId = String(quizRow.RefId || '');
  // آزمون متصل به آیتم/درایو: از طریق آیتم‌های مسیرِ قابل مدیریت بررسی کن
  const linkedTeams = [];
  if (refId) {
    const pathTeam = {};
    rows_('Paths').forEach(p => { pathTeam[String(p.PathId)] = String(p.TeamKey); });
    rows_('PathItems').forEach(i => {
      const sameQuiz = String(i.QuizId || '') === String(quizRow.QuizId);
      const sameRef = refId && String(i.DriveItemId || '') === refId;
      if (sameQuiz || sameRef) {
        const team = pathTeam[String(i.PathId)];
        if (team) linkedTeams.push(team);
      }
    });
  }
  if (linkedTeams.some(t => canManageTeam_(user, t))) return;
  // آزمون مستقل: مدیرِ حداقل یک تیم (با مجوز مدیریت) می‌تواند بسازد/ببیند
  if (refType === 'standalone' && allowedTeams_(user).some(t => canManageTeam_(user, t))) return;
  throw new Error('برای مدیریت این آزمون مجوز ندارید.');
}

function adminSaveQuiz(payload) {
  const user = getUserContext();
  requireRole_(user, ['admin', 'team_manager']);
  if (!payload || !String(payload.title || '').trim()) throw new Error('عنوان آزمون الزامی است.');
  const teamKey = String(payload.teamKey || '');
  if (!TEAM_CONFIG.some(t => t.key === teamKey) || !canManageTeam_(user, teamKey)) throw new Error('تیم آزمون معتبر نیست.');
  const pass = Number(payload.passScore), minutes = Number(payload.timeLimitMin || 0), attempts = Number(payload.maxAttempts || 0);
  if (!Number.isFinite(pass) || pass < 1 || pass > 100 || !Number.isInteger(minutes) || minutes < 0 || minutes > 350 || !Number.isInteger(attempts) || attempts < 0) throw new Error('تنظیمات نمره، زمان یا تعداد تلاش نامعتبر است.');
  if (['always', 'after_pass', 'never'].indexOf(String(payload.showAnswers || 'after_pass')) === -1) throw new Error('تنظیم نمایش پاسخ نامعتبر است.');
  if (payload.refId) {
    const id = parseDriveId_(payload.refId);
    if (!findTeamFolders_(user).some(t => t.key === teamKey && canManageTeam_(user, t.key) && isWithinTeam_(id, t.id))) throw new Error('فایل آزمون باید در پوشه یک تیم مجاز باشد.');
    payload.refId = id;
  }


  const sh = sheet_('Quizzes');
  const id = String(payload.quizId || '').trim();
  const row = [
    id || newId_('QUIZ'), String(payload.refType || 'standalone'), String(payload.refId || '').trim(),
    String(payload.title).trim(), String(payload.description || '').trim(),
    Number(payload.passScore) || 70, Number(payload.timeLimitMin) || 0, Number(payload.maxAttempts) || 0,
    !!payload.shuffleQuestions, !!payload.shuffleOptions,
    String(payload.showAnswers || 'after_pass'), payload.active !== false, nowIso_(), teamKey
  ];
  if (id) {
    const existing = getQuizRow_(id);
    if (!existing) throw new Error('آزمون پیدا نشد.');
    assertQuizManageScope_(user, existing);
    const linkedPathIds = rows_('PathItems').filter(i => String(i.QuizId || '') === id).map(i => String(i.PathId));
    if (rows_('Paths').some(p => linkedPathIds.indexOf(String(p.PathId)) !== -1 && String(p.TeamKey) !== teamKey)) throw new Error('این آزمون به مسیر تیم دیگری متصل است.');
    if (rows_('Attempts').some(a => String(a.QuizId) === id)) throw new Error('پس از ثبت تلاش، ویرایش تنظیمات آزمون مجاز نیست؛ آزمون جدید بسازید.');
    const values = sh.getDataRange().getValues();
    for (let i = 1; i < values.length; i++) {
      if (String(values[i][0]) === id) {
        // CreatedAt اصلی حفظ می‌شود
        row[12] = values[i][12];
        sh.getRange(i + 1, 1, 1, row.length).setValues([row]);
        return ok_({ quizId: id });
      }
    }
  }
  if (user.role !== 'admin' && !allowedTeams_(user).some(t => canManageTeam_(user, t))) {
    throw new Error('برای ساخت آزمون مجوز ندارید.');
  }
  sh.appendRow(row);
  logActivity_(user.email, 'quiz_create', payload.title);
  return ok_({ quizId: row[0], created: true });
}

function adminDeleteQuiz(quizId) {
  const user = getUserContext();
  requireRole_(user, ['admin', 'team_manager']);
  const existing = getQuizRow_(quizId);
  if (!existing) throw new Error('آزمون پیدا نشد.');
  assertQuizManageScope_(user, existing);
  const sh = sheet_('Quizzes');
  const values = sh.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(quizId)) { sh.getRange(i + 1, 12).setValue(false); return ok_({}); }
  }
  return ok_({});
}

function adminSaveQuestion(payload) {
  const user = getUserContext();
  requireRole_(user, ['admin', 'team_manager']);
  const quiz = getQuizRow_(String(payload.quizId || ''));
  if (!quiz) throw new Error('آزمون پیدا نشد.');
  assertQuizManageScope_(user, quiz);

  if (rows_('Attempts').some(a => String(a.QuizId) === String(payload.quizId))) throw new Error('پس از ثبت تلاش، ویرایش سؤال‌های آزمون مجاز نیست؛ آزمون جدید بسازید.');

  const type = String(payload.type || 'mcq');
  if (QUESTION_TYPES.indexOf(type) === -1) throw new Error('نوع سؤال نامعتبر است.');
  const text = String(payload.text || '').trim();
  if (!text) throw new Error('متن سؤال الزامی است.');
  const options = (payload.options || []).map(o => String(o)).filter(o => o.trim() !== '');
  const points = Math.max(0, Number(payload.points) || 1);

  let correct = payload.correct;
  if (type === 'tf') {
    const v = Number(Array.isArray(correct) ? correct[0] : correct);
    if (v !== 0 && v !== 1) throw new Error('پاسخ درست یا غلط را انتخاب کنید.');
    correct = [v];
    if (options.length === 0) { /* گزینه‌ها خودکار */ }
  } else if (type === 'essay') {
    correct = [];
  } else {
    if (options.length < 2 || (type === 'mcq' && options.length !== 4)) throw new Error('برای چهارگزینه‌ای دقیقاً چهار گزینه و برای سایر انواع حداقل دو گزینه وارد کنید.');
    if (type === 'multi') {
      const arr = (Array.isArray(correct) ? correct : [correct]).map(Number)
        .filter(n => !isNaN(n) && n >= 0 && n < options.length);
      if (!arr.length) throw new Error('حداقل یک پاسخ صحیح انتخاب کنید.');
      correct = arr;
    } else { // mcq / scenario
      const idx = Number(Array.isArray(correct) ? correct[0] : correct);
      if (isNaN(idx) || idx < 0 || idx >= options.length) throw new Error('پاسخ صحیح را انتخاب کنید.');
      correct = [idx];
    }
  }
  const finalOptions = type === 'tf' ? ['درست', 'غلط'] : options;

  const sh = sheet_('Questions');
  const id = String(payload.questionId || '').trim();
  const row = [
    id || newId_('Q'), String(payload.quizId), Number(payload.order) || 0, type, text,
    JSON.stringify(finalOptions), JSON.stringify(correct), points
  ];
  if (id) {
    const values = sh.getDataRange().getValues();
    for (let i = 1; i < values.length; i++) {
      if (String(values[i][0]) === id) { sh.getRange(i + 1, 1, 1, row.length).setValues([row]); return ok_({ questionId: id }); }
    }
    throw new Error('سؤال پیدا نشد.');
  }
  sh.appendRow(row);
  return ok_({ questionId: row[0], created: true });
}

function adminDeleteQuestion(questionId) {
  const user = getUserContext();
  requireRole_(user, ['admin', 'team_manager']);
  const sh = sheet_('Questions');
  const values = sh.getDataRange().getValues();
  for (let i = values.length - 1; i >= 1; i--) {
    if (String(values[i][0]) === String(questionId)) {
      const quiz = getQuizRow_(String(values[i][1]));
      if (quiz) assertQuizManageScope_(user, quiz);
      if (rows_('Attempts').some(a => String(a.QuizId) === String(values[i][1]))) throw new Error('پس از ثبت تلاش، حذف سؤال مجاز نیست.');
      sh.deleteRow(i + 1);
      return ok_({});
    }
  }
  throw new Error('سؤال پیدا نشد.');
}

/* ---------- بازبینی آزمون‌ها ---------- */

function adminListAttempts(filter) {
  const user = getUserContext();
  requireRole_(user, ['admin', 'team_manager', 'mentor']);
  filter = filter || {};
  const status = String(filter.status || 'pending_review');

  // محدوده کاربر: ادمین همه؛ مدیر تیم‌هاش؛ منتور تحت‌پوشش‌هاش
  let scopeEmails = null; // null = همه
  if (user.role !== 'admin') {
    const users = activeUsers_();
    if (user.role === 'team_manager') scopeEmails = users.filter(u => u.teams.some(t => allowedTeams_(user).indexOf(t) !== -1)).map(u => u.email);
    else scopeEmails = users.filter(u => u.mentorEmail === user.email).map(u => u.email);
  }

  const quizMap = {};
  rows_('Quizzes').forEach(q => { quizMap[String(q.QuizId)] = q; });
  const userMap = {};
  rows_('Users').forEach(u => { userMap[String(u.Email).toLowerCase()] = String(u.Name || u.Email); });

  const rowsOut = rows_('Attempts')
    .filter(a => String(a.Status) !== 'in_progress')
    .filter(a => !status || String(a.Status) === status)
    .filter(a => !filter.quizId || String(a.QuizId) === String(filter.quizId))
    .filter(a => !scopeEmails || scopeEmails.indexOf(String(a.Email).toLowerCase()) !== -1)
    .sort((a, b) => new Date(b.Timestamp) - new Date(a.Timestamp))
    .slice(0, 200)
    .map(a => ({
      attemptId: String(a.AttemptId), timestamp: a.Timestamp, email: String(a.Email),
      userName: userMap[String(a.Email).toLowerCase()] || String(a.Email),
      quizId: String(a.QuizId), quizTitle: String((quizMap[String(a.QuizId)] || {}).Title || 'آزمون حذف‌شده'),
      score: Number(a.Score) || 0, passed: truthy_(a.Passed), status: String(a.Status), feedback: String(a.Feedback || '')
    }));
  return ok_({ attempts: rowsOut });
}

function adminGetAttemptDetail(attemptId) {
  const user = getUserContext();
  requireRole_(user, ['admin', 'team_manager', 'mentor']);
  const a = rows_('Attempts').find(x => String(x.AttemptId) === String(attemptId));
  if (!a) throw new Error('تلاش پیدا نشد.');
  // کنترل محدوده
  if (user.role !== 'admin') {
    const target = activeUsers_().find(u => u.email === String(a.Email).toLowerCase());
    const okScope = target && (
      (user.role === 'team_manager' && target.teams.some(t => allowedTeams_(user).indexOf(t) !== -1)) ||
      (user.role === 'mentor' && target.mentorEmail === user.email));
    if (!okScope) throw new Error('مجوز مشاهده این تلاش را ندارید.');
  }
  const questions = questionsOf_(String(a.QuizId)).map(q => ({
    questionId: String(q.QuestionId), type: String(q.Type), text: String(q.Text),
    options: parseJsonSafe_(q.OptionsJSON, []), correct: parseJsonSafe_(q.CorrectJSON, []),
    points: Number(q.Points) || 1
  }));
  return ok_({
    attempt: {
      attemptId: String(a.AttemptId), email: String(a.Email), quizId: String(a.QuizId),
      timestamp: a.Timestamp, score: Number(a.Score) || 0, passed: truthy_(a.Passed),
      status: String(a.Status), answers: parseJsonSafe_(a.AnswersJSON, []), feedback: String(a.Feedback || '')
    },
    questions: questions
  });
}

/** نمره‌دهی سوالات تشریحی + ثبت بازخورد — نمره نهایی تلاش بازمحاسبه می‌شود */
function adminReviewAttempt(attemptId, essayMarks, feedback) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try { return adminReviewAttempt_(attemptId, essayMarks, feedback); }
  finally { lock.releaseLock(); }
}

function adminReviewAttempt_(attemptId, essayMarks, feedback) {
  const user = getUserContext();
  requireRole_(user, ['admin', 'team_manager', 'mentor']);
  const a = rows_('Attempts').find(x => String(x.AttemptId) === String(attemptId));
  if (!a) throw new Error('تلاش پیدا نشد.');
  if (String(a.Status) !== 'pending_review') throw new Error('این تلاش در انتظار بازبینی نیست.');
  if (user.role !== 'admin') {
    const target = activeUsers_().find(u => u.email === String(a.Email).toLowerCase());
    const okScope = target && (
      (user.role === 'team_manager' && target.teams.some(t => allowedTeams_(user).indexOf(t) !== -1)) ||
      (user.role === 'mentor' && target.mentorEmail === user.email));
    if (!okScope) throw new Error('مجوز بازبینی این تلاش را ندارید.');
  }

  const quiz = getQuizRow_(String(a.QuizId));
  const questions = questionsOf_(String(a.QuizId));
  const totalPoints = questions.reduce((s, x) => s + (Number(x.Points) || 1), 0);
  questions.filter(q => String(q.Type) === 'essay').forEach(q => { const n = essayMarks && essayMarks[String(q.QuestionId)]; if (n === undefined || n === '' || !Number.isFinite(Number(n)) || Number(n) < 0 || Number(n) > Number(q.Points)) throw new Error('برای همه سؤال‌های تشریحی نمره معتبر وارد کنید.'); });
  const answers = parseJsonSafe_(a.AnswersJSON, []);
  const ansMap = {};
  answers.forEach(x => { if (x && x.questionId) ansMap[String(x.questionId)] = x.value; });

  let earned = 0;
  questions.forEach(question => {
    const qid = String(question.QuestionId);
    const type = String(question.Type);
    const points = Number(question.Points) || 1;
    if (type === 'essay') {
      const m = essayMarks ? Number(essayMarks[qid]) : 0;
      earned += Math.max(0, Math.min(points, isNaN(m) ? 0 : m));
    } else {
      const correct = parseJsonSafe_(question.CorrectJSON, []);
      const given = ansMap[qid];
      if (type === 'multi') {
        const g = Array.isArray(given) ? given.map(Number).sort() : [];
        const c = (Array.isArray(correct) ? correct : [correct]).map(Number).sort();
        if (g.length === c.length && g.every((v, i) => v === c[i])) earned += points;
      } else {
        const g = (given === null || given === undefined || given === '') ? -1 : Number(given);
        const c = Array.isArray(correct) ? Number(correct[0]) : Number(correct);
        if (g === c) earned += points;
      }
    }
  });

  const score = totalPoints ? Math.round((earned / totalPoints) * 100) : 0;
  const passScore = quiz ? Number(quiz.PassScore) || 70 : 70;
  const passed = score >= passScore;

  const sh = sheet_('Attempts');
  sh.getRange(a.__row, 5).setValue(score);
  sh.getRange(a.__row, 6).setValue(passed);
  sh.getRange(a.__row, 7).setValue('reviewed');
  sh.getRange(a.__row, 9).setValue(String(feedback || ''));

  if (passed) {
    const target = activeUsers_().find(u => u.email === String(a.Email).toLowerCase());
    const progress = getMyProgressMap_(String(a.Email).toLowerCase());
    const paths = rows_('Paths');
    if (target) rows_('PathItems').filter(it => truthy_(it.Active) && String(it.QuizId || '') === String(a.QuizId)).forEach(it => {
      const path = paths.find(p => String(p.PathId) === String(it.PathId) && truthy_(p.Active));
      if (!path || target.teams.indexOf(String(path.TeamKey)) === -1) return;
      const previous = itemsOfPath_(String(path.PathId));
      const index = previous.findIndex(x => String(x.ItemId) === String(it.ItemId));
      if (index >= 0) {
        markProgress_(target.email, 'item:' + String(it.ItemId), score);
        notifyPathCompletion_(target.email, String(path.PathId));
      }
    });
  }
  logActivity_(user.email, 'attempt_reviewed', attemptId + ' → ' + score + '٪');
  return ok_({ score: score, passed: passed });
}

/* ---------- اعلانات ---------- */

function announcementVisible_(a, user) {
  if (!truthy_(a.Active)) return false;
  const aud = String(a.Audience || 'all');
  if (aud === 'all') return true;
  if (aud.indexOf('team:') === 0) return allowedTeams_(user).indexOf(aud.slice(5)) !== -1;
  if (aud.indexOf('role:') === 0) return user.role === aud.slice(5);
  return false;
}

function listAnnouncements() {
  const user = getUserContext();
  const items = rows_('Announcements')
    .filter(a => announcementVisible_(a, user))
    .sort((x, y) => new Date(y.CreatedAt || 0) - new Date(x.CreatedAt || 0))
    .slice(0, 10)
    .map(a => ({ id: String(a.Id), title: String(a.Title), body: String(a.Body), audience: String(a.Audience || 'all'), createdAt: a.CreatedAt }));
  return ok_({ announcements: items });
}

function adminListAnnouncements() {
  const user = getUserContext();
  requireRole_(user, ['admin']);
  const items = rows_('Announcements')
    .sort((x, y) => new Date(y.CreatedAt || 0) - new Date(x.CreatedAt || 0))
    .map(a => ({ id: String(a.Id), title: String(a.Title), body: String(a.Body), audience: String(a.Audience || 'all'), active: truthy_(a.Active), createdAt: a.CreatedAt }));
  return ok_({ announcements: items });
}

function adminSaveAnnouncement(payload) {
  const user = getUserContext();
  requireRole_(user, ['admin']);
  if (!payload || !String(payload.title || '').trim()) throw new Error('عنوان اعلان الزامی است.');
  const audience = String(payload.audience || 'all');
  const valid = audience === 'all' || ['role:admin', 'role:learner'].indexOf(audience) !== -1 || TEAM_CONFIG.some(t => audience === 'team:' + t.key);
  if (!valid || !String(payload.body || '').trim()) throw new Error('مخاطب یا متن اطلاعیه نامعتبر است.');
  const sh = sheet_('Announcements');
  const id = String(payload.id || '').trim();
  const row = [id || newId_('ANN'), String(payload.title).trim(), String(payload.body || '').trim(),
    String(payload.audience || 'all'), payload.active !== false, nowIso_()];
  if (id) {
    const values = sh.getDataRange().getValues();
    for (let i = 1; i < values.length; i++) {
      if (String(values[i][0]) === id) { row[5] = values[i][5]; sh.getRange(i + 1, 1, 1, row.length).setValues([row]); return ok_({ id: id }); }
    }
    throw new Error('اعلان پیدا نشد.');
  }
  sh.appendRow(row);
  logActivity_(user.email, 'announcement_create', payload.title);
  return ok_({ id: row[0], created: true });
}

function adminDeleteAnnouncement(id) {
  const user = getUserContext();
  requireRole_(user, ['admin']);
  const sh = sheet_('Announcements');
  const values = sh.getDataRange().getValues();
  for (let i = values.length - 1; i >= 1; i--) {
    if (String(values[i][0]) === String(id)) { sh.deleteRow(i + 1); return ok_({}); }
  }
  throw new Error('اعلان پیدا نشد.');
}

/** اعلان‌های عمومی مجاز و اعلان‌های اختصاصی کاربر، همراه وضعیت خواندن. */
function listNotifications() {
  const user = getUserContext();
  const seen = new Set(rows_('NotificationReads').filter(r => String(r.Email).toLowerCase() === user.email).map(r => String(r.NotificationId)));
  const publicItems = rows_('Announcements').filter(a => announcementVisible_(a, user)).map(a => ({ id: 'ann:' + String(a.Id), title: String(a.Title), body: String(a.Body), createdAt: a.CreatedAt }));
  const privateItems = rows_('Notifications').filter(n => String(n.Email).toLowerCase() === user.email).map(n => ({id: 'personal:' + String(n.Id), title: String(n.Title), body: String(n.Body), createdAt: n.CreatedAt}));
  const items = publicItems.concat(privateItems).sort((a,b) => new Date(b.createdAt || 0) - new Date(a.createdAt || 0)).slice(0, 100);
  items.forEach(n => { n.read = seen.has(n.id); });
  return ok_({ notifications: items, unread: items.filter(n => !n.read).length });
}

function markNotificationRead(id) {
  const user = getUserContext();
  const key = String(id || '');
  const visible = listNotifications().notifications.some(n => n.id === key);
  if (!visible) throw new Error('اطلاعیه در دسترس شما نیست.');
  if (!rows_('NotificationReads').some(r => String(r.Email).toLowerCase() === user.email && String(r.NotificationId) === key)) {
    sheet_('NotificationReads').appendRow([user.email, key, nowIso_()]);
  }
  return ok_({});
}

/** فقط پس از تکمیل تمام مراحل فعال یک مسیر، اعلان یک‌بار برای همان کاربر ایجاد شود. */
function notifyPathCompletion_(email, pathId) {
  const path = rows_('Paths').find(p => String(p.PathId) === String(pathId) && truthy_(p.Active));
  if (!path) return;
  const items = itemsOfPath_(pathId), progress = getMyProgressMap_(email);
  if (!items.length || !items.every(i => progress['item:' + String(i.ItemId)])) return;
  if (rows_('Notifications').some(n => String(n.Email).toLowerCase() === email && String(n.PathId) === String(pathId))) return;
  sheet_('Notifications').appendRow([newId_('NOTICE'), email, 'دوره آموزشی تکمیل شد', 'مسیر «' + String(path.Title) + '» برای شما تکمیل شد.', nowIso_(), String(pathId)]);
}


// ===== Code.gs =====
/*******************************************************
 * آکادمی یادگیری بیمه‌بازار — نقطه ورود
 * Google Apps Script + Google Drive + Google Sheets
 *******************************************************/

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('آکادمی یادگیری | بیمه‌بازار')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include_(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * اجرای اولیه توسط مالک از ویرایشگر. این تابع در فهرست Run نمایش داده می‌شود.
 * از آنجا که توابع عمومی از مرورگر هم قابل فراخوانی‌اند، هویت مالک کنترل می‌شود.
 */
function setupAcademy() {
  const active = String(Session.getActiveUser().getEmail() || '').toLowerCase().trim();
  const owner = String(Session.getEffectiveUser().getEmail() || '').toLowerCase().trim();
  if (!active || !owner || active !== owner) throw new Error('فقط مالک پروژه می‌تواند راه‌اندازی را اجرا کند؛ هویت مالک قابل تأیید نیست.');
  if (!CONFIG.ROOT_FOLDER_ID || CONFIG.ROOT_FOLDER_ID === 'PUT_YOUR_DRIVE_FOLDER_ID_HERE') throw new Error('ابتدا شناسه پوشه اصلی Drive را در ROOT_FOLDER_ID تنظیم کنید.');
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const result = setupAcademy_();
    provisionInitialAdmin_();
    Logger.log('راه‌اندازی انجام شد. فایل داده: ' + result.spreadsheetUrl);
    return result;
  } finally { lock.releaseLock(); }
}

/** اولین مدیر از Script Properties به شکل یک‌باره ایجاد می‌شود؛ قابل فراخوانی توسط کاربر وب نیست. */
function provisionInitialAdmin_() {
  const props = PropertiesService.getScriptProperties();
  const username = normalizeUsername_(props.getProperty('INITIAL_ADMIN_USERNAME'));
  const password = props.getProperty('INITIAL_ADMIN_PASSWORD');
  if (!username && !password) {
    if (!rows_('Credentials').length) Logger.log('برای ساخت مدیر اولیه INITIAL_ADMIN_USERNAME و INITIAL_ADMIN_PASSWORD را در Script Properties تنظیم و setupAcademy را دوباره اجرا کنید.');
    return;
  }
  if (!username || !password) throw new Error('هر دو مقدار INITIAL_ADMIN_USERNAME و INITIAL_ADMIN_PASSWORD الزامی است.');
  savePassword_(username, password);
  const users = sheet_('Users'), existing = rows_('Users').find(u => normalizeUsername_(u.Email) === username);
  if (existing) {
    users.getRange(existing.__row, 3).setValue('admin');
    users.getRange(existing.__row, 6).setValue(true);
    users.getRange(existing.__row, 7).setValue(true);
  } else users.appendRow([username, 'مدیر سامانه', 'admin', TEAM_CONFIG.map(t=>t.key).join(','), '', true, true, nowIso_()]);
  props.deleteProperty('INITIAL_ADMIN_USERNAME'); props.deleteProperty('INITIAL_ADMIN_PASSWORD');
}

/** ساخت ساختار شیت‌ها، بدون درج داده آزمایشی. */
function setupAcademy_() {
  let ss;
  const existing = CONFIG.SPREADSHEET_ID ||
    PropertiesService.getScriptProperties().getProperty('ACADEMY_SPREADSHEET_ID');
  if (existing) {
    ss = SpreadsheetApp.openById(existing);
  } else {
    // در اسکریپت متصل به Sheet از همان فایل استفاده می‌کنیم؛ در پروژه مستقل فایل جدید ساخته می‌شود.
    ss = SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.create('داده‌های آکادمی یادگیری بیمه‌بازار');
  }

  PropertiesService.getScriptProperties().setProperty('ACADEMY_SPREADSHEET_ID', ss.getId());
  Object.keys(SHEET_DEFS).forEach(name => ensureSheet_(name, SHEET_DEFS[name]));
  const usersSheet = sheet_('Users');
  if (usersSheet.getLastRow() < 2) {
    const owner = String(Session.getEffectiveUser().getEmail() || '').toLowerCase().trim();
    if (!owner) throw new Error('ایمیل مالک قابل تشخیص نیست؛ مدیر اولیه را در شیت Users ثبت کنید.');
    usersSheet.appendRow([owner, owner, 'admin', TEAM_CONFIG.map(t => t.key).join(','), '', true, true, nowIso_()]);
  }


  // حذف شیت پیش‌فرض خالی اگر وجود دارد
  ['Sheet1', 'برگه1', 'Sheet 1'].forEach(n => {
    const s = ss.getSheetByName(n);
    if (s && s.getLastRow() === 0) { try { ss.deleteSheet(s); } catch (e) {} }
  });

  PropertiesService.getScriptProperties().setProperty('ACADEMY_SPREADSHEET_ID', ss.getId());
  // کش پوشه‌های تیم را نوسازی کن
  cacheRemove_('team_folders_v1');
  return ok_({ spreadsheetUrl: ss.getUrl(), message: 'ساختار داده‌های آکادمی آماده شد.' });
}

/** داده‌ی اولیه اپ — همه‌چیز از بک‌اند و بر اساس هویت و نقش کاربر */
function getBootstrap() {
  const user = getUserContext();
  if (CONFIG.LOG_LOGINS) logActivity_(user.email, 'login', 'boot');

  // تیم‌ها از Drive (در حالت دمو نیز واقعی است)
  let teams = [], rootName = '', driveError = '';
  try {
    rootName = Drive.Files.get(CONFIG.ROOT_FOLDER_ID, { fields: 'name', supportsAllDrives: true }).name;
    teams = findTeamFolders_(user);
  } catch (e) {
    driveError = e.message || String(e);
  }

  // مسیرهای من + آمار واقعی
  const myPathsResp = listMyPaths('');
  const myPaths = myPathsResp.paths || [];
  const agg = myPaths.reduce((a, p) => ({ total: a.total + p.total, done: a.done + p.done }), { total: 0, done: 0 });

  // آزمون‌های در انتظار (آیتم‌های مسیر که آزمون دارند و هنوز قبول نشده‌اند)
  const progress = getMyProgressMap_(user.email);
  const myAttempts = rows_('Attempts').filter(a => String(a.Email).toLowerCase() === user.email);
  const attempts = myAttempts.filter(a => truthy_(a.Passed));
  const passedQuizzes = {};
  attempts.forEach(a => { passedQuizzes[String(a.QuizId)] = true; });
  let pendingQuizzes = 0;
  const pathIds = new Set(myPaths.map(p => p.pathId));
  rows_('PathItems').filter(i => pathIds.has(String(i.PathId)) && truthy_(i.Active)).forEach(i => {
    const qid = String(i.QuizId || '');
    if (qid && !passedQuizzes[qid] && !progress['item:' + String(i.ItemId)]) pendingQuizzes++;
  });

  // فعالیت‌های اخیر کاربر
  const recent = rows_('Activity')
    .filter(a => String(a.Email).toLowerCase() === user.email)
    .sort((a, b) => new Date(b.Timestamp) - new Date(a.Timestamp))
    .slice(0, 6)
    .map(a => ({ action: String(a.Action), detail: String(a.Detail || ''), timestamp: a.Timestamp }));

  // اعلانات
  const ann = listAnnouncements().announcements;

  return ok_({
    user: user,
    drive: { rootName: rootName, error: driveError },
    teams: teams,
    stats: {
      teams: teams.length,
      paths: myPaths.length,
      totalItems: agg.total,
      completedItems: agg.done,
      completionPercent: agg.total ? Math.round(agg.done / agg.total * 100) : 0,
      pendingQuizzes: pendingQuizzes,
      totalAttempts: myAttempts.filter(a => String(a.Status) !== 'in_progress').length
    },
    activePaths: myPaths.filter(p => p.percent < 100).slice(0, 4),
    allPaths: myPaths,
    recommended: myPaths.filter(p => p.done === 0 && p.total > 0).slice(0, 3),
    recentActivity: recent,
    announcements: ann,
    notifications: listNotifications().notifications,
    pageSize: CONFIG.PAGE_SIZE,
    roleLabels: ROLE_LABELS
  });
}
