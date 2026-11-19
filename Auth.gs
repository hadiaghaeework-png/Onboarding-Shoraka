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
  if (role === 'learner' && (teams.length !== 1 || !teamConfigs_().some(t => t.key === teams[0]))) throw new Error('تیم این کارشناس باید دقیقاً یک تیم معتبر باشد.');
  if (role === 'mentor' && (!teams.length || teams.some(t => !teamConfigs_().some(x => x.key === t)))) throw new Error('تیم‌های منتور نامعتبر است.');
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
    adminListUsers, adminUpsertUser, adminSetUserActive, adminTeamsStatus, adminImportUsers, adminCreateTeam, adminUploadFile, adminTrashFile, adminListPaths, adminSavePath, adminDeletePath,
    adminPathItems, adminSavePathItem, adminDeletePathItem, adminDriveMeta, adminListQuizzes, adminGetQuiz, adminSaveQuiz,
    adminDeleteQuiz, adminSaveQuestion, adminDeleteQuestion, adminListAttempts, adminGetAttemptDetail, adminReviewAttempt,
    adminListAnnouncements, adminSaveAnnouncement, adminDeleteAnnouncement, reportSummary, reportTeamMembers, reportMember,
    reportIncomplete, reportExportCsv, reportMemberFiles, addMentorNote, presencePing
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
  if (user.role === 'admin') return teamConfigs_().map(t => t.key);
  return (user.teams || []).filter(k => teamConfigs_().some(t => t.key === k));
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
  return teamConfigs_().find(t => t.key === key) || null;
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
