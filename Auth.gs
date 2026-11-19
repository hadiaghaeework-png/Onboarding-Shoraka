/*******************************************************
 * هویت و کنترل دسترسی (RBAC مبتنی بر نقش و تیم)
 * همه توابع سرویس‌دهنده باید از این لایه عبور کنند.
 *******************************************************/

function normalizeRole_(role) {
  const r = String(role || '').toLowerCase().trim();
  return r === 'admin' ? 'admin' : 'learner';
}

function roleLabel_(role) {
  return ROLE_LABELS[role] || ROLE_LABELS.learner;
}

/**
 * هویت کاربر جاری را از Session + شیت Users (یا fallback کد) می‌سازد.
 * در هیچ حالتی دسترسی بدون این تابع برقرار نمی‌شود.
 */
function getUserContext() {
  const email = (Session.getActiveUser().getEmail() || '').toLowerCase().trim();


  if (!email) {
    throw new Error('ایمیل حساب Google قابل شناسایی نیست. هویت در این استقرار قابل تأیید نیست؛ تنظیمات دامنه و انتشار Google Workspace را بررسی کنید.');
  }

  // No fallback on read errors: authorization must fail closed.
  const u = rows_('Users').find(x => String(x.Email).toLowerCase().trim() === email);
  if (u) {
    if (!truthy_(u.Active)) throw new Error('حساب شما غیرفعال شده است. با مدیر سامانه تماس بگیرید.');
    const storedRole = String(u.Role || '').trim().toLowerCase();
    if (['admin', 'learner'].indexOf(storedRole) === -1) throw new Error('نقش حساب نامعتبر است؛ با مدیر سامانه تماس بگیرید.');
    const role = normalizeRole_(storedRole);
    const teams = splitCsv_(u.Teams);
    if (role === 'learner' && (teams.length !== 1 || !TEAM_CONFIG.some(t => t.key === teams[0]))) throw new Error('تیم این کارشناس باید دقیقاً یک تیم معتبر باشد؛ با مدیر سامانه تماس بگیرید.');
    return { email: email, name: String(u.Name || email), role: role,
      roleLabel: roleLabel_(role), teams: teams,
      canManage: role === 'admin', mentorEmail: '', demo: false };
  }

  throw new Error('برای حساب شما در آکادمی دسترسی تعریف نشده است. با مدیر سامانه تماس بگیرید.');
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
