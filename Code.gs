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
    Logger.log('راه‌اندازی انجام شد. فایل داده: ' + result.spreadsheetUrl);
    return result;
  } finally { lock.releaseLock(); }
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
