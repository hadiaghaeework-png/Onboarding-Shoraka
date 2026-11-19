/*******************************************************
 * پنل مدیریت — کاربران، محتوا، آزمون‌ها، بازبینی، اعلانات
 * همه توابع مجوز را سرور-ساید بررسی می‌کنند.
 *******************************************************/

/* ---------- کاربران ---------- */

function adminListUsers() {
  const user = getUserContext();
  requireRole_(user, ['admin']);
  const users = rows_('Users').map(u => ({
    email: String(u.Email), name: String(u.Name || ''), role: normalizeRole_(u.Role),
    teams: splitCsv_(u.Teams), mentorEmail: String(u.MentorEmail || ''),
    canManage: truthy_(u.CanManage), active: truthy_(u.Active), createdAt: u.CreatedAt
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
  if (['admin', 'learner'].indexOf(String(payload.role)) === -1) throw new Error('نقش باید مدیر یا کارشناس باشد.');
  const role = normalizeRole_(payload.role);
  if (email === user.email && role !== 'admin') throw new Error('نمی‌توانید نقش خودتان را تغییر دهید.');
  const teams = (payload.teams || []).filter(k => TEAM_CONFIG.some(t => t.key === k));
  if (role === 'learner' && (teams.length !== 1 || (payload.teams || []).length !== 1)) throw new Error('برای هر کارشناس دقیقاً یک تیم معتبر انتخاب کنید.');

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
      if (index >= 0 && previous.slice(0, index).every(x => !truthy_(x.Required) || progress['item:' + String(x.ItemId)])) {
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
