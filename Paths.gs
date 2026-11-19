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
