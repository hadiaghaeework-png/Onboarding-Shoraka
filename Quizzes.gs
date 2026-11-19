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
    const item = { questionId: qid, type: type, points: points, correct: correct, explanation: String(question.Explanation || ''), sourceLocation: String(question.SourceLocation || ''), given: (given === undefined ? null : given), isCorrect: null, awarded: 0 };
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
      isCorrect: r.isCorrect, awarded: r.awarded, points: r.points, explanation: r.explanation, sourceLocation: r.sourceLocation
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
