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
