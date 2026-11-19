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
  Paths:        ['PathId', 'TeamKey', 'Title', 'Description', 'Order', 'Active', 'CreatedAt'],
  PathItems:    ['ItemId', 'PathId', 'Order', 'Title', 'DriveItemId', 'ItemType', 'Required', 'QuizId', 'EstimatedMin', 'Active'],
  Progress:     ['Email', 'PathItemId', 'Status', 'CompletedAt', 'Score'],
  Quizzes:      ['QuizId', 'RefType', 'RefId', 'Title', 'Description', 'PassScore', 'TimeLimitMin', 'MaxAttempts', 'ShuffleQuestions', 'ShuffleOptions', 'ShowAnswers', 'Active', 'CreatedAt'],
  Questions:    ['QuestionId', 'QuizId', 'Order', 'Type', 'Text', 'OptionsJSON', 'CorrectJSON', 'Points'],
  Attempts:     ['AttemptId', 'Timestamp', 'Email', 'QuizId', 'Score', 'Passed', 'Status', 'AnswersJSON', 'Feedback', 'TimeSpentSec'],
  Announcements:['Id', 'Title', 'Body', 'Audience', 'Active', 'CreatedAt'],
  Activity:     ['Id', 'Timestamp', 'Email', 'Action', 'Detail'],
  MentorNotes:  ['Id', 'Timestamp', 'MentorEmail', 'MenteeEmail', 'Note']
};

const ROLE_LABELS = {
  admin: 'مدیر سامانه',
  team_manager: 'مدیر تیم',
  mentor: 'منتور',
  learner: 'کارشناس'
};

const QUESTION_TYPES = ['mcq', 'multi', 'tf', 'essay', 'scenario'];
