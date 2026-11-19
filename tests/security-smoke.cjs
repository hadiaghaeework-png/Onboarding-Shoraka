// Small isolated regression checks; these do not replace a real GAS security audit.
const vm = require('node:vm');
const fs = require('node:fs');
const assert = require('node:assert/strict');
let user = {email:'a@example.com',role:'learner',teams:['account']};
const data = {
  Paths:[{PathId:'p1',TeamKey:'account',Active:true},{PathId:'p2',TeamKey:'mlm',Active:true}],
  PathItems:[{ItemId:'a1',PathId:'p1',Order:1,Required:true,Active:true},{ItemId:'a2',PathId:'p1',Order:2,Required:true,QuizId:'quiz1',Active:true},{ItemId:'b1',PathId:'p2',Order:1,Required:true,QuizId:'quiz2',Active:true}],
  Quizzes:[{QuizId:'quiz1',Active:true},{QuizId:'quiz2',Active:true}],Progress:[],Attempts:[],Questions:[]
};
const sandbox = { rows_: name => data[name] || [], getUserContext:()=>user,
  truthy_: x=>x===true, allowedTeams_: u=>u.teams, assertTeamAccess_:(u,t)=>{if(!u.teams.includes(t))throw Error('denied')},
  teamConfigByKey_: k=>({name:k}), getQuizRefMap_:()=>({}), getQuizMetaById_:()=>({available:true}),
  findTeamFolders_:()=>[], isWithinTeam_:()=>false, ok_: x=>x, logActivity_:()=>{},
  markProgress_:()=>{}, console };
vm.createContext(sandbox);
vm.runInContext(fs.readFileSync('Paths.gs','utf8')+'\n'+fs.readFileSync('Quizzes.gs','utf8'),sandbox);
assert.throws(()=>sandbox.markItemComplete('a2'),/ابتدا باید در آزمون/);
assert.throws(()=>sandbox.getPath('p2'),/denied/);
assert.throws(()=>sandbox.getQuizForAttempt('quiz2'),/دسترسی/);
assert.doesNotThrow(()=>sandbox.assertQuizAccess_(user,data.Quizzes[0]));
data.Progress.push({Email:user.email,PathItemId:'item:a1',Status:'completed'});
assert.equal(sandbox.getPath('p1').items[1].status,'quiz_pending');
assert.doesNotThrow(()=>sandbox.assertQuizAccess_(user,data.Quizzes[0]));
console.log('Security smoke checks passed (mocked GAS).');
