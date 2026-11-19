// Regression test: reading a path with many quizzes must not read each sheet once per item.
const vm = require('node:vm'), fs = require('node:fs'), assert = require('node:assert/strict');
const user = {email:'learner@example.test', role:'learner', teams:['account']};
const data = {Paths:[{PathId:'p',TeamKey:'account',Active:true}], PathItems:[], Progress:[], Attempts:[], Questions:[], Quizzes:[]};
for(let i=0;i<100;i++) { data.PathItems.push({PathId:'p',ItemId:'i'+i,Order:i,Required:false,Active:true,QuizId:'q'+i}); data.Quizzes.push({QuizId:'q'+i,Active:true}); data.Questions.push({QuizId:'q'+i}); }
const counts = {};
const c = {rows_:n=>{counts[n]=(counts[n]||0)+1;return data[n]||[]},getUserContext:()=>user,truthy_:v=>v===true,allowedTeams_:u=>u.teams,assertTeamAccess_:(u,t)=>{if(!u.teams.includes(t))throw Error('denied')},teamConfigByKey_:()=>({name:'Account'}),ok_:x=>x};
vm.createContext(c);vm.runInContext(fs.readFileSync('Paths.gs','utf8'),c);
const result=c.getPath('p');assert.equal(result.items.length,100);assert.equal(result.items[0].quiz.questionCount,1);
for(const key of ['Paths','PathItems','Progress','Attempts','Quizzes','Questions']) assert.equal(counts[key],1,key+' should be read exactly once');
console.log('Path batch-read regression passed (mocked GAS).');
