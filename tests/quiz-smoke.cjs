const vm=require('node:vm'),fs=require('node:fs'),assert=require('node:assert/strict');
const db={Quizzes:[{QuizId:'q',Active:true,Title:'Exam',PassScore:70,MaxAttempts:1,TimeLimitMin:0}],Questions:[{QuizId:'q',QuestionId:'question',Type:'mcq',Text:'?',OptionsJSON:'["a","b"]',CorrectJSON:'[0]',Points:1}],Attempts:[],Paths:[],PathItems:[]};
const tokens={};let seq=0;const user={email:'learner@example.test',teams:['account'],role:'learner'};
const ctx={rows_:n=>db[n]||[], getUserContext:()=>user, truthy_:v=>v===true, ok_:v=>v, Utilities:{getUuid:()=>String(++seq)},cachePut_:(k,v)=>tokens[k]=v,cacheGet_:k=>tokens[k],cacheRemove_:k=>delete tokens[k],allowedTeams_:u=>u.teams,findTeamFolders_:()=>[],isWithinTeam_:()=>false,newId_:()=>String(++seq),nowIso_:()=>new Date().toISOString(),logActivity_:()=>{},markProgress_:()=>{},sheet_:name=>({appendRow:r=>db[name].push({AttemptId:r[0],Timestamp:r[1],Email:r[2],QuizId:r[3],Score:r[4],Passed:r[5],Status:r[6]})}),LockService:{getScriptLock:()=>({waitLock(){},releaseLock(){}})}};
vm.createContext(ctx);vm.runInContext(fs.readFileSync('Quizzes.gs','utf8'),ctx);
// Exercise core submission and replay with authorization mocked separately.
ctx.assertQuizAccess_=()=>{};
const start=ctx.getQuizForAttempt('q');assert(start.attemptToken);assert.equal(start.questions[0].correct,undefined);
const result=ctx.submitQuiz('q',[{questionId:'question',value:0}],1,start.attemptToken);assert.equal(result.score,100);assert.equal(result.passed,true);assert.equal(db.Attempts.length,1);
assert.throws(()=>ctx.submitQuiz('q',[{questionId:'question',value:0}],1,start.attemptToken),/نشست/);
assert.throws(()=>ctx.getQuizForAttempt('q'),/تلاش/);
console.log('Quiz smoke checks passed (mocked GAS).');
