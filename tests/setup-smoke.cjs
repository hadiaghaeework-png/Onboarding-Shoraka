const vm = require('node:vm'), fs = require('node:fs'), assert = require('node:assert/strict');
let active='owner@example.test', effective='owner@example.test', locks=0, releases=0, starts=0;
const ctx = { Session: {getActiveUser:()=>({getEmail:()=>active}),getEffectiveUser:()=>({getEmail:()=>effective})}, CONFIG:{ROOT_FOLDER_ID:'valid-folder'}, LockService:{getScriptLock:()=>({waitLock(){locks++}, releaseLock(){releases++}})}, Logger:{log(){}}, SpreadsheetApp:{}, PropertiesService:{}, console };
vm.createContext(ctx); vm.runInContext(fs.readFileSync('Code.gs','utf8'),ctx);
ctx.setupAcademy_=()=>{starts++;return {spreadsheetUrl:'https://example.test'}};
assert.equal(ctx.setupAcademy().spreadsheetUrl,'https://example.test');assert.equal(starts,1);assert.equal(locks,1);assert.equal(releases,1);
active='someone@example.test';assert.throws(()=>ctx.setupAcademy(),/مالک/);active='';assert.throws(()=>ctx.setupAcademy(),/مالک/);assert.equal(starts,1);
active=effective;ctx.CONFIG.ROOT_FOLDER_ID='PUT_YOUR_DRIVE_FOLDER_ID_HERE';assert.throws(()=>ctx.setupAcademy(),/شناسه/);assert.equal(starts,1);
console.log('Setup owner/placeholder smoke checks passed (mocked GAS).');
