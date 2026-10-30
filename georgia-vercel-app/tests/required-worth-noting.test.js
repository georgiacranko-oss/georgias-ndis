import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import vm from 'node:vm';
const html=fs.readFileSync(new URL('../public/app.html',import.meta.url),'utf8');
const backend=fs.readFileSync(new URL('../apps-script/SupportBackend.gs',import.meta.url),'utf8');
test('blank and whitespace notes are identified on the Finish step',()=>{
 const fn=html.split('\n').find(line=>line.startsWith('function missingNoteQuestions('));
 const c={};vm.createContext(c);vm.runInContext(fn,c);
 for(const notes of ['', '   ']){const missing=c.missingNoteQuestions({elements:{namedItem:key=>({value:key==='notes'?notes:'Completed'})}});assert.equal(missing.length,1);assert.equal(missing[0][0],'notes');assert.equal(missing[0][2],4);}
 assert.equal(c.missingNoteQuestions({elements:{namedItem:()=>({value:'Completed'})}}).length,0);
});
test('backend rejects empty notes before touching spreadsheet data',()=>{
 const c={SUPPORT_WEB_REQUEST_:{email:'ruby.forbes@outlook.com'},Session:{getEffectiveUser:()=>({getEmail:()=> 'georgia.cranko@gmail.com'}),getActiveUser:()=>({getEmail:()=> ''})}};
 vm.createContext(c);vm.runInContext(backend,c);
 for(const notes of ['', '   '])assert.throws(()=>c.SUPPORT_APP_submit_({id:'12345678-1234-1234-1234-123456789abc',fields:{worker:'Ruby',tasks:'Support',notes}}),/something worth mentioning/);
 assert.throws(()=>c.SUPPORT_APP_submit_({id:'12345678-1234-1234-1234-123456789abc',fields:{worker:'Ruby',tasks:'Support',notes:'A completed observation',date:'invalid'}}),/valid support date/);
});
