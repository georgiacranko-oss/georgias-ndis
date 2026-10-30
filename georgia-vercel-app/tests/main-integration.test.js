import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import vm from 'node:vm';
const main=fs.readFileSync(new URL('../apps-script/Main.gs',import.meta.url),'utf8');
const bridge=fs.readFileSync(new URL('../apps-script/WebBridge.gs',import.meta.url),'utf8');
test('all original public Main wrappers require owner or signed bridge',()=>{
 const head=main.split('const NDIA_APP =')[0];
 const wrappers=[...head.matchAll(/^function (\w+)\([^\n]*?\) \{([^\n]+)\}/gm)].filter(m=>m[2].includes('NDIA_APP.'));
 assert.equal(wrappers.filter(m=>!m[1].endsWith('_')).length,60);
 for(const wrapper of wrappers)assert.match(wrapper[2],/^ SUPPORT_WEB_requireOwnerOrBridge_\(\);/);
 assert.doesNotMatch(head,/return NDIA_APP\.doGet/);
});
test('owner gate rejects anonymous and workers, allows owner and verified bridge',()=>{
 const code=bridge.slice(bridge.indexOf('function SUPPORT_WEB_requireOwnerOrBridge_'),bridge.indexOf('function doPost'));
 let email='';const context=vm.createContext({SUPPORT_WEB_REQUEST_:null,Session:{getActiveUser:()=>({getEmail:()=>email})}});
 vm.runInContext(code,context);
 assert.throws(()=>context.SUPPORT_WEB_requireOwnerOrBridge_(),/Owner access required/);
 email='ruby.forbes@outlook.com';assert.throws(()=>context.SUPPORT_WEB_requireOwnerOrBridge_(),/Owner access required/);
 email='georgia.cranko@gmail.com';assert.doesNotThrow(()=>context.SUPPORT_WEB_requireOwnerOrBridge_());
 email='';context.SUPPORT_WEB_REQUEST_={email:'ruby.forbes@outlook.com'};assert.doesNotThrow(()=>context.SUPPORT_WEB_requireOwnerOrBridge_());
});
test('needed helpers are exported from Main closure',()=>{
 const exports=main.slice(main.lastIndexOf('return {'));
 for(const name of ['HANDOVER_findNextVisitor_','GCAL4_writeDayGrid_','GCAL4_renderStackedShiftCards_','EMAIL_sendOnce_'])assert.ok(exports.includes(name),name);
});
test('all Apps Script files parse and include one public doGet and doPost',()=>{
 let combined='';for(const name of ['Main.gs','SupportBackend.gs','WebBridge.gs','GoogleCalendar.gs','CalendarInput.gs']){const code=fs.readFileSync(new URL('../apps-script/'+name,import.meta.url),'utf8');new vm.Script(code,{filename:name});combined+=code+'\n';}
 new vm.Script(combined);
 const top=main.split('const NDIA_APP =')[0]+bridge;
 assert.equal([...top.matchAll(/^function doGet\(/gm)].length,1);
 assert.equal([...top.matchAll(/^function doPost\(/gm)].length,1);
});
