import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import vm from 'node:vm';
const code=fs.readFileSync(new URL('../apps-script/CalendarInput.gs',import.meta.url),'utf8');
function context(email){const c={SUPPORT_WEB_REQUEST_:email?{email}:null,Date,NDIA_APP:{SUPPORT_WEB_calendarId_:()=> 'calendar'},Utilities:{parseDate:value=>new Date(value.replace(' ','T')+':00Z'),formatDate:date=>date.toISOString().slice(0,10)},Calendar:{Events:{}}};vm.createContext(c);vm.runInContext(code,c);return c;}
test('input endpoints reject unsigned users and workers editing notes',()=>{let c=context();assert.throws(()=>c.SUPPORT_APP_calendarEventSave_({}),/Signed website/);c=context('ruby.forbes@outlook.com');assert.throws(()=>c.SUPPORT_APP_calendarNoteSave_({}),/Only Georgia/);});
test('new app events delegate to the Schedule Changes save with the same retry ID',()=>{
 const c=context('ruby.forbes@outlook.com');let written;
 c.SUPPORT_APP_calendarSave_=request=>{written=request;return {saved:true,row:42};};
 const result=c.SUPPORT_APP_calendarEventSave_({requestId:'12345678-1234-1234-1234-123456789abc',title:'Outing',type:'Outing',day:'2026-10-28',until:'2026-10-28',allDay:true});
 assert.equal(result.saved,true);assert.equal(written.id,'12345678-1234-1234-1234-123456789abc');assert.equal(written.fields.type,'Activity');assert.equal(written.fields.details,'Website label: Outing');
});
test('worker cannot edit an imported event without an owned sheet record',()=>{
 const c=context('ruby.forbes@outlook.com');c.Calendar.Events.get=()=>({etag:'tag'});c.SpreadsheetApp={openById:()=>({})};c.SUPPORT_APP_SHEET_ID='sheet';c.SUPPORT_APP_calendarTable_=()=>({header:1,sheet:{getLastRow:()=>1}});
 assert.throws(()=>c.SUPPORT_APP_calendarEventSave_({id:'event',etag:'tag',title:'Test',type:'Other',day:'2026-10-28',until:'2026-10-28',allDay:true}),/Only Georgia/);
});
test('event edits reject stale versions before writing',()=>{const c=context('georgia.cranko@gmail.com');c.Calendar.Events.get=()=>({etag:'new',extendedProperties:{private:{}}});assert.throws(()=>c.SUPPORT_APP_calendarEventSave_({id:'event',etag:'old',title:'Test',day:'2026-10-28',allDay:true}),/event changed/);});
test('calendar input browser script parses',()=>{new vm.Script(fs.readFileSync(new URL('../public/calendar-input.js',import.meta.url),'utf8'));});
