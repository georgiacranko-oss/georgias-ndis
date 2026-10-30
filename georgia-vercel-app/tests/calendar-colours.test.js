import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import vm from 'node:vm';
const html=fs.readFileSync(new URL('../public/app.html',import.meta.url),'utf8');
const style=html.split('\n').find(line=>line.startsWith('function calendarLineStyle('));
function getStyle(type,support,day,attended=false){const text='Sample event';const c={data:{calendar:{days:[{number:day,lineMetadata:{[text]:{type,support,attended}},lineStyles:{}}]}},calendarDateKey:n=>'2026-10-'+String(n).padStart(2,'0'),melbourneToday:()=> '2026-10-28',workerNames:()=>[]};vm.createContext(c);vm.runInContext(style,c);return c.calendarLineStyle(text,day);}
test('future appointments and activities have no event borders',()=>{for(const type of ['Appointment','Activity']){const css=getStyle(type,'Unsupported',29);assert.match(css,/border:0/);assert.doesNotMatch(css,/border-bottom/);}});
test('appointment styles ignore support and supported activities use muted magenta',()=>{assert.equal(getStyle('Appointment','Unsupported',27,true),getStyle('Appointment','Supported',27,true));assert.match(getStyle('Activity','Supported',27),/#a84b8f/);});

test('past appointments use navy regardless of recorded attendance',()=>{assert.match(getStyle('Appointment','Supported',27),/border-bottom:1px solid #082a66/);assert.match(getStyle('Appointment','Supported',27,true),/border-bottom:1px solid #082a66/);});

test('non-funded appointment and other-activity styles do not swap',()=>{
 assert.match(getStyle('Non-Funded Appointment','Unsupported',27,true),/border-bottom:2px solid #b8bec8/);
 assert.match(getStyle('Other','Not recorded',27),/border-bottom:2px solid #000000/);
 assert.match(getStyle('Other','Not recorded',30),/border-bottom:2px solid #000000/);
});

test('label text stays teal while the source keeps its wellbeing underline colour',()=>{const c={data:{calendar:{days:[{number:1,supportLabels:['Rough day'],lineStyles:{'Rough day':{color:'#D32F2F'}}}]}},calendarDateKey:()=> '2026-10-01',melbourneToday:()=> '2026-10-28',workerNames:()=>[]};vm.createContext(c);vm.runInContext(style,c);const result=c.calendarLineStyle('Rough day',1);assert.match(result,/color:#133f52/);assert.match(result,/border-bottom:2px solid #D32F2F/);});
