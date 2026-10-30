import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import vm from 'node:vm';
const code=fs.readFileSync(new URL('../apps-script/GoogleCalendar.gs',import.meta.url),'utf8');
const main=fs.readFileSync(new URL('../apps-script/Main.gs',import.meta.url),'utf8');
const rules=main.slice(main.indexOf('function GCAL4_shiftWellbeingStatus_'),main.indexOf('function GCAL4_shiftTimeRangeLabel_'));
function setup(rows){
 const sheet={getLastRow:()=>rows.length,getRange:()=>({getValues:()=>rows})};
 const c={Date,Set,SUPPORT_APP_noteVersion_:()=> 'version',SUPPORT_APP_SHEET_ID:'test',SpreadsheetApp:{openById:()=>({getSheetByName:()=>sheet})},Utilities:{formatDate:date=>date.toISOString().slice(0,10)},SUPPORT_APP_calendarTable_(){throw Error('not configured');}};
 vm.createContext(c);vm.runInContext(rules,c);c.NDIA_APP={SUPPORT_WEB_noteLabels_:n=>({wellbeing:c.GCAL4_shiftWellbeingLabel_(n),pain:c.GCAL4_shiftPainLabel_(n),outing:n.outing,appointment:n.appointment})};vm.runInContext(code,c);c.SUPPORT_CALENDAR_table_=()=>{throw Error('not configured');};c.SUPPORT_CALENDAR_noteVersion_=()=> 'version';return c;
}
function note(date,hours,values={}){const row=Array(30).fill('');row[1]='Worker';row[2]=new Date(date+'T12:00:00Z');row[3]=hours;row[29]='Support note';for(const [col,value]of Object.entries(values))row[Number(col)]=value;return row;}
const model=()=>({monthKey:'2026-10',days:Array.from({length:31},(_,i)=>({number:i+1,lines:[],lineStyles:{}}))});
test('restores original day classification, always includes worse pain and keeps decimal hours',()=>{
 const rows=[[],note('2026-10-01',2.25,{16:'Worse',17:'Worse',21:'Worse',22:'Head'}),note('2026-10-02',1.7,{16:'Much worse',17:'Worse'}),note('2026-10-03',3,{16:'Better',17:'Better'})];
 const c=setup(rows),m=model();c.SUPPORT_APP_calendarSupport_(m);
 assert.ok(m.days[0].supportLabels.includes('Mediocre day'));assert.ok(m.days[0].supportLabels.includes('Headache'));assert.ok(m.days[1].supportLabels.includes('Rough day'));assert.ok(m.days[2].supportLabels.includes('Good day'));
 const values=Object.fromEntries(m.summary.map(r=>[r.label,r.value]));assert.equal(values['Support Hours'],'6.95');assert.equal(values['Support Days'],'3');assert.equal(values['Outings'],'0');assert.equal('Pain' in values,false);assert.equal(values['Good days'],'1');assert.equal(values['Worse days'],'2');assert.equal(values['Appointments Attended'],'Unavailable');assert.equal(m.supportNotes.length,3);
});
test('another worse note on the same date prevents counting a good day',()=>{
 const c=setup([[],note('2026-10-01',1,{16:'Better',17:'Better'}),note('2026-10-01',1,{16:'Worse'})]),m=model();c.SUPPORT_APP_calendarSupport_(m);assert.equal(m.summary.find(r=>r.label==='Good days').value,'0');assert.equal(m.summary.find(r=>r.label==='Support Days').value,'1');
});
test('attendance excludes cancelled and future appointments and deduplicates submitted notes',()=>{
 const c=setup([]),records=[{day:'2026-10-01',person:'rebecca',type:'appointment',show:true},{day:'2026-10-01',person:'rebecca',type:'missed',show:true},{day:'2026-10-02',person:'physio',type:'appointment',show:true},{day:'2026-10-31',person:'dentist',type:'appointment',show:true}];
 assert.equal(c.SUPPORT_APP_attendanceCount_('2026-10',[{day:'2026-10-02',appointment:'Physio appointment'}],records,[],'2026-10-28'),1);
});
