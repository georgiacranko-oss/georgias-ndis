import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import vm from 'node:vm';
const code=fs.readFileSync(new URL('../apps-script/GoogleCalendar.gs',import.meta.url),'utf8');
function context(events=[]){
 const options=[];
 const c={SUPPORT_WEB_REQUEST_:{email:'georgia.cranko@gmail.com'},NDIA_APP:{SUPPORT_WEB_calendarId_:()=> 'existing-calendar'},Calendar:{Events:{list(id,o){options.push(o);return {items:events};}}},Utilities:{formatDate(date,zone,pattern){const parts=Object.fromEntries(new Intl.DateTimeFormat('en-GB',{timeZone:zone,year:'numeric',month:'2-digit',day:'2-digit',hour:'2-digit',minute:'2-digit',hourCycle:'h23'}).formatToParts(date).map(p=>[p.type,p.value]));if(pattern==='HH:mm')return parts.hour+':'+parts.minute;const key=parts.year+'-'+parts.month;return pattern==='yyyy-MM'?key:key+'-'+parts.day;},parseDate(value){return new Date(value.replace(' ','T')+':00+11:00');}}};
 vm.createContext(c);vm.runInContext(code,c);c.SUPPORT_APP_calendarSupport_=()=>{};return {c,options};
}
test('calendar endpoint reads Calendar without SpreadsheetApp or locks',()=>{
 const {c,options}=context();const result=c.SUPPORT_APP_calendarMonth_('2026-11');assert.equal(result.monthKey,'2026-11');assert.equal(options.length,1);assert.equal(options[0].singleEvents,true);assert.equal(result.source,'google-calendar');assert.equal(result.days.filter(d=>d.number).length,30);
});
test('all-day end is exclusive and spanning events appear on each correct day',()=>{
 const {c}=context();const result=c.SUPPORT_APP_googleCalendarModel_('2026-11',[{summary:'Away',start:{date:'2026-10-31'},end:{date:'2026-11-03'}}],'2026-11-01');const days=result.days.filter(d=>d.number);assert.equal(days[0].lines.length,1);assert.equal(days[1].lines.length,1);assert.equal(days[2].lines.length,0);assert.equal(days[0].today,true);
});
test('timed event is assigned using Melbourne date and clipped to month',()=>{
 const {c}=context();const result=c.SUPPORT_APP_googleCalendarModel_('2026-11',[{summary:'Support',start:{dateTime:'2026-10-31T23:30:00+11:00'},end:{dateTime:'2026-11-01T01:30:00+11:00'}}],'2026-11-01');assert.equal(result.days.find(d=>d.number===1).lines.length,1);assert.equal(result.summary.find(d=>d.label==='Scheduled hours').value,'1.5');
});
test('rejects unsigned requests and invalid month values',()=>{const {c}=context();c.SUPPORT_WEB_REQUEST_=null;assert.throws(()=>c.SUPPORT_APP_calendarMonth_('2026-11'),/Signed/);c.SUPPORT_WEB_REQUEST_={};assert.throws(()=>c.SUPPORT_APP_calendarMonth_('2026-13'),/valid month/);});

test('Rebecca appointments are classified separately from regular shifts',()=>{const {c}=context();assert.equal(c.SUPPORT_APP_calendarKind_('Rebecca',{type:'Regular'}),'Appointment');assert.equal(c.SUPPORT_APP_calendarKind_('Worker A',{type:'Regular'}),'Regular');});
test('attendance styling requires a named recorded appointment',()=>{const {c}=context();assert.equal(c.SUPPORT_APP_appointmentMatches_('Rebecca','Rebecca appointment'),true);assert.equal(c.SUPPORT_APP_appointmentMatches_('Rebecca',''),false);assert.equal(c.SUPPORT_APP_appointmentMatches_('Rebecca','Physio appointment'),false);});
