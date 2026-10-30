import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import vm from 'node:vm';
const backend=fs.readFileSync(new URL('../apps-script/SupportBackend.gs',import.meta.url),'utf8');
const start=backend.indexOf('function SUPPORT_APP_monthWeeks_');
const end=backend.indexOf('function SUPPORT_APP_validateRawClaim_',start);
const context=vm.createContext({Date,Utilities:{formatDate:date=>date.toISOString().slice(0,10)}});vm.runInContext(backend.slice(start,end),context);
test('weeks cover every day for short, leap and long months without overlap',()=>{
 for(const [month,count] of [['2026-02',28],['2028-02',29],['2026-04',30],['2026-10',31]]){
  const weeks=context.SUPPORT_APP_monthWeeks_(month);assert.equal(weeks.length,Math.ceil(count/7));
  const days=weeks.flatMap(w=>Array.from({length:w.end-w.start+1},(_,i)=>w.start+i));assert.equal(days.length,count);assert.equal(new Set(days).size,count);assert.equal(days[0],1);assert.equal(days.at(-1),count);
 }
});
test('last partial week includes 29th to 31st and does not include next month',()=>{
 const header=Array(17).fill('');const make=(date,hours)=>{const row=header.slice();row[2]=new Date(date+'T12:00:00Z');row[3]=hours;row[7]='Mixed assistance';row[13]='Social activity';row[16]='Usual';return row;};
 const result=context.SUPPORT_APP_weeklyChart_(null,'2026-10',[header,make('2026-10-28',2),make('2026-10-29',3),make('2026-10-31',4),make('2026-11-01',100)]);
 assert.equal(result.length,5);assert.equal(result[3].hours,2);assert.equal(result[4].hours,7);assert.equal(result[4].start,29);assert.equal(result[4].end,31);assert.equal(result[4].values[3],2);assert.equal(result[4].values[4],2);
});
test('funding and functionality both use the same helper and chart spacing supports five groups',()=>{
 assert.match(backend,/const weeks=SUPPORT_APP_monthWeeks_\(monthKey\)/);
 const html=fs.readFileSync(new URL('../public/app.html',import.meta.url),'utf8');assert.ok(html.includes('850/weeks.length'));assert.ok(html.includes("ordinalDate(w.start)+' to '+ordinalDate(w.end)"));
});
