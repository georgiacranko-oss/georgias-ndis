import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import vm from 'node:vm';
test('Outings remains a summary total while Pain replaces Outing in the key',()=>{
 const html=fs.readFileSync(new URL('../public/app.html',import.meta.url),'utf8');
 const source=html.split('\n').find(line=>line.startsWith('function calendarSummaryPanel('));
 const c=vm.createContext({data:{calendar:{summary:[{label:'Outings',value:'14'}]}},esc:value=>String(value)});vm.runInContext(source,c);const result=c.calendarSummaryPanel();
 assert.match(result,/>Outings<\/span><strong>14<\/strong>/);
 assert.doesNotMatch(result,/>Pain<\/span>/);
 const key=result.slice(result.indexOf('class="calendar-colour-key"'));
 assert.match(key,/border-bottom:2px solid #df3037;?">Pain<\/div>/);
 assert.doesNotMatch(key,/>Outing<\/div>/);
});
