import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import vm from 'node:vm';
const html=fs.readFileSync(new URL('../public/app.html',import.meta.url),'utf8');
const start=html.indexOf('function calendarLineMarkup('),end=html.indexOf('\nfunction calendarLineStyle(',start);
const c=vm.createContext({data:null,esc:s=>String(s)});vm.runInContext(html.slice(start,end),c);
test('away labels omit time ranges and all-day suffixes',()=>{assert.equal(c.calendarLineMarkup('Ruby away · 08:00–12:30',7),'Ruby away');assert.equal(c.calendarLineMarkup('Fab away · All day',15),'Fab away');});
test('ordinary support shifts retain their smaller time range',()=>{assert.match(c.calendarLineMarkup('Ruby · 08:00–12:30',7),/calendar-time/);assert.match(c.calendarLineMarkup('Ruby · 08:00–12:30',7),/08:00–12:30/);});
