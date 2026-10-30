import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import vm from 'node:vm';
const html=fs.readFileSync(new URL('../public/app.html',import.meta.url),'utf8');
test('month selector safely returns before initial data has loaded',()=>{
 const fn=html.split('\n').find(line=>line.startsWith('function supportMonthControl('));
 const c={data:null,view:'calendar',document:{getElementById:()=>null}};vm.createContext(c);vm.runInContext(fn,c);assert.doesNotThrow(()=>c.supportMonthControl());
});
test('startup message uses requested text and banner has no connection button',()=>{
 const script=[...html.matchAll(/<script\b[^>]*>([\s\S]*?)<\/script>/gi)].map(m=>m[1]).find(s=>s.includes('function display(message,line)'));
 const elements=Object.fromEntries(['banner','banner-title','banner-message','connection'].map(id=>[id,{textContent:'',hidden:true}]));
 const c={document:{readyState:'complete',getElementById:id=>elements[id]},window:{addEventListener(){}},console:{error(){}}};vm.createContext(c);vm.runInContext(script,c);c.window.SUPPORT_APP_showError('Technical detail',12);
 assert.equal(elements['banner-title'].textContent,'Making solid connections is hard. Patience is needed.');assert.equal(elements['banner-message'].textContent,'');assert.equal(elements.banner.hidden,false);assert.doesNotMatch(html,/>Connect workbook<\/button>/);
});
