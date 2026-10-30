import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import vm from 'node:vm';
const html=fs.readFileSync(new URL('../public/app.html',import.meta.url),'utf8');
const source=html.slice(html.indexOf('function loadAppMonth(key)'),html.indexOf('function journalView()'));
function setup(errors){
 const queued=[],alerts=[],calls=[],status={textContent:''};
 const context={invoiceHistoryMode:false,view:'support',journalMonthKey:'2026-10',journalRequest:0,journalLoading:false,data:{journal:{monthKey:'2026-10'}},exampleMode:false,financeToken:'',noteDay:'',selected:null,supportMonthControl(){},render(){},applyAppMonth:(result,key)=>({...result,selectedMonth:key}),document:{getElementById:()=>status},setTimeout:fn=>queued.push(fn),alert:m=>alerts.push(m)};
 const runner={withSuccessHandler(fn){this.success=fn;return this;},withFailureHandler(fn){this.failure=fn;return this;},SUPPORT_APP_data(token,key){calls.push(key);const error=errors.shift();if(error)this.failure({message:error});else this.success({journal:{monthKey:key}});}};
 context.window={google:{script:{run:runner}}};context.google=context.window.google;
 vm.createContext(context);vm.runInContext(source,context);return {context,queued,alerts,calls,status};
}
test('busy month read retries and then renders the selected month',()=>{
 const s=setup(['SUPPORT_MONTH_BUSY: busy']);s.context.loadAppMonth('2026-11');
 assert.equal(s.context.journalLoading,true);assert.match(s.status.textContent,/Standby caller/);assert.equal(s.calls.length,1);
 s.queued.shift()();assert.equal(s.calls.length,2);assert.equal(s.context.data.selectedMonth,'2026-11');assert.equal(s.context.journalLoading,false);assert.equal(s.alerts.length,0);
});
test('retries are bounded and restore the previous month if still busy',()=>{
 const s=setup(Array(4).fill('SUPPORT_MONTH_BUSY: busy'));s.context.loadAppMonth('2026-11');while(s.queued.length)s.queued.shift()();
 assert.equal(s.calls.length,4);assert.equal(s.context.journalMonthKey,'2026-10');assert.equal(s.context.journalLoading,false);assert.equal(s.alerts.length,1);
});
test('ordinary errors do not retry',()=>{const s=setup(['Invalid signature.']);s.context.loadAppMonth('2026-11');assert.equal(s.queued.length,0);assert.equal(s.alerts.length,1);});
test('stale retry does not overwrite a newer request',()=>{const s=setup(['SUPPORT_MONTH_BUSY: busy']);s.context.loadAppMonth('2026-11');s.context.journalRequest++;s.queued.shift()();assert.equal(s.calls.length,1);});
test('inline browser scripts still parse',()=>{for(const script of html.matchAll(/<script\b([^>]*)>([\s\S]*?)<\/script>/gi)){if(script[2].trim())new vm.Script(script[2]);}});
