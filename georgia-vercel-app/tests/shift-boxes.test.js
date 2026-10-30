import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import vm from 'node:vm';
const html=fs.readFileSync(new URL('../public/app.html',import.meta.url),'utf8');
const source=html.slice(html.indexOf('function calendarClockMinutes('),html.indexOf('function calendarDayLines('));
function context(events){const c={data:{calendar:{events}},workerNames:()=>['Worker A','Worker B']};vm.createContext(c);vm.runInContext(source,c);return c;}
test('app groups supported events into one matching shift and keeps unsupported events outside',()=>{
 const events=[{id:'late',time:'17:00',hours:2},{id:'early',time:'08:00',hours:8},{id:'supported',time:'11:00',hours:1},{id:'unsupported',time:'10:00',hours:1}];
 const c=context(events),meta={late:{eventId:'late',type:'Regular',title:'Worker B'},early:{eventId:'early',type:'Regular',title:'Worker A'},supported:{eventId:'supported',type:'Appointment',title:'Physio',support:'Supported'},unsupported:{eventId:'unsupported',type:'Appointment',title:'Dentist',support:'Unsupported'}};
 const result=c.calendarDayLayout({regularLines:['late','supported','unsupported','early'],lineMetadata:meta});
 assert.equal(result.entries.length,3);const boxed=result.entries.filter(entry=>entry.kind==='shift');assert.equal(boxed[0].item.line,'early');assert.equal(boxed[0].item.events[0].line,'supported');assert.equal(result.entries.find(entry=>entry.kind==='event').item.line,'unsupported');assert.equal(boxed[1].item.line,'late');
});
test('overlapping supported events choose the fully containing shift with earliest finish',()=>{
 const c=context([{id:'a',time:'08:00',hours:8},{id:'b',time:'09:00',hours:4},{id:'e',time:'10:00',hours:1}]);
 const result=c.calendarDayLayout({regularLines:['a','b','e'],lineMetadata:{a:{eventId:'a',type:'Regular',title:'Worker A'},b:{eventId:'b',type:'Regular',title:'Worker B'},e:{eventId:'e',type:'Activity',support:'Supported'}}});
 assert.equal(result.entries.filter(e=>e.kind==='event').length,0);assert.equal(result.entries.find(e=>e.item.line==='b').item.events[0].line,'e');
});
test('recorded day labels follow their named worker shift',()=>{
 const c=context([{id:'a',time:'08:00',hours:4},{id:'b',time:'14:00',hours:4}]);
 const result=c.calendarDayLayout({regularLines:['a','b','Rough day'],supportLabels:['Rough day'],supportLabelSources:{'Rough day':[{worker:'Worker B',row:2}]},lineMetadata:{a:{eventId:'a',type:'Regular',title:'Worker A'},b:{eventId:'b',type:'Regular',title:'Worker B'}}});
 assert.equal(result.entries[0].item.labels.length,0);assert.equal(result.entries[1].item.labels[0],'Rough day');
});

test('overlapping shifts sort by finish time rather than their start time',()=>{
 const c=context([{id:'ela',time:'07:45',hours:10},{id:'pea',time:'14:30',hours:2}]);
 const result=c.calendarDayLayout({regularLines:['ela','pea'],lineMetadata:{ela:{eventId:'ela',type:'Regular',title:'Worker A'},pea:{eventId:'pea',type:'Regular',title:'Worker B'}}});
 assert.equal(result.entries[0].item.line,'pea');assert.equal(result.entries[1].item.line,'ela');
});

test('labels from each original note stay together even when shifts share a label',()=>{
 const c=context([{id:'a',time:'07:45',hours:10},{id:'b',time:'14:30',hours:2}]);
 const result=c.calendarDayLayout({regularLines:['a','b','Rough day','Pain'],supportLabels:['Rough day','Pain'],supportLabelSources:{'Rough day':[{row:1,worker:'Worker A'},{row:2,worker:'Worker B'}],Pain:[{row:1,worker:'Worker A'}]},lineMetadata:{a:{eventId:'a',type:'Regular',title:'Worker A'},b:{eventId:'b',type:'Regular',title:'Worker B'}}});
 const a=result.entries.find(e=>e.item.line==='a').item;
 const b=result.entries.find(e=>e.item.line==='b').item;
 assert.equal(a.noteGroups.length,1);assert.equal(a.noteGroups[0].lines.join('|'),'Rough day|Pain');
 assert.equal(b.noteGroups.length,1);assert.equal(b.noteGroups[0].lines.join('|'),'Rough day');
 assert.equal(result.dayNoteGroups.length,0);
});
