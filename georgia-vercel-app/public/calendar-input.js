function calendarEventInputs(day){
 const date=calendarDateKey(Number(day.number));
 const events=(data?.calendar?.events||[]).filter(e=>date>=e.day&&date<=e.until);
 return `<details class="panel" style="margin-top:16px"><summary>Calendar events (${events.length})</summary>${events.map(e=>`<article style="margin-top:12px"><strong>${esc(e.title)}</strong><p>${e.allDay?'All day':esc(e.time)}</p>${e.canEdit?`<button class="button secondary" onclick="openGoogleCalendarInput('',${esc(JSON.stringify(e.id))})">Edit event</button>`:''}</article>`).join('')}</details>`;
}
function calendarInputDialog(title,body){
 document.getElementById('calendar-day-detail')?.close();
 const dialog=document.getElementById('calendar-entry');
 dialog.innerHTML=`<div class="panel-head"><h2>${esc(title)}</h2><button class="textbutton" type="button" onclick="document.getElementById('calendar-entry').close()">Close</button></div>${body}`;
 if(!dialog.open)dialog.showModal();return dialog;
}
function openGoogleCalendarInput(day,id){
 const item=id?(data?.calendar?.events||[]).find(e=>e.id===id):null;
 if(id&&!item)return;
 if(item&&!item.canEdit)return;
 const draft=item?{...item}:{day:day||melbourneToday(),until:day||melbourneToday(),title:'',type:'Appointment',allDay:false,time:'09:00',hours:1,description:'',requestId:crypto.randomUUID()};
 const types=['Appointment','Non NDIS appointment','Outing','Activity','Other'];
 if(!types.includes(draft.type))draft.type='Other';
 const dialog=calendarInputDialog(item?'Edit calendar event':'Add calendar event',`<p>Saves to Schedule Changes in Sheets, then follows your existing Google Calendar sync.</p><form id="direct-event-form" class="form-grid"><label class="form-field"><span>Title</span><input name="title" required maxlength="200" value="${esc(draft.title)}"></label><label class="form-field"><span>Type</span><select name="type">${types.map(type=>`<option ${draft.type===type?'selected':''}>${esc(type)}</option>`).join('')}</select></label><label class="form-field"><span>Date</span><input name="day" required type="date" value="${esc(draft.day)}"></label><label class="consent"><input name="allDay" type="checkbox" ${draft.allDay?'checked':''}><span>All day</span></label><label class="form-field" id="direct-until"><span>Last day (all-day events)</span><input name="until" type="date" value="${esc(draft.until||draft.day)}"></label><label class="form-field" id="direct-time"><span>Start time</span><input name="time" type="time" value="${esc(draft.time||'09:00')}"></label><label class="form-field" id="direct-hours"><span>Hours</span><input name="hours" type="number" min="0.25" max="24" step="0.25" value="${esc(draft.hours||1)}"></label><label class="form-field wide"><span>Support</span><select name="support"><option value="">Not recorded</option><option ${draft.support==='Supported'?'selected':''}>Supported</option><option ${draft.support==='Unsupported'?'selected':''}>Unsupported</option></select></label><label class="form-field wide"><span>Details</span><textarea name="description" maxlength="12000">${esc(draft.description||'')}</textarea></label><p class="wide" role="status" id="direct-result"></p><button class="button" type="submit">Save event</button></form>`);
 const form=dialog.querySelector('form');
 const update=()=>{const allDay=form.elements.allDay.checked;dialog.querySelector('#direct-until').hidden=!allDay;dialog.querySelector('#direct-time').hidden=allDay;dialog.querySelector('#direct-hours').hidden=allDay;form.elements.time.required=!allDay;form.elements.hours.required=!allDay;};
 form.elements.allDay.onchange=update;update();
 form.onsubmit=e=>{e.preventDefault();if(!form.reportValidity())return;const values=Object.fromEntries(new FormData(form));values.allDay=form.elements.allDay.checked;if(!values.allDay)values.until=values.day;
  saveCalendarInput(form,'SUPPORT_APP_calendarEventSave',{...values,id:draft.id||'',etag:draft.etag||'',scheduleVersion:draft.scheduleVersion||'',requestId:draft.requestId||crypto.randomUUID()});
 };
}
function openSupportNoteInput(row){
 if(!data?.calendar?.canEditNotes)return;
 const note=(data.calendar.supportNotes||[]).find(n=>n.row===row);if(!note)return;
 const choices=['','Usual','Much better than usual','Better than usual','A bit better than usual','A bit worse than usual','Worse than usual','Much worse than usual'];
 const fields=[['mood','Mood'],['swallowing','Swallowing'],['anxiety','Anxiety'],['energy','Energy'],['cognition','Thinking'],['pain','Pain']];
 const dialog=calendarInputDialog('Edit support note',`<p>${esc(note.worker)} · ${esc(note.day)}. Previous values are kept in the private edit history. Updating these fields does not regenerate the paragraph or send email.</p><form class="form-grid"><label class="form-field"><span>Support hours</span><input name="hours" type="number" min="0" max="24" step="any" required value="${esc(note.hours)}"></label><label class="form-field"><span>Outing</span><input name="outing" maxlength="2000" value="${esc(note.outing)}"></label><label class="form-field"><span>Appointment supported</span><input name="appointment" maxlength="2000" value="${esc(note.appointment)}"></label>${fields.map(([key,label])=>{const options=[...new Set([...choices,note[key]||''])];return `<label class="form-field"><span>${label}</span><select name="${key}">${options.map(value=>`<option value="${esc(value)}" ${note[key]===value?'selected':''}>${esc(value||'Not recorded')}</option>`).join('')}</select></label>`;}).join('')}<label class="form-field"><span>Pain location</span><input name="painLocation" maxlength="2000" value="${esc(note.painLocation)}"></label><label class="form-field wide"><span>Support note text</span><textarea name="text" maxlength="20000" rows="9">${esc(note.text)}</textarea></label><p class="wide" role="status" id="direct-result"></p><button class="button" type="submit">Save note changes</button></form>`);
 const form=dialog.querySelector('form');form.onsubmit=e=>{e.preventDefault();if(!form.reportValidity())return;saveCalendarInput(form,'SUPPORT_APP_calendarNoteSave',{...Object.fromEntries(new FormData(form)),row:note.row,version:note.version});};
}
function saveCalendarInput(form,action,input){
 const status=form.querySelector('#direct-result'),fields=[...form.elements];
 fields.forEach(field=>field.disabled=true);status.textContent='Saving…';
 google.script.run.withSuccessHandler(result=>{fields.forEach(field=>field.disabled=false);status.textContent=result.message;if(result.saved){form.querySelector('button[type="submit"]').disabled=true;refresh();}}).withFailureHandler(error=>{fields.forEach(field=>field.disabled=false);status.textContent=(error.message||'The save could not be confirmed.')+' Your entry is retained.';})[action](input);
}
