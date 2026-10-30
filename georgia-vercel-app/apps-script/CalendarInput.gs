function SUPPORT_APP_noteVersion_(row){
 return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,JSON.stringify(row),Utilities.Charset.UTF_8).map(b=>('0'+((b+256)%256).toString(16)).slice(-2)).join('');
}
// Website event input always writes Schedule Changes; existing automation syncs Google.
function SUPPORT_APP_calendarEventSave_(input){
 if(!SUPPORT_WEB_REQUEST_)throw Error('Signed website access required.');
 const email=SUPPORT_WEB_REQUEST_.email,id=String(input.id||''),zone='Australia/Melbourne';
 const title=String(input.title||'').trim(),day=String(input.day||''),until=String(input.until||day);
 const parse=value=>{if(!/^\d{4}-\d{2}-\d{2}$/.test(value))throw Error('Choose a valid date.');const date=Utilities.parseDate(value+' 00:00',zone,'yyyy-MM-dd HH:mm');if(Utilities.formatDate(date,zone,'yyyy-MM-dd')!==value)throw Error('Choose a valid date.');return date;};
 const date=parse(day),last=parse(until);if(until<day)throw Error('End date must not be before start date.');
 if(!title||title.length>200)throw Error('Enter an event title up to 200 characters.');
 const label=String(input.type||'Other');if(!['Appointment','Non NDIS appointment','Outing','Activity','Other'].includes(label))throw Error('Choose an event type.');
 const type=/appointment/i.test(label)?'Appointment':'Activity';
 const support=label==='Non NDIS appointment'?'Unsupported':String(input.support||'');
 if(!['','Supported','Unsupported'].includes(support))throw Error('Choose a support status.');
 const allDay=input.allDay===true,time=String(input.time||''),hours=Number(input.hours);
 if(!allDay&&(!/^(?:[01]\d|2[0-3]):[0-5]\d$/.test(time)||!Number.isFinite(hours)||hours<=0||hours>24))throw Error('Enter a start time and duration up to 24 hours.');
 const description=String(input.description||'');if(description.length>12000)throw Error('Event details are too long.');
 const details='Website label: '+label+(description?'\n'+description:'');
 const fields={title:title,type:type,date:day,until:allDay?until:'',allDay:allDay,time:allDay?'':time,hours:allDay?'':hours,support:support,details:details};
 if(!id)return SUPPORT_APP_calendarSave_({id:input.requestId,fields:fields});
 const existing=Calendar.Events.get(NDIA_APP.SUPPORT_WEB_calendarId_(),id);
 if(!input.etag||existing.etag!==input.etag)throw Error('This event changed. Refresh before editing.');
 const ss=SpreadsheetApp.openById(SUPPORT_APP_SHEET_ID),table=SUPPORT_APP_calendarTable_(ss),lastRow=table.sheet.getLastRow();
 const match=lastRow>table.header?table.sheet.getRange(table.header+1,table.start+table.column('Calendar Event ID'),lastRow-table.header,1).createTextFinder(id).matchEntireCell(true).findNext():null;
 if(!match){
  if(email!=='georgia.cranko@gmail.com')throw Error('Only Georgia can edit an event not yet linked to Schedule Changes.');
  fields.googleEventId=id;return SUPPORT_APP_calendarSave_({id:input.requestId,fields:fields});
 }
 const row=match.getRow(),lock=LockService.getScriptLock();if(!lock.tryLock(1000))throw Error('Another update is running. Your entry is retained. Please try again.');
 try{
  const range=table.sheet.getRange(row,table.start,1,table.headers.length),values=range.getValues()[0];
  if(email!=='georgia.cranko@gmail.com'&&String(values[table.column('Added by')]||'').toLowerCase()!==email)throw Error('Only Georgia or the creator can edit this event.');
  if(!input.scheduleVersion||SUPPORT_APP_noteVersion_(values)!==input.scheduleVersion)throw Error('Schedule Changes was edited. Refresh before editing.');
  const literal=value=>/^[=+\-@]/.test(value)?"'"+value:value;
  const put=(name,value)=>{values[table.column(name)]=value;};
  put('Date / from',date);put('Until',allDay?last:'');put('Person / Event',literal(title));put('Type',type);put('Support',support);put('Time',allDay?'':time);put('Hours',allDay?'':hours);put('All day',allDay);put('Extra details',literal(details));put('Sheet Last Edited',new Date().toISOString());
  range.setValues([values]);SpreadsheetApp.flush();
 }finally{lock.releaseLock();}
 SUPPORT_APP_queueCalendarRefresh_();
 return {saved:true,row:row,message:'Updated in Schedule Changes. Google Calendar will follow the existing sync automation.'};
}
function SUPPORT_APP_calendarNoteSave_(input){
 if(!SUPPORT_WEB_REQUEST_||SUPPORT_WEB_REQUEST_.email!=='georgia.cranko@gmail.com')throw Error('Only Georgia can edit saved support notes.');
 const row=Number(input.row);if(!Number.isInteger(row)||row<2)throw Error('Invalid note record.');
 const text=String(input.text||''),hours=Number(input.hours);
 if(text.length>20000||!Number.isFinite(hours)||hours<0||hours>24)throw Error('Enter valid note text and hours between 0 and 24.');
 const ss=SpreadsheetApp.openById(SUPPORT_APP_SHEET_ID),sheet=ss.getSheetByName('Shift Notes');
 const lock=LockService.getScriptLock();if(!lock.tryLock(1000))throw Error('Another update is running. Your edits are retained. Please try again.');
 try{
  if(row>sheet.getLastRow())throw Error('Note record no longer exists.');
  const range=sheet.getRange(row,1,1,30),original=range.getValues()[0];
  if(!input.version||SUPPORT_APP_noteVersion_(original)!==input.version)throw Error('This note changed or moved. Refresh before editing.');
  const fields={outing:14,appointment:15,mood:17,swallowing:18,anxiety:19,energy:20,cognition:21,pain:22,painLocation:23};
  const permitted=['','Usual','Much better','Much better than usual','Better','Better than usual','A bit better than usual','A bit worse','A bit worse than usual','Worse','Worse than usual','Much worse','Much worse than usual','A challenge'];
  for(const [key,column] of Object.entries(fields)){
   const value=String(input[key]||'');if(value.length>2000)throw Error('Field is too long.');
   if(['mood','swallowing','anxiety','energy','cognition','pain'].includes(key)&&!permitted.includes(value)&&value!==String(original[column-1]||''))throw Error('Invalid wellbeing answer.');
  }
  let audit=ss.getSheetByName('__SUPPORT_WEB_NOTE_EDITS__');
  if(!audit){audit=ss.insertSheet('__SUPPORT_WEB_NOTE_EDITS__');audit.appendRow(['Edited at','Edited by','Source row','Original values']);audit.hideSheet();}
  audit.appendRow([new Date(),SUPPORT_WEB_REQUEST_.email,row,JSON.stringify(original)]);
  const literal=value=>/^[=+\-@]/.test(value)?"'"+value:value;
  sheet.getRange(row,4).setValue(hours);
  for(const [key,column] of Object.entries(fields))sheet.getRange(row,column).setValue(literal(String(input[key]||'')));
  sheet.getRange(row,30).setValue(literal(text));SpreadsheetApp.flush();
  return {saved:true,message:'Support note updated. The previous record was saved in the edit history. No email was sent.'};
 }finally{lock.releaseLock();}
}
