// Calendar reader v2: self-contained labels, Shift Notes, summary and schedule reads.
// Replace the existing GoogleCalendar file; do not add a second copy.
// Read-only Google Calendar endpoint. Does not acquire workbook locks or edit Sheets.
function SUPPORT_APP_calendarMonth_(monthKey) {
  if(!SUPPORT_WEB_REQUEST_)throw Error('Signed website access required.');
  const zone='Australia/Melbourne';
  const key=String(monthKey||Utilities.formatDate(new Date(),zone,'yyyy-MM'));
  if(!/^\d{4}-(0[1-9]|1[0-2])$/.test(key)||Number(key.slice(0,4))<1900||Number(key.slice(0,4))>2200)throw Error('Choose a valid month.');
  const year=Number(key.slice(0,4)),month=Number(key.slice(5));
  const next=new Date(Date.UTC(year,month,1)).toISOString().slice(0,7);
  const start=Utilities.parseDate(key+'-01 00:00',zone,'yyyy-MM-dd HH:mm');
  const end=Utilities.parseDate(next+'-01 00:00',zone,'yyyy-MM-dd HH:mm');
  const id=typeof NDIA_APP!=='undefined'&&typeof NDIA_APP.SUPPORT_WEB_calendarId_==='function'?NDIA_APP.SUPPORT_WEB_calendarId_():'c3196a78244332c295765696786facf5e93029a4eddd9bb158fddb15f879d6b0@group.calendar.google.com';
  const events=[];let pageToken;
  do {
    const options={timeMin:start.toISOString(),timeMax:end.toISOString(),singleEvents:true,orderBy:'startTime',showDeleted:false,maxResults:2500,timeZone:zone};
    if(pageToken)options.pageToken=pageToken;
    const page=Calendar.Events.list(id,options);
    events.push.apply(events,page.items||[]);pageToken=page.nextPageToken;
  }while(pageToken);
  const model=SUPPORT_APP_googleCalendarModel_(key,events,Utilities.formatDate(new Date(),zone,'yyyy-MM-dd'));
  SUPPORT_APP_calendarSupport_(model);
  model.supportNotesLoaded=true;model.readerVersion='support-calendar-v2';
  model.canEditNotes=SUPPORT_WEB_REQUEST_.email==='georgia.cranko@gmail.com';
  model.loadedAt=new Date().toISOString();return model;
}

function SUPPORT_APP_googleCalendarModel_(key,events,today) {
  const year=Number(key.slice(0,4)),month=Number(key.slice(5)),zone='Australia/Melbourne';
  const count=new Date(Date.UTC(year,month,0)).getUTCDate(),offset=(new Date(Date.UTC(year,month-1,1)).getUTCDay()+6)%7;
  const first=key+'-01',last=key+'-'+String(count).padStart(2,'0');
  const days=Array.from({length:count},(_,i)=>({number:i+1,lines:[],lineStyles:{},lineMetadata:{},today:key+'-'+String(i+1).padStart(2,'0')===today}));
  const labels=[],editableEvents=[];let eventCount=0,hours=0;
  for(const event of events){
    if(event.status==='cancelled'||!event.start||!event.end)continue;
    const allDay=!!event.start.date;
    const start=allDay?null:new Date(event.start.dateTime),end=allDay?null:new Date(event.end.dateTime);
    if(!allDay&&(!Number.isFinite(start.getTime())||!Number.isFinite(end.getTime())))continue;
    const from=allDay?event.start.date:Utilities.formatDate(start,zone,'yyyy-MM-dd');
    const until=allDay?new Date(Date.parse(event.end.date+'T00:00:00Z')-86400000).toISOString().slice(0,10):Utilities.formatDate(new Date(end.getTime()-1),zone,'yyyy-MM-dd');
    if(until<first||from>last)continue;
    eventCount++;
    if(!allDay){
      const lower=Utilities.parseDate(first+' 00:00',zone,'yyyy-MM-dd HH:mm').getTime();
      const upper=Utilities.parseDate(new Date(Date.UTC(year,month,1)).toISOString().slice(0,10)+' 00:00',zone,'yyyy-MM-dd HH:mm').getTime();
      hours+=Math.max(0,Math.min(end.getTime(),upper)-Math.max(start.getTime(),lower))/3600000;
    }
    const title=String(event.summary||'Untitled event');
    const privateProps=event.extendedProperties&&event.extendedProperties.private||{};
    editableEvents.push({id:event.id,etag:event.etag,title:title,day:from,until:until,allDay:allDay,time:allDay?'':Utilities.formatDate(start,zone,'HH:mm'),hours:allDay?'':Math.round((end-start)/3600000*100)/100,description:String(event.description||''),type:SUPPORT_APP_calendarKind_(title,privateProps)||'Other',canEdit:!!SUPPORT_WEB_REQUEST_&&(SUPPORT_WEB_REQUEST_.email==='georgia.cranko@gmail.com'||privateProps.supportWebCreatedBy===SUPPORT_WEB_REQUEST_.email)});
    const time=allDay?'All day':Utilities.formatDate(start,zone,'HH:mm')+'–'+Utilities.formatDate(end,zone,'HH:mm');
    const line=title+' · '+time;
    const type=SUPPORT_APP_calendarKind_(title,privateProps);
    if(type)labels.push({title:title,type:type,day:from,until:until});
    for(let i=0;i<count;i++){const date=key+'-'+String(i+1).padStart(2,'0');if(date>=from&&date<=until){days[i].lines.push(line);days[i].lineMetadata[line]={eventId:event.id,type:type,support:String(privateProps.support||''),title:title,time:time};}}
  }
  return {source:'google-calendar',monthKey:key,month:['January','February','March','April','May','June','July','August','September','October','November','December'][month-1]+' '+year,
    events:editableEvents,days:Array.from({length:offset},()=>({number:'',lines:[]})).concat(days),labels:labels,
    summary:[{label:'Events',value:String(eventCount)},{label:'Scheduled hours',value:String(Math.round(hours*10)/10)},{label:'Days with events',value:String(days.filter(d=>d.lines.length).length)}]};
}

// Reads source records only; never switches workbook selectors or renders Sheets.
function SUPPORT_APP_calendarSupport_(model){
 const ss=SpreadsheetApp.openById(SUPPORT_APP_SHEET_ID),sheet=ss.getSheetByName('Shift Notes');
 if(!sheet)throw Error('Shift Notes sheet is missing.');
 const rows=sheet.getRange(1,1,Math.max(1,sheet.getLastRow()),30).getValues();
 const dayKey=value=>value instanceof Date&&!isNaN(value.getTime())?Utilities.formatDate(value,'Australia/Melbourne','yyyy-MM-dd'):'';
 const notes=rows.slice(1).map((row,i)=>({row:i+2,day:dayKey(row[2]),worker:String(row[1]||''),hours:row[3],tasks:String(row[5]||''),outing:String(row[13]||''),appointment:String(row[14]||''),detail:String(row[15]||''),mood:String(row[16]||''),swallowing:String(row[17]||''),anxiety:String(row[18]||''),energy:String(row[19]||''),cognition:String(row[20]||''),pain:String(row[21]||''),painLocation:String(row[22]||''),text:String(row[29]||'')})).filter(n=>n.day.slice(0,7)===model.monthKey);
 notes.forEach(n=>n.version=SUPPORT_CALENDAR_noteVersion_(rows[n.row-1]));
 const normalize=v=>String(v||'').trim().toLowerCase();
 const present=v=>!['','none','no','n/a','not applicable'].includes(normalize(v));
 const better=v=>/^(much )?better( than usual)?$/.test(normalize(v));
 const worse=v=>/^(a challenge|a bit worse than usual|(much )?worse( than usual)?)$/.test(normalize(v));
 const supportDays=new Set(),goodCandidates=new Set(),anyWorse=new Set(),worseDays=new Set(),painDays=new Set();
 let hours=0,outings=0,appointments=0;
 for(const note of notes){
  supportDays.add(note.day);const h=Number(note.hours);if(Number.isFinite(h))hours+=h;
  if(present(note.outing))outings++;if(present(note.appointment))appointments++;
  const factors=[note.mood,note.swallowing,note.anxiety,note.energy,note.cognition];
  const w=factors.filter(worse).length;if(w)anyWorse.add(note.day);if(w>=2)worseDays.add(note.day);
  if(factors.filter(better).length>=2&&!w)goodCandidates.add(note.day);
  const day=model.days.find(d=>d.number===Number(note.day.slice(8)));if(!day)continue;
  const labels=SUPPORT_CALENDAR_noteLabels_(note);
  if(labels.pain)painDays.add(note.day);
  day.supportLabels=day.supportLabels||[];
  const add=(text,color)=>{if(!text)return;if(!day.lines.includes(text))day.lines.push(text);if(!day.supportLabels.includes(text))day.supportLabels.push(text);day.supportLabelSources=day.supportLabelSources||{};day.supportLabelSources[text]=day.supportLabelSources[text]||[];if(!day.supportLabelSources[text].some(source=>source.row===note.row))day.supportLabelSources[text].push({worker:note.worker,row:note.row});day.lineStyles[text]={color:color,background:'#ffffff'};};
  if(labels.wellbeing)add(labels.wellbeing.text,labels.wellbeing.colour);
  add(labels.pain,'#D32F2F');
  add(labels.outing,'#9a5fa8');if(!Object.values(day.lineMetadata||{}).some(meta=>/appointment/i.test(meta.type||'')&&SUPPORT_APP_appointmentMatches_(meta.title,note.appointment)))add(labels.appointment,'#123f52');
 }
 let attended='Unavailable',roster=[];
 try{
  const table=SUPPORT_CALENDAR_table_(ss),last=table.sheet.getLastRow();
  const changes=last>table.header?table.sheet.getRange(table.header+1,table.start,last-table.header,table.headers.length).getValues():[];
  const at=(row,name)=>row[table.column(name)];
  const records=changes.map(row=>{const details=String(at(row,'Extra details')||'');return {eventId:String(at(row,'Calendar Event ID')||''),support:String(at(row,'Support')||''),addedBy:String(at(row,'Added by')||''),version:SUPPORT_CALENDAR_noteVersion_(row),details:details,day:dayKey(at(row,'Date / from')),until:dayKey(at(row,'Until')),person:normalize(at(row,'Person / Event')),type:normalize(at(row,'Type')),show:at(row,'Show on calendar?')===true||normalize(at(row,'Show on calendar?'))==='true',originalDay:dayKey(at(row,'Original shift date'))||(details.match(/Original shift date: (\d{4}-\d{2}-\d{2})/)||[])[1]||'',originalPerson:normalize(at(row,'Original person')||(details.match(/Original person: ([^\n]+)/)||[])[1]||at(row,'Person / Event'))};});
  for(const event of model.events||[]){
   const record=records.find(r=>r.eventId&&r.eventId===event.id)||records.find(r=>r.day===event.day&&r.person===normalize(event.title)&&r.show);
   if(!record)continue;
   event.support=record.support;event.scheduleVersion=record.version;
   const label=(record.details.match(/^Website label: (.+)$/m)||[])[1];event.type=label||(record.type==='appointment'?'Appointment':record.type==='activity'?'Activity':event.type);
   if(SUPPORT_WEB_REQUEST_&&(SUPPORT_WEB_REQUEST_.email==='georgia.cranko@gmail.com'||normalize(record.addedBy)===SUPPORT_WEB_REQUEST_.email))event.canEdit=true;
   for(const day of model.days)for(const meta of Object.values(day.lineMetadata||{}))if(meta.eventId===event.id){meta.support=record.support;meta.type=event.type;}
  }
  const regular=ss.getSheetByName('Regular Shifts');
  roster=regular&&regular.getLastRow()>1?regular.getRange(2,1,regular.getLastRow()-1,5).getValues():[];
  attended=SUPPORT_APP_attendanceCount_(model.monthKey,notes,records,roster,Utilities.formatDate(new Date(),'Australia/Melbourne','yyyy-MM-dd'));
 }catch(error){model.attendanceMessage='Appointment attendance could not be calculated: '+error.message;}
 const personKey=value=>{const key=String(value||'').replace(/\s+MISSING\s+NOTE\s*$/i,'').trim().toLowerCase().replace(/\s+/g,' ');return key==='elana'?'ela':key==='craig'?'mr pea':key;};
 const knownWorkers=rows.slice(1).map(row=>String(row[1]||'').trim()).filter(Boolean);
 roster.filter(row=>(row[4]===true||String(row[4]).toLowerCase()==='true')&&!/rebecca|psych/i.test(String(row[2]))).forEach(row=>knownWorkers.push(String(row[2]||'').trim()));
 model.workerNames=Array.from(new Set(knownWorkers.filter(Boolean)));
 const workerKeys=new Set(model.workerNames.map(personKey));
 for(const day of model.days)for(const meta of Object.values(day.lineMetadata||{})){
  const type=String(meta.type||'').toLowerCase();
  const excluded=/appointment|activity|outing|away|missed|cancel/.test(type)||/\baway\b|missing|did not happen|cancelled/i.test(String(meta.title||''))&&!/MISSING NOTE/i.test(String(meta.title||''));
  meta.isShift=!excluded&&(workerKeys.has(personKey(meta.title))||/regular|support|additional|changed|shift/.test(type));
 }
 // Attendance styling requires an explicit matching appointment in a saved note.
 for(const day of model.days)for(const meta of Object.values(day.lineMetadata||{})){
  if(!/appointment/i.test(meta.type||''))continue;
  const date=model.monthKey+'-'+String(day.number).padStart(2,'0');
  meta.attended=notes.some(n=>n.day===date&&SUPPORT_APP_appointmentMatches_(meta.title,n.appointment));
  meta.planned=!meta.attended;
 }
 model.supportNotes=notes;
 model.source='google-calendar-and-support-notes';
 model.summary=[['Support Days',supportDays.size],['Support Hours',Math.round(hours*100)/100],['Good days',Array.from(goodCandidates).filter(day=>!anyWorse.has(day)).length],['Worse days',worseDays.size],['Outings',outings],['Appointments supported',appointments],['Appointments Attended',attended]].map(([label,value])=>({label,value:String(value)}));
}

// Matches the existing summary's past scheduled/submitted attendance convention.
function SUPPORT_APP_attendanceCount_(monthKey,notes,records,roster,today){
 const normalize=value=>String(value||'').trim().toLowerCase();
 const clean=value=>normalize(value).replace(/\s+appointment.*$/i,'');
 const present=value=>!['','none','no','n/a','not applicable'].includes(normalize(value));
 const overrides=new Set(records.filter(r=>r.originalDay).map(r=>r.originalDay+'|'+normalize(r.originalPerson)));
 const count=new Date(Date.UTC(Number(monthKey.slice(0,4)),Number(monthKey.slice(5)),0)).getUTCDate();
 let total=0;
 for(let day=1;day<=count;day++){
  const date=monthKey+'-'+String(day).padStart(2,'0');if(date>today)continue;
  const weekday=(new Date(date+'T12:00:00Z').getUTCDay()+6)%7+1;
  const planned=roster.filter(r=>Number(r[0])===weekday&&(r[4]===true||normalize(r[4])==='true')&&/rebecca|psych/i.test(String(r[2]))&&!overrides.has(date+'|'+normalize(r[2]))).map(r=>normalize(r[2]));
  planned.push(...records.filter(r=>r.day===date&&r.type==='appointment'&&r.show).map(r=>r.person));
  const available=planned.filter(person=>!records.some(r=>r.show&&r.day<=date&&(r.until||r.day)>=date&&/away|missed|cancel|did not happen/.test(r.type)&&(r.person===person||r.person==='georgia')));
  available.push(...notes.filter(n=>n.day===date&&present(n.appointment)).map(n=>n.appointment));
  total+=new Set(available.map(clean).filter(Boolean)).size;
 }
 return total;
}

function SUPPORT_APP_calendarKind_(title,props){
 const type=String(props.supportWebType||props.originalEventType||props.eventType||props.type||'');
 if(/appointment|activity|outing/i.test(type))return type;
 // Same known appointment names used by the existing Regular Shifts renderer.
 if(/rebecca|psych|\bappointment\b/i.test(String(title)))return 'Appointment';
 return type;
}
function SUPPORT_APP_appointmentMatches_(title,recorded){
 const normalize=value=>String(value||'').toLowerCase().replace(/\b(appointment|supported|unsupported|with|the|dr|doctor)\b/g,' ').replace(/[^a-z0-9]+/g,' ').trim();
 const name=normalize(title);if(!name||name.length<3)return false;
 return String(recorded||'').split(/[,;\n]/).some(value=>{const entry=normalize(value);return entry.length>=3&&entry!=='none'&&entry!=='not applicable'&&(entry===name||(' '+entry+' ').includes(' '+name+' ')||(' '+name+' ').includes(' '+entry+' '));});
}

function SUPPORT_CALENDAR_table_(ss){
  const sheet=ss.getSheetByName('Schedule Changes');if(!sheet)throw new Error('Schedule Changes is missing.');
  const start=12,candidates=sheet.getRange(1,start,Math.min(40,sheet.getMaxRows()),3).getDisplayValues();
  const index=candidates.findIndex(r=>r[0]==='Date / from'&&r[1]==='Until'&&r[2]==='Person / Event');
  if(index<0)throw new Error('The saved Schedule Changes table header could not be found. Nothing was saved.');
  const header=index+1,headers=sheet.getRange(header,start,1,23).getDisplayValues()[0];
  const column=name=>{const i=headers.indexOf(name);if(i<0)throw new Error('Required schedule column missing: '+name);return i;};
  return {sheet:sheet,start:start,header:header,headers:headers,column:column};
}

function SUPPORT_CALENDAR_shiftOutingLabel_(value, detailValue) {

  const specificDetail = String(detailValue || '')
    .trim()
    .replace(/[.\s]+$/, '')
    .replace(/^./, function(character) {
      return character.toUpperCase();
    });

  const choices =
    String(
      value || ''
    )
      .split(',')
      .map(
        function(choice) {

          return choice
            .trim()
            .toLowerCase();

        }
      )
      .filter(String);

  const labels = [];

  choices.forEach(
    function(choice) {

      let label = '';

      if (
        choice ===
          'bought groceries' ||
        choice ===
          'ran errands'
      ) {

        label = 'Errands';
      }

      if (
        choice ===
          'saw family and friends' ||
        choice ===
          'social activity' ||
        choice ===
          'did a social activity' ||
        choice ===
          'went to a social activity'
      ) {

        label =
          'Socialised';
      }

      if (
        choice ===
          'community activity' ||
        choice ===
          'community event' ||
        choice ===
          'went to a community event' ||
        choice ===
          'went to a community activity'
      ) {

        label = specificDetail || 'Community';
      }

      if (
        choice ===
        'walked myra'
      ) {

        label = 'Walk';
      }

      if (
        choice ===
          'professional activity' ||
        choice ===
          'work or professional activity'
      ) {

        label = 'Work';
      }

      if (
        label &&
        labels.indexOf(
          label
        ) === -1
      ) {

        labels.push(
          label
        );
      }

    }
  );

  return labels.join(', ');
}

function SUPPORT_CALENDAR_shiftAppointmentLabel_(
  value,
  detailValue,
  supportedTasksValue
) {

  let raw =
    String(
      value || ''
    ).trim();

  const appointmentWasSubmitted =
    /\bappointment\b/i.test(
      String(supportedTasksValue || '')
    );

  const inferredFromDetail = !raw && appointmentWasSubmitted;

  if (inferredFromDetail) {
    raw = String(detailValue || '')
      .trim()
      .replace(/^the\s+/i, '')
      .split(/\s+(?:came|visited|arrived)\b/i)[0]
      .trim()
      .replace(/^./, function(character) {
        return character.toUpperCase();
      });
  }

  if (!raw) return '';

  const choices =
    raw
      .split(',')
      .map(
        function(choice) {

          return choice.trim();

        }
      )
      .filter(String);

  if (
    choices.length > 1
  ) {

    return 'Multiple';
  }

  const lower =
    raw.toLowerCase();

  if (
    lower.indexOf(
      'multiple'
    ) !== -1
  ) {

    return 'Multiple';
  }

  if (
    lower.indexOf(
      'occupational therap'
    ) !== -1 ||
    lower === 'ot'
  ) {

    return 'OT';
  }

  if (
    lower.indexOf(
      'psycholog'
    ) !== -1 ||
    lower === 'psych'
  ) {

    return 'Psych';
  }

  if (
    lower.indexOf(
      'neurophysio'
    ) !== -1
  ) {

    return 'Neurophysio';
  }

  if (
    lower.indexOf(
      'physiotherap'
    ) !== -1 ||
    lower === 'physio'
  ) {

    return 'Physio';
  }

  if (
    lower.indexOf(
      'speech'
    ) !== -1
  ) {

    return 'Speech';
  }

  let custom =
    raw
      .replace(
        /^other\s*:?\s*/i,
        ''
      )
      .trim();

  if (
    !custom ||
    custom.toLowerCase() ===
      'other'
  ) {

    custom =
      String(
        detailValue || ''
      ).trim();
  }

  custom =
    custom
      .split(/[:;]/)[0]
      .trim();

  custom =
    custom
      .replace(
        /\s+appointment.*$/i,
        ''
      )
      .trim();

  if (/\bcrisis support team\b/i.test(custom)) {
    return 'Crisis team';
  }

  return custom;
}

function SUPPORT_CALENDAR_shiftWellbeingStatus_(value) {

  const text =
    String(
      value || ''
    )
      .trim()
      .toLowerCase();

  if (
    text ===
      'a challenge' ||
    text ===
      'much worse' ||
    text ===
      'much worse than usual'
  ) {

    return 'much';
  }

  if (
    text ===
      'worse' ||
    text ===
      'a bit worse than usual' ||
    text ===
      'worse than usual'
  ) {

    return 'worse';
  }

  return '';
}

function SUPPORT_CALENDAR_shiftWellbeingLabel_(values) {

  const factors = [

    {
      key: 'mood',
      better:
        'Better mood',
      worse:
        'Mediocre mood',
      much:
        'Really bad mood'
    },

    {
      key: 'swallowing',
      better:
        'Better swallowing',
      worse:
        'Swallowing difficulty',
      much:
        'Choking risk'
    },

    {
      key: 'anxiety',
      better:
        'Less anxious',
      worse:
        'More anxious',
      much:
        'Extremely overwhelmed'
    },

    {
      key: 'energy',
      better:
        'Better energy',
      worse:
        'Tired',
      much:
        'Exhausted'
    },

    {
      key: 'cognition',
      better:
        'Clearer thinking',
      worse:
        'Brain fog',
      much:
        'Unable to think'
    }

  ];

  const results =
    factors
      .map(
        function(factor) {

          const text =
            String(
              values[
                factor.key
              ] || ''
            )
              .trim()
              .toLowerCase();

          if (
            text ===
              'much better' ||
            text ===
              'much better than usual' ||
            text ===
              'better' ||
            text ===
              'better than usual'
          ) {

            return {

              text:
                factor.better,

              status:
                'better',

              colour:
                '#2E7D32'
            };
          }

          const status =
            SUPPORT_CALENDAR_shiftWellbeingStatus_(
              text
            );

          if (!status) {
            return null;
          }

          return {

            text:
              factor[
                status
              ],

            status:
              status,

            colour:
              status ===
                'much'
                ? '#D32F2F'
                : '#F57C00'
          };

        }
      )
      .filter(
        function(item) {
          return item;
        }
      );

  const worseResults =
    results.filter(
      function(item) {

        return (
          item.status ===
            'worse' ||
          item.status ===
            'much'
        );

      }
    );

  if (
    worseResults.length > 1
  ) {

    const hasMuchWorse =
      worseResults.some(
        function(item) {

          return (
            item.status ===
            'much'
          );

        }
      );

    return {

      text:
        hasMuchWorse
          ? 'Rough day'
          : 'Mediocre day',

      colour:
        hasMuchWorse
          ? '#D32F2F'
          : '#F57C00'
    };
  }

  if (
    worseResults.length === 1
  ) {

    return {

      text:
        worseResults[0]
          .text,

      colour:
        worseResults[0]
          .colour
    };
  }

  const betterResults =
    results.filter(
      function(item) {

        return (
          item.status ===
          'better'
        );

      }
    );

  if (
    betterResults.length > 1
  ) {

    return {

      text:
        'Good day',

      colour:
        '#2E7D32'
    };
  }

  if (
    betterResults.length === 1
  ) {

    return {

      text:
        betterResults[0]
          .text,

      colour:
        '#2E7D32'
    };
  }

  return null;
}

function SUPPORT_CALENDAR_shiftPainLabel_(values) {

  const painText =
    String(
      values.pain || ''
    )
      .trim()
      .toLowerCase();

  const shouldShow =

    painText ===
      'worse' ||

    painText ===
      'much worse' ||

    painText ===
      'much worse than usual' ||

    painText ===
      'a bit worse than usual' ||

    painText ===
      'worse than usual' ||

    painText ===
      'a challenge';

  if (!shouldShow) {
    return '';
  }

  const location =
    SUPPORT_CALENDAR_shiftPainLocation_(
      values.painLocation
    );

  if (!location) {
    return 'Pain';
  }

  if (/^head(?:\s+pain)?$/i.test(location)) {
    return 'Headache';
  }

  return /\bpain\b/i.test(location)
    ? location
    : location + ' pain';
}

function SUPPORT_CALENDAR_shiftPainLocation_(value) {

  let text =
    String(
      value || ''
    ).trim();

  text =
    text.replace(
      /^pain\s*:\s*/i,
      ''
    );

  text =
    text
      .split(
        /\.\s*mood\s*:/i
      )[0]
      .trim();

  return text;
}


function SUPPORT_CALENDAR_noteLabels_(v){
 return {wellbeing:SUPPORT_CALENDAR_shiftWellbeingLabel_(v),pain:SUPPORT_CALENDAR_shiftPainLabel_(v),outing:SUPPORT_CALENDAR_shiftOutingLabel_(v.outing,v.detail),appointment:SUPPORT_CALENDAR_shiftAppointmentLabel_(v.appointment,v.detail,v.tasks)};
}
function SUPPORT_CALENDAR_noteVersion_(row){
 return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,JSON.stringify(row),Utilities.Charset.UTF_8).map(b=>('0'+((b+256)%256).toString(16)).slice(-2)).join('');
}
