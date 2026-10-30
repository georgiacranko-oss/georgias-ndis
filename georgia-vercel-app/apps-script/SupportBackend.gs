// REPLACEMENT support backend: do not add alongside the previous copy.
// Temporary app handover test mode. Set false after testing.
const SUPPORT_APP_HANDOVER_TEST_MODE = true;
// Add alongside the existing Main file. Do not replace Main.
const SUPPORT_APP_SHEET_ID = '1_wSDD6QeqGaQJvpjc9nBHn-pNe3l4G8msxyZaK9ShA8';
function SUPPORT_APP_page_() {
  return HtmlService.createHtmlOutputFromFile('SupportApp')
    .setTitle("Georgia's NDIS All-in-One App").addMetaTag('viewport', 'width=device-width, initial-scale=1');
}
function SUPPORT_APP_submitInit_(){
  SUPPORT_APP_financeOwner_();
  const ss=SpreadsheetApp.openById(SUPPORT_APP_SHEET_ID);
  const sheet=ss.getSheetByName('Shift Notes');
  if(!sheet)throw Error('Missing source sheet: Shift Notes');
  const lastRow=sheet.getLastRow();
  if(lastRow<1)return {workers:[],handovers:[],loadedAt:new Date().toISOString()};
  const headers=sheet.getRange(1,1,1,sheet.getLastColumn()).getDisplayValues()[0].map(v=>String(v).trim().toLowerCase());
  const workerColumn=headers.indexOf('your name')+1;
  if(!workerColumn)throw Error('Required notes column missing: Your name');
  const workers=lastRow>1?[...new Set(sheet.getRange(2,workerColumn,lastRow-1,1).getDisplayValues().flat().map(v=>String(v||'').trim()).filter(Boolean))].sort((a,b)=>a.localeCompare(b)):[];
  const handoverSheet=ss.getSheetByName('__SUPPORT_APP_HANDOVERS__');
  const handovers=handoverSheet&&handoverSheet.getLastRow()>1?handoverSheet.getRange(Math.max(2,handoverSheet.getLastRow()-4),1,Math.min(5,handoverSheet.getLastRow()-1),5).getValues().map(r=>({date:r[1] instanceof Date?Utilities.formatDate(r[1],'Australia/Melbourne','EEE, d MMM yyyy'):String(r[1]||''),worker:String(r[2]||''),comment:String(r[3]||'')})).filter(r=>r.comment).reverse():[];
  return {workers:workers,handovers:handovers,loadedAt:new Date().toISOString()};
}

function SUPPORT_APP_dataForSelectedSheets_(financeToken,journalMonthKey) {
  SUPPORT_APP_financeOwner_();
  // Deploy as the owner with access restricted to Only myself.
  const financeUnlocked=SUPPORT_APP_financeAllowed_(financeToken);
  const ss=SpreadsheetApp.openById(SUPPORT_APP_SHEET_ID);
  const read=(name,range)=>{const sh=ss.getSheetByName(name);if(!sh)throw Error('Missing source sheet: '+name);return sh.getRange(range).getDisplayValues();};
  const fundingAmounts=read('Plan Expenditure','L34:L38');
  const number=v=>Number(String(v||'').replace(/[^0-9.\-]/g,''))||0;
  const available=[0,2,4].reduce((sum,i)=>sum+number(fundingAmounts[i][0]),0);
  const queueHeadings=read('Plan Expenditure','A43:R43')[0];
  const toPay=queueHeadings.find(v=>/^TO PAY\b/i.test(v))||'Open payment queue in Sheets';
  const toClaim=queueHeadings.find(v=>/^TO CLAIM\b/i.test(v))||'Open claim queue in Sheets';
  const support=read('Support and Funding','J8:AS10');
  const month=read('Support and Funding','B5:E5')[0];
  const fundingRows=read('Plan Expenditure','L33:L38');
  const numeric=v=>Number(String(v||'').replace(/[^0-9.\-]/g,''))||0;
  const coreSpent=numeric(fundingRows[0][0]),coreAvailable=numeric(fundingRows[1][0]),cbSpent=numeric(fundingRows[2][0]),cbAvailable=numeric(fundingRows[3][0]);
  const funding={allocations:read('Plan Expenditure','O7:Q10').filter(r=>r[0]).map(r=>({category:r[0],approved:r[1],spent:r[2]})),coreSpent:fundingRows[0][0],coreAvailable:fundingRows[1][0],cbSpent:fundingRows[2][0],cbAvailable:fundingRows[3][0],corePercent:100*coreSpent/(coreSpent+coreAvailable||1),cbPercent:100*cbSpent/(cbSpent+cbAvailable||1)};
  const claimRows=read('Plan Expenditure','B46:I53');
  const payRows=read('Plan Expenditure','K46:R53');
  const queues={claim:claimRows.filter(r=>r[5]).map(r=>({date:r[0],provider:r[2],invoice:r[5],amount:r[6]})),pay:payRows.filter(r=>r[2]).map(r=>({date:r[0],provider:r[1],invoice:r[2],amount:r[3]}))};
  const storyRows=read('Support and Funding','J13:AD13')[0];
  const stories=[storyRows[0],storyRows[10],storyRows[20]];
  const worthNoting=read('Support and Funding','M18')[0][0];
  const metrics=support[0].map((heading,col)=>heading?{label:heading.trim(),hours:support[1][col],detail:support[2][col]}:null).filter(Boolean);
  const benchmarks=[];
  const handoverSheet=ss.getSheetByName('__SUPPORT_APP_HANDOVERS__');
  const handovers=handoverSheet&&handoverSheet.getLastRow()>1?handoverSheet.getRange(Math.max(2,handoverSheet.getLastRow()-99),1,Math.min(100,handoverSheet.getLastRow()-1),5).getValues().map(r=>({date:dateTextForHandover(r[1]),worker:String(r[2]||''),comment:String(r[3]||'')})).filter(r=>r.comment).reverse():[];
  function dateTextForHandover(v){return v instanceof Date?Utilities.formatDate(v,'Australia/Melbourne','EEE, d MMM yyyy'):String(v||'');}
  const journal=SUPPORT_APP_buildJournal_(ss,journalMonthKey||String(month[3])+'-'+String(['january','february','march','april','may','june','july','august','september','october','november','december'].indexOf(String(month[0]).toLowerCase())+1).padStart(2,'0'));
  const notesSheet=ss.getSheetByName('Shift Notes');
  const notesRows=notesSheet.getRange(1,1,Math.max(1,notesSheet.getLastRow()),30).getValues();
  const headers=notesRows[0].map(x=>String(x).trim().toLowerCase());
  const find=(text)=>{const i=headers.indexOf(text.toLowerCase());if(i<0)throw Error('Required notes column missing: '+text);return i;};
  const nc={worker:find('Your name'),date:find('Date support was provided'),hours:find('How many hours of support did you provide?'),text:find('Generated Paragraph')};
  const dateText=v=>v instanceof Date?Utilities.formatDate(v,'Australia/Melbourne','EEE, d MMM yyyy'):String(v||'');
  const notes=notesRows.slice(1).map((r,i)=>({row:i+2,worker:String(r[nc.worker]||''),date:dateText(r[nc.date]),dateKey:r[nc.date] instanceof Date?Utilities.formatDate(r[nc.date],'Australia/Melbourne','d MMMM yyyy'):String(r[nc.date]||''),activities:String(r[5]||''),hours:String(r[nc.hours]||''),worthNoting:String(r[25]||'').trim(),text:String(r[nc.text]||''),sort:r[nc.date] instanceof Date?r[nc.date].getTime():0})).filter(r=>r.text||r.worker).sort((a,b)=>b.sort-a.sort||b.row-a.row);
  const evidenceSheet=ss.getSheetByName('Funding Evidence Data'),evidenceByRow=new Map();
  if(evidenceSheet&&evidenceSheet.getLastRow()>1){evidenceSheet.getRange(2,1,evidenceSheet.getLastRow()-1,9).getDisplayValues().forEach(r=>{const row=Number(r[0]);if(!Number.isInteger(row)||row<2)return;const services=String(r[8]||'').split(/\s*[·•]\s*/).map(x=>x.trim()).filter(Boolean);if(!evidenceByRow.has(row))evidenceByRow.set(row,new Set());services.forEach(service=>evidenceByRow.get(row).add(service));});}
  notes.forEach(note=>{note.evidenceServices=Array.from(evidenceByRow.get(note.row)||[]);});
  const invSheet=ss.getSheetByName('Financial Claims');
  const invoices=invSheet.getRange(2,1,Math.max(1,invSheet.getLastRow()-1),26).getDisplayValues().map((r,i)=>({row:i+2,date:r[0]||r[19],provider:r[4]||r[17],invoice:r[5]||r[20],url:/^https:\/\//i.test(r[10])?r[10]:'',allocation:r[11]||r[1],support:r[15]||r[2]||'',amount:r[6]||r[21]||'',status:String(r[24]).toLowerCase()==='true'?'Paid':(r[25]||r[7]||''),paid:String(r[24]).toLowerCase()==='true'||String(r[25]).toLowerCase()==='paid',claim:r[12],goals:r[3],claimId:r[13],claimDate:r[14]})).filter(r=>r.invoice||r.provider).reverse();
  const calendarRange=ss.getSheetByName('Calendar').getRange('B6:AP64');
  const grid=calendarRange.getDisplayValues(),calendarFills=calendarRange.getBackgrounds(),calendarInks=calendarRange.getFontColors();
  const calMonth=read('Calendar','C1')[0][0],calYear=read('Calendar','V1')[0][0];
  const today=Utilities.formatDate(new Date(),'Australia/Melbourne','d|MMMM|yyyy').split('|');
  const days=[];
  for(let w=0;w<6;w++)for(let d=0;d<7;d++){
    const row=w*10,col=d*6,number=grid[row]?.[col]||'';
    const lines=[],lineStyles={};
    for(let y=0;y<9;y++){
      const cells=(grid[row+y]||[]).slice(col+1,col+5).filter(Boolean);
      if(cells.length){const line=cells.join(' · ');lines.push(line);const positions=[1,2,3,4].filter(i=>grid[row+y]&&grid[row+y][col+i]);const chosen=positions.find(i=>!/^#(?:ffffff|fff)$/i.test(calendarFills[row+y][col+i]))||positions[0];lineStyles[line]={background:calendarFills[row+y][col+chosen],color:calendarInks[row+y][col+chosen]};}
    }
    days.push({number,lines,lineStyles,today:Number(number)===Number(today[0])&&calMonth.toLowerCase()===today[1].toLowerCase()&&String(calYear)===today[2]});
  }
  const result={fundingAwareness:SUPPORT_APP_fundingAwareness_(ss,days),loadedAt:new Date().toISOString(),financeUnlocked:financeUnlocked,handovers:handovers,journal:journal,home:{available:'$'+available.toLocaleString('en-AU',{minimumFractionDigits:2,maximumFractionDigits:2}),toPay:toPay,toClaim:toClaim,audit:'Check missing notes in the Calendar sheet.'},funding:funding,queues:queues,support:{weekly:[],stories:stories,worthNoting:worthNoting,month:month[0]+' '+month[3],metrics},benchmarks,notes,invoices,calendar:{summary:SUPPORT_APP_calendarSummary_(ss),month:calMonth+' '+calYear,days:days,labels:SUPPORT_APP_calendarLabels_(ss)}};
  result.support.weekly=SUPPORT_APP_weeklyChart_(ss,result.fundingAwareness.weekly.monthKey||journal.monthKey,notesRows);
  if(!financeUnlocked){result.home={};result.funding={};result.queues={};result.invoices=[];result.benchmarks=[];}
  return SUPPORT_APP_invoiceDetails_(ss,result);
}

// App submission endpoint. Keep the web app deployment restricted to Only myself.
function SUPPORT_APP_submit_(request) {
  const effective=String(Session.getEffectiveUser().getEmail()||'').toLowerCase();
  const active=String(Session.getActiveUser().getEmail()||'').toLowerCase();
  if(!SUPPORT_WEB_REQUEST_&&(!active||active!==effective))throw new Error('Submit from the private app while signed into the script owner account.');
  if(!request||!/^[a-zA-Z0-9-]{20,80}$/.test(String(request.id||'')))throw new Error('The submission identifier is invalid. Reload the app.');
  const input=request.fields||{};
  const clean=(key,max)=>{const v=String(input[key]||'').trim();if(v.length>(max||12000))throw new Error('The '+key+' answer is too long.');return v;};
  const worker=clean('worker',200),date=clean('date',10),hours=Number(input.hours),tasks=clean('tasks');
  if(!worker||!tasks)throw new Error('Enter your name and the support provided.');
  if(!clean('notes'))throw new Error('Enter something worth mentioning in the notes before submitting.');
  if(!/^\d{4}-\d{2}-\d{2}$/.test(date)||!Number.isFinite(hours)||hours<=0||hours>24)throw new Error('Enter a valid support date and hours greater than zero and no more than 24.');
  const supportDate=Utilities.parseDate(date+' 00:00','Australia/Melbourne','yyyy-MM-dd HH:mm');
  if(Utilities.formatDate(supportDate,'Australia/Melbourne','yyyy-MM-dd')!==date)throw new Error('The support date is invalid.');
  if(date>Utilities.formatDate(new Date(),'Australia/Melbourne','yyyy-MM-dd'))throw new Error('A support note cannot be submitted for a future shift.');
  const requiredContext={goals:'goals worked towards',difficulties:'why support was necessary',assistance:'the assistance provided',enabled:'what support made possible',risk:'risks without support',later:'concerns after the shift'};
  Object.keys(requiredContext).forEach(function(key){if(!clean(key))throw new Error('Enter an answer for '+requiredContext[key]+'. Select a choice or write your own answer.');});
  const nextVisitorComment=clean('nextVisitor',12000);
  const email=clean('email',254);
  if(email&&!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email))throw new Error('Enter one valid email address for the note copy.');
  if(typeof PROJECT_onFormSubmit!=='function')throw new Error('The existing note processing function PROJECT_onFormSubmit is missing. Nothing was saved.');
  const ss=SpreadsheetApp.openById(SUPPORT_APP_SHEET_ID),sheet=ss.getSheetByName('Shift Notes');
  if(!sheet)throw new Error('The Shift Notes source sheet is missing.');
  const normalize=v=>String(v||'').toLowerCase().replace(/[^a-z0-9]/g,'');
  const lock=LockService.getScriptLock();
  if(!lock.tryLock(10000))throw new Error('Another submission is saving. Your draft is retained. Try again.');
  let row,range,values,headers;
  try {
    headers=sheet.getRange(1,1,1,sheet.getLastColumn()).getDisplayValues()[0];
    const keys=headers.map(normalize);
    const column=label=>{const i=keys.indexOf(normalize(label));if(i<0)throw new Error('The required Shift Notes column is missing: '+label);return i;};
    const schema={worker:'Your name',date:'Date support was provided',hours:'How many hours of support did you provide?',tasks:'What did you support Georgia to do?',goals:'Which silly NDIS goals could this support be seen as directly aiding and abetting?',difficulties:'Which "disability-related" difficulties made support necessary?',assistance:'What was the highest level of assistance needed during this support period?',enabled:'What did the support enable Georgia to do that would otherwise not have happened?',risk:'Without this support, what was likely to be affected?',later:'Were you concerned about any of these things after you left?',outing:'What outings did you support Georgia with?',appointment:'Which appointment was supported?',details:'Please elaborate on any outing or appointment.',painLocation:'If pain was worse, what was mostly hurting',moodReason:'If mood was worse, what seemed the primarily reason',observed:'How do you know the information recorded here?',notes:'Something worth mentioning.:',email:'Email for a copy of the generated note.'};
    const mapped={};Object.keys(schema).forEach(key=>{mapped[key]=column(schema[key]);});
    const timestampCol=column('Timestamp'),consentCol=column('Send this to the next person who is visiting?');
    const wellbeing={mood:'General mood',swallowing:'Swallowing',anxiety:'Anxiety',energy:'Energy',cognition:'Cognition',pain:'Pain'};
    const wellbeingColumns={};Object.keys(wellbeing).forEach(key=>{const suffix=normalize(wellbeing[key]);const i=keys.findIndex(k=>k.indexOf('comparedwith')===0&&k.endsWith(suffix));if(i<0)throw new Error('The wellbeing column is missing: '+wellbeing[key]);wellbeingColumns[key]=i;});
    let idIndex=headers.indexOf('Support App submission ID');
    if(idIndex<0){idIndex=headers.length;if(sheet.getMaxColumns()<idIndex+1)sheet.insertColumnsAfter(sheet.getMaxColumns(),idIndex+1-sheet.getMaxColumns());sheet.getRange(1,idIndex+1).setValue('Support App submission ID');headers.push('Support App submission ID');}
    if(sheet.getLastRow()>1){const existing=sheet.getRange(2,idIndex+1,sheet.getLastRow()-1,1).createTextFinder(String(request.id)).matchEntireCell(true).findNext();if(existing){SUPPORT_APP_storeHandover_(ss,String(request.id),supportDate,worker,nextVisitorComment);return {saved:true,row:existing.getRow(),message:'This submission is already saved. A second record was not created, and email processing was not repeated. Open the source record to check its generated note.'};}}
    const literal=v=>/^[=+\-@]/.test(v)?"'"+v:v;
    values=new Array(headers.length).fill('');
    Object.keys(mapped).forEach(key=>values[mapped[key]]=literal(clean(key,key==='worker'?200:key==='email'?254:12000)));
    values[timestampCol]=new Date();values[mapped.worker]=literal(worker);values[mapped.date]=supportDate;values[mapped.hours]=hours;
    const allowed=['','Better than usual','Usual','Worse than usual','A bit worse than usual','Much worse'];
    Object.keys(wellbeingColumns).forEach(key=>{const v=clean(key,100);if(allowed.indexOf(v)<0)throw new Error('Invalid wellbeing answer for '+key);values[wellbeingColumns[key]]=v;});
    const outingCol=keys.indexOf(normalize('Did the support include an outing or appointment?'));
    if(outingCol>=0)values[outingCol]=(clean('outing')||clean('appointment')||clean('details'))?'Yes':'No';
    values[consentCol]=input.handover===true&&!SUPPORT_APP_HANDOVER_TEST_MODE?'Yes':'No';values[idIndex]=String(request.id);
    sheet.appendRow(values);SpreadsheetApp.flush();
    const found=sheet.getRange(2,idIndex+1,sheet.getLastRow()-1,1).createTextFinder(String(request.id)).matchEntireCell(true).findNext();
    if(!found)throw new Error('The save could not be confirmed. Keep this draft and retry with the same submission identifier.');
    row=found.getRow();range=sheet.getRange(row,1,1,headers.length);
    SUPPORT_APP_storeHandover_(ss,String(request.id),supportDate,worker,nextVisitorComment);
  } finally { lock.releaseLock(); }
  try {
    SUPPORT_APP_queueSubmission_(row,String(request.id),SUPPORT_APP_HANDOVER_TEST_MODE&&input.handover===true);
    return {saved:true,row:row,message:'Your shift note was saved in row '+row+'. Note generation, email and calendar updates are queued in the background. They usually start after about a minute. Refresh the app later to see the generated note.'+(SUPPORT_APP_HANDOVER_TEST_MODE&&input.handover===true?' TEST MODE. The handover will be sent only to Georgia. No worker handover was requested.':'')};
  } catch(error) {
    return {saved:true,row:row,message:'Your shift note was saved in row '+row+', but background processing could not be scheduled. Do not submit again. '+String(error.message||error)};
  }
}

function SUPPORT_APP_storeHandover_(ss,id,date,worker,comment){
  if(!comment)return;
  const name='__SUPPORT_APP_HANDOVERS__';
  let sh=ss.getSheetByName(name);
  if(!sh){sh=ss.insertSheet(name);sh.getRange(1,1,1,5).setValues([['Submitted at','Shift date','Worker','Comment for next person','Submission ID']]);sh.hideSheet();}
  if(sh.getLastRow()>1&&sh.getRange(2,5,sh.getLastRow()-1,1).createTextFinder(id).matchEntireCell(true).findNext())return;
  const literal=v=>/^[=+\-@]/.test(v)?"'"+v:v;
  sh.appendRow([new Date(),date,literal(worker),literal(comment),id]);
}


// Finances is a server-gated section inside the existing owner-private app.
function SUPPORT_APP_financeOwner_(){
  if(SUPPORT_WEB_REQUEST_) return;
  const active=String(Session.getActiveUser().getEmail()||'').toLowerCase();
  const effective=String(Session.getEffectiveUser().getEmail()||'').toLowerCase();
  if(!active||active!==effective)throw new Error('Use the private app while signed into the script owner account.');
}
function SUPPORT_APP_financeChallenge_(){
  SUPPORT_APP_financeOwner_();
  const props=PropertiesService.getScriptProperties(),cache=CacheService.getUserCache();
  const existing=props.getProperty('SUPPORT_APP_FINANCE_SALT');
  if(existing)return {configured:true,salt:existing,iterations:600000};
  let salt=cache.get('SUPPORT_APP_FINANCE_SETUP_SALT');
  if(!salt){salt=Utilities.getUuid()+Utilities.getUuid();cache.put('SUPPORT_APP_FINANCE_SETUP_SALT',salt,300);}
  return {configured:false,salt:salt,iterations:600000};
}
function SUPPORT_APP_financeUnlock_(proof,setup){
  SUPPORT_APP_financeOwner_();
  const props=PropertiesService.getScriptProperties(),cache=CacheService.getUserCache();
  const failures=Number(cache.get('SUPPORT_APP_FINANCE_FAILURES')||0);
  if(failures>=5)throw new Error('Too many incorrect attempts. Try again in 10 minutes.');
  if(!/^[A-Za-z0-9+/]{43}=$/.test(String(proof||'')))throw new Error('The password check is invalid.');
  const lock=LockService.getScriptLock();if(!lock.tryLock(5000))throw new Error('Another update is running. Try again.');
  try{
    let expected=props.getProperty('SUPPORT_APP_FINANCE_VERIFIER'),pepper=props.getProperty('SUPPORT_APP_FINANCE_PEPPER');
    if(!expected){
      if(setup!==true)throw new Error('Set a password first.');
      const salt=cache.get('SUPPORT_APP_FINANCE_SETUP_SALT');if(!salt)throw new Error('Password setup expired. Reload this view.');
      pepper=Utilities.getUuid()+Utilities.getUuid();
      const verifier=Utilities.base64Encode(Utilities.computeHmacSha256Signature(String(proof),pepper));
      props.setProperties({SUPPORT_APP_FINANCE_SALT:salt,SUPPORT_APP_FINANCE_PEPPER:pepper,SUPPORT_APP_FINANCE_VERIFIER:verifier});
      cache.remove('SUPPORT_APP_FINANCE_SETUP_SALT');
    }else{
      if(setup===true)throw new Error('A password has already been set. Reload this view to unlock.');
      const actual=Utilities.base64Encode(Utilities.computeHmacSha256Signature(String(proof),pepper));
      let difference=expected.length^actual.length;for(let i=0;i<Math.max(expected.length,actual.length);i++)difference|=(expected.charCodeAt(i)||0)^(actual.charCodeAt(i)||0);
      if(difference){cache.put('SUPPORT_APP_FINANCE_FAILURES',String(failures+1),600);throw new Error('Incorrect password.');}
    }
    cache.remove('SUPPORT_APP_FINANCE_FAILURES');
    const token=Utilities.getUuid()+Utilities.getUuid();cache.put('SUPPORT_APP_FINANCE_TOKEN_'+token,'unlocked',900);
    return {token:token,expiresIn:900};
  }finally{lock.releaseLock();}
}
function SUPPORT_APP_financeAllowed_(token){
  if(!token||!/^[a-zA-Z0-9-]{60,100}$/.test(String(token)))return false;
  try{SUPPORT_APP_financeOwner_();return CacheService.getUserCache().get('SUPPORT_APP_FINANCE_TOKEN_'+token)==='unlocked';}catch(error){return false;}
}
function SUPPORT_APP_financeKeepAlive_(token){
  SUPPORT_APP_financeOwner_();
  if(!SUPPORT_APP_financeAllowed_(token))throw Error('Finance session expired. Enter your password again.');
  CacheService.getUserCache().put('SUPPORT_APP_FINANCE_TOKEN_'+String(token),'unlocked',900);
  return {expiresIn:900};
}
function SUPPORT_APP_financeLock_(token){
  SUPPORT_APP_financeOwner_();if(token)CacheService.getUserCache().remove('SUPPORT_APP_FINANCE_TOKEN_'+String(token));return {locked:true};
}


function SUPPORT_APP_journal_(monthKey){
  SUPPORT_APP_financeOwner_();
  return SUPPORT_APP_buildJournal_(SpreadsheetApp.openById(SUPPORT_APP_SHEET_ID),monthKey);
}
function SUPPORT_APP_buildJournal_(ss,monthKey){
  if(!/^\d{4}-(0[1-9]|1[0-2])$/.test(String(monthKey||'')))throw new Error('Choose a valid journal month.');
  const year=Number(monthKey.slice(0,4)),month=Number(monthKey.slice(5));
  if(year<1900||year>2200)throw new Error('The journal year is outside the supported range.');
  const names=['January','February','March','April','May','June','July','August','September','October','November','December'];
  const today=Utilities.formatDate(new Date(),'Australia/Melbourne','yyyy-MM-dd');
  const count=new Date(Date.UTC(year,month,0)).getUTCDate();
  const days=Array.from({length:count},(_,i)=>{const key=monthKey+'-'+String(i+1).padStart(2,'0');return {number:i+1,weekday:['S','M','T','W','T','F','S'][new Date(Date.UTC(year,month-1,i+1)).getUTCDay()],today:key===today,future:key>today};});
  const labels=['Good day','Rough day','Outing','Appointment','Social','Pain','Worse mood','Increased fatigue','Swallowing risk','Sheet change'];
  const rows=labels.map(label=>({label:label,marks:new Array(count).fill(false)}));
  const sheet=ss.getSheetByName('Shift Notes');if(!sheet)throw new Error('The Shift Notes source sheet is missing.');
  const records=sheet.getRange(1,1,Math.max(1,sheet.getLastRow()),22).getValues();
  const years=new Set([year,Number(today.slice(0,4))]);
  const worse=/^(worse|a bit worse than usual|worse than usual|much worse|much worse than usual|a challenge)$/;
  const better=/^(better|better than usual|much better|much better than usual)$/;
  records.slice(1).forEach(record=>{
    let date=record[2];
    if(!(date instanceof Date)||isNaN(date.getTime()))return;
    const key=Utilities.formatDate(date,'Australia/Melbourne','yyyy-MM-dd');years.add(Number(key.slice(0,4)));
    if(key.slice(0,7)!==monthKey)return;
    const day=Number(key.slice(8))-1;if(day<0||day>=count)return;
    const normalize=value=>String(value||'').trim().toLowerCase();
    const wellbeing=record.slice(16,21).map(normalize),outing=normalize(record[13]);
    const flags=[wellbeing.filter(v=>better.test(v)).length>=2,wellbeing.filter(v=>worse.test(v)).length>=2,/ran errands|social|community|walked myra|work|professional|grocer|family|friend/.test(outing),String(record[14]||'').length>0||/appointment/.test(outing),/social|community|family|friend/.test(outing),worse.test(normalize(record[21])),worse.test(normalize(record[16])),worse.test(normalize(record[19])),worse.test(normalize(record[17])),/changed sheets|change sheets|changing sheets|bedding/.test(normalize(record[5]))];
    flags.forEach((flag,i)=>{if(flag)rows[i].marks[day]=true;});
  });
  const availableMonths=[];Array.from(years).filter(y=>Number.isFinite(y)&&y>=1900&&y<=2200).sort((a,b)=>b-a).forEach(y=>names.forEach((name,i)=>availableMonths.push({key:y+'-'+String(i+1).padStart(2,'0'),label:name+' '+y})));
  return {month:names[month-1]+' '+year,monthKey:monthKey,days:days,rows:rows,availableMonths:availableMonths};
}


function SUPPORT_APP_calendarTable_(ss){
  const sheet=ss.getSheetByName('Schedule Changes');if(!sheet)throw new Error('Schedule Changes is missing.');
  const start=12,candidates=sheet.getRange(1,start,Math.min(40,sheet.getMaxRows()),3).getDisplayValues();
  const index=candidates.findIndex(r=>r[0]==='Date / from'&&r[1]==='Until'&&r[2]==='Person / Event');
  if(index<0)throw new Error('The saved Schedule Changes table header could not be found. Nothing was saved.');
  const header=index+1,headers=sheet.getRange(header,start,1,23).getDisplayValues()[0];
  const column=name=>{const i=headers.indexOf(name);if(i<0)throw new Error('Required schedule column missing: '+name);return i;};
  return {sheet:sheet,start:start,header:header,headers:headers,column:column};
}
function SUPPORT_APP_calendarLabels_(ss){
  const table=SUPPORT_APP_calendarTable_(ss),sheet=table.sheet;if(sheet.getLastRow()<=table.header)return [];
  const rows=sheet.getRange(table.header+1,table.start,sheet.getLastRow()-table.header,23).getValues();
  const textDate=value=>value instanceof Date&&!isNaN(value.getTime())?Utilities.formatDate(value,'Australia/Melbourne','yyyy-MM-dd'):String(value||'');
  return rows.filter(row=>row[2]&&row[0]&&row[7]!==false&&String(row[7]).toLowerCase()!=='false').map(row=>({title:String(row[2]),type:String(row[3]),day:textDate(row[0]),until:textDate(row[1])||textDate(row[0])}));
}
function SUPPORT_APP_calendarSave_(request){
  SUPPORT_APP_financeOwner_();
  if(!request||!/^[a-zA-Z0-9-]{20,80}$/.test(String(request.id||'')))throw new Error('Invalid event submission identifier.');
  const f=request.fields||{},text=(key,max)=>{const value=String(f[key]||'').trim();if(value.length>(max||200))throw new Error('The '+key+' value is too long.');return value;};
  const type=text('type'),title=text('title'),day=text('date'),until=text('until'),originalDay=text('originalDate');
  if(!title||!['Appointment','Activity','Additional','Changed','Away','Missed','Cancelled'].includes(type))throw new Error('Choose an event label and enter a person or event.');
  const parse=value=>{if(!/^\d{4}-\d{2}-\d{2}$/.test(value))throw new Error('Enter a valid date.');const date=Utilities.parseDate(value+' 00:00','Australia/Melbourne','yyyy-MM-dd HH:mm');if(Utilities.formatDate(date,'Australia/Melbourne','yyyy-MM-dd')!==value)throw new Error('Enter a valid date.');return date;};
  const date=parse(day),end=until?parse(until):'';
  if(until&&until<day)throw new Error('Until cannot be before Date.');
  const allDay=f.allDay===true||(type==='Away'&&until>day),time=allDay?'':text('time'),hours=allDay?'':Number(f.hours);
  if(!allDay&&(!/^(?:[01]\d|2[0-3]):[0-5]\d$/.test(time)||!Number.isFinite(hours)||hours<=0||hours>24))throw new Error('Enter a valid start time and a duration greater than zero and no more than 24 hours.');
  if(!allDay&&until&&until!==day)throw new Error('For several days, select All day.');
  if(type==='Changed'){parse(originalDay);if(!text('originalPerson'))throw new Error('Enter the original worker for a changed shift.');}
  const support=text('support');if(!['','Supported','Unsupported'].includes(support))throw new Error('The support selection is invalid.');
  let details=text('details',12000),originalPerson=text('originalPerson')||title;
  if(type==='Changed')details='Original shift date: '+originalDay+'\nOriginal person: '+originalPerson+(details?'\n'+details:'');
  const originalType=text('originalEventType');if(type==='Cancelled'&&['Appointment','Activity'].includes(originalType))details='Original event type: '+originalType+(details?'\n'+details:'');
  if(f.googleEventId&&(!SUPPORT_WEB_REQUEST_||SUPPORT_WEB_REQUEST_.email!=='georgia.cranko@gmail.com'))throw Error('Owner access required to link an existing Google event.');
  const ss=SpreadsheetApp.openById(SUPPORT_APP_SHEET_ID),table=SUPPORT_APP_calendarTable_(ss),key='sheet|'+request.id;
  const lock=LockService.getScriptLock();if(!lock.tryLock(10000))throw new Error('Another workbook update is running. Try again.');
  let row,alreadySaved=false;
  try{
    const keyIndex=table.column('Calendar Source Key'),last=table.sheet.getLastRow();
    const existing=last>table.header?table.sheet.getRange(table.header+1,table.start+keyIndex,last-table.header,1).createTextFinder(key).matchEntireCell(true).findNext():null;
    if(existing){row=existing.getRow();alreadySaved=true;}
    else{
      const values=new Array(table.headers.length).fill(''),literal=value=>/^[=+\-@]/.test(value)?"'"+value:value;
      const put=(name,value)=>values[table.column(name)]=value;
      put('Date / from',date);put('Until',end);put('Person / Event',literal(title));put('Type',type);put('Support',['Appointment','Activity'].includes(type)?support:'');put('Time',time);put('Hours',hours);put('Show on calendar?',true);put('Extra details',literal(details));put('Added by',SUPPORT_WEB_REQUEST_?SUPPORT_WEB_REQUEST_.email:Session.getActiveUser().getEmail());put('Added on',new Date());put('Original shift date',type==='Changed'?originalDay:'');put('Original person',type==='Changed'?literal(originalPerson):'');put('Sheet Last Edited',new Date().toISOString());put('Calendar Source Key',key);put('All day',allDay);if(f.googleEventId)put('Calendar Event ID',String(f.googleEventId));
      row=table.sheet.getLastRow()+1;if(row>table.sheet.getMaxRows())table.sheet.insertRowsAfter(table.sheet.getMaxRows(),row-table.sheet.getMaxRows());
      table.sheet.getRange(row,table.start+table.column('Time')).setNumberFormat('@');
      table.sheet.getRange(row,table.start,1,values.length).setValues([values]);
      table.sheet.getRange(row,table.start+table.column('Date / from')).setNumberFormat('d mmm yyyy');table.sheet.getRange(row,table.start+table.column('Until')).setNumberFormat('d mmm yyyy');
      SpreadsheetApp.flush();
    }
  }finally{lock.releaseLock();}
  SUPPORT_APP_queueCalendarRefresh_();
  return {saved:true,row:row,message:(alreadySaved?'This event was already saved.':'Event saved in Schedule Changes.')+' Google Calendar sync will follow the existing automation.'};
}

// Called only by Main's existing EMAIL_sendOnce_ before its send-once guard.
function SUPPORT_APP_addHandoverComment_(kind,sheet,row,payload){
  if(kind!=='handover'||!sheet||sheet.getName()!=='Shift Notes')return payload;
  payload=SUPPORT_APP_handoverNoteLink_(row,payload);
  const headers=sheet.getRange(1,1,1,sheet.getLastColumn()).getDisplayValues()[0];
  const idColumn=headers.findIndex(h=>String(h).trim()==='Support App submission ID');
  if(idColumn<0)return payload;
  const id=String(sheet.getRange(row,idColumn+1).getValue()||'');
  const comments=sheet.getParent().getSheetByName('__SUPPORT_APP_HANDOVERS__');
  if(!id||!comments||comments.getLastRow()<2)return payload;
  const found=comments.getRange(2,5,comments.getLastRow()-1,1).createTextFinder(id).matchEntireCell(true).findNext();
  if(!found)return payload;
  const comment=String(comments.getRange(found.getRow(),4).getValue()||'').trim();
  if(!comment)return payload;
  const result=Object.assign({},payload);
  result.body=String(result.body||'')+'\n\nA comment or practical message\n'+comment;
  if(result.htmlBody){const escaped=comment.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/\n/g,'<br>');result.htmlBody+='<section><h3>A comment or practical message</h3><p>'+escaped+'</p></section>';}
  return result;
}

function SUPPORT_APP_nextVisitor_(date,worker){
  SUPPORT_APP_financeOwner_();
  if(!/^\d{4}-\d{2}-\d{2}$/.test(String(date))||!String(worker||'').trim())return {name:'',message:'Enter the shift date and your name to find the next visitor.'};
  if(typeof HANDOVER_findNextVisitor_!=='function')return {name:'',message:'The next visitor lookup is unavailable in Main.'};
  const ss=SpreadsheetApp.openById(SUPPORT_APP_SHEET_ID);
  const day=Utilities.parseDate(date+' 00:00','Australia/Melbourne','yyyy-MM-dd HH:mm');
  const next=HANDOVER_findNextVisitor_(ss,ss.getSheetByName('Regular Shifts'),ss.getSheetByName('Shift Notes'),day,String(worker).trim());
  return next?{name:String(next.name),date:Utilities.formatDate(next.date,'Australia/Melbourne','EEE, d MMM, h:mm a')}:{name:'',message:'No next scheduled visitor was found.'};
}

function SUPPORT_APP_handoverNoteLink_(row,payload){
  const url=ScriptApp.getService().getUrl();if(!url)return payload;
  const link=url+'?noteRow='+encodeURIComponent(String(row));
  const result=Object.assign({},payload);result.body=String(result.body||'')+'\n\nOpen note in app\n'+link;
  if(result.htmlBody)result.htmlBody+='<p><a href="'+link.replace(/&/g,'&amp;').replace(/"/g,'&quot;')+'">Open note in app</a></p>';
  return result;
}

// Shared month buckets: 1–7, 8–14, 15–21, 22–28, then 29–month end.
function SUPPORT_APP_monthWeeks_(monthKey){
 const year=Number(monthKey.slice(0,4)),month=Number(monthKey.slice(5));
 const last=new Date(Date.UTC(year,month,0)).getUTCDate();
 return Array.from({length:Math.ceil(last/7)},(_,i)=>({label:'Week '+(i+1),start:i*7+1,end:Math.min(last,i*7+7)}));
}
function SUPPORT_APP_weeklyChart_(ss,monthKey,notesRows){
 if(!monthKey)return [];
 const rows=notesRows||ss.getSheetByName('Shift Notes').getDataRange().getValues();
 const buckets=SUPPORT_APP_monthWeeks_(monthKey).map(w=>Object.assign(w,{outings:0,social:0,appointments:0,moodSum:0,moodCount:0,needSum:0,needCount:0,totalHours:0}));
 const mood={'much better':4,'much better than usual':4,'better':3,'better than usual':3,'a bit better than usual':3,'usual':2,'worse':1,'worse than usual':1,'a bit worse':1,'a bit worse than usual':1};
 const need={'occasional prompting':1,'step-by-step planning or sequencing':1,'keeping routines and essential things going':2,'mixed assistance':2};
 for(const row of rows.slice(1)){
  const date=row[2];if(!(date instanceof Date)||isNaN(date.getTime()))continue;
  const day=Utilities.formatDate(date,'Australia/Melbourne','yyyy-MM-dd');if(day.slice(0,7)!==monthKey)continue;
  const w=buckets[Math.floor((Number(day.slice(8))-1)/7)];if(!w)continue;
  const outing=String(row[13]||'').toLowerCase(),appointment=String(row[14]||'');
  if(outing)w.outings++;
  if(appointment)w.appointments++;
  if(!appointment)for(const phrase of ['saw family','social activity','professional activity','community'])if(outing.includes(phrase))w.social++;
  const moodValue=String(row[16]||'').toLowerCase();if(moodValue){w.moodCount++;w.moodSum+=mood[moodValue]||0;}
  const needValue=String(row[7]||'').toLowerCase();if(needValue){w.needCount++;w.needSum+=need[needValue]||(/tasks done for geo|^helping physically/.test(needValue)?3:/^taking over parts|^doing most essential things|^doing almost everything/.test(needValue)?4:0);}
  const hours=typeof row[3]==='number'?row[3]:/^\d+(?:\.\d+)?$/.test(String(row[3]).trim())?Number(row[3]):NaN;
  if(Number.isFinite(hours)&&hours>=0)w.totalHours+=hours;
 }
 return buckets.map(w=>({label:w.label,start:w.start,end:w.end,values:[Math.max(0,w.outings-1),w.social,w.appointments,w.moodCount?Math.round(w.moodSum/w.moodCount*10)/10:null,w.needCount?Math.round(w.needSum/w.needCount*10)/10:null],usual:2,hours:w.totalHours?Math.round(w.totalHours*10)/10:null}));
}

function SUPPORT_APP_validateRawClaim_(input){
 const fields=['claimId','claimDate','category','startDate','endDate','amount','status'];
 const out={};fields.forEach(k=>out[k]=String(input[k]??'').trim());
 if(fields.some(k=>!out[k]))throw Error('Complete every claim field.');
 if(!/^\d+$/.test(out.claimId))throw Error('Claim ID must contain digits only.');
 for(const k of ['claimDate','startDate','endDate']){if(!/^\d{4}-\d{2}-\d{2}$/.test(out[k]))throw Error('Enter valid dates.');const d=Utilities.parseDate(out[k]+' 00:00','Australia/Melbourne','yyyy-MM-dd HH:mm');if(Utilities.formatDate(d,'Australia/Melbourne','yyyy-MM-dd')!==out[k])throw Error('Enter valid dates.');out[k+'Value']=d;}
 if(out.endDate<out.startDate)throw Error('Support end date must be on or after its start date.');
 out.amount=Number(out.amount);if(!Number.isFinite(out.amount)||out.amount<=0)throw Error('Enter an amount greater than zero.');
 if(out.category.length>200||out.status.length>100)throw Error('The category or status is too long.');
 return out;
}
function SUPPORT_APP_saveRawClaim_(token,input){
 SUPPORT_APP_financeOwner_();if(!SUPPORT_APP_financeAllowed_(token))throw Error('Unlock finances first.');
 const v=SUPPORT_APP_validateRawClaim_(input||{}),ss=SpreadsheetApp.openById(SUPPORT_APP_SHEET_ID),sh=ss.getSheetByName('Raw NDIS Records');
 if(!sh)throw Error('Raw NDIS Records is missing.');
 const expected=['Claim ID','Claim date','Support category','Support start date','Support end date','Amount','NDIS status'];
 const headers=sh.getRange(1,1,1,7).getDisplayValues()[0];if(expected.some((h,i)=>String(headers[i]).trim()!==h))throw Error('Raw claims headers differ from the expected layout. No record was saved.');
 const lock=LockService.getScriptLock();lock.waitLock(20000);
 try{const count=sh.getLastRow();if(count>1){const ids=sh.getRange(2,1,count-1,1).getValues();if(ids.some(r=>String(r[0]).replace(/\.0$/,'')===v.claimId))return {saved:false,message:'That claim ID is already recorded. No duplicate was added.'};}
 const literal=x=>/^[=+@-]/.test(x)?"'"+x:x;
 const row=count+1;sh.getRange(row,1,1,7).setValues([[v.claimId,v.claimDateValue,literal(v.category),v.startDateValue,v.endDateValue,v.amount,literal(v.status)]]);
 sh.getRange(row,2).setNumberFormat('dd mmm yyyy');sh.getRange(row,4,1,2).setNumberFormat('dd mmm yyyy');sh.getRange(row,6).setNumberFormat('$#,##0.00');
 if(row>2){const formula=sh.getRange(row-1,8).getFormulaR1C1();if(formula)sh.getRange(row,8).setFormulaR1C1(formula);}
 SpreadsheetApp.flush();return {saved:true,message:'Raw claim saved in row '+row+'.',row:row};
 }finally{lock.releaseLock();}
}
function SUPPORT_APP_testAutomation_(){
 SUPPORT_APP_financeOwner_();const ss=SpreadsheetApp.openById(SUPPORT_APP_SHEET_ID),checks=[];
 const add=(name,ok,detail)=>checks.push({name:name,status:ok?'Pass':'Needs attention',detail:detail});
 ['Raw NDIS Records','Financial Claims','Plan Expenditure','Shift Notes','Calendar','Schedule Changes','Regular Shifts','Funding Evidence Data'].forEach(name=>add(name,!!ss.getSheetByName(name),'Required source sheet.'));
 const handlers=['PROJECT_onFormSubmit','PROJECT_onEdit','RUN_PARAGRAPHS','CALENDAR_PROCESS_PENDING','GCAL4_RENDER_ONLY','SC5_SYNC_BUTTON','DAILY_NDIS_INVOICE_CHECK','PROCESS_PENDING_INVOICE_PAYMENTS','REFRESH_CURRENT_INVOICE_QUEUES','HANDOVER_findNextVisitor_','EMAIL_sendOnce_'];
 handlers.forEach(name=>add(name,typeof globalThis[name]==='function','Handler availability only. Not executed.'));
 const triggers=ScriptApp.getProjectTriggers();add('Installed triggers',triggers.length>0,triggers.length+' triggers installed.');
 triggers.forEach(t=>{const name=t.getHandlerFunction();add('Trigger '+name,typeof globalThis[name]==='function','Registered handler resolves. Not executed.');});
 const p=PropertiesService.getScriptProperties();checks.push({name:'Calendar pause setting',status:'Info',detail:p.getProperty('CALENDAR_PAUSED')==='1'?'Calendar automation is paused.':'Calendar automation is not paused.'});
 const sh=ss.getSheetByName('Raw NDIS Records');if(sh){const ids=sh.getLastRow()>1?sh.getRange(2,1,sh.getLastRow()-1,1).getValues().map(r=>String(r[0]||'')).filter(Boolean):[];add('Unique raw claim IDs',new Set(ids).size===ids.length,'Duplicate IDs checked.');}
 checks.push({name:'Live workflow execution',status:'Not run',detail:'Email delivery, Gmail invoice import, calendar exports, payment processing and scheduled executions were not run. This report checks configuration, not complete end to end behaviour.'});
 return {checkedAt:new Date().toISOString(),checks:checks};
}

// Run this from the Apps Script editor. It does not send mail or change records.
function SUPPORT_APP_TEST_ALL_AUTOMATIONS_(){
 const report=SUPPORT_APP_testAutomation_();
 const test=(name,fn)=>{try{fn();report.checks.push({name:name,status:'Pass',detail:'Isolated validation test. No records written.'});}catch(e){report.checks.push({name:name,status:'Needs attention',detail:e.message});}};
 const assert=(ok,message)=>{if(!ok)throw Error(message);};
 const sample={claimId:'123456789',claimDate:'2026-10-26',category:'Daily Activities',startDate:'2026-10-25',endDate:'2026-10-26',amount:'100',status:'Paid'};
 test('Valid raw claim',()=>{const v=SUPPORT_APP_validateRawClaim_(sample);assert(v.amount===100,'Amount did not convert correctly.');});
 test('Reject negative amount',()=>{let rejected=false;try{SUPPORT_APP_validateRawClaim_(Object.assign({},sample,{amount:'-1'}));}catch(e){rejected=true;}assert(rejected,'Negative amount accepted.');});
 test('Reject reversed support dates',()=>{let rejected=false;try{SUPPORT_APP_validateRawClaim_(Object.assign({},sample,{endDate:'2026-10-24'}));}catch(e){rejected=true;}assert(rejected,'Reversed dates accepted.');});
 test('Reject invalid date',()=>{let rejected=false;try{SUPPORT_APP_validateRawClaim_(Object.assign({},sample,{claimDate:'2026-02-31'}));}catch(e){rejected=true;}assert(rejected,'Invalid date accepted.');});
 const ss=SpreadsheetApp.openById(SUPPORT_APP_SHEET_ID),sh=ss.getSheetByName('Financial Claims');
 if(sh){const h=sh.getRange(1,1,1,13).getDisplayValues()[0];report.checks.push({name:'Invoice link and funding columns',status:h[10]==='PDF link'&&h[11]==='Budget category'?'Pass':'Needs attention',detail:'Expected PDF link in K and Budget category in L.'});}
 const lines=['AUTOMATION TEST REPORT',report.checkedAt].concat(report.checks.map(c=>c.status+' | '+c.name+' | '+c.detail));
 const output=lines.join('\n');console.log(output);return output;
}

function SUPPORT_APP_testHandoverToGeorgia_(rowNumber) {
  SUPPORT_APP_financeOwner_();
  return SUPPORT_APP_sendTestHandover_(rowNumber);
}
function SUPPORT_APP_sendTestHandover_(rowNumber) {
  const recipient='georgia.cranko@gmail.com';
  const ss=SpreadsheetApp.openById('1_wSDD6QeqGaQJvpjc9nBHn-pNe3l4G8msxyZaK9ShA8');
  const sh=ss.getSheetByName('Shift Notes');
  if(!sh)throw Error('Shift Notes is missing.');
  const row=Number(rowNumber);
  if(!Number.isInteger(row)||row<2||row>sh.getLastRow())throw Error('Select a saved note in the app first.');
  const headers=sh.getRange(1,1,1,sh.getLastColumn()).getDisplayValues()[0];
  const cells=sh.getRange(row,1,1,headers.length).getValues()[0];
  const normal=v=>String(v||'').toLowerCase().replace(/[^a-z0-9]/g,'');
  const read=label=>{const i=headers.findIndex(h=>normal(h)===normal(label));return i<0?'':cells[i];};
  const note=String(read('Something worth mentioning.:')||read('Something worth mentioning')||'').trim();
  const worker=String(read('Your name')||'Not recorded');
  const date=read('Date support was provided');
  const dateText=date instanceof Date?Utilities.formatDate(date,'Australia/Melbourne','EEE, d MMM yyyy'):String(date||'Not recorded');
  const id=String(read('Support App submission ID')||'');
  let comment='';const comments=ss.getSheetByName('__SUPPORT_APP_HANDOVERS__');
  if(id&&comments&&comments.getLastRow()>1){const match=comments.getRange(2,1,comments.getLastRow()-1,5).getValues().find(r=>String(r[4])===id);if(match)comment=String(match[3]||'').trim();}
  if(!note&&!comment)throw Error('No practical message or worth mentioning text was saved. No email sent.');
  const appUrl=ScriptApp.getService().getUrl();
  const link=appUrl?appUrl+'?app=support&noteRow='+row:'';
  const body=['Hi Georgia, this is your test copy of the handover for the next person to see Georgia and Myra.','',comment?worker+' thought it would be helpful for you to know '+comment:'','',note?'They also thought this was worth noting after visiting on '+dateText+'.':'','',note,link?'\nView shift note\n'+link:'','Test copy sent only to Georgia. No support worker was emailed.'].filter(v=>v!=='').join('\n\n');
  const escape=v=>String(v||'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
  const paragraphHtml=v=>String(v||'').replace(/\r\n?/g,'\n').trim().split(/\n[ \t]*\n+/).filter(Boolean).map(p=>'<p style="margin:0 0 18px">'+escape(p.replace(/\n/g,' '))+'</p>').join('');
  const subject='AUTO-EMAIL: '+worker.replace(/[\r\n]+/g,' ')+' thought this was worth noting';
  const htmlBody='<html><body style="margin:0;background:#ffffff;color:#242424;font-family:Arial,Helvetica,sans-serif"><table role="presentation" width="100%" cellpadding="0" cellspacing="0"><tr><td style="padding:28px 20px"><table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="max-width:680px;margin:0 auto"><tr><td style="font-size:18px;line-height:1.55"><p style="margin:0 0 26px">Hi Georgia, this is your test copy of the handover for the next person to see Georgia and Myra.</p>'+(comment?paragraphHtml(worker+' thought it would be helpful for you to know '+comment):'')+(note?'<p style="margin:24px 0 22px">They also thought this was worth noting after visiting on '+escape(dateText)+'.</p><div style="background:#f7f9f9;border-left:4px solid #206a80;padding:20px 18px 2px;margin:0 0 22px">'+paragraphHtml(note)+'</div>':'')+(link?'<p style="margin:24px 0"><a href="'+escape(link)+'" style="display:inline-block;background:#1976b9;color:#ffffff;font-size:18px;font-weight:bold;text-decoration:none;padding:12px 16px;border-radius:5px">View shift note</a></p>':'')+'<p style="font-size:14px;color:#667777;margin:26px 0 0">Test copy sent only to Georgia. No support worker was emailed.</p></td></tr></table></td></tr></table></body></html>';
  const lock=LockService.getUserLock();lock.waitLock(10000);
  try {
    const props=PropertiesService.getUserProperties();
    const fingerprint=[ss.getId(),row,String(read('Timestamp')),note,comment].join('|');
    const digest=Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,fingerprint));
    const key='GEORGIA_HANDOVER_TEST_'+digest;
    if(props.getProperty(key)){return {message:'Test already attempted for this entry. Check your inbox, spam and Sent folder. Not sent again.'};}
    if(MailApp.getRemainingDailyQuota()<1)throw Error('No email quota remaining. Nothing sent.');
    props.setProperty(key,'attempted '+new Date().toISOString());
    try {MailApp.sendEmail({to:recipient,subject:subject,body:body,htmlBody:htmlBody,name:'Georgia Support Hub'});props.setProperty(key,'sent '+new Date().toISOString());return {message:'Test handover sent only to '+recipient+'. No worker emailed.'};}
    catch(e){console.log('Delivery uncertain. Check inbox and Sent before attempting another test.');throw e;}
  } finally {lock.releaseLock();}
}

// Jobs are durable and claimed once. Failed or interrupted jobs are not resent automatically.
function SUPPORT_APP_queueSubmission_(row,id,testHandover) {
  const lock=LockService.getScriptLock();lock.waitLock(10000);
  try {
    const p=PropertiesService.getScriptProperties(),key='SUPPORT_APP_JOB_'+id;
    if(!p.getProperty(key))p.setProperty(key,JSON.stringify({row:row,id:id,testHandover:!!testHandover,state:'queued',createdAt:new Date().toISOString()}));
    SUPPORT_APP_scheduleJobs_();
  } finally {lock.releaseLock();}
}
function SUPPORT_APP_scheduleJobs_() {
  if(!ScriptApp.getProjectTriggers().some(t=>t.getHandlerFunction()==='SUPPORT_APP_PROCESS_JOBS_'))ScriptApp.newTrigger('SUPPORT_APP_PROCESS_JOBS_').timeBased().after(1000).create();
}
function SUPPORT_APP_PROCESS_JOBS_() {
  const p=PropertiesService.getScriptProperties(),lock=LockService.getScriptLock();let key,job;
  lock.waitLock(10000);
  try {
    ScriptApp.getProjectTriggers().filter(t=>t.getHandlerFunction()==='SUPPORT_APP_PROCESS_JOBS_').forEach(t=>ScriptApp.deleteTrigger(t));
    const all=p.getProperties();
    key=Object.keys(all).filter(k=>k.indexOf('SUPPORT_APP_JOB_')===0).sort().find(k=>{try{return JSON.parse(all[k]).state==='queued';}catch(e){return false;}});
    if(!key)return;
    job=JSON.parse(all[key]);job.state='running';job.startedAt=new Date().toISOString();p.setProperty(key,JSON.stringify(job));
  } finally {lock.releaseLock();}
  try {
    const ss=SpreadsheetApp.openById(SUPPORT_APP_SHEET_ID),sh=ss.getSheetByName('Shift Notes');
    const headers=sh.getRange(1,1,1,sh.getLastColumn()).getDisplayValues()[0];
    const idCol=headers.indexOf('Support App submission ID');
    if(idCol<0)throw Error('Submission ID column is missing.');
    const found=sh.getRange(2,idCol+1,sh.getLastRow()-1,1).createTextFinder(job.id).matchEntireCell(true).findNext();
    if(!found)throw Error('Saved submission was not found.');
    const row=found.getRow(),range=sh.getRange(row,1,1,headers.length),values=range.getValues()[0],namedValues={};
    headers.forEach((h,i)=>{if(h)namedValues[h]=[values[i] instanceof Date?Utilities.formatDate(values[i],'Australia/Melbourne','dd/MM/yyyy'):String(values[i]??'')];});
    let handoverError='';
    if(job.testHandover){try{const result=SUPPORT_APP_sendTestHandover_(row);job.message=result.message;console.log(result.message);}catch(error){handoverError=String(error.message||error);console.error('Test handover failed. '+handoverError);}}
    PROJECT_onFormSubmit({source:ss,range:range,values:values,namedValues:namedValues});SpreadsheetApp.flush();
    if(handoverError)throw Error('Note processing ran, but the test handover failed. '+handoverError);
    job.state='completed';job.finishedAt=new Date().toISOString();p.setProperty(key,JSON.stringify(job));
    console.log('App background processing completed for row '+row+'. Check preceding Main logs for any skipped or caught workflow errors.');
  } catch(error) {job.state='failed';job.error=String(error.message||error);p.setProperty(key,JSON.stringify(job));console.error('App background processing failed. '+job.error);throw error;}
  finally {const remaining=p.getProperties();if(Object.keys(remaining).some(k=>{if(k.indexOf('SUPPORT_APP_JOB_')!==0)return false;try{return JSON.parse(remaining[k]).state==='queued';}catch(e){return false;}}))SUPPORT_APP_scheduleJobs_();}
}
function SUPPORT_APP_BACKGROUND_REPORT_() {
  SUPPORT_APP_financeOwner_();const all=PropertiesService.getScriptProperties().getProperties();
  const jobs=Object.keys(all).filter(k=>k.indexOf('SUPPORT_APP_JOB_')===0).map(k=>JSON.parse(all[k]));
  const report=JSON.stringify(jobs,null,2);console.log(report);return report;
}

function SUPPORT_APP_calendarSummary_(ss){
 const sh=ss.getSheetByName('Calendar');
 const labels=['Support Days','Support Hours','Good days','Worse days','Outings','Appointments supported','Appointments Attended'];
 if(!sh||sh.getMaxColumns()<47)return labels.map(label=>({label:label,value:''}));
 const rows=sh.getRange(1,47,Math.min(16,sh.getMaxRows()),Math.min(14,sh.getMaxColumns()-46)).getDisplayValues();
 const normal=v=>String(v||'').trim().toLowerCase().replace(/\s+/g,' ');
 return labels.map(label=>{const row=rows.find(r=>r.some(c=>normal(c)===normal(label)));if(!row)return {label:label,value:''};const i=row.findIndex(c=>normal(c)===normal(label));return {label:label,value:row.slice(i+1).find(v=>String(v).trim()!=='')||''};});
}

function SUPPORT_APP_fundingAwareness_(ss,calendarDays){
 const sh=ss.getSheetByName('Plan Expenditure');if(!sh)return null;
 const value=a=>{const v=sh.getRange(a).getValue();return typeof v==='number'&&Number.isFinite(v)?v:null;};
 return {weekly:SUPPORT_APP_weeklyFunding_(ss,sh.getRange('F12').getValue(),calendarDays),period:sh.getRange('F7').getDisplayValue(),month:sh.getRange('F12').getDisplayValue(),categories:[
 {label:'Core',remaining:value('L34'),spent:value('F20'),budget:value('H20')},
 {label:'Capacity Building',remaining:value('L36'),spent:value('K20'),budget:value('M20')},
 {label:'Assistive Technology',remaining:value('L38'),spent:null,budget:null}]};
}

function SUPPORT_APP_weeklyFunding_(ss,monthDate,calendarDays){
 if(!(monthDate instanceof Date)||isNaN(monthDate.getTime()))return {weeks:[],message:'The monthly funding selector is not a date.'};
 const monthKey=Utilities.formatDate(monthDate,'Australia/Melbourne','yyyy-MM');
 const year=Number(monthKey.slice(0,4)),month=Number(monthKey.slice(5)),last=new Date(Date.UTC(year,month,0)).getUTCDate();
 const today=Utilities.formatDate(new Date(),'Australia/Melbourne','yyyy-MM-dd');
 const weeks=SUPPORT_APP_monthWeeks_(monthKey).map(w=>Object.assign(w,{core:0,cb:0,hours:0,future:monthKey+'-'+String(w.start).padStart(2,'0')>today}));
 const key=d=>d instanceof Date&&!isNaN(d.getTime())?Utilities.formatDate(d,'Australia/Melbourne','yyyy-MM-dd'):'';
 let excluded=0;
 const invoices=ss.getSheetByName('Financial Claims');
 if(invoices&&invoices.getLastRow()>1)invoices.getRange(2,1,invoices.getLastRow()-1,13).getValues().forEach(r=>{
  const day=key(r[0]);if(day.slice(0,7)!==monthKey||!String(r[5]||'').trim()||/cancel|void/i.test(String(r[7]||'')))return;
  const category=String(r[11]||r[1]||'').trim().toLowerCase();const field=/^(core|flexible core)$/.test(category)?'core':category==='capacity building'?'cb':'';
  if(!field||typeof r[6]!=='number'||!Number.isFinite(r[6])){excluded++;return;}
  weeks[Math.floor((Number(day.slice(8))-1)/7)][field]+=r[6];
 });
 const notes=ss.getSheetByName('Shift Notes');
 if(notes&&notes.getLastRow()>1){const headers=notes.getRange(1,1,1,notes.getLastColumn()).getDisplayValues()[0];const normal=v=>String(v).toLowerCase().replace(/[^a-z0-9]/g,'');const dateCol=headers.findIndex(h=>normal(h)===normal('Date support was provided')),hourCol=headers.findIndex(h=>normal(h)===normal('How many hours of support did you provide?'));
 if(dateCol>=0&&hourCol>=0)notes.getRange(2,1,notes.getLastRow()-1,headers.length).getValues().forEach(r=>{const day=key(r[dateCol]);if(day.slice(0,7)!==monthKey)return;const hours=typeof r[hourCol]==='number'?r[hourCol]:/^\d+(?:\.\d+)?$/.test(String(r[hourCol]).trim())?Number(r[hourCol]):NaN;if(Number.isFinite(hours)&&hours>=0)weeks[Math.floor((Number(day.slice(8))-1)/7)].hours+=hours;});
 else return {weeks:weeks,message:'Support hour columns could not be found. Hours are unavailable.',hoursUnavailable:true};}
 const therapy=SUPPORT_APP_therapyHours_(ss,monthKey,calendarDays,today);
 weeks.forEach(w=>{w.coreHours=w.hours;w.cbHours=0;for(let day=w.start;day<=w.end;day++){if(therapy[day]===null){w.cbHours=null;break;}w.cbHours+=therapy[day]||0;}});
 return {monthKey:monthKey,weeks:weeks,excluded:excluded,message:'Spending uses invoice dates. Hours use support dates and are not split by funding category. Weekly invoice totals may differ from workbook monthly spending.'};
}

function SUPPORT_APP_data_(financeToken,monthKey){
  SUPPORT_APP_financeOwner_();
  if(!monthKey)return SUPPORT_APP_dataForSelectedSheets_(financeToken,monthKey);
  if(!/^\d{4}-(0[1-9]|1[0-2])$/.test(String(monthKey)))throw Error('Choose a valid month.');
  const lock=LockService.getScriptLock();
  if(!lock.tryLock(1000))throw Error('SUPPORT_MONTH_BUSY: Another spreadsheet update is running. Please wait.');
  const ss=SpreadsheetApp.openById(SUPPORT_APP_SHEET_ID),saved=[];let calendarRendered=false;
  const names=['January','February','March','April','May','June','July','August','September','October','November','December'];
  const year=Number(monthKey.slice(0,4)),month=Number(monthKey.slice(5));
  if(year<1900||year>2200){lock.releaseLock();throw Error('Choose a supported year.');}
  try{
    [['Support and Funding','B5',names[month-1]],['Support and Funding','E5',year],['Calendar','C1',names[month-1]],['Calendar','V1',year],['Plan Expenditure','F12',Utilities.parseDate(monthKey+'-01 12:00','Australia/Melbourne','yyyy-MM-dd HH:mm')]].forEach(([name,address,value])=>{
      const sh=ss.getSheetByName(name);if(!sh)throw Error('Missing source sheet: '+name);
      const cell=sh.getRange(address);saved.push({cell:cell,formula:cell.getFormula(),value:cell.getValue()});cell.setValue(value);
    });
    SpreadsheetApp.flush();
    SUPPORT_APP_renderCalendarMonth_(ss);
    calendarRendered=true;
    const calendar=ss.getSheetByName('Calendar');
    const selectedCalendarMonth=String(calendar.getRange('C1').getDisplayValue()).trim();
    const selectedCalendarYear=String(calendar.getRange('V1').getDisplayValue()).trim();
    if(selectedCalendarMonth.toLowerCase()!==names[month-1].toLowerCase()||selectedCalendarYear!==String(year))throw Error('The Calendar sheet did not switch to '+names[month-1]+' '+year+'. Check its month and year selectors in cells C1 and V1.');
    const result=SUPPORT_APP_dataForSelectedSheets_(financeToken,monthKey);result.selectedMonth=monthKey;return result;
  }finally{
    try{
      saved.reverse().forEach(item=>{
        const keepSelectedCalendar=calendarRendered&&item.cell.getSheet().getName()==='Calendar'&&['C1','V1'].includes(item.cell.getA1Notation());
        if(!keepSelectedCalendar)item.formula?item.cell.setFormula(item.formula):item.cell.setValue(item.value);
      });
      SpreadsheetApp.flush();
    }finally{lock.releaseLock();}
  }
}

function SUPPORT_APP_renderCalendarMonth_(ss){
  const calendar=ss.getSheetByName('Calendar');
  if(!calendar)throw Error('Missing source sheet: Calendar');
  if(typeof GCAL4_writeDayGrid_==='function'&&typeof GCAL4_renderStackedShiftCards_==='function'){
    GCAL4_writeDayGrid_(calendar);
    SpreadsheetApp.flush();
    GCAL4_renderStackedShiftCards_();
  }else if(typeof GCAL4_RENDER_ONLY==='function'){
    GCAL4_RENDER_ONLY();
  }else if(typeof NDIA_APP!=='undefined'&&NDIA_APP&&typeof NDIA_APP.GCAL4_RENDER_ONLY==='function'){
    NDIA_APP.GCAL4_RENDER_ONLY();
  }else throw Error('The Calendar renderer is missing. Add GCAL4_RENDER_ONLY to Main.gs, then try again.');
  SpreadsheetApp.flush();
}

function SUPPORT_APP_checkForInvoices_(token){
 SUPPORT_APP_financeOwner_();if(!SUPPORT_APP_financeAllowed_(token))throw Error('Unlock finances first.');
 if(typeof RUN_CHECK_FOR_NEW_INVOICES_BUTTON!=='function')throw Error('The existing DAILY_NDIS_INVOICE_CHECK handler is missing from your Main script.');
 RUN_CHECK_FOR_NEW_INVOICES_BUTTON();
 return {message:'Invoice check completed. Refreshing your records.'};
}

function SUPPORT_APP_invoiceKey_(invoice){return JSON.stringify([invoice.provider,invoice.invoice,invoice.date,invoice.amount].map(v=>String(v||'').trim()));}
function SUPPORT_APP_spendingBySupport_(ss,plan){
 // Count spending only in the current funding year.
 const formula=plan.getRange('L33').getFormula();
 const datePattern=/DATE\(\s*(\d{4})\s*,\s*(\d{1,2})\s*,\s*(\d{1,2})\s*\)/gi,dates=[];let match;while(dates.length<2&&(match=datePattern.exec(formula)))dates.push(match);
 if(dates.length!==2)throw Error('Funding period is unavailable in the expenditure formula.');
 const zone=ss.getSpreadsheetTimeZone(),day=d=>{if(!(d instanceof Date)||isNaN(d))return NaN;const parts=Utilities.formatDate(d,zone,'yyyy-MM-dd').split('-').map(Number);return Date.UTC(parts[0],parts[1]-1,parts[2]);};
 const endYear=Number(dates[1][1]);const start=Date.UTC(Number(dates[0][1]),+dates[0][2]-1,+dates[0][3]),end=Date.UTC(endYear,+dates[1][2]-1,+dates[1][3]);
 const groups={core:new Map(),capacity:new Map()},add=(group,label,amount)=>{const m=groups[group];m.set(label,(m.get(label)||0)+amount);};
 const claims=ss.getSheetByName('Financial Claims');if(!claims)throw Error('Financial Claims is unavailable.');
 if(claims.getLastRow()>1)claims.getRange(2,1,claims.getLastRow()-1,26).getValues().forEach(r=>{
   const category=String(r[11]||'').trim(),label=String(r[15]||r[2]||'').trim(),date=day(r[19]),amount=Number(r[21]),status=String(r[25]||'').trim();
   if(!(date>=start&&date<end)||!['Paid','To Be Paid','To Be Claimed'].includes(status)||!Number.isFinite(amount))return;
   let group=category==='Core'?'core':category==='Capacity Building'?'capacity':null;
   if(!category){if(/Assistance with daily tasks|Professional and Community Engagement|One-off Domestic Task Assistance/i.test(label))group='core';else if(/Psychology|Physiology|Physiotherapy|Neurophysiology|Speech/i.test(label))group='capacity';}
   if(group){const cbSupport=/psycholog|physiotherap|physiology|occupational therap|speech|swallow|therapy/i.test(label);const reallocated=group==='core'&&(/core used for cb/i.test(String(r[9]||''))||/capacity building/i.test(String(r[1]||''))||cbSupport);add(group,(label||'Support not recorded')+(reallocated?' - reallocated':''),amount);}
 });
 const transport=ss.getSheetByName('Transport');
 if(transport&&transport.getLastRow()>1)transport.getRange(2,1,Math.min(157,transport.getLastRow()-1),4).getValues().forEach(r=>{const date=day(r[0]),amount=Number(r[2]);if(date>=start&&date<end&&r[3]==='Paid'&&Number.isFinite(amount))add('core','Transport',amount);});
 // Include planned Capacity Building supports even when no invoice has been recorded.
 const normalSupport=s=>String(s||'').toLowerCase().replace(/neurological\s+physiotherapy/g,'neurophysiotherapy').replace(/[^a-z0-9]/g,'');
 plan.getRange('K68:K72').getDisplayValues().forEach(r=>{const label=String(r[0]||'').split('\n')[0].trim();if(!label)return;const key=normalSupport(label);if(![...groups.capacity.keys()].some(existing=>normalSupport(existing)===key))groups.capacity.set(label,0);});
 const format=n=>{const fixed=(Math.round(n*100)/100).toFixed(2).split('.');return '$'+fixed[0].replace(/\B(?=(\d{3})+(?!\d))/g,',')+'.'+fixed[1];};
 return ['core','capacity'].map(group=>{const items=[...groups[group]].map(([label,amount])=>({label,amount}));
   const warning='';
   return {rows:items.filter(x=>group==='capacity'||Math.abs(x.amount)>=0.005).sort((a,b)=>b.amount-a.amount).map(x=>({label:x.label,amount:format(x.amount)})),warning};
 });
}

function SUPPORT_APP_invoiceDetails_(ss,result){
 if(!result.financeUnlocked)return result;
 const normalizeInvoice=v=>String(v||'').toLowerCase().replace(/\.pdf$/,'').replace(/^invoice[\s_-]*/,'').replace(/[^a-z0-9]/g,'');
 for(const kind of ['pay','claim'])result.queues[kind]=(result.queues[kind]||[]).filter(q=>!result.invoices.some(i=>String(i.provider).trim()===String(q.provider).trim()&&normalizeInvoice(i.invoice)===normalizeInvoice(q.invoice)&&(i.paid||String(i.status).toLowerCase()==='paid')));
 result.invoices.forEach(invoice=>{const claimId=String(invoice.claimId||'').trim();invoice.claimUrl=/^\d+$/.test(claimId)?'https://myplace.ndis.gov.au/ndisstorefront/participant/claims/getPaymentReqestDetails/'+encodeURIComponent(claimId)+'?currentPage=1':'';invoice.recordUrl=ss.getUrl()+'#gid='+ss.getSheetByName('Financial Claims').getSheetId()+'&range=A'+invoice.row+':O'+invoice.row;});
 const plan=ss.getSheetByName('Plan Expenditure'),v=a=>plan.getRange(a).getDisplayValue(),breakdown=plan.getRange('U7:X15').getDisplayValues();
 result.planExpenditure={funding:v('F6'),period:v('F7'),month:v('F12'),core:{available:v('F9'),spent:v('F10'),remaining:v('F16'),monthSpent:v('F20'),budget:v('H20')},capacity:{available:v('K9'),spent:v('K10'),remaining:v('K16'),monthSpent:v('K20'),budget:v('M20')},coreBreakdown:breakdown.filter(r=>r[0]&&!/amount remaining/i.test(r[0])).map(r=>({label:r[0],amount:r[1]})),capacityBreakdown:breakdown.filter(r=>r[2]&&!/amount remaining/i.test(r[2])).map(r=>({label:r[2],amount:r[3]}))};
 try{const grouped=SUPPORT_APP_spendingBySupport_(ss,plan);result.planExpenditure.coreBreakdown=grouped[0].rows;result.planExpenditure.capacityBreakdown=grouped[1].rows;result.planExpenditure.coreBreakdownWarning=grouped[0].warning;result.planExpenditure.capacityBreakdownWarning=grouped[1].warning;}catch(e){result.planExpenditure.coreBreakdown=[];result.planExpenditure.capacityBreakdown=[];result.planExpenditure.breakdownError='Spending by support could not be verified against the sheet totals.';}
 const needs=plan.getRange('D68:O72').getDisplayValues();
 result.planExpenditure.needs=[{label:'Core',estimatedHours:v('F73'),estimatedCost:v('H73'),allocatedHours:v('F75'),allocatedCost:v('D75'),usedHours:v('I75'),usedCost:v('G75'),supports:needs.filter(r=>r[0]).map(r=>({support:r[0],goal:r[2],risk:r[4]}))},{label:'Capacity Building',estimatedHours:v('M73'),estimatedCost:v('O73'),allocatedHours:v('M75'),allocatedCost:v('K75'),usedHours:v('P75'),usedCost:v('N75'),supports:needs.filter(r=>r[7]).map(r=>({support:r[7],goal:r[9],risk:r[11]}))}];
 const history=plan.getRange('AI99:BA104').getValues(),number=v=>typeof v==='number'&&Number.isFinite(v)?v:0;
 result.planExpenditure.history=[{label:'Core expenditure',max:450000,years:history.filter(r=>r[0]).map(r=>({year:String(r[0]),spent:number(r[1]),remaining:number(r[2]),overspent:number(r[3]),rollover:number(r[16]),deficit:number(r[12])}))},{label:'CB expenditure',max:30000,years:history.filter(r=>r[6]).map(r=>({year:String(r[6]),spent:number(r[7]),remaining:number(r[8]),overspent:number(r[9]),rollover:number(r[14]),deficit:number(r[15])}))}];
 return result;
}

function SUPPORT_APP_setInvoicePaid_(token,input){
 SUPPORT_APP_financeOwner_();if(!SUPPORT_APP_financeAllowed_(token))throw Error('Unlock finances first.');
 const row=Number(input?.row);if(!Number.isInteger(row)||row<2||typeof input.paid!=='boolean')throw Error('Choose a valid invoice.');
 const lock=LockService.getScriptLock();lock.waitLock(20000);try{
  const ss=SpreadsheetApp.openById(SUPPORT_APP_SHEET_ID),sh=ss.getSheetByName('Financial Claims');if(row>sh.getLastRow())throw Error('The invoice moved. Refresh and try again.');
  const r=sh.getRange(row,1,1,26).getDisplayValues()[0];const key=SUPPORT_APP_invoiceKey_({provider:r[4]||r[17],invoice:r[5]||r[20],date:r[0]||r[19],amount:r[6]||r[21]||''});if(key!==input.key)throw Error('The invoice changed. Refresh and try again.');
  if(sh.getRange('Y1').getDisplayValue()!=='✓ Paid')throw Error('The paid column changed. No record was updated.');
  const cell=sh.getRange(row,25),rule=cell.getDataValidation();if(!rule||rule.getCriteriaType()!==SpreadsheetApp.DataValidationCriteria.CHECKBOX)throw Error('This invoice has no paid checkbox. No record was updated.');
  if(!input.paid&&String(r[25]).trim().toLowerCase()==='paid'){if(sh.getRange('H1').getDisplayValue()!=='Status')throw Error('The invoice status column changed. No record was updated.');sh.getRange(row,8).setValue('To Be Paid');}
  cell.setValue(input.paid);SpreadsheetApp.flush();return {saved:true,paid:input.paid};
 }finally{lock.releaseLock();}
}
function SUPPORT_APP_therapyHours_(ss,monthKey,days,today){
 const normal=v=>String(v||'').toLowerCase().trim();
 const therapy=v=>/rebecca|\bot\b|occupational therap|physiotherap|speech path|speech therap|psycholog|\bfca\b/.test(normal(v));
 const name=v=>{const s=normal(v);return /rebecca/.test(s)?'rebecca':/\bot\b|occupational therap|\bfca\b/.test(s)?'ot':s.split(/[·,]/)[0].replace(/\s+\d.*$/,'').trim();};
 const date=v=>v instanceof Date&&!isNaN(v)?Utilities.formatDate(v,'Australia/Melbourne','yyyy-MM-dd'):String(v||'').slice(0,10);
 const duration=v=>v!==''&&v!==null&&v!==undefined&&Number.isFinite(Number(v))&&Number(v)>0?Number(v):null;
 const regular=ss.getSheetByName('Regular Shifts'),defaults={};
 if(regular&&regular.getLastRow()>1)regular.getRange(2,1,regular.getLastRow()-1,5).getValues().forEach(r=>{if(r[4]!==false&&therapy(r[2]))defaults[name(r[2])]=duration(r[3]);});
 const table=SUPPORT_APP_calendarTable_(ss),rows=table.sheet.getLastRow()>table.header?table.sheet.getRange(table.header+1,table.start,table.sheet.getLastRow()-table.header,23).getValues():[];
 const shown=rows.filter(r=>r[7]!==false&&String(r[7]).toLowerCase()!=='false');
 const result={};
 (days||[]).forEach(day=>{const n=Number(day.number);if(!n)return;const key=monthKey+'-'+String(n).padStart(2,'0');if(key>today){result[n]=0;return;}
 const blocked=person=>shown.some(r=>date(r[0])<=key&&(date(r[1])||date(r[0]))>=key&&/away|missed|cancel|did not happen/.test(normal(r[3]))&&(name(r[2])===person||normal(r[2])==='georgia'));
 const seen=new Set();let total=0,missing=false;
 (day.lines||[]).forEach(line=>{if(!therapy(line)||/away|missed|cancel|did not happen|not attend/.test(normal(line)))return;const person=name(line);if(seen.has(person)||blocked(person))return;seen.add(person);
 const event=shown.find(r=>date(r[0])===key&&name(r[2])===person&&normal(r[3])==='appointment');
 let hours=event?duration(event[6]):null;if(hours===null)hours=defaults[person]??null;
 if(hours===null){missing=true;return;}total+=hours;
 });result[n]=missing?null:total;
 });return result;
}

// Render schedule views in the background so a website save does not wait for them.
function SUPPORT_APP_queueCalendarRefresh_(){
 const properties=PropertiesService.getScriptProperties();properties.setProperty('SUPPORT_WEB_CALENDAR_PENDING',Utilities.getUuid());
 if(!ScriptApp.getProjectTriggers().some(t=>t.getHandlerFunction()==='SUPPORT_APP_processCalendarChanges_'))ScriptApp.newTrigger('SUPPORT_APP_processCalendarChanges_').timeBased().after(1000).create();
}
function SUPPORT_APP_processCalendarChanges_(){
 const properties=PropertiesService.getScriptProperties(),version=properties.getProperty('SUPPORT_WEB_CALENDAR_PENDING');
 ScriptApp.getProjectTriggers().filter(t=>t.getHandlerFunction()==='SUPPORT_APP_processCalendarChanges_').forEach(t=>ScriptApp.deleteTrigger(t));
 if(!version)return;
 NDIA_APP.SC5_REBUILD();NDIA_APP.GCAL4_RENDER_ONLY();
 if(properties.getProperty('SUPPORT_WEB_CALENDAR_PENDING')===version)properties.deleteProperty('SUPPORT_WEB_CALENDAR_PENDING');
 else SUPPORT_APP_queueCalendarRefresh_();
}
