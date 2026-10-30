// Add this file alongside the replacement SupportBackend.gs and existing Main.
// Review Main's public functions before making a deployment accessible to Anyone.
var SUPPORT_WEB_REQUEST_ = null;
function SUPPORT_WEB_requireOwnerOrBridge_() {
  if(SUPPORT_WEB_REQUEST_)return;
  const active=String(Session.getActiveUser().getEmail()||'').toLowerCase();
  if(active!=='georgia.cranko@gmail.com')throw Error('Owner access required.');
}
function doPost(e) {
  const output=value=>ContentService.createTextOutput(JSON.stringify(value)).setMimeType(ContentService.MimeType.JSON);
  try {
    const raw=String(e&&e.postData&&e.postData.contents||'');
    if(raw.length>180000)throw Error('Invalid request.');
    const envelope=JSON.parse(raw), properties=PropertiesService.getScriptProperties();
    const secret=properties.getProperty('SUPPORT_WEB_BRIDGE_SECRET');
    if(!secret||secret.length<32)throw Error('The connection is not configured.');
    if(typeof envelope.payload!=='string'||!/^[a-f0-9]{64}$/.test(envelope.signature||''))throw Error('Invalid request.');
    const signature=Utilities.computeHmacSha256Signature(envelope.payload,secret,Utilities.Charset.UTF_8).map(b=>('0'+((b+256)%256).toString(16)).slice(-2)).join('');
    let difference=0;for(let i=0;i<64;i++)difference|=signature.charCodeAt(i)^envelope.signature.charCodeAt(i);
    if(difference)throw Error('Invalid signature.');
    const request=JSON.parse(envelope.payload);
    if(!Number.isFinite(request.timestamp)||Math.abs(Date.now()-request.timestamp)>60000||!/^[-a-f0-9]{36}$/.test(request.nonce||''))throw Error('Expired request.');
    const members=['georgia.cranko@gmail.com','busyandimportant01@gmail.com','elanagrunberg@protonmail.com','fcschmid95@gmail.com','freya.f.harvey@gmail.com','hannahgwarnesmith@gmail.com','ruby.forbes@outlook.com'];
    if(members.indexOf(request.email)<0)throw Error('Access denied.');
    const actions={
      SUPPORT_APP_calendarEventSave:SUPPORT_APP_calendarEventSave_,
      SUPPORT_APP_calendarNoteSave:SUPPORT_APP_calendarNoteSave_,
      SUPPORT_APP_calendarMonth:SUPPORT_APP_calendarMonth_,
      SUPPORT_APP_calendarSave:SUPPORT_APP_calendarSave_,SUPPORT_APP_checkForInvoices:SUPPORT_APP_checkForInvoices_,
      SUPPORT_APP_data:SUPPORT_APP_data_,SUPPORT_APP_financeChallenge:SUPPORT_APP_financeChallenge_,
      SUPPORT_APP_financeKeepAlive:SUPPORT_APP_financeKeepAlive_,SUPPORT_APP_financeLock:SUPPORT_APP_financeLock_,
      SUPPORT_APP_financeUnlock:SUPPORT_APP_financeUnlock_,SUPPORT_APP_journal:SUPPORT_APP_journal_,
      SUPPORT_APP_nextVisitor:SUPPORT_APP_nextVisitor_,SUPPORT_APP_saveRawClaim:SUPPORT_APP_saveRawClaim_,
      SUPPORT_APP_setInvoicePaid:SUPPORT_APP_setInvoicePaid_,SUPPORT_APP_submit:SUPPORT_APP_submit_,
      SUPPORT_APP_submitInit:SUPPORT_APP_submitInit_,SUPPORT_APP_testAutomation:SUPPORT_APP_testAutomation_
    };
    if(!Object.prototype.hasOwnProperty.call(actions,request.action)||!Array.isArray(request.args)||request.args.length>10)throw Error('Unsupported action.');
    const owner=request.email==='georgia.cranko@gmail.com';
    if(/finance|Invoice|RawClaim|testAutomation|calendarNoteSave/.test(request.action)&&!owner)throw Error('Owner access required.');
    if(request.action==='SUPPORT_APP_data'&&!owner)request.args[0]='';
    // Nonce checks must not queue behind long spreadsheet rebuilds.
    // All bridge requests execute as the deployment owner, so this user lock
    // serializes nonce check/insert independently from the workbook script lock.
    const lock=LockService.getUserLock();
    if(!lock.tryLock(1000))throw Error('SUPPORT_MONTH_BUSY: The connection is busy. Please wait.');
    try{const cache=CacheService.getScriptCache(),key='web-nonce:'+request.nonce;if(cache.get(key))throw Error('Request already received.');cache.put(key,'1',180);}finally{lock.releaseLock();}
    SUPPORT_WEB_REQUEST_={email:request.email};
    const result=actions[request.action].apply(null,request.args);
    return output({ok:true,result:result===undefined?null:result});
  } catch(error) {
    return output({ok:false,error:String(error.message||'Request failed.')});
  } finally { SUPPORT_WEB_REQUEST_=null; }
}
