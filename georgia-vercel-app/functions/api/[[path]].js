const owner = 'georgia.cranko@gmail.com';
const members = new Set([owner,'busyandimportant01@gmail.com','elanagrunberg@protonmail.com','fcschmid95@gmail.com','freya.f.harvey@gmail.com','hannahgwarnesmith@gmail.com','ruby.forbes@outlook.com']);
const actions = new Set(['SUPPORT_APP_calendarEventSave','SUPPORT_APP_calendarNoteSave','SUPPORT_APP_calendarMonth','SUPPORT_APP_calendarSave','SUPPORT_APP_checkForInvoices','SUPPORT_APP_data','SUPPORT_APP_financeChallenge','SUPPORT_APP_financeKeepAlive','SUPPORT_APP_financeLock','SUPPORT_APP_financeUnlock','SUPPORT_APP_journal','SUPPORT_APP_nextVisitor','SUPPORT_APP_saveRawClaim','SUPPORT_APP_setInvoicePaid','SUPPORT_APP_submit','SUPPORT_APP_submitInit','SUPPORT_APP_testAutomation']);
const ownerActions = new Set([...actions].filter(name => /finance|Invoice|RawClaim|testAutomation|calendarNoteSave/.test(name)));
function json(value,status=200) { return new Response(JSON.stringify(value),{status,headers:{'Content-Type':'application/json','Cache-Control':'no-store','X-Content-Type-Options':'nosniff'}}); }
export async function handle(request,env,fetcher=fetch) {
  const url = new URL(request.url);
  if (url.pathname === '/api/config' && request.method === 'GET') {
    if (!env.SUPABASE_URL || !env.SUPABASE_ANON_KEY) return json({error:'Sign-in has not been configured yet.'},503);
    return json({url:env.SUPABASE_URL,key:env.SUPABASE_ANON_KEY});
  }
  if (url.pathname !== '/api/rpc') return json({error:'Not found.'},404);
  if (request.method !== 'POST') return json({error:'Method not allowed.'},405);
  if (request.headers.get('Origin') !== url.origin) return json({error:'Request origin rejected.'},403);
  if (!env.SUPABASE_URL || !env.SUPABASE_ANON_KEY || !env.APPS_SCRIPT_URL || !env.BRIDGE_SECRET || env.BRIDGE_SECRET.length<32) return json({error:'The live connection has not been configured yet.'},503);
  const authorization=request.headers.get('Authorization')||'';
  if (!authorization.startsWith('Bearer ')) return json({error:'Please sign in.'},401);
  let user;
  try {
    const response=await fetcher(env.SUPABASE_URL+'/auth/v1/user',{headers:{Authorization:authorization,apikey:env.SUPABASE_ANON_KEY},signal:AbortSignal.timeout(10000)});
    if(!response.ok) return json({error:'Your session expired. Please sign in again.'},401);
    user=await response.json();
  } catch { return json({error:'Sign-in service is temporarily unavailable.'},503); }
  const email=String(user.email||'').toLowerCase();
  if (!user.email_confirmed_at || !members.has(email)) return json({error:'This account is not approved for access.'},403);
  let input;
  try { const text=await request.text(); if(text.length>150000) return json({error:'Request is too large.'},413); input=JSON.parse(text); } catch { return json({error:'Invalid request.'},400); }
  if (!actions.has(input.action) || !Array.isArray(input.args) || input.args.length>10) return json({error:'Unsupported action.'},400);
  if (ownerActions.has(input.action) && email!==owner) return json({error:'Only Georgia can use financial and administration features.'},403);
  if (input.action==='SUPPORT_APP_data' && email!==owner) input.args[0]='';
  const payload=JSON.stringify({action:input.action,args:input.args,email,timestamp:Date.now(),nonce:crypto.randomUUID()});
  const key=await crypto.subtle.importKey('raw',new TextEncoder().encode(env.BRIDGE_SECRET),{name:'HMAC',hash:'SHA-256'},false,['sign']);
  const signature=Array.from(new Uint8Array(await crypto.subtle.sign('HMAC',key,new TextEncoder().encode(payload))),b=>b.toString(16).padStart(2,'0')).join('');
  try {
    const response=await fetcher(env.APPS_SCRIPT_URL,{method:'POST',headers:{'Content-Type':'text/plain;charset=utf-8'},body:JSON.stringify({payload,signature}),redirect:'follow',signal:AbortSignal.timeout(55000)});
    if(!response.ok) return json({error:'The spreadsheet service is unavailable.'},502);
    const result=await response.json();
    if(!result.ok) return json({error:result.error||'The request could not be completed.'},502);
    return json({result:result.result});
  } catch { return json({error:'The request could not be confirmed. Check the app before retrying a save.'},502); }
}
export const onRequest=({request,env})=>handle(request,env);
