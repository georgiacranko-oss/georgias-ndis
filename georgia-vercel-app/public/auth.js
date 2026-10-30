const key='georgia-support-session';
let config;
async function settings(){if(!config){const r=await fetch('/api/config');config=await r.json();if(!r.ok){const message=config.error;config=null;throw Error(message);}}return config;}
export function session(){try{return JSON.parse(sessionStorage.getItem(key)||'null');}catch{return null;}}
function save(value){sessionStorage.setItem(key,JSON.stringify(value));}
export async function signIn(email,password){const c=await settings();const r=await fetch(c.url+'/auth/v1/token?grant_type=password',{method:'POST',headers:{apikey:c.key,'Content-Type':'application/json'},body:JSON.stringify({email,password})});const result=await r.json();if(!r.ok)throw Error('Sign-in failed. Check your email and password.');save({...result,expires_at:Date.now()+result.expires_in*1000});}
let refreshing;
export async function accessToken(){let s=session();if(!s)throw Error('Please sign in.');if(s.expires_at>Date.now()+60000)return s.access_token;if(!refreshing)refreshing=(async()=>{const c=await settings();const r=await fetch(c.url+'/auth/v1/token?grant_type=refresh_token',{method:'POST',headers:{apikey:c.key,'Content-Type':'application/json'},body:JSON.stringify({refresh_token:s.refresh_token})});const fresh=await r.json();if(!r.ok){sessionStorage.removeItem(key);throw Error('Your session expired. Sign in again.');}save({...fresh,expires_at:Date.now()+fresh.expires_in*1000});return fresh.access_token;})().finally(()=>{refreshing=null;});return refreshing;}
export async function signOut(){const s=session();sessionStorage.removeItem(key);try{const c=await settings();if(s)await fetch(c.url+'/auth/v1/logout',{method:'POST',headers:{apikey:c.key,Authorization:'Bearer '+s.access_token}});}finally{location.replace('/');}}
