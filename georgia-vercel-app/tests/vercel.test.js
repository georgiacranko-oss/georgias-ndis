import test from 'node:test';
import assert from 'node:assert/strict';
import {Readable} from 'node:stream';
import {createVercelHandler} from '../lib/vercel-api.js';
const env={SUPABASE_URL:'https://identity.example',SUPABASE_ANON_KEY:'public-key',APPS_SCRIPT_URL:'https://script.example/exec',BRIDGE_SECRET:'a'.repeat(64)};
async function invoke(route,{body,headers={},method='POST',fetcher,settings=env,stream}={}){
 const req=stream?Readable.from(stream):{};
 Object.assign(req,{method,headers:{host:'support.example',origin:'https://support.example',authorization:'Bearer sample','content-type':'application/json',...headers}});
 if(body!==undefined)req.body=body;
 const res={headers:{},setHeader(name,value){this.headers[name.toLowerCase()]=value;},end(body){this.body=body;}};
 await createVercelHandler(route,{env:settings,fetcher})(req,res);return {...res,json:JSON.parse(res.body)};
}
test('Vercel config returns JSON and only public settings',async()=>{const r=await invoke('/api/config',{method:'GET'});assert.equal(r.statusCode,200);assert.deepEqual(r.json,{url:env.SUPABASE_URL,key:env.SUPABASE_ANON_KEY});assert.equal(r.headers['cache-control'],'no-store');});
test('Vercel config fails closed without environment variables',async()=>{const r=await invoke('/api/config',{method:'GET',settings:{}});assert.equal(r.statusCode,503);assert.match(r.json.error,/not been configured/);});
test('common pasted REST suffix is normalized before authentication',async()=>{const r=await invoke('/api/config',{method:'GET',settings:{...env,SUPABASE_URL:env.SUPABASE_URL+'/rest/v1/'}});assert.equal(r.json.url,env.SUPABASE_URL);});
test('logged-out and cross-origin requests never contact Sheets',async()=>{for(const headers of [{authorization:''},{origin:'https://other.example'}]){const r=await invoke('/api/rpc',{body:{action:'SUPPORT_APP_data',args:[]},headers,fetcher:()=>assert.fail('Unexpected request')});assert.ok([401,403].includes(r.statusCode));}});
test('parsed JSON body retains confirmed-user validation and signed bridge',async()=>{let calls=0;const r=await invoke('/api/rpc',{body:{action:'SUPPORT_APP_data',args:[]},fetcher:async(url,options)=>{calls++;if(calls===1){assert.equal(url,env.SUPABASE_URL+'/auth/v1/user');return Response.json({email:'georgia.cranko@gmail.com',email_confirmed_at:'2026-01-01'});}const message=JSON.parse(options.body);assert.equal(JSON.parse(message.payload).action,'SUPPORT_APP_data');assert.match(message.signature,/^[a-f0-9]{64}$/);return Response.json({ok:true,result:{notes:[]}});}});assert.equal(r.statusCode,200);assert.deepEqual(r.json,{result:{notes:[]}});assert.equal(calls,2);});
test('workers cannot use financial actions through the Vercel adapter',async()=>{let calls=0;const r=await invoke('/api/rpc',{body:{action:'SUPPORT_APP_setInvoicePaid',args:[]},fetcher:async()=>{calls++;return Response.json({email:'ruby.forbes@outlook.com',email_confirmed_at:'2026-01-01'});}});assert.equal(r.statusCode,403);assert.equal(calls,1);});
test('streamed oversized bodies are rejected before authentication',async()=>{const r=await invoke('/api/rpc',{stream:['x'.repeat(150001)],fetcher:()=>assert.fail('Unexpected request')});assert.equal(r.statusCode,413);});
test('malformed JSON returns a JSON error rather than an HTML page',async()=>{const r=await invoke('/api/rpc',{body:'{broken',fetcher:async()=>Response.json({email:'georgia.cranko@gmail.com',email_confirmed_at:'2026-01-01'})});assert.equal(r.statusCode,400);assert.equal(r.json.error,'Invalid request.');});
