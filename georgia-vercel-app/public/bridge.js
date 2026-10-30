(function(){
 const auth=import('/auth.js');
 function runner(success,failure){return new Proxy({}, {get(_,name){
   if(name==='withSuccessHandler')return fn=>runner(fn,failure);
   if(name==='withFailureHandler')return fn=>runner(success,fn);
   return (...args)=>{auth.then(async a=>{if(!a.session()){location.replace('/');throw Error('Please sign in.');}const token=await a.accessToken();const response=await fetch('/api/rpc',{method:'POST',headers:{'Content-Type':'application/json',Authorization:'Bearer '+token},body:JSON.stringify({action:String(name),args})});const body=await response.json();if(!response.ok)throw Error(body.error);return body.result;}).then(result=>{if(success)success(result);}).catch(error=>{if(failure)failure(error);else if(window.SUPPORT_APP_showError)window.SUPPORT_APP_showError(error.message);});};
 }});}
 window.google={script:{run:runner(),url:{getLocation(callback){callback({parameter:Object.fromEntries(new URLSearchParams(location.search))});}}}};
 auth.then(a=>{if(!a.session())location.replace('/');});
 document.addEventListener('DOMContentLoaded',()=>{const button=document.createElement('button');button.className='button secondary';button.textContent='Sign out';button.addEventListener('click',()=>auth.then(a=>a.signOut()));document.querySelector('.topbar')?.append(button);});
})();
