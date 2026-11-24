(function(root,factory){
  if(typeof module==='object'&&module.exports)module.exports=factory();
  else root.DBNSession=factory();
})(typeof globalThis!=='undefined'?globalThis:this,function(){
  'use strict';
  function create(initialToken,options={}){
    let token=initialToken,refresh;
    const transport=options.fetch||((...args)=>fetch(...args));
    const readToken=options.readToken||(html=>new DOMParser().parseFromString(html,'text/html').querySelector('meta[name="app-token"]')?.content);
    async function renew(){
      if(!refresh)refresh=(async()=>{
        const page=await transport('/',{cache:'no-store',credentials:'same-origin'});
        if(!page.ok)throw new Error('Reopen the app to reconnect this window.');
        const next=readToken(await page.text());
        if(!/^[A-Za-z0-9_-]{43}$/.test(next||''))throw new Error('Reopen the app to reconnect this window.');
        token=next;
      })().finally(()=>{refresh=null;});
      await refresh;
    }
    async function request(path,data){
      if(!/^[a-z][a-z0-9/-]*$/.test(path))throw new Error('Invalid local API path.');
      const send=async()=>{
        const response=await transport('/api/'+path,{method:data===undefined?'GET':'POST',cache:'no-store',credentials:'same-origin',
          headers:{'X-App-Token':token,...(data===undefined?{}:{'Content-Type':'application/json'})},
          ...(data===undefined?{}:{body:JSON.stringify(data)})});
        return {response,value:await response.json()};
      };
      const sentToken=token;
      let result=await send();
      // This exact 403 is returned before any API action executes. Never retry
      // generation failures, timeouts, other denials or successful mutations.
      if(result.response.status===403&&result.value.error==='Reopen the app to reconnect this window.'){
        if(token===sentToken)await renew();
        result=await send();
      }
      return result;
    }
    return {request};
  }
  return {create};
});
