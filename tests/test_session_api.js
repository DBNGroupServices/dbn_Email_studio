const assert=require('node:assert/strict');
const session=require('../web/session-api.js');
const stale='a'.repeat(43),fresh='b'.repeat(43);
const reply=(status,value)=>({status,ok:status>=200&&status<300,json:async()=>value});
(async()=>{
  const calls=[];let executed=0;
  const client=session.create(stale,{readToken:html=>html,fetch:async(url,options)=>{
    calls.push({url,options});
    if(url==='/')return {ok:true,text:async()=>fresh};
    if(options.headers['X-App-Token']===stale)return reply(403,{error:'Reopen the app to reconnect this window.'});
    executed++;return reply(200,{body:'Fictional result'});
  }});
  const result=await client.request('generate',{points:'Fictional test'});
  assert.equal(result.value.body,'Fictional result');assert.equal(executed,1);
  assert.equal(calls.length,3);assert.equal(calls[0].options.body,calls[2].options.body);
  assert.equal(calls[2].options.headers['X-App-Token'],fresh);
  await client.request('settings');assert.equal(calls.length,4);
  await assert.rejects(()=>client.request('https://example.com'),/Invalid local/);
  for(const [code,error] of [[403,'This request must come from the app.'],[500,'Failed'],[429,'Busy']]){
    let count=0;
    const denied=session.create(stale,{fetch:async()=>{count++;return reply(code,{error});}});
    assert.equal((await denied.request('generate',{points:'Test'})).response.status,code);
    assert.equal(count,1);
  }
  let attempts=0;
  const failing=session.create(stale,{readToken:x=>x,fetch:async(url)=>{
    if(url==='/')return {ok:true,text:async()=>fresh};
    attempts++;return reply(403,{error:'Reopen the app to reconnect this window.'});
  }});
  assert.equal((await failing.request('settings')).response.status,403);assert.equal(attempts,2);
  console.log('Session reconnection: preserved POST, single execution, bounded retry and origin denial checks passed.');
})().catch(error=>{console.error(error);process.exitCode=1;});
