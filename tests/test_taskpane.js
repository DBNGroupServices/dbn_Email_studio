const assert=require('node:assert/strict');
const fs=require('node:fs');
const vm=require('node:vm');
const bridge=require('../web/office/office-client.js');
const source=fs.readFileSync(require.resolve('../web/office/taskpane.js'),'utf8');
async function panel(office){
  const elements=new Map(),calls=[],ready=[];
  function element(id){
    if(!elements.has(id))elements.set(id,{value:'',hidden:id==='result',disabled:false,files:[],listeners:{},textContent:'',classList:{toggle(){}},
      addEventListener(name,fn){this.listeners[name]=fn;},replaceChildren(){},append(){},
      async fire(name='click'){return this.listeners[name]?.();}});
    return elements.get(id);
  }
  element('action').value='rewrite';element('tone').value='Professional';element('length').value='Standard';
  const state={fail:false,pending:null};
  const api={request:async(path,data)=>{
    calls.push({path,data});let value;
    if(path==='settings')value={connected:true,settings:{provider:'chatgpt',model:'fixture'}};
    else if(path==='templates')value=[];
    else if(path==='generate'||path==='assist'){
      if(state.fail)throw new Error('Fictional generation failure');
      value={subject:'Fictional subject',body:'Fictional generated result',checks:[]};
    }else if(path==='outlook/draft'){if(state.pending)await state.pending;value={ok:true};}
    return {response:{ok:true},value};
  }};
  if(office){office.HostType={Outlook:'Outlook'};office.onReady=cb=>ready.push(Promise.resolve().then(()=>cb({host:'Outlook'})));}
  vm.runInNewContext(source,{window:{Office:office},Office:office,DBNOffice:bridge,DBNSession:{create:()=>api},
    document:{getElementById:element,querySelector:()=>({content:'a'.repeat(43)}),createElement:()=>({})},
    Option:function(){},navigator:{clipboard:{}},setTimeout});
  await new Promise(setImmediate);await Promise.all(ready);
  return {element,calls,state};
}
function readItem(id){return {itemId:id,subject:'Fictional '+id,body:{getAsync:(type,cb)=>cb({status:'succeeded',value:'Fictional email '+id})},displayReplyForm(value){this.reply=value;}};}
(async()=>{
  const browser=await panel();assert.equal(browser.element('apply').disabled,true);
  browser.element('points').value='Fictional typed notes';await browser.element('generate').fire();
  assert.equal(browser.element('apply').disabled,false);assert.equal(browser.element('apply').textContent,'Create draft in Outlook Classic');
  browser.element('draft-body').value='Edited result <review>';await browser.element('draft-body').fire('input');
  let release;browser.state.pending=new Promise(resolve=>{release=resolve;});
  const opening=browser.element('apply').fire();await browser.element('apply').fire();
  assert.equal(browser.element('apply').disabled,true);release();await opening;
  const drafts=browser.calls.filter(c=>c.path==='outlook/draft');assert.equal(drafts.length,1);assert.equal(drafts[0].data.body,'Edited result <review>');assert.equal(drafts[0].data.context,undefined);
  browser.element('draft-body').value=' ';await browser.element('draft-body').fire('input');assert.equal(browser.element('apply').disabled,true);
  const original=readItem('original'),office={context:{mailbox:{item:original,addHandlerAsync(){}}}};
  const linked=await panel(office);await linked.element('generate').fire();
  assert.equal(linked.element('apply').textContent,'Open reply in Outlook');
  const other=readItem('other');office.context.mailbox.item=other;await linked.element('read-email').fire();
  linked.state.fail=true;await linked.element('generate').fire();await linked.element('apply').fire();
  assert.match(linked.element('status').textContent,/changed/);assert.equal(other.reply,undefined);assert.equal(original.reply,undefined);
  linked.state.fail=false;await linked.element('generate').fire();await linked.element('apply').fire();assert.ok(other.reply.htmlBody.includes('Fictional generated result'));
  const compose={subject:{getAsync:cb=>cb({status:'failed',error:{message:'Cannot read subject'}})},body:{setSelectedDataAsync:(text,opts,cb)=>{compose.inserted=text;cb({status:'succeeded'});}}};
  const unreadable=await panel({context:{mailbox:{item:compose,addHandlerAsync(){}}}});
  unreadable.element('points').value='Notes without readable email';await unreadable.element('generate').fire();
  assert.equal(unreadable.element('apply').textContent,'Insert at cursor');await unreadable.element('apply').fire();assert.equal(compose.inserted,'Fictional generated result');
  await linked.element('example').fire();await linked.element('generate').fire();assert.equal(linked.element('apply').textContent,'Create draft in Outlook Classic');
  console.log('Task pane: notes-only draft, body-read failure, edited output, duplicate click, failed generation and changed-item protection passed.');
})().catch(error=>{console.error(error);process.exitCode=1;});
