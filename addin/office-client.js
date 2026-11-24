/* Office.js adapter. User text is always encoded before HTML insertion.
   v2: reads sender/recipient metadata, supports Reply all, inserts formatted
   HTML (paragraphs and lists) and protects against applying a result to a
   different email from the one it was written for. */
((root)=>{
  const tools=root.DBNDraft||(typeof require==='function'?require('./draft-tools.js'):null);
  const html=text=>tools?tools.toHtml(text):String(text).replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c])).replace(/\r?\n/g,'<br>');
  const call=(object,method,...args)=>new Promise((resolve,reject)=>object[method](...args,result=>{
    if(result.status==='succeeded')resolve(result.value);else reject(new Error(result.error?.message||'Outlook could not complete this action.'));
  }));
  const person=value=>value&&typeof value==='object'?{name:String(value.displayName||''),email:String(value.emailAddress||'')}:null;
  const people=value=>Array.isArray(value)?value.map(person).filter(Boolean):[];
  function target(office){
    const item=office?.context?.mailbox?.item;
    if(!item)return null;
    const compose=typeof item.subject!=='string';
    const meta=compose?{}:{from:person(item.from)||person(item.sender)||{name:'',email:''},to:people(item.to),cc:people(item.cc),
      date:item.dateTimeCreated instanceof Date?item.dateTimeCreated.toISOString():String(item.dateTimeCreated||'')};
    return {item,compose,id:item.itemId||'',conversationId:item.conversationId||'',subject:compose?'':item.subject,
      from:meta.from?.name||'',meta,attachments:item.attachments||[]};
  }
  async function read(office){
    const context=target(office);
    if(!context)throw new Error('Open Email Studio from the Outlook ribbon to read the current email. You can still create a new Outlook Classic draft from your notes here.');
    const {item,compose}=context;
    const subject=compose?await call(item.subject,'getAsync'):item.subject;
    let text='',selectionWarning='',selection=false;
    if(compose&&item.getSelectedDataAsync){
      try{const selected=await call(item,'getSelectedDataAsync','text');if(selected?.sourceProperty!=='subject')text=selected?.data||'';selection=!!text;}
      catch(error){selectionWarning='Outlook could not read the selection. The full draft was loaded; review Email context before generating.';}
    }
    if(!text)text=await call(item.body,'getAsync','text');
    return {...context,subject,text,selection,selectionWarning,meta:{...context.meta,subject}};
  }
  async function attachments(context){
    if(!context?.item?.getAttachmentContentAsync)throw new Error('This Outlook version cannot read attachments here. Save and add the files manually.');
    const files=[];const supported=/\.(png|jpe?g|webp|pdf|docx|eml|msg|txt|html?)$/i;
    for(const attachment of context.attachments.filter(a=>!a.isInline)){
      if(!supported.test(attachment.name))continue;
      if(attachment.size>10*1024*1024)throw new Error(attachment.name+' is over 10 MB. Add a smaller file.');
      const value=await call(context.item,'getAttachmentContentAsync',attachment.id);
      if(value.format!=='base64')throw new Error('Save '+attachment.name+' and add it manually; Outlook returned an unsupported attachment format.');
      files.push({name:attachment.name,data:value.content});
    }
    return files;
  }
  /* how: 'reply' | 'replyAll' | 'insert' (compose) | 'replace' (compose body) */
  async function apply(office,context,text,how='reply'){
    const current=office?.context?.mailbox?.item;
    if(!current||!context||current!==context.item||(context.id&&current.itemId!==context.id))throw new Error('The selected email changed. Read the current email before applying a result.');
    const body=html(text);
    if(context.compose){
      if(how==='replace'&&current.body.setAsync){await call(current.body,'setAsync',body,{coercionType:'html'});return 'replaced';}
      if(!current.body.setSelectedDataAsync)throw new Error('This Outlook version cannot insert the text. Use Copy instead.');
      await call(current.body,'setSelectedDataAsync',body,{coercionType:'html'});
      return 'inserted';
    }
    if(how==='replyAll'&&current.displayReplyAllForm){current.displayReplyAllForm({htmlBody:body});return 'replyAll';}
    current.displayReplyForm({htmlBody:body});
    return 'reply';
  }
  const api={html,target,read,attachments,apply};
  if(typeof module!=='undefined'&&module.exports)module.exports=api;else root.DBNOffice=api;
})(typeof window==='undefined'?globalThis:window);
