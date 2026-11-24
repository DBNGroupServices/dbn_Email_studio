/* Small Office.js adapter; user text is always encoded before HTML insertion. */
((root)=>{
  const html=text=>String(text).replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c])).replace(/\r?\n/g,'<br>');
  const call=(object,method,...args)=>new Promise((resolve,reject)=>object[method](...args,result=>{
    if(result.status==='succeeded')resolve(result.value);else reject(new Error(result.error?.message||'Outlook could not complete this action.'));
  }));
  function target(office){
    const item=office?.context?.mailbox?.item;
    if(!item)return null;
    const compose=typeof item.subject!=='string';
    return {item,compose,id:item.itemId||'',subject:compose?'':item.subject,from:item.from?.displayName||'',attachments:item.attachments||[]};
  }
  async function read(office){
    const context=target(office);
    if(!context)throw new Error('Open Email Studio from the Outlook ribbon to read the current email. You can still create a new Outlook Classic draft from your notes here.');
    const {item,compose}=context;
    const subject=compose?await call(item.subject,'getAsync'):item.subject;
    let text='',selectionWarning='';
    if(compose&&item.getSelectedDataAsync){
      try{const selected=await call(item,'getSelectedDataAsync','text');if(selected?.sourceProperty!=='subject')text=selected?.data||'';}
      catch(error){selectionWarning='Outlook could not read the selection. The full draft was loaded; review Email context before generating.';}
    }
    if(!text)text=await call(item.body,'getAsync','text');
    return {...context,subject,text,selectionWarning};
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
  async function apply(office,context,text){
    const current=office?.context?.mailbox?.item;
    if(!current||!context||current!==context.item||(context.id&&current.itemId!==context.id))throw new Error('The selected email changed. Read the current email before applying a result.');
    if(context.compose){
      if(!current.body.setSelectedDataAsync)throw new Error('This Outlook version cannot insert the text. Use Copy instead.');
      await call(current.body,'setSelectedDataAsync',html(text),{coercionType:'html'});
    }else{
      current.displayReplyForm({htmlBody:html(text)});
    }
  }
  const api={html,target,read,attachments,apply};
  if(typeof module!=='undefined'&&module.exports)module.exports=api;else root.DBNOffice=api;
})(typeof window==='undefined'?{}:window);
