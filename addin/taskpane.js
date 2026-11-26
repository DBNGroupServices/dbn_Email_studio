'use strict';
/* dbn Email Studio — Outlook task pane, "one box" design (v3.1). */
const $=id=>document.getElementById(id);
const all=sel=>Array.from(document.querySelectorAll(sel));
const Draft=window.DBNDraft,Claude=window.DBNClaude;

const QUICK=[
  {intent:'acknowledge',title:'Acknowledge',hint:'Confirm you’ve got it and are on it'},
  {intent:'confirm',title:'Confirm or approve',hint:'Restate the key details and say yes'},
  {intent:'more-info',title:'Ask for more information',hint:'List what you need to act'},
  {intent:'decline',title:'Decline politely',hint:'Say no clearly and courteously'}];
const AUDIENCE_VOICE={tenant:'builtin-tenant',contractor:'builtin-contractor',executive:'builtin-executive',team:'builtin-team'};
const TOOL_NAMES={summary:'Summary',actions:'Action items',proofread:'Proofread',translate:'Translation'};
const REPLY_ICON='<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M9 14L4 9l5-5"></path><path d="M4 9h10a6 6 0 0 1 6 6v4"></path></svg>';
const CHEV='<svg class="chev" viewBox="0 0 24 24" aria-hidden="true"><path d="M9 6l6 6-6 6"></path></svg>';

const state={store:null,context:null,profiles:[],templates:[],versions:Draft.history(),resultContext:null,kind:'reply',
  lastTask:null,busy:false,applying:false,dirty:false,usingExample:false,armed:null,editing:null,
  cache:new Map(),suggestToken:0,voiceTouched:false};

/* ---------- small utilities ---------- */
const prefs={get(k,d){try{const v=localStorage.getItem('dbn.'+k);return v===null?d:JSON.parse(v);}catch(e){return d;}},
  set(k,v){try{localStorage.setItem('dbn.'+k,JSON.stringify(v));}catch(e){/* storage unavailable */}}};
function status(text,kind='',where='status'){const el=$(where);el.textContent=text;el.className='status'+(kind?' '+kind:'');}
function draftStatus(text,kind=''){status(text,kind,'draft-status');}
let toastTimer=null;
function settingsStatus(text,failed=false){
  const el=$('settings-status');el.textContent=text;el.className='status toast'+(failed?' error':' ok');el.hidden=!text;
  clearTimeout(toastTimer);if(!failed&&text)toastTimer=setTimeout(()=>{el.hidden=true;},4000);
}
function flash(id,text='Saved ✓'){
  const b=$(id);if(!b)return;const original=b.dataset.label||b.textContent;b.dataset.label=original;
  b.textContent=text;b.classList.add('done');clearTimeout(b._t);b._t=setTimeout(()=>{b.textContent=original;b.classList.remove('done');},2000);
}
function banner(text,action,handler,bad=false){
  $('banner').hidden=!text;$('banner').classList.toggle('bad',bad);$('banner-text').textContent=text||'';
  $('banner-action').hidden=!action;$('banner-action').textContent=action||'';$('banner-action').onclick=handler||null;
}
function option(value,label){const o=document.createElement('option');o.value=value;o.textContent=label;return o;}
function fill(select,items,selected){select.replaceChildren(...items.map(([v,l])=>option(v,l)));if(selected!==undefined)select.value=selected;}
function errorText(error){return error.message+(error.detail?'\n'+error.detail:'');}
function el(tag,cls,text){const e=document.createElement(tag);if(cls)e.className=cls;if(text!==undefined)e.textContent=text;return e;}
function show(view){$('view-compose').hidden=view!=='compose';$('view-draft').hidden=view!=='draft';}
function firstName(){return Draft.firstName(state.context?.meta?.from?.name||'')||'';}

/* ---------- connection & settings ---------- */
function modelName(id){return (Claude.MODELS.find(m=>m[0]===id)||[id,id])[1].split(' — ')[0];}
function connection(){
  const s=state.store.settings(),ok=!!state.store.apiKey();
  if(ok&&/Add your Claude API key/.test($('banner-text').textContent))banner('');
  $('connection').classList.toggle('ok',ok);$('connection').classList.toggle('bad',!ok);
  $('connection-label').textContent=ok?modelName(s.model).replace('Claude ',''):'Add key';
  $('settings-connection').textContent=ok?'Connected with your Claude API key · '+modelName(s.model)+'.':'Not connected yet. Paste your Claude API key below.';
  $('security-guard').checked=s.security_guard!==false;
  $('auto-suggest').checked=s.auto_suggest!==false;
  $('include-attachments').checked=!!s.include_attachments;
  fill($('model'),Claude.MODELS,s.model);
  $('api-key').value='';$('api-key').placeholder=ok?'Saved (…'+state.store.apiKey().slice(-4)+') · paste to replace':'sk-ant-…';
  $('remove-key').hidden=!ok;
  ['name','role','organisation','signature','style'].forEach(k=>{const e=$('me-'+(k==='organisation'?'org':k));if(e)e.value=s[k]||'';});
  const u=state.store.usage();
  $('storage-status').textContent=(state.store.kind==='mailbox'?'Saved in your Outlook mailbox (follows you to any PC). ':'Browser preview: saved in this browser only. ')+
    `Using ${(u.used/1024).toFixed(1)} KB of ${(u.limit/1024).toFixed(0)} KB.`;
}
function loadAll(){connection();loadProfiles();loadTemplates();}

/* ---------- voices (profiles) ---------- */
function currentProfile(){return state.profiles.find(p=>p.id===$('profile').value)||null;}
function loadProfiles(){
  state.profiles=state.store.profiles();
  const items=state.profiles.map(p=>[p.id,p.name]),def=state.store.settings().default_profile;
  const keep=$('profile').value;
  fill($('profile'),items,items.some(i=>i[0]===keep)?keep:(items.some(i=>i[0]===def)?def:items[0]?.[0]||''));
  fill($('default-profile'),[['','None'],...items],items.some(i=>i[0]===def)?def:'');
  fill($('edit-profile-pick'),items,state.editing&&items.some(i=>i[0]===state.editing)?state.editing:items[0]?.[0]);
  editProfile($('edit-profile-pick').value);
}
function setVoiceFor(audience){
  if(state.voiceTouched)return;
  const def=state.store.settings().default_profile;
  const want=AUDIENCE_VOICE[audience];
  const id=want&&state.profiles.some(p=>p.id===want)?want:(state.profiles.some(p=>p.id===def)?def:$('profile').value);
  $('profile').value=id;
  $('voice-note').textContent=want&&id===want?'auto-picked for '+audience+'s':'';
}
function exampleBox(text=''){
  const wrap=el('div','example');
  const area=document.createElement('textarea');area.rows=4;area.maxLength=2500;area.value=text;area.placeholder='Paste an email you wrote and liked';
  const remove=el('button','icon-btn sm','✕');remove.type='button';remove.setAttribute('aria-label','Remove example');
  remove.addEventListener('click',()=>wrap.remove());
  wrap.append(area,remove);return wrap;
}
function syncOutput(id){const out=document.querySelector('output[for="'+id+'"]');if(out)out.textContent=$(id).value;}
function editProfile(id){
  const p=state.profiles.find(x=>x.id===id)||{name:'',purpose:'',guidance:'',greeting:'',signoff:'',formality:3,warmth:3,directness:3,length:'Standard',format:'Paragraphs',avoid:'',signature:'',include_signature:true,examples:[]};
  state.editing=p.id||null;
  ['name','purpose','guidance','greeting','signoff','avoid','signature'].forEach(k=>$('pf-'+k).value=p[k]||'');
  ['formality','warmth','directness'].forEach(k=>{$('pf-'+k).value=p[k]||3;syncOutput('pf-'+k);});
  $('pf-length').value=p.length||'Standard';$('pf-format').value=p.format||'Paragraphs';
  $('pf-include-signature').checked=p.include_signature!==false;
  $('pf-examples').replaceChildren(...(p.examples||[]).map(exampleBox));
  $('pf-delete').hidden=!p.id;
}
async function saveProfile(event){
  event.preventDefault();
  const data={id:state.editing||'',include_signature:$('pf-include-signature').checked,
    examples:all('#pf-examples textarea').map(t=>t.value).filter(v=>v.trim()).slice(0,3)};
  ['name','purpose','guidance','greeting','signoff','avoid','signature'].forEach(k=>data[k]=$('pf-'+k).value);
  ['formality','warmth','directness'].forEach(k=>data[k]=Number($('pf-'+k).value));
  data.length=$('pf-length').value;data.format=$('pf-format').value;
  try{const r=await state.store.saveProfile(data);state.editing=r.saved.id;loadProfiles();connection();flash('pf-save');settingsStatus('Voice "'+r.saved.name+'" saved.');}
  catch(error){settingsStatus(error.message,true);}
}

/* ---------- templates ---------- */
function loadTemplates(){
  state.templates=state.store.templates();
  fill($('templates'),[['','Template…'],...state.templates.map(t=>[t.id,t.name])],'');
  $('templates').hidden=!state.templates.length;
  $('template-list').replaceChildren(...state.templates.map(t=>{
    const row=el('div','item');const name=el('span','',t.name);
    const del=el('button','link','Delete');del.type='button';
    del.addEventListener('click',async()=>{try{await state.store.deleteTemplate(t.id);loadTemplates();connection();}catch(error){settingsStatus(error.message,true);}});
    row.append(name,del);return row;
  }));
  if(!state.templates.length)$('template-list').append(el('p','muted small','No templates yet.'));
}

/* ---------- the open email ---------- */
function mode(){
  const c=state.usingExample?{compose:false}:state.context;
  if(!c)return 'new';
  if(c.compose)return c.selection?'rewrite':'compose';
  return 'reply';
}
function showEmail(ctx){
  const line=$('email-line');line.replaceChildren();
  const m=mode();
  if(m==='reply'){const who=ctx?.meta?.from?.name||ctx?.meta?.from?.email||'the sender';line.append('Replying to ',el('strong','',who));}
  else if(m==='new')line.textContent='New email';
  else line.textContent=m==='rewrite'?'Rewriting your selected text':'Writing in your draft';
  $('email-subject').textContent=ctx?.subject||'';
  $('generate').textContent={reply:'Write reply',new:'Write email',compose:'Write it',rewrite:'Rewrite'}[m];
  $('points-label').textContent=m==='reply'?'Or say what you want':m==='rewrite'?'How should it change?':'What should it say?';
  $('suggest-label').hidden=$('suggestions').hidden=m!=='reply';
}
async function readCurrent(){
  const previous=state.context,candidate=DBNOffice.target(window.Office);
  try{
    state.context=await DBNOffice.read(window.Office);state.usingExample=false;
    $('source').value=state.context.text;
  }catch(error){state.context=previous||(!state.usingExample?candidate:null);status(error.message,'error');}
  state.voiceTouched=false;
  showEmail(state.context);banner('');
  if(mode()==='reply')loadSuggestions();else{$('summary').hidden=true;setVoiceFor('other');}
}
function onItemChanged(){
  if(state.dirty){banner('You moved to another email. Your draft is kept until you load it.','Load new email',()=>{reset();readCurrent();});}
  else{reset();readCurrent();}
}
function reset(){state.versions.reset();state.dirty=false;state.resultContext=null;state.lastTask=null;state.armed=null;show('compose');status('');}

/* ---------- suggested replies ---------- */
function renderCards(items,loading){
  const box=$('suggestions');box.replaceChildren();
  if(loading){for(let i=0;i<3;i++){const b=el('div','card-btn skeleton');b.innerHTML='<span class="tile"></span><span class="txt"><span class="bar1"></span><span class="bar2"></span></span>';box.append(b);}return;}
  for(const item of items){
    const b=el('button','card-btn');b.type='button';
    const tile=el('span','tile');tile.innerHTML=REPLY_ICON;
    const txt=el('span','txt');txt.append(el('strong','',item.title),el('span','hint2',item.hint));
    b.append(tile,txt);b.insertAdjacentHTML('beforeend',CHEV);
    b.addEventListener('click',()=>generate(item.intent?{intent:item.intent}:{instruction:item.instruction,label:item.title}));
    box.append(b);
  }
}
async function loadSuggestions(){
  const ctx=state.usingExample?{itemId:'example',text:$('source').value,meta:{from:{name:'Alex (fictional)'}},subject:$('email-subject').textContent}:state.context;
  const s=state.store.settings();
  const key=(ctx&&(ctx.id||ctx.itemId))||'';
  const fallback=()=>{renderCards(QUICK);$('summary').hidden=true;};
  if(!state.store.apiKey()||s.auto_suggest===false||!ctx||!String(ctx.text||'').trim()){fallback();setVoiceFor('other');return;}
  if(key&&state.cache.has(key)){apply(state.cache.get(key));return;}
  const token=++state.suggestToken;
  renderCards(null,true);$('summary').hidden=false;$('summary').className='summary loading';$('summary').textContent='Reading the email…';
  try{
    const r=await Claude.suggest(state.store.apiKey(),s.model,{text:ctx.text,meta:ctx.meta,subject:ctx.subject},s);
    if(token!==state.suggestToken)return;
    if(key)state.cache.set(key,r);
    apply(r);
  }catch(error){
    if(token!==state.suggestToken)return;
    fallback();status('Suggestions unavailable: '+error.message,'error');
  }
  function apply(r){
    $('summary').hidden=false;$('summary').className='summary';$('summary').textContent=r.summary;
    renderCards(r.suggestions);setVoiceFor(r.audience);
  }
}

/* ---------- writing ---------- */
function setBusy(value,text){
  state.busy=value;
  ['generate','points','refine','refine-text','profile','mic','mic2','back'].forEach(id=>{if($(id))$(id).disabled=value;});
  all('.card-btn,.chips button,.more button').forEach(b=>b.disabled=value);
  if(value){status(text||'Writing…','busy');draftStatus('');}
  else if($('status').classList.contains('busy'))status('');
  syncButtons();
}
async function readUploads(){
  const files=[];for(const file of $('uploads').files){
    if(file.size>10*1024*1024)throw new Error(file.name+' is over 10 MB.');
    const data=await new Promise((resolve,reject)=>{const reader=new FileReader();reader.onload=()=>resolve(String(reader.result).split(',')[1]||'');reader.onerror=()=>reject(new Error('Could not read '+file.name));reader.readAsDataURL(file);});
    files.push({name:file.name,data});
  }return files;
}
async function gatherFiles(ctx){
  const files=await readUploads();
  if(state.store.settings().include_attachments&&ctx&&!ctx.compose&&ctx.attachments?.length)files.push(...await DBNOffice.attachments(ctx));
  if(files.length>10)throw new Error('Choose up to 10 files.');
  return files;
}
async function generate(opts={}){
  if(state.busy)return;
  if(!state.store)return status('Still connecting to Outlook… try again in a moment.','error');
  if(!state.store.apiKey()){openSettings();return status('Add your Claude API key first (Settings → AI connection).','error');}
  const m=mode(),typed=$('points').value.trim();
  if(!opts.tool&&!opts.intent&&!opts.instruction&&!typed&&m!=='rewrite')return status(m==='reply'?'Pick a suggested reply, or type what the reply should say.':'Type what the email should say.','error'),$('points').focus();
  setBusy(true,opts.tool?'Working…':'Writing…');
  try{
    const ctx=state.usingExample?null:(state.context||DBNOffice.target(window.Office));
    const files=await gatherFiles(ctx);
    const points=[opts.instruction,typed].filter(Boolean).join('\n');
    const task={source:$('source').value,points,mode:m==='new'?'compose':m==='compose'?'compose':m==='rewrite'?'rewrite':'reply',
      files,tone:'Use profile',length:(currentProfile()||{}).length||'Standard',format:(currentProfile()||{}).format||'Paragraphs',
      intent:opts.intent||'',variants:1};
    if(ctx&&!ctx.compose)task.meta={...ctx.meta,reply_all:false};
    if(opts.tool){task.action=opts.tool;task.language=typed||'English';task.points=opts.tool==='translate'?'':typed;}
    const s=state.store.settings();
    const result=await Claude.generate(state.store.apiKey(),s.model,task,s,currentProfile());
    state.kind=opts.tool?'tool':m;state.tool=opts.tool||'';
    state.resultContext=m==='new'?null:ctx;
    state.lastTask=opts.tool?null:task;
    state.versions.reset();
    setBusy(false);
    showResult(result);
  }catch(error){setBusy(false);status(errorText(error),'error');}
}
async function refine(instruction){
  if(state.busy||!instruction.trim())return;
  if(!state.lastTask)return draftStatus('Write a reply first, then ask for changes.','error');
  setBusy(true);draftStatus('Revising…','busy');
  try{
    const s=state.store.settings();
    const result=await Claude.generate(state.store.apiKey(),s.model,{...state.lastTask,intent:'',variants:1,refinement:instruction,current:{subject:$('draft-subject').value,body:$('draft-body').value}},s,currentProfile());
    setBusy(false);showResult(result);$('refine-text').value='';draftStatus('Revised. Use ‹ › to compare versions.');
  }catch(error){setBusy(false);draftStatus(errorText(error),'error');}
}

/* ---------- draft view ---------- */
function showResult(result){
  state.versions.push({subject:result.subject,body:result.body,checks:result.checks||[],warnings:result.warnings||[],alternatives:result.alternatives&&result.alternatives.length>1?result.alternatives:null,variant:0});
  state.dirty=state.kind!=='tool';
  const fn=firstName();
  $('draft-title').textContent=state.kind==='tool'?TOOL_NAMES[state.tool]:state.kind==='reply'?'Reply'+(fn?' to '+fn:''):state.kind==='new'?'New email':state.kind==='rewrite'?'Rewritten text':'For your draft';
  $('subject-row').hidden=state.kind!=='new';
  document.querySelector('.refine').hidden=$('refine-text').parentElement.hidden=state.kind==='tool';
  renderVersion();show('draft');banner('');
  $('draft-body').focus();$('draft-body').setSelectionRange(0,0);$('draft-body').scrollTop=0;
}
function renderVersion(){
  const v=state.versions.current();if(!v)return;
  $('draft-subject').value=v.subject;$('draft-body').value=v.body;
  const tabs=$('variant-tabs');tabs.hidden=!v.alternatives;
  tabs.replaceChildren(...(v.alternatives||[]).map((alt,i)=>{
    const b=el('button','',alt.label||('Option '+(i+1)));b.type='button';b.classList.toggle('on',i===v.variant);
    b.addEventListener('click',()=>{state.versions.update({variant:i,subject:alt.subject,body:alt.body});renderVersion();});
    return b;
  }));
  $('ver-label').textContent=state.versions.size>1?`${state.versions.position}/${state.versions.size}`:'';
  $('ver-back').hidden=$('ver-forward').hidden=state.versions.size<2;
  renderReview();
}
function renderReview(){
  const v=state.versions.current();if(!v)return;
  const items=state.kind==='tool'?[]:Draft.review($('draft-body').value+'\n'+$('draft-subject').value,{securityGuard:$('security-guard').checked,hasAttachments:!!state.context?.compose&&(state.context.attachments||[]).length>0});
  const list=[...items.map(i=>[i.level,i.message]),...v.checks.map(c=>['ai',c]),...v.warnings.map(w=>['warn',w])];
  const stops=list.filter(x=>x[0]==='stop');
  $('review').replaceChildren(...list.map(([level,text])=>el('li',level,text)));
  const line=$('review-line');line.hidden=!list.length;line.className='review-line'+(stops.length?' stop':'');
  if(list.length){
    const first=(stops[0]||list[0])[1];
    $('review-text').textContent=stops.length===1?first:list.length===1?first:stops.length?`${stops.length} to fix before sending · ${first}`:`${list.length} thing${list.length>1?'s':''} to check · ${first}`;
  }
  $('review-toggle').hidden=list.length<2;$('review-toggle').textContent=$('review').hidden?'Show':'Hide';
  state.armed=null;syncButtons();
}
function hasStops(){return all('#review li.stop').length>0;}
function canNewMessage(){return !!window.Office?.context?.mailbox?.displayNewMessageForm;}
function recipientsCount(){return (state.resultContext?.meta?.to?.length||0)+(state.resultContext?.meta?.cc?.length||0);}
function syncButtons(){
  const empty=!$('draft-body').value.trim(),k=state.kind;
  let label=k==='tool'?'Copy result':k==='reply'?'Insert reply':k==='compose'?'Insert into draft':k==='rewrite'?'Replace selection':(canNewMessage()?'Open as new email':'Copy email');
  if(!state.resultContext&&k!=='new'&&k!=='tool')label=canNewMessage()?'Open as new email':'Copy';
  if(state.applying)label='Opening in Outlook…';
  $('apply').textContent=state.armed==='apply'?'Fix the red item first — insert anyway?':label;
  $('apply-all').textContent=state.armed==='apply-all'?'Reply all anyway?':'Reply all';
  $('apply-all').hidden=k!=='reply'||!state.resultContext||recipientsCount()<2;
  $('copy').hidden=k==='tool';
  $('apply').disabled=$('apply-all').disabled=state.busy||state.applying||empty;
  $('copy').disabled=empty;
  $('apply-help').textContent=k==='tool'?'':'Opens in Outlook for you to review. Nothing is sent automatically.';
}
async function applyResult(how,button){
  if(state.busy||state.applying||!$('draft-body').value.trim())return;
  if(state.kind==='tool')return copyResult();
  if(hasStops()&&state.armed!==button){state.armed=button;syncButtons();return;}
  state.armed=null;state.applying=true;syncButtons();
  try{
    const linked=state.resultContext;
    if(linked&&state.kind!=='new'){await DBNOffice.apply(window.Office,linked,$('draft-body').value,how);}
    else if(canNewMessage()){Office.context.mailbox.displayNewMessageForm({subject:$('draft-subject').value,htmlBody:Draft.toHtml($('draft-body').value)});}
    else{state.applying=false;await copyResult();return;}
    state.dirty=false;
    draftStatus(linked&&state.kind!=='new'?(linked.compose?'Inserted into your draft. Nothing was sent.':'Opened in Outlook for your review. Nothing was sent.'):'Opened as a new email. Add recipients and review it. Nothing was sent.');
  }catch(error){draftStatus(error.message,'error');}
  finally{state.applying=false;syncButtons();}
}
async function copyResult(){
  const text=$('draft-body').value;
  try{await navigator.clipboard.writeText(text);draftStatus('Copied.');}
  catch(error){$('draft-body').focus();$('draft-body').select();try{document.execCommand('copy');draftStatus('Copied.');}catch(e){draftStatus('Press Ctrl+C to copy the selected text.','error');}}
}

/* ---------- settings drawer ---------- */
function openSettings(){$('drawer').hidden=false;$('close-settings').focus();if(state.store)connection();else settingsStatus('Still connecting to Outlook… settings will appear in a moment.',true);}
function closeSettings(){$('drawer').hidden=true;$('open-settings').focus();}
async function saveKey(){
  const typed=$('api-key').value.trim(),model=$('model').value;
  settingsStatus('Checking your key…');
  try{
    const key=typed||state.store.apiKey();
    if(!key)throw new Error('Paste your Claude API key first.');
    const info=await Claude.testKey(key,model);
    if(typed)await state.store.setApiKey(typed);
    await state.store.saveSettings({model});
    connection();flash('save-key','Connected ✓');settingsStatus('Connected. '+info.name+' is ready.');
    if(/API key/.test($('status').textContent))status('');
    if(mode()==='reply'&&!$('view-compose').hidden)loadSuggestions();
  }catch(error){settingsStatus(errorText(error),true);}
}
async function saveFlag(key,value){
  try{await state.store.saveSettings({[key]:value});settingsStatus('Saved.');}catch(error){settingsStatus(error.message,true);}
}

/* ---------- theme ---------- */
function applyOfficeTheme(){
  try{
    const bg=Office.context.officeTheme?.bodyBackgroundColor;if(!bg)return;
    const n=parseInt(bg.replace('#',''),16),lum=((n>>16)&255)*.299+((n>>8)&255)*.587+(n&255)*.114;
    document.documentElement.dataset.theme=lum<110?'dark':'light';
  }catch(e){/* keep light */}
}

/* ---------- wiring ---------- */
function wire(){
  $('generate').addEventListener('click',()=>generate());
  $('points').addEventListener('keydown',e=>{if((e.ctrlKey||e.metaKey)&&e.key==='Enter'){e.preventDefault();generate();}});
  $('profile').addEventListener('change',()=>{state.voiceTouched=true;$('voice-note').textContent='';});
  all('.more [data-tool]').forEach(b=>b.addEventListener('click',()=>generate({tool:b.dataset.tool})));
  $('uploads').addEventListener('change',()=>{const n=$('uploads').files.length;$('files-note').hidden=!n;$('files-note').textContent=n?`${n} file${n>1?'s':''} attached for the next draft.`:'';});
  $('templates').addEventListener('change',()=>{const t=state.templates.find(x=>x.id===$('templates').value);if(t)$('points').value=[$('points').value,t.body].filter(Boolean).join('\n\n');$('templates').value='';$('points').focus();});
  $('back').addEventListener('click',()=>{show('compose');if(state.dirty)banner('Your draft is kept.','Open draft',()=>{banner('');show('draft');});});
  $('draft-body').addEventListener('input',()=>{state.versions.update({body:$('draft-body').value});if(state.kind!=='tool')state.dirty=true;renderReview();});
  $('draft-subject').addEventListener('input',()=>state.versions.update({subject:$('draft-subject').value}));
  $('review-toggle').addEventListener('click',()=>{$('review').hidden=!$('review').hidden;$('review-toggle').textContent=$('review').hidden?'Show':'Hide';});
  all('.refine button').forEach(b=>b.addEventListener('click',()=>refine(b.dataset.refine)));
  $('refine').addEventListener('click',()=>refine($('refine-text').value));
  $('refine-text').addEventListener('keydown',e=>{if(e.key==='Enter'){e.preventDefault();refine($('refine-text').value);}});
  $('ver-back').addEventListener('click',()=>{state.versions.back();renderVersion();});
  $('ver-forward').addEventListener('click',()=>{state.versions.forward();renderVersion();});
  $('apply').addEventListener('click',()=>applyResult('reply','apply'));
  $('apply-all').addEventListener('click',()=>applyResult('replyAll','apply-all'));
  $('copy').addEventListener('click',copyResult);
  $('example').addEventListener('click',()=>{
    state.context=null;state.usingExample=true;reset();state.voiceTouched=false;
    $('source').value='Hi Sam,\n\nCould you confirm whether contractors can access the loading dock this Saturday from 7am for the chiller replacement? They will need the goods lift on standby.\n\nThanks,\nAlex';
    $('email-line').replaceChildren('Replying to ',el('strong','','Alex (fictional example)'));
    $('email-subject').textContent='Saturday dock access';
    $('suggest-label').hidden=$('suggestions').hidden=false;$('generate').textContent='Write reply';
    loadSuggestions();
  });
  if(window.DBNDictation){
    const hint=$('mic-hint'),base=hint.textContent;
    DBNDictation.create({button:$('mic'),field:$('points'),onState:on=>{hint.textContent=on?'Listening… tap the mic to stop':base;},onMessage:t=>{if(t)status(t);}});
    DBNDictation.create({button:$('mic2'),field:$('refine-text'),onMessage:t=>{if(t)draftStatus(t);}});
  }
  document.addEventListener('keydown',e=>{if(e.key==='Escape'&&!$('drawer').hidden)closeSettings();});
  // Settings
  $('open-settings').addEventListener('click',openSettings);$('connection').addEventListener('click',openSettings);
  $('close-settings').addEventListener('click',closeSettings);
  $('drawer').addEventListener('click',e=>{if(e.target===$('drawer'))closeSettings();});
  $('show-key').addEventListener('click',()=>{const k=$('api-key');k.type=k.type==='password'?'text':'password';$('show-key').textContent=k.type==='password'?'Show':'Hide';});
  $('save-key').addEventListener('click',saveKey);
  $('remove-key').addEventListener('click',async()=>{try{await state.store.setApiKey('');connection();settingsStatus('Key removed from your mailbox settings.');}catch(error){settingsStatus(error.message,true);}});
  $('save-me').addEventListener('click',async()=>{try{await state.store.saveSettings({name:$('me-name').value,role:$('me-role').value,organisation:$('me-org').value,signature:$('me-signature').value,style:$('me-style').value});connection();flash('save-me');settingsStatus('Your details are saved.');}catch(error){settingsStatus(error.message,true);}});
  $('default-profile').addEventListener('change',async()=>{try{await state.store.saveSettings({default_profile:$('default-profile').value});settingsStatus('Default voice saved.');}catch(error){settingsStatus(error.message,true);}});
  $('security-guard').addEventListener('change',()=>{saveFlag('security_guard',$('security-guard').checked);renderReview();});
  $('auto-suggest').addEventListener('change',()=>saveFlag('auto_suggest',$('auto-suggest').checked));
  $('include-attachments').addEventListener('change',()=>saveFlag('include_attachments',$('include-attachments').checked));
  $('edit-profile-pick').addEventListener('change',()=>editProfile($('edit-profile-pick').value));
  $('new-profile').addEventListener('click',()=>{editProfile(null);$('pf-name').focus();});
  $('profile-form').addEventListener('submit',saveProfile);
  ['pf-formality','pf-warmth','pf-directness'].forEach(id=>$(id).addEventListener('input',()=>syncOutput(id)));
  $('pf-add-example').addEventListener('click',()=>{if(all('#pf-examples .example').length<3)$('pf-examples').append(exampleBox());});
  $('pf-from-email').addEventListener('click',()=>{
    const text=$('source').value.trim();if(!text)return settingsStatus('Open one of your sent emails first.',true);
    if(all('#pf-examples .example').length>=3)return settingsStatus('Remove an example first (maximum 3).',true);
    $('pf-examples').append(exampleBox(text.slice(0,2500)));settingsStatus('Added. Remove any quoted older messages, then Save voice.');
  });
  $('pf-delete').addEventListener('click',async()=>{if(!state.editing)return;try{await state.store.deleteProfile(state.editing);state.editing=null;loadProfiles();connection();settingsStatus('Voice deleted.');}catch(error){settingsStatus(error.message,true);}});
  $('save-template').addEventListener('click',async()=>{try{await state.store.saveTemplate($('template-name').value,$('template-body').value);$('template-name').value=$('template-body').value='';loadTemplates();connection();flash('save-template');settingsStatus('Template saved.');}catch(error){settingsStatus(error.message,true);}});
}

/* ---------- start-up ----------
   Buttons are wired straight away; mailbox settings and the open email are
   loaded once Outlook confirms the add-in is ready (Office.onReady). */
let started=false;
function init(inOutlook){
  if(started)return;started=true;window.__dbnStarted=true;
  state.store=DBNStore.create();
  loadAll();syncButtons();
  if(!state.store.apiKey())banner('Welcome! Add your Claude API key to start drafting.','Open settings',openSettings);
  if(!inOutlook){state.context=null;showEmail(null);$('email-subject').textContent='Open Email Studio from the Outlook ribbon to work with your emails.';return;}
  applyOfficeTheme();
  readCurrent().then(()=>{if(!state.store.apiKey())banner('Welcome! Add your Claude API key to start drafting.','Open settings',openSettings);});
  try{Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged,()=>{onItemChanged();});}catch(error){/* older Outlook: reopen the pane */}
}
function boot(){
  try{wire();show('compose');renderCards(QUICK);}
  catch(error){status('Start-up error: '+error.message,'error');throw error;}
  if(window.Office&&typeof Office.onReady==='function'){
    Office.onReady(info=>{
      window.__dbnOfficeReady=true;
      const inOutlook=!!(info&&info.host===Office.HostType.Outlook)||!!(Office.context&&Office.context.mailbox);
      try{init(inOutlook);}catch(error){status('Start-up error: '+error.message,'error');}
    });
  }else init(false);
}
boot();
