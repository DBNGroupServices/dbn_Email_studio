'use strict';
/* DBN Email Studio — Outlook task pane v2. */
const $=id=>document.getElementById(id);
const all=sel=>Array.from(document.querySelectorAll(sel));
const token=document.querySelector('meta[name="app-token"]').content;
const sessionApi=DBNSession.create(token);
const Draft=window.DBNDraft;

const INTENTS=[['acknowledge','Acknowledge'],['confirm','Confirm'],['more-info','Ask for info'],['decline','Decline'],
  ['schedule','Schedule'],['chase','Chase up'],['escalate','Escalate'],['delegate','Passing on'],['incident','Incident reply'],['thanks','Thanks']];
const TONES=['Use profile','Professional','Warm & professional','Friendly','Formal','Firm & clear','Diplomatic','Reassuring','Apologetic','Assertive','Customer focused','Brief & direct'];
const LABELS={reply:'Generate reply',write:'Generate draft',tools:'Run tool'};

const state={snapshot:null,context:null,mode:'reply',writeKind:'compose',intent:'',profiles:[],templates:[],
  versions:Draft.history(),resultContext:null,lastTask:null,busy:false,applying:false,dirty:false,usingExample:false,
  armed:null,offline:false,editing:null};

/* ---------- small utilities ---------- */
const prefs={get(k,d){try{const v=localStorage.getItem('dbn.'+k);return v===null?d:JSON.parse(v);}catch(e){return d;}},
  set(k,v){try{localStorage.setItem('dbn.'+k,JSON.stringify(v));}catch(e){/* storage unavailable */}}};
async function api(path,data){
  let result;
  try{result=await sessionApi.request(path,data);}
  catch(error){engineDown();throw new Error('Email Studio on this computer is not responding. Reopen DBN Email Studio, then try again.');}
  if(state.offline)engineBack();
  if(!result.response.ok)throw new Error(result.value.error||'This action could not be completed.');
  return result.value;
}
function status(text,kind=''){const el=$('status');el.textContent=text;el.className='status'+(kind?' '+kind:'');}
function settingsStatus(text,failed=false){$('settings-status').textContent=text;$('settings-status').className='status'+(failed?' error':'');}
function banner(text,action,handler,bad=false){
  $('banner').hidden=!text;$('banner').classList.toggle('bad',bad);$('banner-text').textContent=text||'';
  $('banner-action').hidden=!action;$('banner-action').textContent=action||'';$('banner-action').onclick=handler||null;
}
function option(value,label){const o=document.createElement('option');o.value=value;o.textContent=label;return o;}
function fill(select,items,selected){select.replaceChildren(...items.map(([v,l])=>option(v,l)));if(selected!==undefined)select.value=selected;}

/* ---------- engine connectivity ---------- */
let healthTimer=null;
function engineDown(){
  if(state.offline)return;state.offline=true;
  banner('Lost connection to Email Studio on this computer. Your work is kept.','Retry',()=>pingHealth(),true);
  healthTimer=setInterval(pingHealth,5000);
}
function engineBack(){state.offline=false;clearInterval(healthTimer);banner('');loadSettings();}
async function pingHealth(){try{const r=await fetch('/health',{cache:'no-store'});if(r.ok)engineBack();}catch(e){/* still offline */}}

/* ---------- connection & settings ---------- */
function connection(value){
  state.snapshot=value;
  const ok=!!value.connected;
  $('connection').classList.toggle('ok',ok);$('connection').classList.toggle('bad',!ok);
  $('connection-label').textContent=ok?(value.settings.provider==='chatgpt'?'ChatGPT plan':'API key'):'Not connected';
  $('settings-connection').textContent=ok?(value.settings.provider==='chatgpt'?'Using your ChatGPT plan · model '+value.settings.model:'Using separate API billing · model '+value.settings.model):'Not connected. Sign in to draft with your ChatGPT plan.';
  $('signin').hidden=ok&&value.settings.provider==='chatgpt';
  $('version').textContent=value.version?'v'+value.version:'';
  $('security-guard').checked=value.settings.security_guard!==false;
  $('save-history').checked=value.settings.save_history!==false;
}
async function loadSettings(){
  try{connection(await api('settings'));}
  catch(error){status(error.message,'error');return;}
  try{await Promise.all([loadProfiles(),loadTemplates()]);}catch(error){status(error.message,'error');}
}
async function saveSetting(patch){
  try{connection(await api('settings',patch));settingsStatus('Saved.');}catch(error){settingsStatus(error.message,true);}
}

/* ---------- profiles ---------- */
function currentProfile(){return state.profiles.find(p=>p.id===$('profile').value)||null;}
function applyProfileDefaults(){
  const p=currentProfile()||{formality:3,warmth:3,directness:3,length:'Standard',format:'Paragraphs'};
  ['formality','warmth','directness'].forEach(k=>{$(k).value=p[k]||3;syncOutput(k);});
  setSeg('length',p.length||'Standard');$('format').value=p.format||'Paragraphs';
}
async function loadProfiles(){
  state.profiles=await api('profiles');
  const items=state.profiles.map(p=>[p.id,p.name]);
  const preferred=prefs.get('profile',state.snapshot?.settings?.default_profile||items[0]?.[0]||'');
  fill($('profile'),[['','No profile'],...items],items.some(i=>i[0]===preferred)?preferred:(items[0]?.[0]||''));
  fill($('default-profile'),[['','None'],...items],state.snapshot?.settings?.default_profile||'');
  fill($('edit-profile-pick'),items,state.editing&&items.some(i=>i[0]===state.editing)?state.editing:items[0]?.[0]);
  applyProfileDefaults();editProfile($('edit-profile-pick').value);
}
function exampleBox(text=''){
  const wrap=document.createElement('div');wrap.className='example';
  const area=document.createElement('textarea');area.rows=4;area.maxLength=5000;area.value=text;area.placeholder='Paste an email you wrote and liked';
  const remove=document.createElement('button');remove.type='button';remove.className='icon-btn small';remove.textContent='✕';remove.setAttribute('aria-label','Remove example');
  remove.addEventListener('click',()=>wrap.remove());
  wrap.append(area,remove);return wrap;
}
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
  try{
    const before=new Set(state.profiles.map(p=>p.id));
    state.profiles=await api('profiles',data);
    state.editing=data.id||state.profiles.find(p=>!before.has(p.id))?.id||null;
    await loadProfiles();settingsStatus('Profile saved.');
  }catch(error){settingsStatus(error.message,true);}
}

/* ---------- templates ---------- */
async function loadTemplates(){
  state.templates=await api('templates');
  fill($('templates'),[['','Choose…'],...state.templates.map(t=>[t.id,t.name])],'');
  $('template-list').replaceChildren(...state.templates.map(t=>{
    const row=document.createElement('div');row.className='item';
    const name=document.createElement('span');name.textContent=t.name;
    const del=document.createElement('button');del.type='button';del.className='link small';del.textContent='Delete';
    del.addEventListener('click',async()=>{try{await api('templates/delete',{id:t.id});await loadTemplates();}catch(error){settingsStatus(error.message,true);}});
    row.append(name,del);return row;
  }));
  if(!state.templates.length){const p=document.createElement('p');p.className='muted small';p.textContent='No templates yet.';$('template-list').append(p);}
}

/* ---------- reading Outlook ---------- */
function showEmail(ctx){
  const from=ctx?.meta?.from?.name||ctx?.meta?.from?.email||(ctx?.compose?'Your draft':'');
  $('email-from').textContent=from||'No email loaded';
  $('email-subject').textContent=ctx?.subject||'';
  $('avatar').textContent=(Draft.firstName(from)||from||'?').charAt(0).toUpperCase();
  const others=(ctx?.meta?.to?.length||0)+(ctx?.meta?.cc?.length||0);
  $('reply-all').checked=false;$('apply-all').dataset.allowed=!ctx||ctx.compose||others<2?'0':'1';
}
async function readCurrent(){
  const previous=state.context,candidate=DBNOffice.target(window.Office);
  try{
    state.context=await DBNOffice.read(window.Office);state.usingExample=false;
    $('source').value=state.context.text;showEmail(state.context);
    $('host-status').textContent=state.context.selectionWarning||(state.context.compose?
      (state.context.selection?'Compose · your selected text will be rewritten.':'Compose · select text in your draft to rewrite just that part.'):
      'Ready. Pick a quick reply or add your points.');
    if(state.context.compose&&state.mode==='reply')setMode('write');
    if(state.context.compose)setWriteKind(state.context.selection||state.context.text.trim()?'rewrite':'compose');
  }catch(error){state.context=previous||(!state.usingExample?candidate:null);$('host-status').textContent=error.message;}
  banner('');syncButtons();
}
function onItemChanged(){
  if(!$('result').hidden&&state.dirty){
    banner('You moved to another email. This draft is kept until you load it.','Load new email',()=>{clearResult();readCurrent();});
  }else{clearResult();readCurrent();}
}

/* ---------- mode, segments, sliders ---------- */
function setMode(mode){
  state.mode=mode;
  all('.tab').forEach(t=>t.setAttribute('aria-selected',String(t.dataset.mode===mode)));
  $('intents').hidden=mode!=='reply';$('write-kind').hidden=mode!=='write';$('tool-kind').hidden=mode!=='tools';
  $('points-label').textContent=mode==='reply'?'What should the reply say?':mode==='write'?(state.writeKind==='rewrite'?'What should change?':'Your notes or dot points'):'Extra instructions (optional)';
  $('generate').firstChild.textContent=LABELS[mode]+' ';
  syncButtons();
}
function setWriteKind(kind){state.writeKind=kind;all('#write-kind button').forEach(b=>b.classList.toggle('on',b.dataset.kind===kind));if(state.mode==='write')setMode('write');}
function setSeg(id,value){all('#'+id+' button').forEach(b=>b.classList.toggle('on',b.dataset.value===value));}
function segValue(id){return all('#'+id+' button.on')[0]?.dataset.value;}
function syncOutput(id){const out=document.querySelector('output[for="'+id+'"]');if(out)out.textContent=$(id).value;}

/* ---------- generation ---------- */
function setBusy(value,text){
  state.busy=value;
  ['generate','read-email','points','source','uploads','include-attachments','refine','refine-text','profile','tone'].forEach(id=>$(id).disabled=value);
  all('.chips button').forEach(b=>b.disabled=value);
  if(value)status(text||'Writing…','busy');
  syncButtons();
}
async function readUploads(){
  const files=[];for(const file of $('uploads').files){
    if(file.size>10*1024*1024)throw new Error(file.name+' is over 10 MB.');
    const data=await new Promise((resolve,reject)=>{const reader=new FileReader();reader.onload=()=>resolve(String(reader.result).split(',')[1]);reader.onerror=()=>reject(new Error('Could not read '+file.name));reader.readAsDataURL(file);});
    files.push({name:file.name,data});
  }return files;
}
function buildTask(files){
  const mode=state.mode==='reply'?'reply':state.mode==='write'?state.writeKind:'compose';
  const ctx=state.context&&!state.usingExample?state.context:null;
  const task={source:$('source').value,points:$('points').value,mode,files,tone:$('tone').value,length:segValue('length')||'Standard',
    format:$('format').value,model:state.snapshot.settings.model,profile_id:$('profile').value,intent:state.mode==='reply'?state.intent:'',
    variants:state.mode==='tools'?1:Number($('variants').value),
    sliders:{formality:Number($('formality').value),warmth:Number($('warmth').value),directness:Number($('directness').value)}};
  if(ctx&&!ctx.compose)task.meta={...ctx.meta,reply_all:$('reply-all').checked};
  if(state.mode==='tools'){task.action=$('action').value;task.language=$('language').value;}
  return task;
}
async function generate(){
  if(state.busy)return;
  if(!state.snapshot?.connected){openSettings();return status('Connect your ChatGPT plan first (Settings).','error');}
  setBusy(true,state.mode==='tools'?'Working…':'Writing your draft…');
  try{
    const ctx=state.usingExample?null:(state.context||DBNOffice.target(window.Office));
    const files=await readUploads();if($('include-attachments').checked&&ctx)files.push(...await DBNOffice.attachments(ctx));
    if(files.length>10)throw new Error('Choose up to 10 files.');
    const task=buildTask(files);
    const result=await api(state.mode==='tools'?'assist':'generate',task);
    // A new email from notes in a read window is a new message, not a reply.
    state.resultContext=state.mode==='write'&&state.writeKind==='compose'&&!ctx?.compose?null:ctx;
    state.lastTask=state.mode==='tools'?null:task;
    state.versions.reset();
    showResult(result,'Draft 1');
    status(result.alternatives?.length>1?'Ready. Compare the options above the draft.':'Ready to review.');
  }catch(error){status(error.message,'error');}
  finally{setBusy(false);}
}
async function refine(instruction){
  if(state.busy||!instruction.trim())return;
  if(!state.lastTask)return status('Refinements apply to drafts. Generate a reply or draft first.','error');
  setBusy(true,'Revising…');
  try{
    const result=await api('generate',{...state.lastTask,variants:1,refinement:instruction,current:{subject:$('draft-subject').value,body:$('draft-body').value}});
    showResult(result,instruction.length>28?instruction.slice(0,26)+'…':instruction);
    $('refine-text').value='';status('Revised. Use ‹ › to compare versions.');
  }catch(error){status(error.message,'error');}
  finally{setBusy(false);}
}

/* ---------- result ---------- */
function showResult(result,label){
  const alternatives=result.alternatives&&result.alternatives.length>1?result.alternatives:null;
  state.versions.push({subject:result.subject,body:result.body,checks:result.checks||[],warnings:result.warnings||[],alternatives,variant:0,label});
  renderVersion();$('result').hidden=false;state.dirty=true;
  $('result').scrollIntoView({behavior:'smooth',block:'start'});
}
function renderVersion(){
  const v=state.versions.current();if(!v)return;
  $('draft-subject').value=v.subject;$('draft-body').value=v.body;
  const tabs=$('variant-tabs');tabs.hidden=!v.alternatives;
  tabs.replaceChildren(...(v.alternatives||[]).map((alt,i)=>{
    const b=document.createElement('button');b.type='button';b.textContent=alt.label||('Option '+(i+1));b.classList.toggle('on',i===v.variant);
    b.addEventListener('click',()=>{state.versions.update({variant:i,subject:alt.subject,body:alt.body});renderVersion();});
    return b;
  }));
  $('ver-label').textContent=state.versions.size>1?`${state.versions.position}/${state.versions.size}`:'';
  $('ver-back').hidden=$('ver-forward').hidden=state.versions.size<2;
  renderReview();syncButtons();
}
function renderReview(){
  const v=state.versions.current();if(!v)return;
  const items=Draft.review($('draft-body').value+'\n'+$('draft-subject').value,{securityGuard:$('security-guard').checked,hasAttachments:!!state.context?.compose&&(state.context.attachments||[]).length>0});
  const list=[...items.map(i=>[i.level,i.message]),...v.checks.map(c=>['ai',c]),...v.warnings.map(w=>['warn',w])];
  $('review').replaceChildren(...list.map(([level,text])=>{const li=document.createElement('li');li.className=level;li.textContent=text;return li;}));
  state.armed=null;syncButtons();
}
function clearResult(){$('result').hidden=true;state.versions.reset();state.dirty=false;state.resultContext=null;state.lastTask=null;syncButtons();}
function hasStops(){return all('#review li.stop').length>0;}
function syncButtons(){
  const linked=state.resultContext,empty=$('result').hidden||!$('draft-body').value.trim();
  const label=state.applying?'Opening in Outlook…':linked?(linked.compose?'Insert into draft':'Open reply'):'Create Outlook draft';
  $('apply').textContent=state.armed==='apply'?'Send-check warnings — open anyway?':label;
  $('apply-all').textContent=state.armed==='apply-all'?'Open reply all anyway?':'Reply all';
  $('apply-all').hidden=!linked||linked.compose||$('apply-all').dataset.allowed!=='1';
  $('apply').disabled=$('apply-all').disabled=state.busy||state.applying||empty;
  $('copy').disabled=empty;
  $('apply-help').textContent=linked?'Opens in Outlook for you to review. Nothing is sent automatically.':'Creates a new editable email with no recipients. Nothing is sent automatically.';
}
async function applyResult(how,button){
  if(state.busy||state.applying||$('result').hidden||!$('draft-body').value.trim())return;
  if(hasStops()&&state.armed!==button){state.armed=button;syncButtons();return;}
  state.armed=null;state.applying=true;syncButtons();
  try{
    const linked=state.resultContext;
    if(linked)await DBNOffice.apply(window.Office,linked,$('draft-body').value,how);
    else await api('outlook/draft',{subject:$('draft-subject').value,body:$('draft-body').value});
    state.dirty=false;
    status(linked?(linked.compose?'Inserted into your draft. Nothing was sent.':'Opened in Outlook for your review. Nothing was sent.'):'New draft opened in Outlook Classic. Add recipients and review it. Nothing was sent.');
  }catch(error){status(error.message,'error');}
  finally{state.applying=false;syncButtons();}
}
async function copyResult(){
  const text=$('draft-body').value;
  try{await navigator.clipboard.writeText(text);status('Copied.');}
  catch(error){$('draft-body').focus();$('draft-body').select();try{document.execCommand('copy');status('Copied.');}catch(e){status('Press Ctrl+C to copy the selected draft.','error');}}
}

/* ---------- settings drawer ---------- */
async function openSettings(){
  $('drawer').hidden=false;$('close-settings').focus();
  try{const a=await api('autostart');$('autostart').checked=!!a.enabled;$('autostart').disabled=!a.supported;
    if(!a.supported)$('autostart').parentElement.title='Available in the installed Windows app.';}catch(e){$('autostart').disabled=true;}
  try{const o=await api('office/status');$('engine-status').textContent='Engine '+($('version').textContent||'')+' running. Local certificate renews automatically'+(o.server_certificate_expires?' (current one valid to '+new Date(o.server_certificate_expires).toLocaleDateString('en-AU')+').':'.');}catch(e){$('engine-status').textContent='';}
}
function closeSettings(){$('drawer').hidden=true;$('open-settings').focus();}

/* ---------- theme ---------- */
function applyOfficeTheme(){
  try{
    const bg=Office.context.officeTheme?.bodyBackgroundColor;if(!bg)return;
    const n=parseInt(bg.replace('#',''),16),lum=((n>>16)&255)*.299+((n>>8)&255)*.587+(n&255)*.114;
    document.documentElement.dataset.theme=lum<110?'dark':'light';
  }catch(e){/* keep system theme */}
}

/* ---------- wiring ---------- */
function wire(){
  fill($('tone'),TONES.map(t=>[t,t]),prefs.get('tone','Use profile'));
  $('intents').replaceChildren(...INTENTS.map(([id,label])=>{
    const b=document.createElement('button');b.type='button';b.dataset.intent=id;b.textContent=label;
    b.title='One-click: '+label.toLowerCase()+' (adds your notes if any)';
    b.addEventListener('click',()=>{state.intent=id;all('#intents button').forEach(x=>x.classList.toggle('on',x===b));generate();});
    return b;
  }));
  all('.tab').forEach(t=>t.addEventListener('click',()=>setMode(t.dataset.mode)));
  all('#write-kind button').forEach(b=>b.addEventListener('click',()=>setWriteKind(b.dataset.kind)));
  all('#length button').forEach(b=>b.addEventListener('click',()=>{setSeg('length',b.dataset.value);prefs.set('length',b.dataset.value);}));
  ['formality','warmth','directness','pf-formality','pf-warmth','pf-directness'].forEach(id=>$(id).addEventListener('input',()=>syncOutput(id)));
  $('profile').addEventListener('change',()=>{prefs.set('profile',$('profile').value);applyProfileDefaults();});
  $('tone').addEventListener('change',()=>prefs.set('tone',$('tone').value));
  $('variants').value=String(prefs.get('variants',1));$('variants').addEventListener('change',()=>prefs.set('variants',Number($('variants').value)));
  $('action').addEventListener('change',()=>{$('language-row').hidden=$('action').value!=='translate';});
  $('generate').addEventListener('click',()=>{state.intent='';all('#intents button').forEach(x=>x.classList.remove('on'));generate();});
  $('read-email').addEventListener('click',()=>{clearResult();readCurrent();});
  $('draft-body').addEventListener('input',()=>{state.versions.update({body:$('draft-body').value});state.dirty=true;renderReview();});
  $('draft-subject').addEventListener('input',()=>state.versions.update({subject:$('draft-subject').value}));
  all('.refine button').forEach(b=>b.addEventListener('click',()=>refine(b.dataset.refine)));
  $('refine').addEventListener('click',()=>refine($('refine-text').value));
  $('refine-text').addEventListener('keydown',e=>{if(e.key==='Enter'){e.preventDefault();refine($('refine-text').value);}});
  $('ver-back').addEventListener('click',()=>{state.versions.back();renderVersion();});
  $('ver-forward').addEventListener('click',()=>{state.versions.forward();renderVersion();});
  $('apply').addEventListener('click',()=>applyResult($('reply-all').checked?'replyAll':'reply','apply'));
  $('apply-all').addEventListener('click',()=>applyResult('replyAll','apply-all'));
  $('reply-all').addEventListener('change',syncButtons);
  $('copy').addEventListener('click',copyResult);
  $('use-template').addEventListener('click',()=>{const t=state.templates.find(x=>x.id===$('templates').value);if(t)$('points').value=[$('points').value,t.body].filter(Boolean).join('\n\n');});
  $('example').addEventListener('click',()=>{
    state.context=null;state.usingExample=true;clearResult();setMode('reply');
    $('source').value='Hi Sam,\n\nCould you confirm whether contractors can access the loading dock this Saturday from 7am for the chiller replacement? They will need a lift on standby.\n\nThanks,\nAlex';
    $('points').value='Yes from 7am. Contractors to sign in at the security desk and complete induction. Goods lift booked 7am to 12pm.';
    showEmail({meta:{from:{name:'Alex (fictional)',email:''}},subject:'Saturday dock access — fictional example'});
    $('host-status').textContent='Example mode — fictional email, nothing loaded from Outlook.';
  });
  document.addEventListener('keydown',e=>{
    if(e.key==='Escape'&&!$('drawer').hidden)closeSettings();
    if((e.ctrlKey||e.metaKey)&&e.key==='Enter'&&$('drawer').hidden){e.preventDefault();generate();}
  });
  // Settings
  $('open-settings').addEventListener('click',openSettings);$('connection').addEventListener('click',openSettings);
  $('close-settings').addEventListener('click',closeSettings);
  $('drawer').addEventListener('click',e=>{if(e.target===$('drawer'))closeSettings();});
  $('default-profile').addEventListener('change',()=>saveSetting({default_profile:$('default-profile').value}));
  $('security-guard').addEventListener('change',()=>{saveSetting({security_guard:$('security-guard').checked});renderReview();});
  $('save-history').addEventListener('change',()=>saveSetting({save_history:$('save-history').checked}));
  $('autostart').addEventListener('change',async()=>{try{const a=await api('autostart',{enabled:$('autostart').checked});$('autostart').checked=a.enabled;settingsStatus(a.enabled?'Email Studio will start when you sign in to Windows.':'Start at sign-in turned off.');}catch(error){$('autostart').checked=!$('autostart').checked;settingsStatus(error.message,true);}});
  $('edit-profile-pick').addEventListener('change',()=>editProfile($('edit-profile-pick').value));
  $('new-profile').addEventListener('click',()=>{editProfile(null);$('pf-name').focus();});
  $('profile-form').addEventListener('submit',saveProfile);
  $('pf-add-example').addEventListener('click',()=>{if(all('#pf-examples .example').length<3)$('pf-examples').append(exampleBox());});
  $('pf-from-email').addEventListener('click',()=>{
    const text=$('source').value.trim();if(!text)return settingsStatus('Open one of your sent emails first.',true);
    if(all('#pf-examples .example').length>=3)return settingsStatus('Remove an example first (maximum 3).',true);
    $('pf-examples').append(exampleBox(text.slice(0,5000)));settingsStatus('Added. Remove any quoted older messages, then Save profile.');
  });
  $('pf-delete').addEventListener('click',async()=>{if(!state.editing)return;try{await api('profiles/delete',{id:state.editing});state.editing=null;await loadProfiles();settingsStatus('Profile deleted.');}catch(error){settingsStatus(error.message,true);}});
  $('save-template').addEventListener('click',async()=>{try{await api('templates',{name:$('template-name').value,body:$('template-body').value});$('template-name').value=$('template-body').value='';await loadTemplates();settingsStatus('Template saved on this computer.');}catch(error){settingsStatus(error.message,true);}});
  $('signin').addEventListener('click',async()=>{try{await api('chatgpt/start',{});settingsStatus('Complete sign-in in your browser.');const poll=async()=>{try{const value=await api('chatgpt/status');connection(value);if(value.auth.pending)setTimeout(poll,1800);else settingsStatus(value.auth.message||'');}catch(error){settingsStatus(error.message,true);}};setTimeout(poll,1800);}catch(error){settingsStatus(error.message,true);}});
  setSeg('length',prefs.get('length','Standard'));
}

function boot(){
  wire();setMode('reply');syncButtons();
  loadSettings().then(()=>{const l=prefs.get('length',null);if(l)setSeg('length',l);});
  const browserMode=()=>{$('host-status').textContent='Browser mode · drafts open as new Outlook Classic emails. Use the Outlook ribbon button for replies.';syncButtons();};
  if(window.Office&&Office.onReady){
    Office.onReady(async info=>{
      if(info&&info.host===Office.HostType.Outlook){
        applyOfficeTheme();
        await readCurrent();
        try{Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged,()=>{onItemChanged();});}catch(error){/* older hosts: reload button */}
      }else browserMode();
    });
  }else browserMode();
}
boot();
