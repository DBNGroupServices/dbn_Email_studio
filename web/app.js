'use strict';

const $ = (id) => document.getElementById(id);
const token = document.querySelector('meta[name="app-token"]').content;
const sessionApi = DBNSession.create(token);
const modelNames = {'gpt-6.1-sol': 'GPT-6.1 Sol', 'gpt-6-astra': 'GPT-6 Astra', 'gpt-6.1-terra': 'GPT-6.1 Terra'};
const supported = new Set(['png','jpg','jpeg','webp','pdf','docx','eml','msg','txt','html','htm']);
let state = {mode:'reply', files:[], settings:null, hasKey:false, busy:false, attaching:false, result:null, history:[], outlookContext:'', openingOutlook:false};
let toastTimer;
let busyTimer;
let fileSequence = Promise.resolve();
let authPoll;
let outlookPicker;
// Refresh only this product's native window branding; this does not reload work.
api('windows/icon',{}).catch(()=>{});

async function api(path, data) {
  let response, value;
  try {
    ({response,value} = await sessionApi.request(path,data));
  } catch (_) {
    throw new Error('The app is not reachable. Reopen DBN Email Studio and try again.');
  }
  if (!response.ok) {
    const failure = new Error(value.error || 'This request could not be completed.');
    failure.diagnostics = value.diagnostics;
    throw failure;
  }
  return value;
}

function diagnosticText(details) {
  if (!details) return '';
  return [['code','OpenAI code'],['organisation','Organisation'],['project','Project'],['request_id','Request ID']]
    .filter(([key]) => details[key]).map(([key,label]) => label + ': ' + details[key]).join('\n');
}
function error(message, details) {
  $('error-message').textContent = message;
  $('error-details').hidden = !diagnosticText(details);
  $('error-diagnostics').textContent = diagnosticText(details);
  $('error-manage-usage').hidden = details?.code !== 'subscription_sharing_usage_limit_exceeded';
  $('app-error').hidden = false;
}
function toast(message) {
  clearTimeout(toastTimer);
  $('toast').textContent = message;
  $('toast').hidden = false;
  toastTimer = setTimeout(() => {$('toast').hidden = true;}, 3600);
}
function settingsFeedback(message, failed=false, details) {
  $('settings-feedback').textContent = message + (diagnosticText(details) ? '\n\n' + diagnosticText(details) : '');
  $('settings-feedback').classList.toggle('error',failed);
  $('settings-feedback').hidden = false;
  $('settings-feedback').scrollIntoView({block:'nearest'});
}
function updateConnection(snapshot) {
  state.settings = snapshot.settings;
  state.hasKey = snapshot.connected;
  state.chatgpt = snapshot.chatgpt;
  const plan = snapshot.settings.provider === 'chatgpt';
  if (plan) Object.assign(modelNames,snapshot.models);
  $('connection').classList.toggle('connected',snapshot.connected);
  $('connection-label').textContent = snapshot.connected ? (plan ? 'Using ChatGPT plan' : 'Using API billing') : (plan ? 'Connect ChatGPT' : 'Connect your API');
  $('setup-banner').hidden = snapshot.connected;
  $('model-label').textContent = modelNames[snapshot.settings.model] || snapshot.settings.model;
  $('key-status').textContent = snapshot.has_key ? 'Key already available' : 'Not connected yet';
  $('provider').value = snapshot.settings.provider;
  $('chatgpt-connection').hidden = !plan;
  $('api-connection').hidden = plan;
  $('forget-key').hidden = plan;
  $('model').replaceChildren();
  Object.entries(snapshot.models).forEach(([id,label]) => {
    const option=document.createElement('option');option.value=id;option.textContent=plan?label:(modelNames[id]||id)+' · '+label;$('model').append(option);
  });
  if (!$('model').options.length) {
    const option=document.createElement('option');option.value=snapshot.settings.model;option.textContent='Sign in to load your models';$('model').append(option);
  }
  $('model').value = snapshot.settings.model;
  const account=snapshot.chatgpt;
  $('chatgpt-account').textContent = account.signed_in ? (account.email || account.name || 'ChatGPT account connected') : 'Sign in to connect your plan.';
  $('chatgpt-signout').hidden = !account.signed_in;
  $('chatgpt-accounts').replaceChildren();
  account.accounts.forEach(item=>{const option=document.createElement('option');option.value=item.id;option.textContent=item.email||item.name||'Saved ChatGPT account';$('chatgpt-accounts').append(option);});
  if (!account.accounts.length) {const option=document.createElement('option');option.value='';option.textContent='Add an account';$('chatgpt-accounts').append(option);}
  $('chatgpt-accounts').value=account.active;
  $('chatgpt-models').disabled=!account.enabled;
  $('test-drafting').disabled=!snapshot.connected;
  $('billing-label').textContent=plan?(snapshot.connected?'Using ChatGPT plan':'ChatGPT plan · sign in required'):'Using separate API billing';
  $('compose-usage').hidden=!plan;
  if (plan&&account.enabled&&!account.welcomed&&!$('plan-welcome').open) $('plan-welcome').showModal();
}
function setMode(mode) {
  state.mode = mode;
  document.querySelectorAll('[data-mode]').forEach((button) => {
    const selected = button.dataset.mode === mode;
    button.classList.toggle('selected',selected);
    button.setAttribute('aria-pressed',String(selected));
  });
  const labels = {
    reply:['The email you received','Paste the email here, or add a screenshot below…','What would you like to say?','Create my reply'],
    rewrite:['Your current draft','Paste the draft you would like rewritten…','Anything to change or keep?','Rewrite my email'],
    compose:['Background or reference material','Add background or a reference email, if helpful…','What should your email cover?','Write my email']
  }[mode];
  $('source-label').firstChild.textContent = labels[0] + ' ';
  $('source').placeholder = labels[1];
  $('points-label').firstChild.textContent = labels[2] + ' ';
  $('generate-label').textContent = labels[3];
  updateOutlookLabel();
  saveControls();
}
function saveControls() {
  try {localStorage.setItem('dbn-email-controls',JSON.stringify({mode:state.mode,tone:$('tone').value,length:$('length').value,format:$('format').value}));} catch (_) {}
}
function restoreControls() {
  try {
    const saved = JSON.parse(localStorage.getItem('dbn-email-controls') || '{}');
    ['tone','length','format'].forEach((key) => {if ([...$(key).options].some(o => o.value === saved[key])) $(key).value = saved[key];});
    setMode(['reply','rewrite','compose'].includes(saved.mode) ? saved.mode : 'reply');
  } catch (_) {setMode('reply');}
}
function setBusy(busy, refining=false) {
  state.busy = busy;
  document.body.classList.toggle('busy',busy);
  ['generate','new-draft','history-nav','sample-button','setup-button','settings-nav','connection','source','points','recipient','tone','length','format','file-input','refine','refinement','outlook-import','outlook-open'].forEach(id => $(id).disabled = busy);
  document.querySelectorAll('[data-mode],[data-refine],.attachment-remove,.attachment-preview').forEach(button => button.disabled = busy);
  $('drop-zone').setAttribute('aria-disabled',String(busy));
  $('busy-state').hidden = !busy;
  $('empty-state').hidden = busy || !!state.result;
  $('draft-result').hidden = busy || !state.result;
  $('output-badge').textContent = busy ? 'WRITING…' : state.result ? 'READY TO REVIEW' : 'READY WHEN YOU ARE';
  clearTimeout(busyTimer);
  if (busy) {
    $('busy-label').textContent = refining ? 'Applying your changes to the current draft.' : 'Reading your email and writing your draft.';
    busyTimer = setTimeout(() => {$('busy-label').textContent = 'Still working. Longer emails and documents can take a little more time.';},25000);
  }
}

function showDraft(result) {
  state.result = result;
  $('draft-subject').value = result.subject;
  $('draft-body').value = result.body;
  $('draft-model').textContent = modelNames[result.model] || result.model || '';
  const checks = [...(result.checks || []), ...(result.warnings || [])];
  $('checks-list').replaceChildren();
  checks.forEach(note => {const li = document.createElement('li'); li.textContent = note; $('checks-list').append(li);});
  $('checks').hidden = checks.length === 0;
  $('draft-result').hidden = false;
  $('empty-state').hidden = true;
  $('output-badge').textContent = 'READY TO REVIEW';
  updateWordCount();
}
function updateWordCount() {
  const text = $('draft-body').value.trim();
  $('word-count').textContent = (text ? text.split(/\s+/).length : 0) + ' words';
}
async function generate(refining=false, instruction='') {
  if (state.busy) return;
  $('app-error').hidden = true;
  if (state.attaching) return toast('Your files are still being read. Try again in a moment.');
  if (!state.settings) return error('The app is still loading. If it stays here, reopen DBN Email Studio.');
  if (!state.hasKey) return openSettings();
  const refinement = refining ? instruction || $('refinement').value.trim() : '';
  if (refining && !refinement) return toast('Add the change you would like, or choose a suggestion.');
  if (!refining && !$('source').value.trim() && !$('points').value.trim() && !state.files.length) {
    error('Paste an email, attach a file or add your dot points first.');
    $('points').focus();
    return;
  }
  const payload = {mode:state.mode,source:$('source').value,points:$('points').value,recipient:$('recipient').value,
    tone:$('tone').value,length:$('length').value,format:$('format').value,model:state.settings.model,
    files:state.files.map(({name,data})=>({name,data})),refinement,
    current:refining ? {subject:$('draft-subject').value,body:$('draft-body').value} : {}};
  setBusy(true,refining);
  try {
    const result = await api('generate',payload);
    showDraft(result);
    if (refining) $('refinement').value = '';
    await loadHistory();
    toast(refining ? 'Draft updated.' : 'Your draft is ready.');
  } catch (err) {error(err.message, err.diagnostics);}
  finally {setBusy(false);}
}

function readFile(file) {
  return new Promise((resolve,reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result).split(',')[1]);
    reader.onerror = () => reject(new Error('Could not read ' + file.name + '. Try selecting it again.'));
    reader.readAsDataURL(file);
  });
}
function queueFiles(files) {
  fileSequence = fileSequence.then(() => addFiles(files)).catch(err=>error(err.message));
}
async function addFiles(files) {
  if (state.busy || !files.length) return;
  state.attaching = true;
  $('generate').disabled = true;
  $('upload-progress').hidden = false;
  try {
    let rejected = [];
    for (const file of files) {
      $('upload-progress').textContent = 'Reading ' + file.name + '…';
      let ext = file.name.split('.').pop().toLowerCase();
      if (!supported.has(ext)) {rejected.push(file.name + ': unsupported file type.'); continue;}
      if (!file.size || file.size > 10 * 1024 * 1024) {rejected.push(file.name + ': use a non-empty file under 10 MB.'); continue;}
      if (state.files.length >= 10 || state.files.reduce((sum,f)=>sum+f.size,0) + file.size > 25 * 1024 * 1024) {rejected.push('Use up to 10 files, totalling no more than 25 MB.'); break;}
      const data = await readFile(file);
      const inspected = await api('inspect',{files:[{name:file.name,data}]});
      ext = {'image/png':'png','image/jpeg':'jpg','image/webp':'webp'}[inspected.files[0]?.mime] || ext;
      state.files.push({id:crypto.randomUUID(),name:file.name,size:file.size,data,ext,inspected});
      renderAttachments();
    }
    if (rejected.length) error(rejected.join(' '));
  } catch (err) {error(err.message);}
  finally {state.attaching = false; $('upload-progress').hidden = true; $('generate').disabled = state.busy; $('file-input').value = '';}
}
function renderAttachments() {
  $('attachments').replaceChildren();
  state.files.forEach(file => {
    const row = document.createElement('div'); row.className = 'attachment';
    const kind = document.createElement('span'); kind.className = 'attachment-type'; kind.textContent = file.ext.toUpperCase();
    const preview = document.createElement('button'); preview.className = 'attachment-preview'; preview.title = 'Preview ' + file.name;
    const name = document.createElement('strong'); name.textContent = file.name;
    const info = document.createElement('small'); info.textContent = (file.size < 1024*1024 ? Math.ceil(file.size/1024)+' KB' : (file.size/1024/1024).toFixed(1)+' MB') + ' · click to preview';
    preview.append(name,info); preview.addEventListener('click',()=>previewFile(file));
    const remove = document.createElement('button'); remove.className = 'attachment-remove'; remove.textContent = '×'; remove.setAttribute('aria-label','Remove '+file.name);
    remove.addEventListener('click',()=>{state.files=state.files.filter(f=>f.id!==file.id);renderAttachments();});
    row.append(kind,preview,remove); $('attachments').append(row);
  });
  const warnings = state.files.flatMap(file=>file.inspected.warnings);
  $('attachment-notes').hidden = !warnings.length;
  $('attachment-notes').textContent = warnings.join('\n');
}
function previewFile(file) {
  $('preview-title').textContent = file.name;
  const image = ['png','jpg','jpeg','webp'].includes(file.ext);
  $('preview-image').hidden = !image;
  $('preview-text').hidden = image;
  if (image) {
    const mime = file.inspected.files[0]?.mime || {png:'image/png',jpg:'image/jpeg',jpeg:'image/jpeg',webp:'image/webp'}[file.ext];
    $('preview-image').src = 'data:' + mime + ';base64,' + file.data;
  } else {
    $('preview-image').removeAttribute('src');
    $('preview-text').textContent = file.inspected.files.map(f=>f.name+'\n\n'+f.preview).join('\n\n────────────────────\n\n');
  }
  $('preview-dialog').showModal();
}

function switchView(view) {
  if (state.busy) return;
  const history = view === 'history';
  $('compose-view').hidden = history;
  $('history-view').hidden = !history;
  $('compose-nav').classList.toggle('active',!history);
  $('history-nav').classList.toggle('active',history);
  $('page-label').textContent = history ? 'Recent drafts' : 'Compose';
  if (history) loadHistory().catch(err=>error(err.message));
}
async function loadHistory() {
  state.history = await api('history');
  $('history-count').textContent = state.history.length;
  $('history-list').replaceChildren();
  if (!state.history.length) {
    const empty = document.createElement('div'); empty.className='history-empty';
    const title = document.createElement('h3'); title.textContent = 'A fresh start.';
    const note = document.createElement('p'); note.textContent = 'Your generated drafts will appear here when local history is enabled.';
    empty.append(title,note); $('history-list').append(empty); return;
  }
  state.history.forEach(record=>{
    const button = document.createElement('button'); button.className='history-card';
    const subject = document.createElement('strong'); subject.textContent = record.subject || 'Untitled email';
    const preview = document.createElement('p'); preview.textContent = record.body.slice(0,200);
    const meta = document.createElement('small'); meta.textContent = new Date(record.created).toLocaleString('en-AU',{dateStyle:'medium',timeStyle:'short'}) + ' · ' + (modelNames[record.model] || record.model);
    button.append(subject,preview,meta);
    button.addEventListener('click',()=>{
      if (hasWork() && !confirm('Open this saved draft? The email currently in your workspace will be cleared.')) return;
      resetWorkspace();
      setMode(record.mode || 'reply');
      showDraft(record); switchView('compose');
      toast('Draft restored. Add source files again if a revision needs them.');
    });
    $('history-list').append(button);
  });
}
function hasWork() {return !!($('source').value || $('points').value || state.files.length || state.result);}
function resetWorkspace() {
  ['source','points','recipient','refinement','draft-subject','draft-body'].forEach(id=>$(id).value='');
  state.files=[];state.result=null;state.outlookContext='';renderAttachments();updateOutlookLabel();
  $('empty-state').hidden=false; $('draft-result').hidden=true; $('checks').hidden=true;
  $('output-badge').textContent='READY WHEN YOU ARE'; $('app-error').hidden=true;
}
function newDraft() {
  if (state.busy || state.attaching) return;
  if (hasWork() && !confirm('Start a new email? The current workspace will be cleared. Generated drafts already saved to history will stay there.')) return;
  resetWorkspace();switchView('compose');$('source').focus();
}

function updateOutlookLabel() {
  const reply = !!state.outlookContext && state.mode === 'reply';
  $('outlook-open-label').textContent = reply ? 'Open reply in Outlook Classic' : 'Open in Outlook Classic';
  $('outlook-export-note').textContent = reply ? 'Replies to the imported email. Opens for review; nothing is sent automatically.' : 'Opens a new, editable email for you to address, review and send.';
}

async function importOutlook() {
  if (state.busy || state.attaching) return;
  outlookPicker={session:'',data:null,selected:'',busy:false,importing:false};
  $('outlook-search').value='';$('outlook-email-list').replaceChildren();$('outlook-folders').replaceChildren();
  $('outlook-mailbox').replaceChildren();
  $('outlook-folder-path').textContent='';$('outlook-page-label').textContent='';
  $('outlook-picker-dialog').showModal();
  await browseOutlook();
}

function pickerBusy(busy) {
  const picker=outlookPicker;
  if (!picker)return;
  picker.busy=busy;
  ['outlook-mailbox','outlook-inbox','outlook-parent','outlook-search','outlook-search-button','outlook-refresh','outlook-previous','outlook-next'].forEach(id=>$(id).disabled=busy);
  $('outlook-email-list').disabled=busy;
  $('outlook-folders').querySelectorAll('button').forEach(button=>button.disabled=busy);
  $('outlook-picker-import').disabled=busy||!picker.selected;
  $('outlook-picker-close').disabled=picker.importing;
  $('outlook-picker-cancel').disabled=picker.importing;
  if(!busy){
    $('outlook-inbox').disabled=!picker.data?.inbox;
    $('outlook-parent').disabled=!picker.data?.parent;
    $('outlook-previous').disabled=!(picker.data?.offset>0);
    $('outlook-next').disabled=!picker.data?.has_more;
  }
  $('outlook-email-list').setAttribute('aria-busy',String(busy));
}

async function browseOutlook(folder='',offset=0,query) {
  const picker=outlookPicker;
  if(!picker||picker.busy)return;
  picker.selected='';pickerBusy(true);
  $('outlook-email-list').querySelectorAll('input').forEach(input=>input.checked=false);
  $('outlook-picker-status').textContent='Loading emails from Outlook Classic…';
  $('outlook-picker-status').classList.remove('error');
  try {
    const result=await api('outlook/browse',{picker:picker.session,folder,offset,query:query===undefined?(picker.data?.query||''):query});
    if(outlookPicker!==picker||!$('outlook-picker-dialog').open)return;
    picker.session=result.picker;picker.data=result;
    renderOutlookPicker(result);
  } catch(err) {
    if(outlookPicker===picker){$('outlook-picker-status').textContent=err.message;$('outlook-picker-status').classList.add('error');}
  } finally {if(outlookPicker===picker)pickerBusy(false);}
}

function renderOutlookPicker(data) {
  $('outlook-folder-path').textContent=data.folder?.path||data.folder?.name||'Outlook';
  const chosenMailbox=$('outlook-mailbox').value;
  $('outlook-mailbox').replaceChildren();
  data.mailboxes.forEach(folder=>{const option=document.createElement('option');option.value=folder.id;option.textContent=folder.name;$('outlook-mailbox').append(option);});
  if([...$('outlook-mailbox').options].some(option=>option.value===chosenMailbox))$('outlook-mailbox').value=chosenMailbox;
  const root=data.mailboxes.find(folder=>data.folder?.path===folder.path||data.folder?.path?.startsWith(folder.path+'\\'));
  if(root)$('outlook-mailbox').value=root.id;
  $('outlook-folders').replaceChildren();
  data.folders.forEach(folder=>{
    const button=document.createElement('button');button.type='button';button.className='outlook-folder';button.textContent=folder.name;
    button.addEventListener('click',()=>{$('outlook-search').value='';browseOutlook(folder.id,0,'');});$('outlook-folders').append(button);
  });
  if(!data.folders.length){const note=document.createElement('p');note.className='help';note.textContent='No subfolders. Use Parent folder to browse other folders.';$('outlook-folders').append(note);}
  $('outlook-email-list').replaceChildren();
  data.emails.forEach(email=>{
    const row=document.createElement('label');row.className='outlook-email'+(email.unread?' unread':'');
    const radio=document.createElement('input');radio.type='radio';radio.name='outlook-email';radio.value=email.id;
    radio.disabled=email.size>10*1024*1024;
    const content=document.createElement('span');content.className='outlook-email-content';
    const subject=document.createElement('strong');subject.textContent=email.subject||'(No subject)';
    const sender=document.createElement('span');sender.textContent=email.sender||'Unknown sender';
    const meta=document.createElement('small');
    const received=new Date(email.received);meta.textContent=(!Number.isNaN(received.getTime())?received.toLocaleString('en-AU',{dateStyle:'medium',timeStyle:'short'}):'No date')+(email.has_attachments?' · Attachments':'')+(radio.disabled?' · Over 10 MB':'');
    content.append(subject,sender,meta);row.append(radio,content);$('outlook-email-list').append(row);
    radio.addEventListener('change',()=>{outlookPicker.selected=radio.value;pickerBusy(false);});
  });
  $('outlook-picker-status').textContent=data.emails.length?'Newest emails first. Select one message to import.':(data.query?'No matching emails in this folder. Try another search or folder.':'No emails in this folder. Choose another folder.');
  $('outlook-page-label').textContent=data.emails.length?`${data.offset+1}–${data.offset+data.emails.length}`:'0 emails';
}

async function importPickedOutlook() {
  const picker=outlookPicker;
  if(!picker||picker.busy||!picker.selected)return;
  if (($('source').value || state.files.length || state.result) && !confirm('Replace the current source email and draft with the selected Outlook email? Your dot points will stay. Saved drafts remain in history.')) return;
  picker.importing=true;pickerBusy(true);
  $('outlook-picker-status').textContent='Importing the selected email…';
  state.attaching=true;
  $('outlook-import').disabled=true;$('generate').disabled=true;
  $('upload-progress').textContent='Reading the selected email from Outlook Classic…';$('upload-progress').hidden=false;
  $('app-error').hidden=true;
  try {
    const result=await api('outlook/import',{picker:picker.session,message:picker.selected});
    const points=$('points').value;
    resetWorkspace();$('points').value=points;setMode('reply');
    const inspected={files:result.files,warnings:result.warnings};
    state.files=[{id:crypto.randomUUID(),name:result.file.name,data:result.file.data,size:result.files[0].size,ext:'msg',inspected}];
    state.outlookContext=result.context || '';
    renderAttachments();updateOutlookLabel();
    $('outlook-picker-dialog').close();
    toast('Email imported from Outlook. Add your points, then create a reply.');
  } catch(err) {$('outlook-picker-status').textContent=err.message;$('outlook-picker-status').classList.add('error');}
  finally {state.attaching=false;picker.importing=false;if(outlookPicker===picker)pickerBusy(false);$('outlook-import').disabled=state.busy;$('generate').disabled=state.busy;$('upload-progress').hidden=true;}
}

async function openOutlookDraft() {
  if (state.busy || state.openingOutlook) return;
  state.openingOutlook=true;$('outlook-open').disabled=true;$('app-error').hidden=true;
  try {
    const result=await api('outlook/draft',{subject:$('draft-subject').value,body:$('draft-body').value,
      context:state.mode==='reply'?state.outlookContext:''});
    toast(result.action==='reply'?'Reply opened in Outlook for your review.':'Email draft opened in Outlook for your review.');
  } catch(err) {error(err.message);}
  finally {state.openingOutlook=false;$('outlook-open').disabled=state.busy;}
}

function openSettings() {
  if (state.busy) return;
  if (!state.settings) return error('Settings could not be loaded. Please reopen the app.');
  const values = {'model':'model','sender-name':'name','sender-role':'role','organisation':'organisation','signature':'signature','writing-style':'style'};
  Object.entries(values).forEach(([id,key])=>$(id).value=state.settings[key]);
  $('save-history').checked=state.settings.save_history;
  $('api-key').value=''; $('settings-feedback').hidden=true;
  $('settings-dialog').showModal();
}
async function saveSettings(event) {
  event.preventDefault();
  $('save-settings').disabled=true;
  try {
    const snapshot = await api('settings',{provider:$('provider').value,api_key:$('api-key').value.trim(),model:$('model').value,
      name:$('sender-name').value,role:$('sender-role').value,organisation:$('organisation').value,
      signature:$('signature').value,style:$('writing-style').value,save_history:$('save-history').checked});
    updateConnection(snapshot);$('api-key').value='';$('settings-dialog').close();toast('Settings saved. Applied to every new draft.');
  } catch(err) {settingsFeedback(err.message,true);}
  finally {$('save-settings').disabled=false;}
}
async function copyText(text) {
  if (!text.trim()) return toast('Create a draft first.');
  try {await navigator.clipboard.writeText(text); toast('Copied. Ready to paste into your email.');}
  catch (_) {error('Clipboard access was unavailable. Select the draft text and press Ctrl C.');}
}

$('generate').addEventListener('click',()=>generate());
$('outlook-import').addEventListener('click',importOutlook);
$('outlook-picker-import').addEventListener('click',importPickedOutlook);
['outlook-picker-close','outlook-picker-cancel'].forEach(id=>$(id).addEventListener('click',()=>{if(!outlookPicker?.importing)$('outlook-picker-dialog').close();}));
$('outlook-picker-dialog').addEventListener('cancel',event=>{if(outlookPicker?.importing)event.preventDefault();});
$('outlook-picker-dialog').addEventListener('close',()=>{outlookPicker=null;});
$('outlook-search-form').addEventListener('submit',event=>{event.preventDefault();browseOutlook(outlookPicker?.data?.folder?.id||'',0,$('outlook-search').value.trim());});
$('outlook-refresh').addEventListener('click',()=>browseOutlook(outlookPicker?.data?.folder?.id||'',0,$('outlook-search').value.trim()));
$('outlook-mailbox').addEventListener('change',()=>{$('outlook-search').value='';browseOutlook($('outlook-mailbox').value,0,'');});
$('outlook-inbox').addEventListener('click',()=>{$('outlook-search').value='';browseOutlook(outlookPicker?.data?.inbox?.id||'',0,'');});
$('outlook-parent').addEventListener('click',()=>{$('outlook-search').value='';browseOutlook(outlookPicker?.data?.parent?.id||'',0,'');});
$('outlook-previous').addEventListener('click',()=>browseOutlook(outlookPicker?.data?.folder?.id||'',Math.max(0,(outlookPicker?.data?.offset||0)-50)));
$('outlook-next').addEventListener('click',()=>browseOutlook(outlookPicker?.data?.folder?.id||'',(outlookPicker?.data?.offset||0)+50));
$('outlook-open').addEventListener('click',openOutlookDraft);
$('refine').addEventListener('click',()=>generate(true));
document.querySelectorAll('[data-refine]').forEach(button=>button.addEventListener('click',()=>generate(true,button.dataset.refine)));
document.querySelectorAll('[data-mode]').forEach(button=>button.addEventListener('click',()=>setMode(button.dataset.mode)));
['tone','length','format'].forEach(id=>$(id).addEventListener('change',saveControls));
$('new-draft').addEventListener('click',newDraft);
$('compose-nav').addEventListener('click',()=>switchView('compose'));
$('history-nav').addEventListener('click',()=>switchView('history'));
['settings-nav','connection','setup-button'].forEach(id=>$(id).addEventListener('click',openSettings));
$('close-settings').addEventListener('click',()=>$('settings-dialog').close());
$('settings-dialog').addEventListener('close',()=>{$('api-key').value='';});
$('settings-form').addEventListener('submit',saveSettings);
$('dismiss-error').addEventListener('click',()=>$('app-error').hidden=true);
$('close-preview').addEventListener('click',()=>{$('preview-dialog').close();$('preview-image').removeAttribute('src');});
$('draft-body').addEventListener('input',updateWordCount);
$('copy-body').addEventListener('click',()=>copyText($('draft-body').value));
$('copy-all').addEventListener('click',()=>copyText('Subject: '+$('draft-subject').value+'\n\n'+$('draft-body').value));
$('download').addEventListener('click',()=>{
  const blob=new Blob(['Subject: '+$('draft-subject').value+'\r\n\r\n'+$('draft-body').value.replace(/\r?\n/g,'\r\n')],{type:'text/plain;charset=utf-8'});
  const link=document.createElement('a');link.href=URL.createObjectURL(blob);
  link.download=($('draft-subject').value.replace(/[<>:"/\\|?*\x00-\x1f]/g,'').trim().slice(0,90)||'Email draft')+'.txt';
  link.click();setTimeout(()=>URL.revokeObjectURL(link.href),1000);
});
$('test-connection').addEventListener('click',async()=>{
  const button=$('test-connection');button.disabled=true;button.textContent='Checking…';
  let saved=false;
  try {
    const key=$('api-key').value.trim();
    if(key){updateConnection(await api('settings',{api_key:key,model:$('model').value}));$('api-key').value='';saved=true;}
    const result=await api('test',{model:$('model').value});
    settingsFeedback('Drafting test passed with '+modelNames[result.model]+'. Your key is saved. Use Save settings below for any other writing-preference changes.');
  }
  catch(err){settingsFeedback((saved?'Your key has been saved on this computer. ':'')+err.message,true,err.diagnostics);}
  finally{button.disabled=false;button.textContent='Save & test drafting';}
});
$('provider').addEventListener('change',async()=>{
  try {
    const provider=$('provider').value;
    const models=provider==='chatgpt'?state.chatgpt.models:{'gpt-6.1-sol':true,'gpt-6-astra':true,'gpt-6.1-terra':true};
    const model=Object.hasOwn(models,state.settings.model)?state.settings.model:(Object.keys(models)[0]||'gpt-6.1-sol');
    updateConnection(await api('settings',{provider,model}));
  } catch(err) {settingsFeedback(err.message,true);$('provider').value=state.settings.provider;}
});
async function pollChatGPT() {
  clearTimeout(authPoll);
  try {
    const snapshot=await api('chatgpt/status');updateConnection(snapshot);
    $('chatgpt-status').textContent=snapshot.auth.message;
    if(snapshot.auth.pending) authPoll=setTimeout(pollChatGPT,1800);
    else if(snapshot.chatgpt.enabled&&!snapshot.chatgpt.welcomed&&!$('plan-welcome').open) $('plan-welcome').showModal();
  } catch(err) {$('chatgpt-status').textContent=err.message;}
}
async function signInChatGPT(newAccount=false) {
  $('chatgpt-signin').disabled=true;$('chatgpt-status').textContent='Opening OpenAI sign-in in your browser…';
  try {
    const snapshot=await api('chatgpt/start',{new:newAccount,account:newAccount?'':$('chatgpt-accounts').value});
    updateConnection(snapshot);$('chatgpt-status').textContent=snapshot.auth.message;
    authPoll=setTimeout(pollChatGPT,1800);
  } catch(err) {$('chatgpt-status').textContent=err.message;}
  finally {$('chatgpt-signin').disabled=false;}
}
$('chatgpt-signin').addEventListener('click',()=>signInChatGPT(false));
$('chatgpt-new').addEventListener('click',()=>signInChatGPT(true));
$('chatgpt-models').addEventListener('click',async()=>{
  try {updateConnection(await api('chatgpt/models',{}));$('chatgpt-status').textContent='Models refreshed from your ChatGPT account.';}
  catch(err){$('chatgpt-status').textContent=err.message;}
});
$('chatgpt-signout').addEventListener('click',async()=>{
  clearTimeout(authPoll);
  try {const snapshot=await api('chatgpt/signout',{});updateConnection(snapshot);$('chatgpt-status').textContent=snapshot.auth.message;}
  catch(err){$('chatgpt-status').textContent=err.message;}
});
$('plan-welcome-close').addEventListener('click',async()=>{
  $('plan-welcome').close();
  try {updateConnection(await api('chatgpt/welcome',{}));}catch(err){error(err.message);}
});
$('test-drafting').addEventListener('click',async()=>{
  const button=$('test-drafting');button.disabled=true;button.textContent='Testing drafting…';
  try {await api('test',{model:$('model').value});settingsFeedback('Drafting test passed. Your selected connection can write emails.');}
  catch(err){settingsFeedback(err.message,true,err.diagnostics);}
  finally{button.disabled=false;button.textContent='Test drafting with fictional text';}
});
$('check-network').addEventListener('click',async()=>{
  const button=$('check-network');button.disabled=true;button.textContent='Checking…';
  try{const result=await api('network',{});settingsFeedback(result.message,!result.ok);}
  catch(err){settingsFeedback(err.message,true);}
  finally{button.disabled=false;button.textContent='Check network';}
});
$('forget-key').addEventListener('click',async()=>{
  try {const snapshot=await api('settings',{remove_key:true});updateConnection(snapshot);$('api-key').value='';settingsFeedback(snapshot.environment_key?'The saved key was removed. An OPENAI_API_KEY environment variable is still available.':'Saved API key removed.');}
  catch(err){settingsFeedback(err.message,true);}
});
$('clear-history').addEventListener('click',async()=>{
  if(!state.history.length)return;
  if(!confirm('Delete all saved drafts from this computer? This cannot be undone.'))return;
  try{await api('history/clear',{});await loadHistory();toast('Local draft history cleared.');}
  catch(err){error(err.message);}
});
$('shutdown').addEventListener('click',async()=>{
  if(hasWork()&&!confirm('Quit DBN Email Studio? Any unsaved workspace text will be lost.'))return;
  try{await api('shutdown',{});document.body.replaceChildren();const message=document.createElement('p');message.textContent='DBN Email Studio is closed. You can close this window. Open the app again whenever you need it.';message.className='history-empty';document.body.append(message);}
  catch(err){settingsFeedback(err.message,true);}
});

const drop=$('drop-zone');
function chooseFiles(){if(!state.busy)$('file-input').click();}
drop.addEventListener('click',chooseFiles);
drop.addEventListener('keydown',event=>{if(event.key==='Enter'||event.key===' '){event.preventDefault();chooseFiles();}});
$('file-input').addEventListener('change',event=>queueFiles([...event.target.files]));
drop.addEventListener('dragover',event=>{event.preventDefault();if(!state.busy)drop.classList.add('drag-over');});
drop.addEventListener('dragleave',()=>drop.classList.remove('drag-over'));
drop.addEventListener('drop',event=>{
  event.preventDefault();drop.classList.remove('drag-over');
  if(state.busy)return;
  const files=[...event.dataTransfer.files];
  if(files.length)queueFiles(files);
  else error('Outlook did not pass a readable message file. Use Import from Outlook Classic above, or save the email as MSG and attach that file.');
});
// Stop browser navigation when a file is accidentally dropped outside the target.
document.addEventListener('dragover',event=>{if(event.dataTransfer.types.includes('Files'))event.preventDefault();});
document.addEventListener('drop',event=>{if(event.dataTransfer.files.length)event.preventDefault();});
document.addEventListener('paste',event=>{
  if(state.busy||$('settings-dialog').open||$('preview-dialog').open||$('outlook-picker-dialog').open)return;
  const images=[...event.clipboardData.items].filter(item=>item.type.startsWith('image/')).map(item=>item.getAsFile()).filter(Boolean);
  if(images.length){event.preventDefault();queueFiles(images.map((file,index)=>{const ext={'image/png':'png','image/jpeg':'jpg','image/webp':'webp'}[file.type]||'png';return new File([file],'Screenshot-'+new Date().toISOString().replace(/[:.]/g,'-')+'-'+(index+1)+'.'+ext,{type:file.type});}));}
});
document.addEventListener('keydown',event=>{
  if((event.ctrlKey||event.metaKey)&&event.key.toLowerCase()==='n'&&!$('settings-dialog').open&&!$('outlook-picker-dialog').open){event.preventDefault();newDraft();}
  if((event.ctrlKey||event.metaKey)&&event.key==='Enter'&&!$('settings-dialog').open&&!$('outlook-picker-dialog').open){event.preventDefault();generate(document.activeElement===$('refinement'));}
});
window.addEventListener('beforeunload',event=>{if(state.busy||state.attaching){event.preventDefault();event.returnValue='';}});
$('sample-button').addEventListener('click',()=>{
  if(hasWork()&&!confirm('Load a sample email? This will clear the current workspace.'))return;
  resetWorkspace();setMode('reply');
  $('source').value='Hi,\n\nCould you please confirm whether you can provide two security officers for our tenant event on Friday, from 5:00 pm to 9:00 pm? We are expecting approximately 80 guests. Please also let us know what information you need from us.\n\nKind regards,\nAlex';
  $('points').value='• Thank Alex for the enquiry.\n• We will check officer availability and provide a quote.\n• Ask them to confirm the event date, location and any special access requirements.\n• Do not confirm the booking yet.';
  toast('Example loaded. Create a reply to try your AI connection.');
});

(async()=>{
  restoreControls();
  try{updateConnection(await api('settings'));await loadHistory();}
  catch(err){$('connection-label').textContent='Reconnect app';error(err.message);}
})();
