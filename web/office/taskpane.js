'use strict';
const $=id=>document.getElementById(id);
const token=document.querySelector('meta[name="app-token"]').content;
const sessionApi=DBNSession.create(token);
let snapshot,context,busy=false,applying=false,usingExample=false,resultContext,templates=[];
async function api(path,data){
  const {response,value}=await sessionApi.request(path,data);
  if(!response.ok)throw new Error(value.error||'This action could not be completed.');return value;
}
function status(text,failed=false){$('status').textContent=text;$('status').classList.toggle('error',failed);}
function connection(value){snapshot=value;$('connection-label').textContent=value.connected?(value.settings.provider==='chatgpt'?'Using ChatGPT plan':'Using separate API billing'):'Connect your ChatGPT plan';$('signin').hidden=value.connected;}
function syncApply(){
  const linked=resultContext;
  $('apply').textContent=applying?'Opening in Outlook…':linked?(linked.compose?'Insert at cursor':'Open reply in Outlook'):'Create draft in Outlook Classic';
  $('apply').disabled=busy||applying||$('result').hidden||!$('draft-body').value.trim();
  $('apply-help').textContent=linked?'Review the result before applying. Nothing is sent automatically.':'Creates a new editable email with no recipients. For a reply to an existing message, open Email Studio from the Outlook ribbon before generating. Nothing is sent automatically.';
}
async function readCurrent(){
  const previous=context,candidate=DBNOffice.target(window.Office);
  try{context=await DBNOffice.read(window.Office);usingExample=false;$('source').value=context.text;$('context-subject').textContent=context.subject;$('host-status').textContent=context.selectionWarning||(context.compose?'Compose mode · select text in your draft to rewrite only that selection.':'Reading the selected email.');}
  catch(error){context=previous||(!usingExample?candidate:null);$('host-status').textContent=error.message;}
  syncApply();
}
function setBusy(value){busy=value;['generate','read-email','action','tone','length','points','source','uploads','include-attachments','example'].forEach(id=>$(id).disabled=value);$('generate').textContent=value?'Working…':'Generate';syncApply();}
async function readUploads(){
  const files=[];for(const file of $('uploads').files){
    if(file.size>10*1024*1024)throw new Error(file.name+' is over 10 MB.');
    const data=await new Promise((resolve,reject)=>{const reader=new FileReader();reader.onload=()=>resolve(String(reader.result).split(',')[1]);reader.onerror=()=>reject(new Error('Could not read '+file.name));reader.readAsDataURL(file);});files.push({name:file.name,data});
  }return files;
}
$('generate').addEventListener('click',async()=>{
  if(busy)return;if(!snapshot?.connected)return status('Connect your ChatGPT plan before generating.',true);
  setBusy(true);status('Reading your context and preparing the result…');
  try{
    const action=$('action').value;
    const candidate=context||(!usingExample?DBNOffice.target(window.Office):null);
    const targetForResult=action==='compose'&&!candidate?.compose?null:candidate;
    const files=await readUploads();if($('include-attachments').checked)files.push(...await DBNOffice.attachments(context));
    if(files.length>10)throw new Error('Choose up to 10 files.');
    const task={source:$('source').value,points:$('points').value,mode:action,action,files,tone:$('tone').value,length:$('length').value,format:'Paragraphs',model:snapshot.settings.model,language:$('language').value};
    const result=await api(['reply','rewrite','compose'].includes(action)?'generate':'assist',task);
    resultContext=targetForResult;
    $('draft-subject').value=result.subject;$('draft-body').value=result.body;$('checks').replaceChildren();
    [...(result.checks||[]),...(result.warnings||[])].forEach(note=>{const p=document.createElement('p');p.textContent=note;$('checks').append(p);});
    $('result').hidden=false;status('Ready to review.');
  }catch(error){status(error.message,true);}finally{setBusy(false);}
});
$('action').addEventListener('change',()=>{$('language-row').hidden=$('action').value!=='translate';});
$('read-email').addEventListener('click',readCurrent);
$('example').addEventListener('click',()=>{context=null;usingExample=true;$('source').value='Hi Sam, could you confirm the meeting location? Thanks, Alex.';$('points').value='Fictional example: say we will confirm the location shortly. Do not invent a venue.';$('context-subject').textContent='Fictional example';$('host-status').textContent='Example mode. No Outlook email has been loaded.';});
$('draft-body').addEventListener('input',syncApply);
$('apply').addEventListener('click',async()=>{
  if(busy||applying||$('result').hidden||!$('draft-body').value.trim())return;
  applying=true;syncApply();
  try{
    if(resultContext)await DBNOffice.apply(window.Office,resultContext,$('draft-body').value);
    else await api('outlook/draft',{subject:$('draft-subject').value,body:$('draft-body').value});
    status(resultContext?'Opened in Outlook for your review. Nothing was sent.':'New draft opened in Outlook Classic. Add recipients and review it there. Nothing was sent.');
  }catch(error){status(error.message,true);}finally{applying=false;syncApply();}
});
$('copy').addEventListener('click',async()=>{try{await navigator.clipboard.writeText($('draft-body').value);status('Copied.');}catch(error){status('Select the result text and press Ctrl+C.',true);}});
async function loadTemplates(){templates=await api('templates');$('templates').replaceChildren(new Option('Choose a template',''));templates.forEach(t=>$('templates').append(new Option(t.name,t.id)));}
$('use-template').addEventListener('click',()=>{const template=templates.find(t=>t.id===$('templates').value);if(template)$('points').value=[$('points').value,template.body].filter(Boolean).join('\n\n');});
$('save-template').addEventListener('click',async()=>{try{await api('templates',{name:$('template-name').value,body:$('template-body').value});await loadTemplates();status('Template saved on this computer.');}catch(error){status(error.message,true);}});
$('signin').addEventListener('click',async()=>{try{await api('chatgpt/start',{});status('Complete sign-in in your browser.');const poll=async()=>{try{const value=await api('chatgpt/status');connection(value);if(value.auth.pending)setTimeout(poll,1800);else status(value.auth.message);}catch(error){status(error.message,true);}};setTimeout(poll,1800);}catch(error){status(error.message,true);}});
(async()=>{try{connection(await api('settings'));await loadTemplates();}catch(error){status('The local studio is not available. Open DBN Email Studio on this computer.',true);}})();
function browserMode(){$('host-status').textContent='Browser mode · create a new draft in Outlook Classic. Open Email Studio from the Outlook ribbon for replies or insertion into an existing draft.';syncApply();}
syncApply();
if(window.Office){Office.onReady(async info=>{if(info.host===Office.HostType.Outlook){await readCurrent();try{Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged,()=>{$('host-status').textContent='The Outlook selection changed. Choose Read current email to load it; your current work is kept.';});}catch(error){/* Older hosts can use Read current email manually. */}}else browserMode();});}
else browserMode();
