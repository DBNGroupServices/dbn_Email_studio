/* Unit tests for the hosted add-in (no network): prompt building, file
   handling, response parsing, error messages and mailbox settings storage. */
const assert=require('node:assert/strict');
const fs=require('node:fs');
const path=require('node:path');
const C=require('../addin/claude-client.js');
const S=require('../addin/store.js');
const ROOT=path.resolve(__dirname,'..');

// Shared helpers must stay identical to the v2 pane's copies.
for(const f of ['draft-tools.js','office-client.js'])
  assert.equal(fs.readFileSync(path.join(ROOT,'addin',f),'utf8'),fs.readFileSync(path.join(ROOT,'web','office',f),'utf8'),f+' differs between addin/ and web/office/');

const b64=s=>Buffer.from(s).toString('base64');
const settings={...S.DEFAULT_SETTINGS,name:'Sam Fictional'};
const profile=S.validateProfile({name:'Tenant',guidance:'Reassuring',formality:2,warmth:5,directness:3,examples:['Hi Jo, sorted. Cheers, S'],signature:'Sam\nFictional Pty Ltd'});

// Prompt / request
let r=C.buildRequest({mode:'reply',source:'Fictional email',intent:'confirm',tone:'Use profile',sliders:{formality:1},
  meta:{from:{name:'Jo',email:'jo@example.com'},to:[],cc:[],reply_all:true}},settings,profile);
let ctx=JSON.parse(r.content[0].text);
assert.equal(ctx.task,'reply');assert.match(ctx.quick_intent,/Confirm/);
assert.deepEqual(ctx.tone_controls,{formality:1,warmth:5,directness:3});
assert.deepEqual(ctx.voice_profile.example_emails,['Hi Jo, sorted. Cheers, S']);
assert.equal(ctx.sender.signature,'Sam\nFictional Pty Ltd');assert.equal(ctx.email_metadata.reply_all,true);
assert.match(r.system,/untrusted reference material/);assert.match(r.system,/Security disclosure guard/);
assert.equal(r.schema,C.SCHEMA);assert.equal(r.schema.additionalProperties,false);
r=C.buildRequest({mode:'reply',source:'x',variants:3},{...settings,security_guard:false},null);
assert.equal(r.schema,C.VARIANT_SCHEMA);assert.equal(JSON.parse(r.content[0].text).number_of_variants,3);assert.doesNotMatch(r.system,/Security disclosure guard/);
r=C.buildRequest({mode:'compose',points:'x'},settings,S.validateProfile({name:'n',signature:'Sig',include_signature:false}));
assert.equal(JSON.parse(r.content[0].text).sender.signature,'');
r=C.buildRequest({action:'translate',source:'Bonjour',language:'French'},settings,null);
assert.match(r.system,/French/);assert.match(r.system,/Translate/);assert.equal(r.variants,1);
for(const bad of [{mode:'send',source:'x'},{mode:'reply',source:'x',intent:'nuke'},{mode:'reply',source:'x',tone:'Sarcastic'},{mode:'reply',source:'x',variants:4},{mode:'reply'},{action:'hack',source:'x'}])
  assert.throws(()=>C.buildRequest(bad,settings,null),C.StudioError,JSON.stringify(bad));

// Files
const png='\x89PNG\r\n\x1a\n'+'rest';
let f=C.fileBlocks([{name:'shot.png',data:Buffer.from(png,'latin1').toString('base64')},{name:'a.pdf',data:b64('%PDF-1.7 fake')},{name:'n.txt',data:b64('hello')},
  {name:'p.html',data:b64('<p>Hi&amp;bye</p><script>x</script>')},{name:'w.docx',data:b64('PK')},{name:'fake.jpg',data:b64('notimage')}]);
assert.deepEqual(f.blocks.filter(b=>b.type==='image').map(b=>b.source.media_type),['image/png']);
assert.equal(f.blocks.find(b=>b.type==='document').source.media_type,'application/pdf');
assert.ok(f.blocks.some(b=>b.type==='text'&&b.text==='REFERENCE FILE: n.txt\nhello'));
assert.ok(f.blocks.some(b=>b.type==='text'&&b.text==='REFERENCE FILE: p.html\nHi&bye'));
assert.equal(f.warnings.length,2);assert.match(f.warnings.join(' '),/DOCX/);assert.match(f.warnings.join(' '),/fake\.jpg/);
assert.throws(()=>C.fileBlocks(Array(11).fill({name:'a.txt',data:b64('x')})),/up to 10/);

// Responses
const msg=(text,stop='end_turn')=>({stop_reason:stop,model:'claude-opus-5-5',content:[{type:'thinking',thinking:''},{type:'text',text}]});
assert.deepEqual(C.readResult(msg(JSON.stringify({subject:'S',body:'B',checks:[]}))),{subject:'S',body:'B',checks:[]});
const v=C.readResult(msg(JSON.stringify({variants:[{label:'A',subject:'s',body:'b1'},{label:'B',subject:'s',body:'b2'}],checks:['c']})));
assert.equal(v.body,'b1');assert.equal(v.alternatives.length,2);
assert.throws(()=>C.readResult(msg('{}','refusal')),/declined/);
assert.throws(()=>C.readResult(msg('{"subj','max_tokens')),/cut off/);
assert.throws(()=>C.readResult(msg('not json')),/unreadable/);
assert.throws(()=>C.readResult(msg(JSON.stringify({subject:'S',body:'  ',checks:[]}))),/unreadable/);

// Errors -> plain-English messages
class APIError extends Error{constructor(status,m){super(m);this.status=status;this.requestID='req_123';}}
class BadRequestError extends APIError{}class AuthenticationError extends APIError{}class RateLimitError extends APIError{}
class APIConnectionError extends Error{}class InternalServerError extends APIError{}
const SDK={APIError,BadRequestError,AuthenticationError,RateLimitError,APIConnectionError,InternalServerError};
assert.match(C.friendlyError(new AuthenticationError(401,'x'),SDK).message,/key was rejected/);
assert.equal(C.friendlyError(new AuthenticationError(401,'x'),SDK).detail,'Request ID: req_123');
assert.match(C.friendlyError(new BadRequestError(400,'Your credit balance is too low'),SDK).message,/no API credit/);
assert.match(C.friendlyError(new RateLimitError(429,'x'),SDK).message,/rate limit/);
assert.match(C.friendlyError(new APIConnectionError('x'),SDK).message,/api\.anthropic\.com/);
assert.match(C.friendlyError(new InternalServerError(529,'overloaded'),SDK).message,/busy/);

// Store with mocked Outlook roaming settings
(async()=>{
  const bag={};let saves=0;
  global.Office={context:{roamingSettings:{get:k=>bag[k],set:(k,v)=>{bag[k]=v;},remove:k=>{delete bag[k];},saveAsync:cb=>{saves++;cb({status:'succeeded'});}}}};
  const st=S.create();assert.equal(st.kind,'mailbox');
  assert.equal(st.profiles().length,5);assert.equal(st.settings().model,'claude-opus-5-5');
  await assert.rejects(()=>st.setApiKey('sk-wrong'),/sk-ant-/);
  await st.setApiKey('sk-ant-api03-'+'a'.repeat(40));assert.equal(st.apiKey().length,53);assert.equal(saves,1);
  await st.saveSettings({name:'Sam',security_guard:false});assert.equal(st.settings().security_guard,false);
  await assert.rejects(()=>st.saveSettings({security_guard:'no'}),/on or off/);
  const {saved}=await st.saveProfile({name:'Mine',warmth:5,examples:['x']});assert.equal(st.profiles().length,6);
  await st.saveProfile({...saved,name:'Renamed'});assert.equal(st.profiles().find(p=>p.id===saved.id).name,'Renamed');
  await st.deleteProfile(saved.id);assert.equal(st.profiles().length,5);
  for(const bad of [{name:''},{name:'x',formality:9},{name:'x',examples:['a','b','c','d']},{name:'x',id:'../x'},{name:'x',length:'Epic'}])
    await assert.rejects(()=>st.saveProfile(bad),Error,JSON.stringify(bad));
  const t=await st.saveTemplate('T','Body');await st.deleteTemplate(t[0].id);assert.deepEqual(st.templates(),[]);
  // 32 KB limit is enforced and the previous value is restored
  for(let i=0;i<9;i++)await st.saveProfile({name:'Big '+i,examples:['y'.repeat(2500)]});
  await assert.rejects(()=>st.saveProfile({name:'Overflow',examples:['z'.repeat(2500),'z'.repeat(2500),'z'.repeat(2500)]}),/storage is full/);
  assert.ok(!st.profiles().some(p=>p.name==='Overflow'));assert.ok(st.usage().used<=S.LIMIT);
  await st.setApiKey('');assert.equal(st.apiKey(),'');
  // Manifest points at HTTPS hosting with the same add-in ID as before
  const manifest=fs.readFileSync(path.join(ROOT,'addin','manifest.xml'),'utf8');
  assert.match(manifest,/<Id>7877fe62-01f2-4af8-b073-dfa805099a63<\/Id>/);assert.doesNotMatch(manifest,/localhost/);
  assert.match(manifest,/https:\/\/dbngroupservices\.github\.io\/dbn_Email_studio\/taskpane\.html/);assert.match(manifest,/<Permissions>ReadWriteItem<\/Permissions>/);
  const html=fs.readFileSync(path.join(ROOT,'addin','taskpane.html'),'utf8');
  for(const src of html.match(/src="([^"]+)"/g).map(s=>s.slice(5,-1)).filter(s=>!s.startsWith('https://')))assert.ok(fs.existsSync(path.join(ROOT,'addin',src)),'missing '+src);
  assert.match(html,/connect-src 'self' https:\/\/api\.anthropic\.com/);
  console.log('Hosted add-in: prompts, files, responses, errors, mailbox storage limits and manifest passed.');
})().catch(e=>{console.error(e);process.exitCode=1;});
