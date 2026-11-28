/* End-to-end test of the hosted add-in (v3.1 "one box" design) in Chromium:
   static site + mocked Outlook (Office.js) + mocked Claude API + mocked speech
   recognition. Run: node tests/test_addin_browser.js  (SCREENSHOT_DIR=... optional) */
const assert=require('node:assert/strict');
const http=require('node:http');
const fs=require('node:fs');
const path=require('node:path');
const {chromium}=require('playwright');
const ROOT=path.resolve(__dirname,'..','addin');
const TYPES={'.html':'text/html','.js':'text/javascript','.css':'text/css','.png':'image/png','.xml':'application/xml','.svg':'image/svg+xml','.woff2':'font/woff2'};

const OFFICE_MOCK=`
window.__calls=[];const bag={};
const item={itemId:'fixture-1',subject:'Your whitelabel email is not functioning',
  from:{displayName:'Jordan Example',emailAddress:'jordan@example.com'},
  to:[{displayName:'Sam',emailAddress:'sam@example.com'},{displayName:'Facilities',emailAddress:'fm@example.com'}],cc:[],
  dateTimeCreated:new Date('2026-11-23T05:42:00Z'),attachments:[],
  body:{getAsync:(type,cb)=>cb({status:'succeeded',value:'Hi Sam,\\n\\nYour sending domain is missing SPF and DKIM records. Can we re-test?\\n\\nThanks,\\nJordan'})},
  displayReplyForm(v){window.__calls.push(['reply',v.htmlBody]);},displayReplyAllForm(v){window.__calls.push(['replyAll',v.htmlBody]);}};
window.Office={HostType:{Outlook:'Outlook'},EventType:{ItemChanged:'olkItemSelectedChanged'},
  context:{officeTheme:{bodyBackgroundColor:'#ffffff'},roamingSettings:{get:k=>bag[k],set:(k,v)=>{bag[k]=v;},remove:k=>{delete bag[k];},saveAsync:cb=>cb({status:'succeeded'})},
    mailbox:{item,addHandlerAsync(t,h){window.__itemChanged=h;},displayNewMessageForm(v){window.__calls.push(['new',v.subject,v.htmlBody]);}}},
  onReady(cb){setTimeout(()=>cb({host:'Outlook'}),0);return Promise.resolve({host:'Outlook'});}};
window.__bag=bag;`;
const SPEECH_MOCK=`
window.SpeechRecognition=window.webkitSpeechRecognition=class{start(){window.__sr=this;setTimeout(()=>{this.onresult&&this.onresult({results:[Object.assign([{transcript:'propose re-testing Tuesday at ten'}],{isFinal:true})]});},30);}stop(){this.onend&&this.onend();}};`;

const reply={subject:'Re: whitelabel',body:'Hi Jordan,\n\nAgreed, we will add the records.\n\n- SPF record\n- DKIM record\n\nCould we re-test on [confirm date]?\n\nKind regards,\nSam',checks:['Test date not supplied.']};
const pv=t=>({formal:'Formal: '+t,balanced:'Balanced: '+t,friendly:'Friendly: '+t});
const suggestions={summary:'Jordan says your sending domain is missing SPF and DKIM records and asks to re-test.',audience:'contractor',priority:'High',need:'Fix + re-test',mood:'Neutral',
  suggestions:[{title:'Agree and propose a new date',tag:'BEST FIT',preview:pv('agreed'),hint:'Ask which day suits them',instruction:'Agree the records are missing and propose re-testing; ask which day suits.'},
    {title:'Ask them to fix it',tag:'FASTEST',preview:pv('please fix'),hint:'Request the records be added',instruction:'Ask them to add the records on their side.'},
    {title:'Acknowledge for now',tag:'HOLDING',preview:pv('on it'),hint:'Short holding reply',instruction:'Acknowledge and say we are looking into it.'}]};
const message=obj=>({id:'msg_1',type:'message',role:'assistant',model:'claude-opus-5-5',stop_reason:'end_turn',stop_sequence:null,
  content:[{type:'thinking',thinking:'',signature:'x'},{type:'text',text:JSON.stringify(obj)}],usage:{input_tokens:10,output_tokens:10}});

(async()=>{
  const server=http.createServer((req,res)=>{
    const file=path.join(ROOT,decodeURIComponent(req.url.split('?')[0]).replace(/^\/+/,'')||'index.html');
    if(!file.startsWith(ROOT)||!fs.existsSync(file)){res.writeHead(404);return res.end();}
    res.writeHead(200,{'Content-Type':TYPES[path.extname(file)]||'application/octet-stream'});fs.createReadStream(file).pipe(res);
  }).listen(0);
  const port=server.address().port;
  const browser=await chromium.launch();const errors=[],requests=[];
  let mode='ok';
  const route=async r=>{
    const req=r.request(),url=new URL(req.url());
    const cors={'access-control-allow-origin':'*','access-control-allow-headers':'*','content-type':'application/json'};
    if(req.method()==='OPTIONS')return r.fulfill({status:204,headers:cors});
    const entry={method:req.method(),path:url.pathname,headers:await req.allHeaders(),body:req.postData()?JSON.parse(req.postData()):null};requests.push(entry);
    if(mode==='auth')return r.fulfill({status:401,headers:cors,body:JSON.stringify({type:'error',error:{type:'authentication_error',message:'invalid x-api-key'}})});
    if(url.pathname.startsWith('/v1/models/'))return r.fulfill({status:200,headers:cors,body:JSON.stringify({id:'claude-opus-5-5',display_name:'Claude Opus 5.5',type:'model'})});
    if(entry.body.output_config.format.schema.properties.suggestions)return r.fulfill({status:200,headers:cors,body:JSON.stringify(message(suggestions))});
    const ctx=JSON.parse(entry.body.messages[0].content[0].text);
    const out=ctx.requested_revision?{...reply,body:ctx.current_draft.body+'\n\n(shorter)',checks:[]}:/Summarise/.test(entry.body.system)?{subject:'Summary',body:'- Records missing\n- Re-test requested',checks:[]}:reply;
    return r.fulfill({status:200,headers:cors,body:JSON.stringify(message(out))});
  };
  try{
    const page=await browser.newPage({viewport:{width:400,height:900},deviceScaleFactor:process.env.SCREENSHOT_DIR?2:1});
    page.on('pageerror',e=>errors.push(e.message));
    page.on('console',m=>{if(m.type()==='error'&&!/Failed to load resource/.test(m.text()))errors.push(m.text());});
    await page.addInitScript(SPEECH_MOCK);
    await page.route('https://appsforoffice.microsoft.com/**',r=>r.fulfill({contentType:'text/javascript',body:OFFICE_MOCK}));
    await page.route('https://api.anthropic.com/**',route);
    await page.goto(`http://127.0.0.1:${port}/taskpane.html`);
    const shot=async n=>{if(process.env.SCREENSHOT_DIR)await page.screenshot({path:path.join(process.env.SCREENSHOT_DIR,n+'.png'),fullPage:true});};

    // 1. Brand, email context, quick replies before a key exists
    assert.equal(await page.getAttribute('.brand','aria-label'),'Bulwark HQ Assistant');
    assert.ok(await page.$eval('.brand img',i=>i.complete&&i.naturalWidth>0),'logo loads');
    await page.waitForFunction(()=>/Jordan Example/.test(document.getElementById('email-line').textContent));
    assert.equal(await page.textContent('#email-subject'),'Your whitelabel email is not functioning');
    assert.equal(await page.textContent('#email-subject .hl'),'not functioning');
    assert.match(await page.textContent('#banner-text'),/Add your Claude API key/);
    assert.equal(await page.$$eval('#suggestions .card-btn',e=>e.length),4,'quick replies shown without a key');
    assert.equal(await page.isVisible('#insights'),false);
    await page.click('#generate');
    await page.waitForSelector('#drawer:not([hidden])');
    assert.equal(requests.length,0,'no API call without a key');

    // 2. Bad key rejected, good key saved, then insights and replies load for this email
    mode='auth';await page.fill('#api-key','sk-ant-api03-'+'b'.repeat(40));await page.click('#save-key');
    await page.waitForFunction(()=>/rejected/.test(document.getElementById('settings-status').textContent));
    assert.equal(await page.evaluate(()=>window.__bag.apiKey),undefined);
    mode='ok';const key='sk-ant-api03-'+'a'.repeat(40);
    await page.fill('#api-key',key);await page.click('#save-key');
    await page.waitForFunction(()=>/Connected/.test(document.getElementById('settings-status').textContent));
    assert.equal(await page.evaluate(()=>window.__bag.apiKey),key);
    await page.click('#close-settings');
    await page.waitForFunction(()=>/missing SPF/.test(document.getElementById('summary').textContent));
    assert.equal(await page.textContent('#status'),'','old "add your key" message cleared');
    const sug=requests.find(r=>r.body&&r.body.output_config.format.schema.properties.suggestions);
    assert.equal(sug.body.output_config.effort,'low');assert.match(sug.body.messages[0].content[0].text,/SPF and DKIM/);
    assert.deepEqual(await page.$$eval('#suggestions .card-btn .txt',e=>e.map(x=>x.textContent)),suggestions.suggestions.map(s=>s.title));
    assert.equal(await page.textContent('#in-priority'),'High');assert.equal(await page.getAttribute('#in-priority','class'),'high');
    assert.equal(await page.textContent('#in-need'),'Fix + re-test');
    assert.match(await page.textContent('#email-line'),/Jordan Example · contractor/);
    assert.equal(await page.inputValue('#profile'),'builtin-contractor','voice auto-picked from sender');

    // 3. Pick a reply and tone: live preview changes, then one request carries both
    assert.equal(await page.textContent('#preview-text'),'Balanced: agreed');
    await page.click('#suggestions .card-btn >> nth=1');
    assert.equal(await page.textContent('#preview-text'),'Balanced: please fix');
    await page.click('#tone button >> nth=0');
    assert.equal(await page.textContent('#tone-label'),'Very formal');
    assert.equal(await page.textContent('#preview-text'),'Formal: please fix');
    assert.equal(await page.textContent('#preview-greeting'),'Dear Jordan,');
    await page.click('#tone button >> nth=4');assert.equal(await page.textContent('#preview-text'),'Friendly: please fix');
    await page.click('#suggestions .card-btn >> nth=0');
    await shot('v32-1-open');
    const before=requests.length;
    await page.click('#generate');
    await page.waitForSelector('#view-draft:not([hidden])');
    const gen=requests.slice(before).find(r=>r.path==='/v1/messages');
    const ctx=JSON.parse(gen.body.messages[0].content[0].text);
    assert.match(ctx.user_dot_points,/propose re-testing/);assert.equal(ctx.task,'reply');
    assert.equal(ctx.tone,'Friendly');assert.deepEqual(ctx.tone_controls,{formality:1,warmth:5,directness:5});
    assert.match(ctx.voice_profile.name,/Contractors/);assert.equal(gen.body.fallbacks,'default');
    assert.equal(await page.textContent('#draft-title'),'Reply to Jordan');
    assert.match(await page.textContent('#review-text'),/placeholder/);assert.match(await page.getAttribute('#review-line','class'),/stop/);
    await shot('v32-2-draft');

    // 4. Red item guards insert; reply HTML has real list items
    await page.click('#apply');assert.match(await page.textContent('#apply'),/anyway/);
    await page.click('#apply');await page.waitForFunction(()=>window.__calls.length===1);
    assert.match((await page.evaluate(()=>window.__calls))[0][1],/<li>SPF record<\/li>/);
    await page.fill('#draft-body',(await page.inputValue('#draft-body')).replace('[confirm date]','Tuesday 10am'));
    assert.doesNotMatch(await page.getAttribute('#review-line','class'),/stop/);
    assert.equal(await page.isVisible('#apply-all'),true);
    await page.click('#apply-all');await page.waitForFunction(()=>window.__calls.length===2);
    assert.equal((await page.evaluate(()=>window.__calls))[1][0],'replyAll');
    assert.equal(await page.textContent('#saved-time'),'8m');assert.equal(await page.textContent('#saved-count'),'2 REPLIES');

    // 5. Refine with history; dictation into the change box
    await page.click('.refine button:has-text("Shorter")');
    await page.waitForFunction(()=>/\(shorter\)/.test(document.getElementById('draft-body').value));
    assert.equal(await page.textContent('#ver-label'),'2/2');
    await page.click('#mic2');
    await page.waitForFunction(()=>/Tuesday at ten/.test(document.getElementById('refine-text').value));

    // 6. Back keeps the draft
    await page.click('#back');assert.equal(await page.isVisible('#view-compose'),true);
    assert.match(await page.textContent('#banner-text'),/draft is kept/);
    await page.click('#banner-action');assert.equal(await page.isVisible('#view-draft'),true);
    await page.click('#back');

    // 7. Dictation into the main box, then write from it
    await page.fill('#points','');await page.click('#mic');
    await page.waitForFunction(()=>/Tuesday at ten/.test(document.getElementById('points').value));
    assert.equal(await page.getAttribute('#mic','aria-pressed'),'true');
    await page.click('#mic');assert.equal(await page.getAttribute('#mic','aria-pressed'),'false');

    // 8. Tools: Summarise shows a copyable result
    await page.click('.more [data-tool="summary"]');
    await page.waitForFunction(()=>document.getElementById('draft-title').textContent==='Summary');
    assert.equal(await page.textContent('#apply'),'Copy result');
    await page.click('#back');

    // 9. Haiku: no server-side fallback parameter
    await page.click('#open-settings');await page.selectOption('#model','claude-haiku-5-5');await page.click('#save-key');
    await page.waitForFunction(()=>/Connected/.test(document.getElementById('settings-status').textContent));
    await page.click('#close-settings');
    const n=requests.length;await page.fill('#points','thank them');await page.click('#generate');
    await page.waitForSelector('#view-draft:not([hidden])');
    const haiku=requests.slice(n).find(r=>r.path==='/v1/messages'&&!r.body.output_config.format.schema.properties.suggestions);
    assert.equal(haiku.body.model,'claude-haiku-5-5');assert.equal(haiku.body.fallbacks,undefined);

    // 10. Voice saved to mailbox; settings confirm visibly
    await page.click('#open-settings');await page.click('#new-profile');
    await page.fill('#pf-name','Fictional strata voice');await page.click('#pf-save');
    await page.waitForFunction(()=>(window.__bag.profiles||[]).some(p=>p.name==='Fictional strata voice'));
    assert.equal(await page.textContent('#pf-save'),'Saved ✓');
    await page.click('#close-settings');
    assert.deepEqual(errors,[],'no console errors');

    // 11. No speech recognition: the mic explains Windows dictation
    const plain=await browser.newPage({viewport:{width:400,height:900}});
    await plain.route('https://appsforoffice.microsoft.com/**',r=>r.fulfill({contentType:'text/javascript',body:OFFICE_MOCK}));
    await plain.route('https://api.anthropic.com/**',route);
    await plain.goto(`http://127.0.0.1:${port}/taskpane.html`);
    await plain.waitForFunction(()=>window.__dbnStarted);
    await plain.addInitScript('window.SpeechRecognition=undefined;window.webkitSpeechRecognition=undefined;');await plain.reload();await plain.waitForFunction(()=>window.__dbnStarted);
    await plain.click('#mic');
    await plain.waitForFunction(()=>/Win \+ H/.test(document.getElementById('status').textContent));
    await plain.close();

    // 12. Outlook never signals ready: buttons still work and the watchdog explains why
    const stuck=await browser.newPage({viewport:{width:400,height:900}});
    await stuck.route('https://appsforoffice.microsoft.com/**',r=>r.fulfill({contentType:'text/javascript',
      body:"window.Office={HostType:{Outlook:'Outlook'},onReady(){return new Promise(()=>{});}};"}));
    await stuck.goto(`http://127.0.0.1:${port}/taskpane.html`);
    await stuck.waitForFunction(()=>document.querySelectorAll('#suggestions .card-btn').length===4&&document.querySelectorAll('#tone button').length===5);
    await stuck.click('#open-settings');assert.match(await stuck.textContent('#settings-status'),/Still connecting/);
    await stuck.click('#close-settings');
    await stuck.waitForFunction(()=>/could not finish starting/.test(document.getElementById('status').textContent),null,{timeout:20000});
    await stuck.close();

    console.log('Bulwark HQ Assistant v3.2 (browser): brand + logo, quick replies, key flow, insights, reply + tone live preview, auto voice, time saved, send-check guard, reply/reply all HTML, refine history, dictation (both boxes + Win+H fallback), back keeps draft, tools, Haiku, voices, watchdog passed.');
  }finally{await browser.close();server.close();}
})().catch(e=>{console.error(e);process.exitCode=1;});
