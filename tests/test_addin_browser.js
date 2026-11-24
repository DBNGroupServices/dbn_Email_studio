/* End-to-end test of the hosted add-in in Chromium: static site + mocked
   Outlook (Office.js) + mocked Claude API. Checks the requests the bundled SDK
   actually sends. Run: node tests/test_addin_browser.js  (SCREENSHOT_DIR=... optional) */
const assert=require('node:assert/strict');
const http=require('node:http');
const fs=require('node:fs');
const path=require('node:path');
const {chromium}=require('playwright');
const ROOT=path.resolve(__dirname,'..','addin');
const TYPES={'.html':'text/html','.js':'text/javascript','.css':'text/css','.png':'image/png','.xml':'application/xml'};

const OFFICE_MOCK=`
window.__calls=[];window.__saves=0;const bag={};
const item={itemId:'fixture-1',subject:'Saturday dock access',
  from:{displayName:'Jordan Example',emailAddress:'jordan@example.com'},
  to:[{displayName:'Sam',emailAddress:'sam@example.com'},{displayName:'Facilities',emailAddress:'fm@example.com'}],cc:[],
  dateTimeCreated:new Date('2026-11-23T05:42:00Z'),attachments:[],
  body:{getAsync:(type,cb)=>cb({status:'succeeded',value:'Hi Sam,\\n\\nCan contractors use the loading dock this Saturday from 7am?\\n\\nThanks,\\nJordan'})},
  displayReplyForm(v){window.__calls.push(['reply',v.htmlBody]);},displayReplyAllForm(v){window.__calls.push(['replyAll',v.htmlBody]);}};
window.Office={HostType:{Outlook:'Outlook'},EventType:{ItemChanged:'olkItemSelectedChanged'},
  context:{officeTheme:{bodyBackgroundColor:'#ffffff'},roamingSettings:{get:k=>bag[k],set:(k,v)=>{bag[k]=v;},remove:k=>{delete bag[k];},saveAsync:cb=>{window.__saves++;cb({status:'succeeded'});}},
    mailbox:{item,addHandlerAsync(t,h){window.__itemChanged=h;},displayNewMessageForm(v){window.__calls.push(['new',v.subject,v.htmlBody]);}}},
  onReady(cb){setTimeout(()=>cb({host:'Outlook'}),0);return Promise.resolve({host:'Outlook'});}};
window.__bag=bag;`;

const reply={subject:'Re: Saturday dock access',body:'Hi Jordan,\n\nYes, the dock is available from 7am.\n\n- Sign in at the security desk\n- Goods lift booked 7am to 12pm\n\nPlease send names by [confirm deadline].\n\nKind regards,\nSam',checks:['Deadline not supplied.']};
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
  try{
    const page=await browser.newPage({viewport:{width:420,height:900},deviceScaleFactor:process.env.SCREENSHOT_DIR?2:1});
    page.on('pageerror',e=>errors.push(e.message));
    page.on('console',m=>{if(m.type()==='error'&&!/Failed to load resource/.test(m.text()))errors.push(m.text());});
    await page.route('https://appsforoffice.microsoft.com/**',r=>r.fulfill({contentType:'text/javascript',body:OFFICE_MOCK}));
    await page.route('https://api.anthropic.com/**',async route=>{
      const req=route.request(),url=new URL(req.url());
      const cors={'access-control-allow-origin':'*','access-control-allow-headers':'*','content-type':'application/json'};
      if(req.method()==='OPTIONS')return route.fulfill({status:204,headers:cors});
      const entry={method:req.method(),path:url.pathname,headers:await req.allHeaders(),body:req.postData()?JSON.parse(req.postData()):null};requests.push(entry);
      if(mode==='auth')return route.fulfill({status:401,headers:{...cors,'request-id':'req_fixture'},body:JSON.stringify({type:'error',error:{type:'authentication_error',message:'invalid x-api-key'}})});
      if(url.pathname.startsWith('/v1/models/'))return route.fulfill({status:200,headers:cors,body:JSON.stringify({id:'claude-opus-5-5',display_name:'Claude Opus 5.5',type:'model'})});
      const ctx=JSON.parse(entry.body.messages[0].content[0].text);
      const out=ctx.number_of_variants?{variants:[1,2,3].slice(0,ctx.number_of_variants).map(i=>({label:'Option '+i,subject:'Re: x',body:'Hi Jordan,\n\nVersion '+i+'.\n\nSam'})),checks:[]}
        :ctx.requested_revision?{...reply,body:ctx.current_draft.body+'\n\n(shorter)',checks:[]}:reply;
      return route.fulfill({status:200,headers:cors,body:JSON.stringify(message(out))});
    });
    await page.goto(`http://127.0.0.1:${port}/taskpane.html`);
    const shot=async n=>{if(process.env.SCREENSHOT_DIR)await page.screenshot({path:path.join(process.env.SCREENSHOT_DIR,n+'.png'),fullPage:true});};

    // 1. Loads with no engine, reads the email, prompts for a key
    await page.waitForFunction(()=>document.getElementById('email-from').textContent==='Jordan Example');
    assert.match(await page.textContent('#banner-text'),/Add your Claude API key/);
    assert.equal(await page.textContent('#connection-label'),'Add API key');
    await page.click('#intents button[data-intent="confirm"]');
    await page.waitForSelector('#drawer:not([hidden])');
    assert.equal(requests.length,0,'no API call without a key');
    await shot('addin-1-welcome');

    // 2. Bad key is rejected with a plain message and is not saved
    mode='auth';
    await page.fill('#api-key','sk-ant-api03-'+'b'.repeat(40));await page.click('#save-key');
    await page.waitForFunction(()=>/rejected/.test(document.getElementById('settings-status').textContent));
    assert.equal(await page.evaluate(()=>window.__bag.apiKey),undefined);

    // 3. Good key: tested via the Models API, saved to mailbox roaming settings
    mode='ok';const key='sk-ant-api03-'+'a'.repeat(40);
    await page.fill('#api-key',key);await page.click('#save-key');
    await page.waitForFunction(()=>/Connected/.test(document.getElementById('settings-status').textContent));
    const test=requests.at(-1);assert.equal(test.path,'/v1/models/claude-opus-5-5');assert.ok(test.headers['x-api-key']===key||test.headers['authorization']==='Bearer '+key,'key sent');
    assert.equal(test.headers['anthropic-dangerous-direct-browser-access'],'true');
    assert.equal(await page.evaluate(()=>window.__bag.apiKey),key);
    assert.equal(await page.textContent('#connection-label'),'Opus 5.5');
    await shot('addin-2-settings');
    await page.click('#close-settings');

    // 4. One-click reply: request shape sent by the SDK
    await page.click('#intents button[data-intent="confirm"]');
    await page.waitForSelector('#result:not([hidden])');
    const gen=requests.at(-1);
    assert.equal(gen.method,'POST');assert.equal(gen.path,'/v1/messages');
    assert.equal(gen.body.model,'claude-opus-5-5');assert.equal(gen.body.max_tokens,16000);
    assert.equal(gen.body.output_config.format.type,'json_schema');assert.equal(gen.body.output_config.effort,'medium');
    assert.equal(gen.body.fallbacks,'default');assert.match(gen.headers['anthropic-beta']||'',/server-side-fallback-2026-07-01/);
    assert.equal(gen.body.thinking,undefined);assert.equal(gen.body.temperature,undefined);
    const ctx=JSON.parse(gen.body.messages[0].content[0].text);
    assert.match(ctx.source_email_or_draft,/loading dock/);assert.equal(ctx.email_metadata.from.name,'Jordan Example');assert.match(ctx.quick_intent,/Confirm/);
    assert.match(gen.body.system,/Security disclosure guard/);
    assert.ok(await page.$$eval('#review li.stop',e=>e.some(x=>/placeholder/.test(x.textContent))));
    await shot('addin-3-draft');

    // 5. Insert guard, then reply with Outlook HTML
    await page.click('#apply');assert.match(await page.textContent('#apply'),/anyway/);
    await page.click('#apply');await page.waitForFunction(()=>window.__calls.length===1);
    const [how,html]=(await page.evaluate(()=>window.__calls))[0];
    assert.equal(how,'reply');assert.match(html,/<li>Sign in at the security desk<\/li>/);

    // 6. Refine keeps history; three options
    await page.click('.refine button:has-text("Shorter")');
    await page.waitForFunction(()=>/\(shorter\)/.test(document.getElementById('draft-body').value));
    assert.equal(await page.textContent('#ver-label'),'2/2');
    assert.equal(requests.at(-1).body.fallbacks,'default');
    await page.click('#more-controls summary');await page.selectOption('#variants','3');await page.click('#generate');
    await page.waitForFunction(()=>document.querySelectorAll('#variant-tabs button').length===3);
    await page.selectOption('#variants','1');

    // 7. New email from notes opens a new Outlook message (no engine needed)
    await page.click('.tab[data-mode="write"]');await page.click('#write-kind button[data-kind="compose"]');
    await page.fill('#points','Fictional: remind tenants about Saturday works.');await page.click('#generate');
    await page.waitForFunction(()=>document.getElementById('apply').textContent==='Open as new email');
    await page.fill('#draft-body','Hi all,\n\nReminder about Saturday works.\n\nSam');
    await page.click('#apply');await page.waitForFunction(()=>window.__calls.some(c=>c[0]==='new'));

    // 8. Haiku: no server-side fallback parameter
    await page.click('#open-settings');await page.selectOption('#model','claude-haiku-5-5');await page.click('#save-key');
    await page.waitForFunction(()=>/Connected/.test(document.getElementById('settings-status').textContent));
    await page.click('#close-settings');await page.click('.tab[data-mode="reply"]');
    const before=requests.length;
    await page.click('#intents button[data-intent="thanks"]');
    for(let i=0;i<50&&requests.length===before;i++)await page.waitForTimeout(100);
    const haiku=requests.at(-1);assert.equal(haiku.body.model,'claude-haiku-5-5');assert.equal(haiku.body.fallbacks,undefined);

    // 9. Profile saved to mailbox storage
    await page.click('#open-settings');await page.click('#new-profile');
    await page.fill('#pf-name','Fictional strata voice');await page.click('#pf-save');
    await page.waitForFunction(()=>(window.__bag.profiles||[]).some(p=>p.name==='Fictional strata voice'));

    assert.deepEqual(errors,[],'no console errors');
    console.log('Hosted add-in (browser): no-engine load, key test & mailbox storage, SDK request shape (structured output, effort, fallbacks, browser header), reply HTML, refine, variants, new email, Haiku and profiles passed.');
  }finally{await browser.close();server.close();}
})().catch(e=>{console.error(e);process.exitCode=1;});
