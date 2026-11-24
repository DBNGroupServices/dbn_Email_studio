/* End-to-end check of the Outlook task pane in a real browser (Chromium via
   Playwright) against the real local engine with simulated AI and a mocked
   Outlook item. Run: DBN_EMAIL_TEST_PLAINTEXT_STORE=1 node tests/test_taskpane_browser.js
   (on Windows the env var is not needed). Optional: SCREENSHOT_DIR=path */
const assert=require('node:assert/strict');
const {spawn}=require('node:child_process');
const path=require('node:path');
const {chromium}=require('playwright');

const PORT=8790+Math.floor(Math.random()*100);
const ROOT=path.resolve(__dirname,'..');
const OFFICE_MOCK=`
window.__calls=[];
const item={itemId:'fixture-1',subject:'Saturday dock access',conversationId:'c1',
  from:{displayName:'Jordan Example',emailAddress:'jordan@example.com'},
  to:[{displayName:'Sam',emailAddress:'sam@example.com'},{displayName:'Facilities',emailAddress:'fm@example.com'}],cc:[],
  dateTimeCreated:new Date('2026-11-23T05:42:00Z'),attachments:[],
  body:{getAsync:(type,cb)=>cb({status:'succeeded',value:'Hi Sam,\\n\\nCan contractors use the loading dock this Saturday from 7am? They need the goods lift.\\n\\nThanks,\\nJordan'})},
  displayReplyForm(v){window.__calls.push(['reply',v.htmlBody]);},
  displayReplyAllForm(v){window.__calls.push(['replyAll',v.htmlBody]);}};
window.Office={HostType:{Outlook:'Outlook'},EventType:{ItemChanged:'olkItemSelectedChanged'},
  context:{officeTheme:{bodyBackgroundColor:'#ffffff'},mailbox:{item,addHandlerAsync(type,handler){window.__itemChanged=handler;}}},
  onReady(cb){setTimeout(()=>cb({host:'Outlook'}),0);return Promise.resolve({host:'Outlook'});}};`;

function startServer(){
  return new Promise((resolve,reject)=>{
    const proc=spawn(process.platform==='win32'?'python':'python3',[path.join(ROOT,'tests','preview_server.py'),'--signed-in','--port',String(PORT)],
      {cwd:ROOT,env:{...process.env,DBN_EMAIL_TEST_PLAINTEXT_STORE:process.env.DBN_EMAIL_TEST_PLAINTEXT_STORE||'1'}});
    let out='';
    proc.stdout.on('data',d=>{out+=d;if(out.includes('TEST ONLY'))resolve(proc);});
    proc.stderr.on('data',d=>{out+=d;});
    proc.on('exit',code=>reject(new Error('preview server exited '+code+'\n'+out)));
    setTimeout(()=>reject(new Error('preview server timeout\n'+out)),15000);
  });
}

(async()=>{
  const server=await startServer();
  const browser=await chromium.launch();
  const errors=[];
  try{
    const page=await browser.newPage({viewport:{width:420,height:900},deviceScaleFactor:process.env.SCREENSHOT_DIR?2:1});
    page.on('pageerror',e=>errors.push(e.message));
    page.on('console',m=>{if(m.type()==='error')errors.push(m.text());});
    await page.route('https://appsforoffice.microsoft.com/**',route=>route.fulfill({contentType:'text/javascript',body:OFFICE_MOCK}));
    await page.goto(`http://127.0.0.1:${PORT}/office/taskpane.html`);
    const shot=async name=>{if(process.env.SCREENSHOT_DIR)await page.screenshot({path:path.join(process.env.SCREENSHOT_DIR,name+'.png'),fullPage:true});};

    // 1. Reads the open email automatically with sender details
    await page.waitForFunction(()=>document.getElementById('email-from').textContent==='Jordan Example');
    assert.equal(await page.textContent('#email-subject'),'Saturday dock access');
    assert.equal(await page.textContent('#avatar'),'J');
    await page.waitForFunction(()=>document.getElementById('profile').options.length>3);
    assert.match(await page.textContent('#connection-label'),/ChatGPT plan/);
    await shot('1-ready');

    // 2. One-click quick reply produces a draft with review checks
    await page.click('#intents button[data-intent="confirm"]');
    await page.waitForSelector('#result:not([hidden])');
    assert.match(await page.inputValue('#draft-body'),/loading dock is available/);
    const review=await page.$$eval('#review li',els=>els.map(e=>[e.className,e.textContent]));
    assert.ok(review.some(([c,t])=>c==='stop'&&/placeholder/.test(t)),'placeholder flagged');
    assert.ok(review.some(([c])=>c==='ai'),'AI check shown');
    await shot('2-draft');

    // 3. Insert is guarded by a second click when there are stop-level warnings
    await page.click('#apply');
    assert.match(await page.textContent('#apply'),/anyway/);
    assert.equal((await page.evaluate(()=>window.__calls)).length,0);
    await page.click('#apply');
    await page.waitForFunction(()=>window.__calls.length===1);
    const [how,htmlBody]=(await page.evaluate(()=>window.__calls))[0];
    assert.equal(how,'reply');assert.match(htmlBody,/<li>Sign in at the security desk<\/li>/);

    // 4. Fixing the placeholder clears the stop; Reply all is offered (2+ recipients)
    await page.fill('#draft-body',(await page.inputValue('#draft-body')).replace('[confirm deadline]','Thursday 5pm'));
    assert.equal(await page.$$eval('#review li.stop',e=>e.length),0);
    assert.equal(await page.isVisible('#apply-all'),true);
    await page.click('#apply-all');
    await page.waitForFunction(()=>window.__calls.length===2);
    assert.equal((await page.evaluate(()=>window.__calls))[1][0],'replyAll');

    // 5. Refine keeps version history
    await page.click('.refine button:has-text("Shorter")');
    await page.waitForFunction(()=>/Simulated revision/.test(document.getElementById('draft-body').value));
    assert.equal(await page.textContent('#ver-label'),'2/2');
    await page.click('#ver-back');
    assert.equal(await page.textContent('#ver-label'),'1/2');
    assert.ok(!(await page.inputValue('#draft-body')).includes('Simulated revision'));

    // 6. Security guard flags codes typed into a draft
    await page.fill('#draft-body','Hi Jordan,\n\nThe dock door code is 4821.\n\nThanks');
    assert.ok(await page.$$eval('#review li.stop',els=>els.some(e=>/code or PIN/.test(e.textContent))));

    // 7. Compare three options
    await page.click('#more-controls summary');
    await page.selectOption('#variants','3');
    await page.click('#generate');
    await page.waitForFunction(()=>document.querySelectorAll('#variant-tabs button').length===3);
    await page.click('#variant-tabs button:nth-child(2)');
    assert.match(await page.inputValue('#draft-body'),/Option 2/);
    await page.selectOption('#variants','1');

    // 8. Moving to another email keeps an unsent draft and offers to load
    await page.evaluate(()=>window.__itemChanged());
    await page.waitForSelector('#banner:not([hidden])');
    assert.match(await page.textContent('#banner-text'),/another email/);

    // 9. Voice profile editor saves a new profile that becomes selectable
    await page.click('#open-settings');
    await page.click('#new-profile');
    await page.fill('#pf-name','Fictional strata voice');
    await page.fill('#pf-guidance','Calm, factual, short.');
    await page.click('#pf-add-example');
    await page.fill('#pf-examples textarea','Hi team, all done for today. Thanks, S');
    await page.click('#pf-save');
    await page.waitForFunction(()=>[...document.getElementById('profile').options].some(o=>o.text==='Fictional strata voice'));
    await shot('3-settings');
    await page.click('#close-settings');

    assert.deepEqual(errors,[],'no console errors');
    console.log('Task pane (browser): auto-read, quick reply, send-check guard, reply/reply all HTML, refine history, security guard, variants, item change and profiles passed.');
  }finally{
    await browser.close();server.kill();
  }
})().catch(error=>{console.error(error);process.exitCode=1;});
