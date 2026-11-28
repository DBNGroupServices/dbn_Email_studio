// End-to-end: the real Worker (handle + D1 schema) serving the real add-in to Chromium,
// with a mocked Outlook (Office.js + single sign-on), Claude and Stripe.
// Checks the free trial -> paywall -> checkout -> paid plan journey, and that iPhone/Android
// never shows upgrade or purchase links. Run: node test/pane.browser.mjs (SCREENSHOT_DIR=... optional)
import assert from 'node:assert/strict';
import http from 'node:http';
import fs from 'node:fs';
import path from 'node:path';
import { createRequire } from 'node:module';
import { fileURLToPath } from 'node:url';
import { handle } from '../src/index.js';
import { AuthError } from '../src/auth.js';
import { makeD1 } from './d1-shim.mjs';

const require = createRequire(import.meta.url);
const { chromium } = require('playwright');
const here = path.dirname(fileURLToPath(import.meta.url));
const ADDIN = path.resolve(here, '..', '..', 'addin');
const TYPES = { '.html': 'text/html', '.js': 'text/javascript', '.css': 'text/css', '.png': 'image/png', '.svg': 'image/svg+xml', '.woff2': 'font/woff2' };

const db = makeD1(path.join(here, '..', 'migrations', '0001_init.sql'));
const calls = { claude: [], checkout: [], portal: [] };
const reply = (o) => ({ stop_reason: 'end_turn', model: 'x', content: [{ type: 'text', text: JSON.stringify(o) }] });
const pv = (t) => ({ formal: 'Formal ' + t, balanced: 'Balanced ' + t, friendly: 'Friendly ' + t });
const SUGGEST = { summary: 'Jordan asks to re-test the email records.', audience: 'contractor', priority: 'High', need: 'Re-test', mood: 'Neutral',
  suggestions: [{ title: 'Agree and propose a date', tag: 'BEST FIT', hint: 'h', instruction: 'Agree and propose a date.', preview: pv('agree') },
    { title: 'Ask them to fix it', tag: 'FASTEST', hint: 'h', instruction: 'Ask them to fix it.', preview: pv('fix') }] };
const DRAFT = { subject: 'Re: records', body: 'Hi Jordan,\n\nAgreed. Could we re-test on Tuesday?\n\nKind regards,\nSam', checks: [] };
const claude = (p) => { calls.claude.push(p.model); return reply(p.output_config.format.schema.properties.suggestions ? SUGGEST : DRAFT); };
const subs = {};
const deps = {
  now: () => Date.now(),
  verify: async (req) => {
    const who = /^Bearer (\w+)$/.exec(req.headers.get('authorization') || '')?.[1];
    if (!who) throw new AuthError('Sign-in required.');
    return { id: 't1:' + who, tid: 't1', email: who + '@example.com', name: who };
  },
  anthropic: { messages: { create: async (p) => claude(p) }, beta: { messages: { create: async (p) => claude(p) } } },
  stripe: {
    prices: { list: async ({ lookup_keys }) => ({ data: [{ id: 'price_' + lookup_keys[0] }] }) },
    customers: { create: async (p) => ({ id: 'cus_' + p.metadata.account_id.split(':')[1] }) },
    checkout: { sessions: { create: async (p) => { calls.checkout.push(p); return { url: 'https://checkout.stripe.test/s' }; } } },
    billingPortal: { sessions: { create: async (p) => { calls.portal.push(p); return { url: 'https://billing.stripe.test/p' }; } } },
    subscriptions: { retrieve: async (id) => subs[id] },
    webhooks: { constructEventAsync: async (body) => JSON.parse(body) },
  },
};
const ASSETS = {
  fetch: async (req) => {
    const p = decodeURIComponent(new URL(req.url).pathname).replace(/^\/+/, '') || 'index.html';
    const file = path.join(ADDIN, p);
    if (!file.startsWith(ADDIN) || !fs.existsSync(file) || fs.statSync(file).isDirectory()) return new Response('Not found', { status: 404 });
    return new Response(fs.readFileSync(file), { headers: { 'content-type': TYPES[path.extname(file)] || 'application/octet-stream' } });
  },
};
let port = 0;
const env = () => ({ DB: db, ASSETS, PUBLIC_URL: `http://127.0.0.1:${port}`, MS_CLIENT_ID: 'client-123', TRIAL_TZ: 'Australia/Melbourne' });

const server = http.createServer(async (req, res) => {
  const chunks = []; for await (const c of req) chunks.push(c);
  const body = chunks.length ? Buffer.concat(chunks) : undefined;
  const r = await handle(new Request(`http://127.0.0.1:${port}${req.url}`, { method: req.method, headers: req.headers, body: req.method === 'GET' ? undefined : body }), env(), deps);
  res.writeHead(r.status, Object.fromEntries(r.headers)); res.end(Buffer.from(await r.arrayBuffer()));
}).listen(0);
port = server.address().port;

const officeMock = (user, platform) => `
window.__calls=[];const bag={};
const item={itemId:'fixture-1',subject:'Email records not working',from:{displayName:'Jordan Example',emailAddress:'jordan@example.com'},
  to:[{displayName:'Sam',emailAddress:'sam@example.com'}],cc:[],dateTimeCreated:new Date('2026-11-23T05:42:00Z'),attachments:[],
  body:{getAsync:(t,cb)=>cb({status:'succeeded',value:'Hi Sam,\\n\\nThe SPF and DKIM records are missing. Can we re-test?\\n\\nThanks,\\nJordan'})},
  displayReplyForm(v){window.__calls.push(['reply',v.htmlBody]);},displayReplyAllForm(v){window.__calls.push(['replyAll',v.htmlBody]);}};
window.Office={HostType:{Outlook:'Outlook'},EventType:{ItemChanged:'x'},PlatformType:{PC:'PC',Mac:'Mac',iOS:'iOS',Android:'Android',OfficeOnline:'OfficeOnline'},
  auth:{getAccessToken:async()=>{window.__tokens=(window.__tokens||0)+1;return '${user}';}},
  context:{platform:'${platform}',officeTheme:{bodyBackgroundColor:'#ffffff'},ui:{openBrowserWindow(u){window.__calls.push(['browser',u]);}},
    roamingSettings:{get:k=>bag[k],set:(k,v)=>{bag[k]=v;},remove:k=>{delete bag[k];},saveAsync:cb=>cb({status:'succeeded'})},
    mailbox:{item,addHandlerAsync(t,h){},displayNewMessageForm(v){window.__calls.push(['new',v.subject]);}}},
  onReady(cb){setTimeout(()=>cb({host:'Outlook'}),0);return Promise.resolve({host:'Outlook'});}};`;

const browser = await chromium.launch();
const errors = [];
async function openPane(user, platform = 'PC') {
  const page = await browser.newPage({ viewport: { width: 400, height: 900 }, deviceScaleFactor: process.env.SCREENSHOT_DIR ? 2 : 1 });
  page.on('pageerror', (e) => errors.push(e.message));
  page.on('console', (m) => { if (m.type() === 'error' && !/Failed to load resource/.test(m.text())) errors.push(m.text()); });
  await page.route('https://appsforoffice.microsoft.com/**', (r) => r.fulfill({ contentType: 'text/javascript', body: officeMock(user, platform) }));
  await page.route('https://api.anthropic.com/**', (r) => { errors.push('pane called Claude directly'); return r.abort(); });
  await page.goto(`http://127.0.0.1:${port}/taskpane.html`);
  return page;
}
const shot = async (page, n) => { if (process.env.SCREENSHOT_DIR) await page.screenshot({ path: path.join(process.env.SCREENSHOT_DIR, n + '.png') }); };

try {
  // 1. Sign in starts the free trial; no API key settings; suggestions come from the service.
  const page = await openPane('alice');
  await page.waitForFunction(() => /^Trial · \dd$/.test(document.getElementById('connection-label').textContent));
  assert.match(await page.textContent('#connection-label'), /^Trial · 3d$/);
  await page.waitForSelector('#suggestions .card-btn:not(.skeleton) >> text=Agree and propose a date');
  assert.equal(await page.textContent('#in-priority'), 'High');
  assert.equal(calls.claude[0], 'claude-haiku-5-5', 'insights use Haiku');
  await page.click('#open-settings');
  assert.equal(await page.isHidden('#byok-fields'), true, 'no API key fields in subscription mode');
  assert.equal(await page.textContent('#ai-title'), 'Your plan');
  assert.match(await page.textContent('#settings-connection'), /alice@example\.com · Free trial · 3 working days · 50 replies left/);
  await page.click('#close-settings');

  // 2. Draft a reply on Opus; the countdown updates.
  await page.click('#generate');
  await page.waitForSelector('#view-draft:not([hidden])');
  assert.match(await page.inputValue('#draft-body'), /re-test on Tuesday/);
  assert.equal(calls.claude.at(-1), 'claude-opus-5-5', 'trial drafts use Opus');
  await page.click('#open-settings');
  assert.match(await page.textContent('#settings-connection'), /49 replies left/);
  assert.match(await page.textContent('#usage-line'), /1 of 50 trial replies used/);
  await page.click('#close-settings');
  await page.click('#back');

  // 3. Trial used up: the next draft opens the plans sheet instead (card asked for only now).
  db.raw.prepare('UPDATE accounts SET trial_replies = 50 WHERE id = ?').run('t1:alice');
  const before = calls.claude.length;
  await page.fill('#points', 'Agree and propose Tuesday');
  await page.click('#generate');
  await page.waitForSelector('#upgrade:not([hidden])');
  assert.equal(calls.claude.length, before, 'no AI call after the trial');
  assert.match(await page.textContent('#upgrade-reason'), /free trial has ended/);
  assert.match(await page.textContent('#connection-label'), /Upgrade/);
  const prices = await page.$$eval('.plan .price', (n) => n.map((x) => x.textContent));
  assert.deepEqual(prices, ['A$15 / month', 'A$29 / month', 'A$25 per user / month']);
  await page.click('.interval [data-interval="year"]');
  assert.deepEqual(await page.$$eval('.plan .price', (n) => n.map((x) => x.textContent)), ['A$150 / year', 'A$290 / year', 'A$250 per user / year']);
  await shot(page, 'service-plans');
  await page.click('.plan.best .primary');
  await page.waitForFunction(() => window.__calls.some((c) => c[0] === 'browser'));
  assert.deepEqual(await page.evaluate(() => window.__calls.find((c) => c[0] === 'browser')), ['browser', 'https://checkout.stripe.test/s']);
  assert.equal(calls.checkout[0].line_items[0].price, 'price_professional_year');
  assert.match(await page.textContent('#upgrade-status'), /Checkout opened in your browser/);

  // 4. Stripe confirms payment (webhook); back in Outlook the plan updates on focus/refresh.
  subs.sub_1 = { id: 'sub_1', status: 'active', customer: 'cus_alice', metadata: { account_id: 't1:alice', tid: 't1', plan: 'professional' },
    items: { data: [{ quantity: 1, current_period_end: 1830000000, price: { lookup_key: 'professional_year', recurring: { interval: 'year' } } }] } };
  const hook = await fetch(`http://127.0.0.1:${port}/api/stripe/webhook`, { method: 'POST', headers: { 'stripe-signature': 'x' },
    body: JSON.stringify({ id: 'evt_1', type: 'checkout.session.completed', data: { object: { mode: 'subscription', subscription: 'sub_1' } } }) });
  assert.equal(hook.status, 200);
  await page.click('#close-upgrade');
  await page.click('#open-settings');
  await page.click('#refresh-plan');
  await page.waitForFunction(() => document.getElementById('connection-label').textContent === 'Professional');
  assert.match(await page.textContent('#settings-connection'), /Professional · annual · 0 of 3000 replies this month/);
  assert.equal(await page.isVisible('#manage-billing'), true);
  assert.equal(await page.isHidden('#see-plans'), true);
  await shot(page, 'service-settings');
  await page.click('#manage-billing');
  await page.waitForFunction(() => window.__calls.filter((c) => c[0] === 'browser').length === 2);
  assert.equal(calls.portal.length, 1);
  await page.click('#close-settings');
  await page.click('#generate');
  await page.waitForSelector('#view-draft:not([hidden])');
  assert.equal(calls.claude.at(-1), 'claude-opus-5-5', 'Professional drafts use Opus');
  await page.close();

  // 5. iPhone: an expired trial shows no prices, plans or purchase links.
  db.raw.prepare(`INSERT INTO accounts (id, tid, email, name, created_at, trial_ends_at, trial_replies) VALUES ('t1:bob','t1','bob@example.com','bob',0,1,0)`).run();
  const phone = await openPane('bob', 'iOS');
  await phone.waitForFunction(() => document.getElementById('connection-label').textContent === 'Trial ended');
  await phone.fill('#points', 'Agree');
  await phone.click('#generate');
  await phone.waitForFunction(() => /trial has ended/.test(document.getElementById('status').textContent));
  assert.equal(await phone.isHidden('#upgrade'), true, 'no plans sheet on iOS');
  assert.equal(await phone.isHidden('#banner-action'), true, 'no upgrade button on iOS');
  await phone.click('#open-settings');
  assert.equal(await phone.isHidden('#see-plans'), true);
  assert.equal(await phone.isHidden('#manage-billing'), true);
  assert.doesNotMatch(await phone.textContent('body'), /A\$\d/);

  assert.deepEqual(errors, []);
  console.log('Subscription pane (browser): sign-in trial, Haiku insights, Opus drafts, countdown, paywall, plans + annual toggle, Stripe checkout, webhook upgrade, billing portal, no purchase UI on iOS passed.');
} finally {
  await browser.close();
  server.close();
}
