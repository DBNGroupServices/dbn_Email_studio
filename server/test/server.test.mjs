// Server tests: real request handler + real SQL schema on an in-memory database,
// with Microsoft sign-in, Claude and Stripe mocked. Run: npm test
import { test } from 'node:test';
import assert from 'node:assert/strict';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { handle } from '../src/index.js';
import { trialEndsAt, effectivePlan, replyDecision, MODELS } from '../src/plans.js';
import { AuthError } from '../src/auth.js';
import { makeD1 } from './d1-shim.mjs';

const here = path.dirname(fileURLToPath(import.meta.url));
const MIGRATION = path.join(here, '..', 'migrations', '0001_init.sql');
const DAY = 86400000;
// Monday 30 Nov 2026, 10:00 in Melbourne (AEDT, UTC+11) = Sunday 29 Nov 23:00 UTC.
const MON_10AM = Date.UTC(2026, 10, 29, 23, 0, 0);

function reply(text) {
  return { stop_reason: 'end_turn', model: 'x', content: [{ type: 'text', text: JSON.stringify(text) }] };
}
const DRAFT = { subject: 'Re: test', body: 'Hi Jordan,\n\nThanks.\n\nSam', checks: [] };
const SUGGEST = { summary: 'They want a re-test.', audience: 'supplier', priority: 'High', need: 'Fix', mood: 'Neutral',
  suggestions: [{ title: 'Agree', tag: 'BEST FIT', hint: 'h', instruction: 'Agree and propose a date.', preview: { formal: 'F', balanced: 'B', friendly: 'R' } }] };

function setup(opts = {}) {
  const db = makeD1(MIGRATION);
  const calls = { claude: [], stripe: [] };
  let now = opts.now ?? MON_10AM;
  const users = { alice: { id: 't1:alice', tid: 't1', email: 'alice@example.com', name: 'Alice' },
    bob: { id: 't1:bob', tid: 't1', email: 'bob@example.com', name: 'Bob' },
    carol: { id: 't1:carol', tid: 't1', email: 'carol@example.com', name: 'Carol' },
    dave: { id: 't1:dave', tid: 't1', email: 'dave@example.com', name: 'Dave' } };
  const respond = (params) => {
    calls.claude.push(params);
    return params.output_config.format.schema.properties.suggestions ? reply(SUGGEST) : reply(DRAFT);
  };
  const subscriptions = {};
  const deps = {
    now: () => now,
    verify: async (req) => {
      const who = /^Bearer (\w+)$/.exec(req.headers.get('authorization') || '')?.[1];
      if (!users[who]) throw new AuthError('Sign-in required.');
      return users[who];
    },
    anthropic: { messages: { create: async (p) => respond(p) }, beta: { messages: { create: async (p) => respond(p) } } },
    stripe: {
      prices: { list: async ({ lookup_keys }) => ({ data: [{ id: 'price_' + lookup_keys[0] }] }) },
      customers: { create: async (p) => { calls.stripe.push(['customer', p]); return { id: 'cus_' + p.metadata.account_id.split(':')[1] }; } },
      checkout: { sessions: { create: async (p) => { calls.stripe.push(['checkout', p]); return { url: 'https://checkout.stripe.test/s' }; } } },
      billingPortal: { sessions: { create: async (p) => { calls.stripe.push(['portal', p]); return { url: 'https://billing.stripe.test/p' }; } } },
      subscriptions: { retrieve: async (id) => subscriptions[id] },
      webhooks: { constructEventAsync: async (body, sig) => { if (sig !== 'good') throw new Error('bad sig'); return JSON.parse(body); } },
    },
  };
  const env = { DB: db, PUBLIC_URL: 'https://assistant.bulwarkhq.com.au', MS_CLIENT_ID: 'client-123', TRIAL_TZ: 'Australia/Melbourne', STRIPE_WEBHOOK_SECRET: 'whsec' };
  const api = async (who, p, body, method) => {
    const init = { method: method || (body === undefined ? 'GET' : 'POST'), headers: { authorization: 'Bearer ' + who, 'content-type': 'application/json' } };
    if (body !== undefined) init.body = JSON.stringify(body);
    const res = await handle(new Request('https://assistant.bulwarkhq.com.au' + p, init), env, deps);
    return { status: res.status, body: res.headers.get('content-type')?.includes('json') ? await res.json() : await res.text() };
  };
  const webhook = async (event, sig = 'good') => {
    const res = await handle(new Request('https://assistant.bulwarkhq.com.au/api/stripe/webhook', { method: 'POST', headers: { 'stripe-signature': sig }, body: JSON.stringify(event) }), env, deps);
    return { status: res.status, body: await res.json() };
  };
  const sub = (id, account, lookup, status = 'active', quantity = 1, tid = 't1') => {
    const s = { id, status, customer: 'cus_' + account.split(':')[1], metadata: { account_id: account, tid, plan: lookup.split('_')[0] },
      items: { data: [{ quantity, current_period_end: 1800000000, price: { lookup_key: lookup, recurring: { interval: lookup.split('_')[1] } } }] } };
    subscriptions[id] = s; return s;
  };
  return { db, env, deps, calls, api, webhook, sub, setNow: (t) => { now = t; } };
}
const TASK = { task: { mode: 'reply', source: 'Hi Sam, can we re-test?', points: 'Agree', tone: 'Use profile' }, settings: { name: 'Sam' } };

test('trial ends at the end of the third working day (Melbourne time)', () => {
  assert.equal(new Date(MON_10AM + 11 * 3600000).getUTCDay(), 1, 'fixture is a Monday in Melbourne');
  // Monday 10am -> end of Wednesday 2 Dec, 23:59:59.999 AEDT = 12:59:59.999 UTC.
  assert.equal(trialEndsAt(MON_10AM), Date.UTC(2026, 11, 2, 12, 59, 59, 999));
  // Friday 4 Dec, 4pm -> Fri, Mon, Tue -> end of Tuesday 8 Dec.
  assert.equal(trialEndsAt(Date.UTC(2026, 11, 4, 5, 0, 0)), Date.UTC(2026, 11, 8, 12, 59, 59, 999));
  // Saturday 5 Dec -> Mon, Tue, Wed -> end of Wednesday 9 Dec.
  assert.equal(trialEndsAt(Date.UTC(2026, 11, 5, 1, 0, 0)), Date.UTC(2026, 11, 9, 12, 59, 59, 999));
  // Winter (AEST, UTC+10): Monday 1 Jun 2026, 9am -> end of Wednesday 3 Jun = 13:59:59.999 UTC.
  assert.equal(trialEndsAt(Date.UTC(2026, 4, 31, 23, 0, 0)), Date.UTC(2026, 5, 3, 13, 59, 59, 999));
});

test('plan rules pick the right model and stop at the limits', () => {
  const acc = { plan: 'none', status: '', team_member: 0, trial_ends_at: 10, trial_replies: 0 };
  assert.equal(effectivePlan(acc, null, 5), 'trial');
  assert.equal(effectivePlan({ ...acc, trial_replies: 50 }, null, 5), 'expired');
  assert.equal(effectivePlan(acc, null, 11), 'expired');
  assert.equal(effectivePlan({ ...acc, plan: 'essentials', status: 'active' }, null, 11), 'essentials');
  assert.equal(effectivePlan({ ...acc, plan: 'essentials', status: 'canceled' }, null, 11), 'expired');
  assert.equal(effectivePlan({ ...acc, team_member: 1 }, { status: 'active' }, 11), 'team');
  assert.equal(effectivePlan({ ...acc, stripe_subscription_id: 'sub_old', status: 'canceled' }, null, 5), 'expired', 'no second trial');
  assert.equal(replyDecision('essentials', { replies: 10, premium: 0 }, acc).model, MODELS.standard);
  assert.equal(replyDecision('essentials', { replies: 300, premium: 0 }, acc).code, 'limit_reached');
  assert.equal(replyDecision('professional', { replies: 10, premium: 599 }, acc).model, MODELS.premium);
  assert.equal(replyDecision('professional', { replies: 700, premium: 600 }, acc).model, MODELS.standard);
  assert.equal(replyDecision('trial', { replies: 0, premium: 0 }, acc).model, MODELS.premium);
});

test('sign-in is required, config.js advertises service mode', async () => {
  const t = setup();
  assert.equal((await t.api('nobody', '/api/me')).status, 401);
  const cfg = await handle(new Request('https://assistant.bulwarkhq.com.au/config.js'), t.env, t.deps);
  assert.match(await cfg.text(), /"mode":"service".*"clientId":"client-123"/);
});

test('new user gets a trial; suggestions use Haiku; replies use Opus and count down', async () => {
  const t = setup();
  const me = (await t.api('alice', '/api/me')).body;
  assert.equal(me.plan, 'trial');assert.equal(me.trial.replies_left, 50);
  const s = await t.api('alice', '/api/suggest', { email: { text: 'Please re-test', subject: 'Test' }, settings: {} });
  assert.equal(s.status, 200);assert.equal(s.body.priority, 'High');
  assert.equal(t.calls.claude.at(-1).model, 'claude-haiku-5-5');
  assert.equal(t.calls.claude.at(-1).output_config.effort, 'low');
  const g = await t.api('alice', '/api/generate', TASK);
  assert.equal(g.status, 200);assert.equal(g.body.body, DRAFT.body);
  const call = t.calls.claude.at(-1);
  assert.equal(call.model, 'claude-opus-5-5');assert.equal(call.fallbacks, 'default');assert.equal(call.output_config.effort, 'medium');
  assert.match(call.system, /Security disclosure guard/);
  assert.equal(g.body.account.trial.replies_left, 49);
});

test('trial ends after 50 replies or after the third working day', async () => {
  const t = setup();
  await t.api('alice', '/api/me');
  t.db.raw.prepare("UPDATE accounts SET trial_replies = 49 WHERE id = 't1:alice'").run();
  assert.equal((await t.api('alice', '/api/generate', TASK)).status, 200);
  const blocked = await t.api('alice', '/api/generate', TASK);
  assert.equal(blocked.status, 402);assert.equal(blocked.body.code, 'plan_required');
  assert.equal((await t.api('alice', '/api/suggest', { email: { text: 'x' } })).status, 402);
  const t2 = setup();
  await t2.api('bob', '/api/me');
  t2.setNow(MON_10AM + 3 * DAY); // Thursday
  assert.equal((await t2.api('bob', '/api/me')).body.plan, 'expired');
});

test('checkout, webhook and Essentials limits', async () => {
  const t = setup();
  await t.api('alice', '/api/me');
  const c = await t.api('alice', '/api/billing/checkout', { plan: 'essentials', interval: 'month' });
  assert.equal(c.status, 200);assert.equal(c.body.url, 'https://checkout.stripe.test/s');
  const session = t.calls.stripe.find((x) => x[0] === 'checkout')[1];
  assert.equal(session.mode, 'subscription');assert.equal(session.customer, 'cus_alice');
  assert.equal(session.line_items[0].price, 'price_essentials_month');assert.equal(session.line_items[0].quantity, 1);
  assert.equal(session.subscription_data.metadata.account_id, 't1:alice');
  assert.match(session.success_url, /billing\/done\.html\?result=success/);
  assert.equal((await t.api('alice', '/api/billing/checkout', { plan: 'gold', interval: 'month' })).status, 400);
  // Webhook: bad signature refused; good one activates the plan; replays are ignored.
  t.sub('sub_1', 't1:alice', 'essentials_month');
  const event = { id: 'evt_1', type: 'checkout.session.completed', data: { object: { mode: 'subscription', subscription: 'sub_1' } } };
  assert.equal((await t.webhook(event, 'bad')).status, 400);
  assert.equal((await t.webhook(event)).status, 200);
  assert.equal((await t.webhook(event)).body.duplicate, true);
  const me = (await t.api('alice', '/api/me')).body;
  assert.equal(me.plan, 'essentials');assert.equal(me.can_manage_billing, true);assert.equal(me.trial, undefined);
  await t.api('alice', '/api/generate', TASK);
  assert.equal(t.calls.claude.at(-1).model, 'claude-sonnet-5-5');
  t.db.raw.prepare("UPDATE usage SET replies = 300 WHERE account_id = 't1:alice'").run();
  const over = await t.api('alice', '/api/generate', TASK);
  assert.equal(over.status, 402);assert.equal(over.body.code, 'limit_reached');
  // Portal and cancellation.
  assert.equal((await t.api('alice', '/api/billing/portal', {})).body.url, 'https://billing.stripe.test/p');
  await t.webhook({ id: 'evt_2', type: 'customer.subscription.deleted', data: { object: { ...t.sub('sub_1', 't1:alice', 'essentials_month', 'canceled') } } });
  assert.equal((await t.api('alice', '/api/me')).body.plan, 'expired');
});

test('Professional: 600 premium replies then standard model', async () => {
  const t = setup();
  await t.api('alice', '/api/me');
  await t.webhook({ id: 'evt_p', type: 'customer.subscription.created', data: { object: t.sub('sub_p', 't1:alice', 'professional_year') } });
  const me = (await t.api('alice', '/api/me')).body;
  assert.equal(me.plan, 'professional');assert.equal(me.interval, 'year');assert.equal(me.usage.premium_limit, 600);
  await t.api('alice', '/api/generate', TASK);assert.equal(t.calls.claude.at(-1).model, 'claude-opus-5-5');
  t.db.raw.prepare("UPDATE usage SET premium = 600, replies = 600 WHERE account_id = 't1:alice'").run();
  const g = await t.api('alice', '/api/generate', TASK);
  assert.equal(g.status, 200);assert.equal(t.calls.claude.at(-1).model, 'claude-sonnet-5-5');assert.equal(g.body.premium, false);
});

test('Team: minimum 3 seats, colleagues in the same tenant claim seats until full', async () => {
  const t = setup();
  await t.api('alice', '/api/me');
  await t.api('alice', '/api/billing/checkout', { plan: 'team', interval: 'month', seats: 1 });
  const li = t.calls.stripe.find((x) => x[0] === 'checkout')[1].line_items[0];
  assert.equal(li.quantity, 3);assert.equal(li.adjustable_quantity.minimum, 3);
  await t.webhook({ id: 'evt_t', type: 'customer.subscription.created', data: { object: t.sub('sub_t', 't1:alice', 'team_month', 'active', 3) } });
  assert.equal((await t.api('alice', '/api/me')).body.plan, 'team');
  assert.equal((await t.api('bob', '/api/me')).body.plan, 'team');
  assert.equal((await t.api('carol', '/api/me')).body.plan, 'team');
  const dave = (await t.api('dave', '/api/me')).body; // 4th person: no seat left, so the trial applies
  assert.equal(dave.plan, 'trial');assert.equal(dave.team.member, false);
  assert.equal((await t.api('bob', '/api/billing/checkout', { plan: 'team', interval: 'month' })).status, 409);
  assert.equal((await t.api('alice', '/api/me')).body.team.owner, true);
});

test('invalid input is refused without calling Claude', async () => {
  const t = setup();
  const before = t.calls.claude.length;
  assert.equal((await t.api('alice', '/api/generate', { task: { mode: 'send-now', source: 'x' } })).status, 400);
  assert.equal((await t.api('alice', '/api/generate', { ...TASK, profile: { name: '' } })).status, 400);
  assert.equal((await t.api('alice', '/api/suggest', { email: { text: '' } })).status, 400);
  assert.equal((await t.api('alice', '/api/me', {}, 'DELETE')).status, 405);
  assert.equal(t.calls.claude.length, before);
});
