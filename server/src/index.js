// Bulwark HQ Assistant: Cloudflare Worker entry point.
// Serves /api/* and /config.js; everything else is the add-in's static files (../addin).
import Anthropic from '@anthropic-ai/sdk';
import Stripe from 'stripe';
import Claude from '../../addin/claude-client.js';
import Store from '../../addin/store.js';
import { makeVerifier, AuthError } from './auth.js';
import { getOrCreateAccount, getOrg, claimSeat, getUsage, countUse } from './db.js';
import { effectivePlan, replyDecision, suggestDecision, summary, MODELS } from './plans.js';
import { createCheckout, createPortal, handleWebhook } from './billing.js';

const MAX_BODY = 32 * 1024 * 1024;
const MESSAGES = {
  plan_required: 'Your free trial has ended. Choose a plan to keep using Bulwark HQ Assistant.',
  limit_reached: 'You have used this month’s replies on your plan. Upgrade, or wait until the 1st of next month.',
};

function json(status, body, extra = {}) {
  return new Response(JSON.stringify(body), {
    status,
    headers: { 'content-type': 'application/json; charset=utf-8', 'cache-control': 'no-store', 'x-content-type-options': 'nosniff', ...extra },
  });
}

export function buildDeps(env) {
  const stripe = new Stripe(env.STRIPE_SECRET_KEY, { httpClient: Stripe.createFetchHttpClient() });
  stripe.__cryptoProvider = Stripe.createSubtleCryptoProvider();
  return {
    stripe,
    anthropic: new Anthropic({ apiKey: env.ANTHROPIC_API_KEY, maxRetries: 2, timeout: 180000 }),
    verify: makeVerifier(env),
    now: () => Date.now(),
  };
}

// One Claude call with the service key. Opus/Sonnet use server-side refusal fallback.
async function callClaude(anthropic, model, system, content, schema, effort, maxTokens) {
  const params = { model, max_tokens: maxTokens, system, messages: [{ role: 'user', content }],
    output_config: { format: { type: 'json_schema', schema }, effort } };
  try {
    return model === MODELS.insights
      ? await anthropic.messages.create(params)
      : await anthropic.beta.messages.create({ ...params, betas: ['server-side-fallback-2026-07-01'], fallbacks: 'default' });
  } catch (error) {
    const status = error && error.status;
    const e = new Error(status === 429 || status === 529 || status >= 500
      ? 'The AI service is busy. Please try again in a moment.'
      : 'The AI service could not process this request. Try fewer or smaller attachments.');
    e.httpStatus = status === 400 ? 400 : 503;
    throw e;
  }
}

function cleanSettings(s) {
  const src = s && typeof s === 'object' ? s : {};
  const str = (v, max) => (typeof v === 'string' ? v.slice(0, max) : '');
  return { name: str(src.name, 200), role: str(src.role, 200), organisation: str(src.organisation, 200),
    signature: str(src.signature, 1000), style: str(src.style, 2000), security_guard: src.security_guard !== false };
}

async function readJson(request) {
  const len = Number(request.headers.get('content-length') || 0);
  if (len > MAX_BODY) { const e = new Error('Request too large. Attach up to 24 MB of files.'); e.httpStatus = 413; throw e; }
  try { return await request.json(); } catch { const e = new Error('Invalid request.'); e.httpStatus = 400; throw e; }
}

export async function handle(request, env, deps) {
  const url = new URL(request.url);
  const path = url.pathname;

  if (path === '/config.js') {
    const cfg = { mode: 'service', clientId: env.MS_CLIENT_ID, publicUrl: env.PUBLIC_URL };
    return new Response(`window.DBN_CONFIG=${JSON.stringify(cfg)};`, { headers: { 'content-type': 'text/javascript; charset=utf-8', 'cache-control': 'no-store' } });
  }
  if (path === '/api/health') return json(200, { ok: true });
  if (path === '/api/stripe/webhook' && request.method === 'POST') {
    const r = await handleWebhook(deps.stripe, env.DB, env, request);
    return json(r.status, r.body);
  }
  if (!path.startsWith('/api/')) return env.ASSETS ? env.ASSETS.fetch(request) : json(404, { error: 'Not found.' });

  try {
    const user = await deps.verify(request);
    const now = deps.now();
    let account = await getOrCreateAccount(env.DB, user, now, env.TRIAL_TZ || 'Australia/Melbourne');
    const org = await getOrg(env.DB, account.tid);
    account = await claimSeat(env.DB, account, org);
    const usage = await getUsage(env.DB, account.id, now);
    const plan = effectivePlan(account, org, now);

    if (path === '/api/me' && request.method === 'GET') return json(200, summary(account, org, usage, now));

    if (request.method !== 'POST') return json(405, { error: 'Method not allowed.' });
    const body = await readJson(request);

    if (path === '/api/suggest') {
      const d = suggestDecision(plan, usage);
      if (!d.ok) return json(402, { code: d.code, error: MESSAGES[d.code] });
      const req = Claude.buildSuggest(body.email || {}, cleanSettings(body.settings));
      const message = await callClaude(deps.anthropic, d.model, req.system, req.content, req.schema, 'low', 4000);
      const result = Claude.readSuggestions(message);
      await countUse(env.DB, account, now, { suggest: 1 });
      return json(200, result);
    }

    if (path === '/api/generate') {
      const d = replyDecision(plan, usage, account);
      if (!d.ok) return json(402, { code: d.code, error: MESSAGES[d.code], account: summary(account, org, usage, now) });
      const settings = cleanSettings(body.settings);
      const profile = body.profile ? Store.validateProfile(body.profile) : null;
      const task = body.task && typeof body.task === 'object' ? body.task : {};
      const req = Claude.buildRequest(task, settings, profile);
      const message = await callClaude(deps.anthropic, d.model, req.system, req.content, req.schema, 'medium', 16000);
      const result = Claude.readResult(message);
      result.warnings = req.warnings;
      result.premium = d.premium;
      await countUse(env.DB, account, now, { reply: 1, premium: d.premium && plan !== 'trial' ? 1 : 0, trial: plan === 'trial' });
      const after = await getUsage(env.DB, account.id, now);
      const fresh = plan === 'trial' ? { ...account, trial_replies: account.trial_replies + 1 } : account;
      result.account = summary(fresh, org, after, now);
      return json(200, result);
    }

    if (path === '/api/billing/checkout') {
      const r = await createCheckout(deps.stripe, env.DB, env, account, org, body);
      return json(r.status, r.body);
    }
    if (path === '/api/billing/portal') {
      const r = await createPortal(deps.stripe, env, account, org);
      return json(r.status, r.body);
    }
    return json(404, { error: 'Not found.' });
  } catch (error) {
    if (error instanceof AuthError) return json(401, { code: 'auth', error: error.message });
    if (error instanceof Claude.StudioError || (error && error.name === 'StudioError')) return json(400, { error: error.message });
    if (error && error.httpStatus) return json(error.httpStatus, { error: error.message });
    if (error && /profile|Tone controls|example emails|valid length/i.test(error.message || '')) return json(400, { error: error.message });
    console.error('request failed', path, error && error.message); // never logs request bodies
    return json(500, { error: 'Something went wrong. Your inputs are kept; please try again.' });
  }
}

let cached = null;
export default {
  fetch(request, env) {
    if (env.__deps) return handle(request, env, env.__deps);
    if (!cached || cached.env !== env) cached = { env, deps: buildDeps(env) };
    return handle(request, env, cached.deps);
  },
};
