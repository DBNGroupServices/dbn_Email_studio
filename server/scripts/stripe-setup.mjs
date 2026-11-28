// Creates (or updates) the Stripe products and prices Bulwark HQ Assistant uses.
// Safe to run more than once: prices are found by lookup key and only created if missing.
//
//   STRIPE_SECRET_KEY=sk_test_... node scripts/stripe-setup.mjs
//
// Run it with a test key first, then again with the live key.
// Prices are in AUD and include GST (tax_behavior "inclusive").
import Stripe from 'stripe';
import { PRICES, PLANS } from '../src/plans.js';

const key = process.env.STRIPE_SECRET_KEY;
if (!key) {
  console.error('Set STRIPE_SECRET_KEY first, for example:\n  STRIPE_SECRET_KEY=sk_test_... node scripts/stripe-setup.mjs');
  process.exit(1);
}
const stripe = new Stripe(key);

const PRODUCTS = {
  essentials: { name: 'Bulwark HQ Assistant – Essentials', description: `${PLANS.essentials.replies} replies a month in your voice.` },
  professional: { name: 'Bulwark HQ Assistant – Professional', description: `${PLANS.professional.premium} premium replies a month, then standard replies (fair use).` },
  team: { name: 'Bulwark HQ Assistant – Team', description: 'Professional for every seat in your organisation. Minimum 3 seats.' },
};

async function product(plan) {
  const found = await stripe.products.search({ query: `metadata['bulwark_plan']:'${plan}'` });
  if (found.data[0]) {
    const p = await stripe.products.update(found.data[0].id, { name: PRODUCTS[plan].name, description: PRODUCTS[plan].description });
    console.log(`product  ${plan.padEnd(13)} ${p.id} (updated)`);
    return p;
  }
  const p = await stripe.products.create({
    name: PRODUCTS[plan].name, description: PRODUCTS[plan].description,
    metadata: { bulwark_plan: plan }, tax_code: 'txcd_10103001', // Software as a service, business use
  });
  console.log(`product  ${plan.padEnd(13)} ${p.id} (created)`);
  return p;
}

const products = {};
for (const plan of Object.keys(PRODUCTS)) products[plan] = await product(plan);

for (const [lookup, p] of Object.entries(PRICES)) {
  const existing = (await stripe.prices.list({ lookup_keys: [lookup], limit: 1 })).data[0];
  const same = existing && existing.unit_amount === p.amount && existing.currency === 'aud' &&
    existing.recurring?.interval === p.interval && existing.product === products[p.plan].id;
  if (same) {
    if (!existing.active) await stripe.prices.update(existing.id, { active: true });
    console.log(`price    ${lookup.padEnd(19)} ${existing.id} A$${(p.amount / 100).toFixed(2)}/${p.interval} (exists)`);
    continue;
  }
  // Prices can't be edited: create a new one and move the lookup key onto it.
  const price = await stripe.prices.create({
    product: products[p.plan].id, currency: 'aud', unit_amount: p.amount,
    recurring: { interval: p.interval }, tax_behavior: 'inclusive',
    lookup_key: lookup, transfer_lookup_key: true,
    nickname: `${PLANS[p.plan].label} ${p.interval === 'year' ? 'annual' : 'monthly'}${p.plan === 'team' ? ' (per seat)' : ''}`,
    metadata: { bulwark_plan: p.plan },
  });
  if (existing) await stripe.prices.update(existing.id, { active: false });
  console.log(`price    ${lookup.padEnd(19)} ${price.id} A$${(p.amount / 100).toFixed(2)}/${p.interval} (created)`);
}

console.log(`
Done. Next, in the Stripe Dashboard:
  1. Settings → Tax: turn on Stripe Tax and add your Australian GST registration (or leave tax off if not registered).
  2. Settings → Billing → Customer portal: allow cancelling, switching between the three products and updating seat quantity.
  3. Developers → Webhooks: add https://assistant.bulwarkhq.com.au/api/stripe/webhook with the events
     checkout.session.completed, customer.subscription.created/updated/deleted/paused/resumed,
     then save its signing secret with:  npx wrangler secret put STRIPE_WEBHOOK_SECRET`);
