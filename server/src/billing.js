// Stripe: Checkout for upgrades, the customer portal for changes/cancellation, and the
// webhook that keeps plans in the database in step with Stripe.
import { PRICES, TEAM_MIN_SEATS, lookupKey } from './plans.js';
import { setStripeCustomer } from './db.js';

const priceCache = new Map();

async function priceId(stripe, key) {
  if (priceCache.has(key)) return priceCache.get(key);
  const list = await stripe.prices.list({ lookup_keys: [key], active: true, limit: 1 });
  const id = list.data[0]?.id;
  if (!id) throw new Error(`Stripe price "${key}" is missing. Run scripts/stripe-setup.mjs.`);
  priceCache.set(key, id);
  return id;
}

async function ensureCustomer(stripe, db, account) {
  if (account.stripe_customer_id) return account.stripe_customer_id;
  const customer = await stripe.customers.create({
    email: account.email || undefined, name: account.name || undefined,
    metadata: { account_id: account.id, tid: account.tid },
  });
  await setStripeCustomer(db, account.id, customer.id);
  return customer.id;
}

export async function createCheckout(stripe, db, env, account, org, { plan, interval, seats }) {
  const key = lookupKey(plan, interval);
  if (!key) return { status: 400, body: { error: 'Choose a valid plan.' } };
  if (plan === 'team' && org && ['active', 'trialing', 'past_due'].includes(org.status)) {
    return { status: 409, body: { error: 'Your organisation already has a Team subscription. Ask its owner to add seats.' } };
  }
  const customer = await ensureCustomer(stripe, db, account);
  const quantity = plan === 'team' ? Math.min(500, Math.max(TEAM_MIN_SEATS, Number(seats) || TEAM_MIN_SEATS)) : 1;
  const metadata = { account_id: account.id, tid: account.tid, plan };
  const session = await stripe.checkout.sessions.create({
    mode: 'subscription',
    customer,
    client_reference_id: account.id,
    line_items: [{
      price: await priceId(stripe, key), quantity,
      ...(plan === 'team' ? { adjustable_quantity: { enabled: true, minimum: TEAM_MIN_SEATS, maximum: 500 } } : {}),
    }],
    subscription_data: { metadata },
    metadata,
    allow_promotion_codes: true,
    tax_id_collection: { enabled: true },
    customer_update: { name: 'auto', address: 'auto' },
    success_url: `${env.PUBLIC_URL}/billing/done.html?result=success`,
    cancel_url: `${env.PUBLIC_URL}/billing/done.html?result=cancelled`,
  });
  return { status: 200, body: { url: session.url } };
}

export async function createPortal(stripe, env, account, org) {
  const customer = account.stripe_customer_id || (org && org.owner_id === account.id ? org.stripe_customer_id : '');
  if (!customer) return { status: 404, body: { error: 'There is no billing account to manage yet.' } };
  const session = await stripe.billingPortal.sessions.create({ customer, return_url: `${env.PUBLIC_URL}/billing/done.html?result=updated` });
  return { status: 200, body: { url: session.url } };
}

function fromSubscription(sub) {
  const item = sub.items?.data?.[0] || {};
  const key = item.price?.lookup_key || '';
  const price = PRICES[key];
  return {
    plan: price?.plan || sub.metadata?.plan || '',
    interval: price?.interval || item.price?.recurring?.interval || '',
    quantity: item.quantity || 1,
    status: sub.status,
    period_end: ((item.current_period_end ?? sub.current_period_end) || 0) * 1000,
    customer: typeof sub.customer === 'string' ? sub.customer : sub.customer?.id || '',
    id: sub.id,
    account_id: sub.metadata?.account_id || '',
    tid: sub.metadata?.tid || '',
  };
}

export async function applySubscription(db, sub) {
  const s = fromSubscription(sub);
  if (!s.account_id) return;
  if (s.plan === 'team') {
    await db.batch([
      db.prepare(`INSERT INTO orgs (tid, owner_id, seats, status, interval, period_end, stripe_customer_id, stripe_subscription_id)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        ON CONFLICT(tid) DO UPDATE SET owner_id = excluded.owner_id, seats = excluded.seats, status = excluded.status,
          interval = excluded.interval, period_end = excluded.period_end, stripe_customer_id = excluded.stripe_customer_id,
          stripe_subscription_id = excluded.stripe_subscription_id`)
        .bind(s.tid, s.account_id, s.quantity, s.status, s.interval, s.period_end, s.customer, s.id),
      db.prepare('UPDATE accounts SET team_member = 1, stripe_customer_id = ? WHERE id = ?').bind(s.customer, s.account_id),
    ]);
    return;
  }
  await db.prepare(`UPDATE accounts SET plan = ?, status = ?, interval = ?, period_end = ?, stripe_customer_id = ?, stripe_subscription_id = ?
    WHERE id = ? AND (stripe_subscription_id = '' OR stripe_subscription_id = ? OR ? IN ('active','trialing'))`)
    .bind(s.plan, s.status, s.interval, s.period_end, s.customer, s.id, s.account_id, s.id, s.status).run();
}

export async function handleWebhook(stripe, db, env, request) {
  const signature = request.headers.get('stripe-signature');
  const body = await request.text();
  let event;
  try {
    event = await stripe.webhooks.constructEventAsync(body, signature, env.STRIPE_WEBHOOK_SECRET, undefined, stripe.__cryptoProvider);
  } catch {
    return { status: 400, body: { error: 'Invalid signature.' } };
  }
  const fresh = await db.prepare('INSERT OR IGNORE INTO stripe_events (id, received_at) VALUES (?, ?)').bind(event.id, Date.now()).run();
  if (fresh.meta && fresh.meta.changes === 0) return { status: 200, body: { received: true, duplicate: true } };
  const obj = event.data.object;
  switch (event.type) {
    case 'checkout.session.completed':
      if (obj.mode === 'subscription' && obj.subscription) {
        const sub = await stripe.subscriptions.retrieve(typeof obj.subscription === 'string' ? obj.subscription : obj.subscription.id);
        await applySubscription(db, sub);
      }
      break;
    case 'customer.subscription.created':
    case 'customer.subscription.updated':
    case 'customer.subscription.deleted':
    case 'customer.subscription.paused':
    case 'customer.subscription.resumed':
      await applySubscription(db, obj);
      break;
    default:
      break;
  }
  return { status: 200, body: { received: true } };
}
