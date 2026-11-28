// Plans, limits and the trial. Pure functions: no I/O, easy to test.
// Agreed 28 Nov 2026 (see the "Pricing & Marketplace Plan" doc). Prices include GST.

export const MODELS = {
  insights: 'claude-haiku-5-5',      // insights + suggested replies, every opened email
  standard: 'claude-sonnet-5-5',     // Essentials drafts; Professional/Team after the premium allowance
  premium: 'claude-opus-5-5',        // Professional/Team drafts up to the premium allowance; trial
};

export const PLANS = {
  trial:        { label: 'Trial',        replies: 50,   premium: 50,  suggests: 300 },
  essentials:   { label: 'Essentials',   replies: 300,  premium: 0,   suggests: 1500 },
  professional: { label: 'Professional', replies: 3000, premium: 600, suggests: 3000 }, // 3000 = fair-use ceiling
  team:         { label: 'Team',         replies: 3000, premium: 600, suggests: 3000 },
};

// Stripe price lookup keys (created by scripts/stripe-setup.mjs). AUD cents, GST inclusive.
export const PRICES = {
  essentials_month:   { plan: 'essentials',   interval: 'month', amount: 1500 },
  essentials_year:    { plan: 'essentials',   interval: 'year',  amount: 15000 },
  professional_month: { plan: 'professional', interval: 'month', amount: 2900 },
  professional_year:  { plan: 'professional', interval: 'year',  amount: 29000 },
  team_month:         { plan: 'team',         interval: 'month', amount: 2500 },  // per seat
  team_year:          { plan: 'team',         interval: 'year',  amount: 25000 }, // per seat
};
export const TEAM_MIN_SEATS = 3;
export const TRIAL_WORKING_DAYS = 3;
const LIVE = new Set(['active', 'trialing', 'past_due']); // past_due: Stripe is still retrying the card

export function lookupKey(plan, interval) {
  const key = `${plan}_${interval}`;
  return PRICES[key] ? key : null;
}

// Local calendar date parts and UTC offset (minutes) for a time zone at an instant.
function zoned(ms, tz) {
  const parts = Object.fromEntries(new Intl.DateTimeFormat('en-AU', {
    timeZone: tz, year: 'numeric', month: '2-digit', day: '2-digit', weekday: 'short', timeZoneName: 'longOffset',
  }).formatToParts(new Date(ms)).map((p) => [p.type, p.value]));
  const m = /GMT([+-])(\d{2}):?(\d{2})?/.exec(parts.timeZoneName || '') || [];
  const offset = m[1] ? (m[1] === '-' ? -1 : 1) * (Number(m[2]) * 60 + Number(m[3] || 0)) : 0;
  return { y: Number(parts.year), mo: Number(parts.month), d: Number(parts.day), weekday: parts.weekday, offset };
}

// End of the 3rd working day (Mon-Fri) in the given time zone, counting the sign-up day
// when it is a weekday. Public holidays are not excluded.
export function trialEndsAt(startMs, tz = 'Australia/Melbourne', days = TRIAL_WORKING_DAYS) {
  let day = zoned(startMs, tz);
  let cursor = Date.UTC(day.y, day.mo - 1, day.d);
  let counted = 0;
  for (let i = 0; i < 14; i++) {
    const wd = new Date(cursor).getUTCDay(); // calendar weekday of the local date
    if (wd >= 1 && wd <= 5) counted++;
    if (counted === days) break;
    cursor += 86400000;
  }
  // 23:59:59.999 local on that date, converted to UTC with that date's offset.
  const localEnd = cursor + 86400000 - 1;
  const offset = zoned(localEnd - 12 * 3600000, tz).offset;
  return localEnd - offset * 60000;
}

export function monthKey(ms) {
  const d = new Date(ms);
  return `${d.getUTCFullYear()}-${String(d.getUTCMonth() + 1).padStart(2, '0')}`;
}

// Which plan applies right now. org = the tenant's Team subscription row (or null).
export function effectivePlan(account, org, now) {
  if ((account.plan === 'essentials' || account.plan === 'professional') && LIVE.has(account.status)) return account.plan;
  if (org && LIVE.has(org.status) && account.team_member) return 'team';
  // The trial is once only: anyone who has had a subscription does not fall back to it.
  if (!account.stripe_subscription_id && now < account.trial_ends_at && account.trial_replies < PLANS.trial.replies) return 'trial';
  return 'expired';
}

// Model for the next reply, or a reason it is refused.
export function replyDecision(plan, usage, account) {
  if (plan === 'expired') return { ok: false, code: 'plan_required' };
  const limits = PLANS[plan];
  const used = plan === 'trial' ? account.trial_replies : usage.replies;
  if (used >= limits.replies) return { ok: false, code: plan === 'trial' ? 'plan_required' : 'limit_reached' };
  const premium = plan === 'trial' || usage.premium < limits.premium;
  return { ok: true, model: premium ? MODELS.premium : MODELS.standard, premium };
}

export function suggestDecision(plan, usage) {
  if (plan === 'expired') return { ok: false, code: 'plan_required' };
  if (usage.suggests >= PLANS[plan].suggests) return { ok: false, code: 'limit_reached' };
  return { ok: true, model: MODELS.insights };
}

// What the add-in shows: plan badge, trial countdown, usage meters.
export function summary(account, org, usage, now) {
  const plan = effectivePlan(account, org, now);
  const limits = PLANS[plan] || null;
  const out = {
    email: account.email, name: account.name, plan, label: limits ? limits.label : 'Trial ended',
    interval: plan === 'team' ? (org && org.interval) || '' : account.interval || '',
    status: plan === 'team' ? (org && org.status) || '' : account.status || '',
    can_manage_billing: !!account.stripe_customer_id || !!(org && org.owner_id === account.id),
    usage: {
      replies: plan === 'trial' ? account.trial_replies : usage.replies,
      replies_limit: limits ? limits.replies : 0,
      premium: plan === 'trial' ? account.trial_replies : usage.premium,
      premium_limit: limits ? limits.premium : 0,
    },
  };
  if (plan === 'trial' || plan === 'expired') {
    out.trial = { ends_at: account.trial_ends_at, replies_left: Math.max(0, PLANS.trial.replies - account.trial_replies) };
  }
  if (org) out.team = { seats: org.seats, owner: org.owner_id === account.id, member: !!account.team_member };
  return out;
}
