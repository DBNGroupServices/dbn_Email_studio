-- Accounts are keyed by Microsoft tenant + object id. No email content is ever stored.
CREATE TABLE accounts (
  id TEXT PRIMARY KEY,                 -- "<tid>:<oid>"
  tid TEXT NOT NULL,
  email TEXT NOT NULL DEFAULT '',
  name TEXT NOT NULL DEFAULT '',
  created_at INTEGER NOT NULL,         -- unix ms
  trial_ends_at INTEGER NOT NULL,      -- unix ms, end of the 3rd working day
  trial_replies INTEGER NOT NULL DEFAULT 0,
  plan TEXT NOT NULL DEFAULT 'none',   -- none | essentials | professional (personal subscriptions)
  status TEXT NOT NULL DEFAULT '',     -- Stripe subscription status
  interval TEXT NOT NULL DEFAULT '',   -- month | year
  period_end INTEGER NOT NULL DEFAULT 0,
  stripe_customer_id TEXT NOT NULL DEFAULT '',
  stripe_subscription_id TEXT NOT NULL DEFAULT '',
  team_member INTEGER NOT NULL DEFAULT 0 -- holds a seat on the tenant's Team subscription
);
CREATE INDEX accounts_tid ON accounts(tid);
CREATE INDEX accounts_customer ON accounts(stripe_customer_id);

-- One Team subscription per Microsoft tenant (organisation).
CREATE TABLE orgs (
  tid TEXT PRIMARY KEY,
  owner_id TEXT NOT NULL,
  seats INTEGER NOT NULL DEFAULT 0,
  status TEXT NOT NULL DEFAULT '',
  interval TEXT NOT NULL DEFAULT '',
  period_end INTEGER NOT NULL DEFAULT 0,
  stripe_customer_id TEXT NOT NULL DEFAULT '',
  stripe_subscription_id TEXT NOT NULL DEFAULT ''
);

-- Monthly counters (calendar month, UTC).
CREATE TABLE usage (
  account_id TEXT NOT NULL,
  month TEXT NOT NULL,                 -- "YYYY-MM"
  replies INTEGER NOT NULL DEFAULT 0,
  premium INTEGER NOT NULL DEFAULT 0,  -- replies written on the premium model
  suggests INTEGER NOT NULL DEFAULT 0,
  PRIMARY KEY (account_id, month)
);

-- Stripe webhook events already processed (idempotency).
CREATE TABLE stripe_events (id TEXT PRIMARY KEY, received_at INTEGER NOT NULL);
