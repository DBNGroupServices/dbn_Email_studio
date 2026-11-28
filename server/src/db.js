// Data access on Cloudflare D1. Only account, plan and usage data; never email content.
import { trialEndsAt, monthKey } from './plans.js';

export async function getOrCreateAccount(db, user, now, tz) {
  const existing = await db.prepare('SELECT * FROM accounts WHERE id = ?').bind(user.id).first();
  if (existing) {
    if ((user.email && user.email !== existing.email) || (user.name && user.name !== existing.name)) {
      await db.prepare('UPDATE accounts SET email = ?, name = ? WHERE id = ?').bind(user.email || existing.email, user.name || existing.name, user.id).run();
      existing.email = user.email || existing.email; existing.name = user.name || existing.name;
    }
    return existing;
  }
  const row = { id: user.id, tid: user.tid, email: user.email, name: user.name, created_at: now, trial_ends_at: trialEndsAt(now, tz) };
  await db.prepare('INSERT OR IGNORE INTO accounts (id, tid, email, name, created_at, trial_ends_at) VALUES (?, ?, ?, ?, ?, ?)')
    .bind(row.id, row.tid, row.email, row.name, row.created_at, row.trial_ends_at).run();
  return db.prepare('SELECT * FROM accounts WHERE id = ?').bind(user.id).first();
}

export function getOrg(db, tid) {
  return db.prepare('SELECT * FROM orgs WHERE tid = ?').bind(tid).first();
}

// Gives the account a seat on its organisation's Team plan if one is free.
export async function claimSeat(db, account, org) {
  if (!org || account.team_member) return account;
  const { n } = await db.prepare('SELECT COUNT(*) AS n FROM accounts WHERE tid = ? AND team_member = 1').bind(account.tid).first();
  if (n >= org.seats) return account;
  await db.prepare('UPDATE accounts SET team_member = 1 WHERE id = ?').bind(account.id).run();
  return { ...account, team_member: 1 };
}

export async function getUsage(db, accountId, now) {
  const month = monthKey(now);
  const row = await db.prepare('SELECT * FROM usage WHERE account_id = ? AND month = ?').bind(accountId, month).first();
  return row || { account_id: accountId, month, replies: 0, premium: 0, suggests: 0 };
}

// Trial replies count towards the 50-reply trial only, never towards a later paid month.
export async function countUse(db, account, now, { reply = 0, premium = 0, suggest = 0, trial = false }) {
  const month = monthKey(now);
  const monthly = trial ? 0 : reply;
  const stmts = [db.prepare(`INSERT INTO usage (account_id, month, replies, premium, suggests) VALUES (?, ?, ?, ?, ?)
    ON CONFLICT(account_id, month) DO UPDATE SET replies = replies + excluded.replies, premium = premium + excluded.premium, suggests = suggests + excluded.suggests`)
    .bind(account.id, month, monthly, premium, suggest)];
  if (trial && reply) stmts.push(db.prepare('UPDATE accounts SET trial_replies = trial_replies + ? WHERE id = ?').bind(reply, account.id));
  await db.batch(stmts);
}

export async function setStripeCustomer(db, accountId, customerId) {
  await db.prepare('UPDATE accounts SET stripe_customer_id = ? WHERE id = ?').bind(customerId, accountId).run();
}
