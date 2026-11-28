# Bulwark HQ Assistant – go-live guide

One Cloudflare Worker at **https://assistant.bulwarkhq.com.au** serves the Outlook add-in, the
pricing/legal pages and the subscription API. Claude, Stripe and Microsoft sign-in keys stay on the
server; no customer needs their own Claude key.

```
Outlook pane ──Microsoft sign-in token──▶ Worker /api/*  ──▶ Claude API (Haiku / Sonnet / Opus)
                                             │  D1 database: accounts, plans, monthly counts (no email content)
Stripe Checkout / portal ──webhook──────────▶┘
```

Do the steps in order. Use **Stripe test mode** first, then repeat steps 3 and 5 with live keys.

## 1. Cloudflare (about 15 minutes)

1. Create a Cloudflare account and add the **bulwarkhq.com.au** zone (or ask whoever runs that domain's DNS to move it to Cloudflare, or delegate an `assistant` subdomain).
2. On your PC, in this repository:

```bash
cd server
```

```bash
npm ci
```

```bash
npx wrangler login
```

```bash
npx wrangler d1 create assistant
```

3. Copy the `database_id` it prints into `server/wrangler.jsonc` (replace `REPLACE_WITH_D1_DATABASE_ID`), then create the tables:

```bash
npx wrangler d1 migrations apply assistant --remote
```

## 2. Microsoft sign-in (Entra app registration, about 15 minutes)

In the **Microsoft Entra admin centre** (entra.microsoft.com) → **App registrations** → **New registration**:

| Setting | Value |
|---|---|
| Name | Bulwark HQ Assistant |
| Supported account types | **Accounts in any organizational directory and personal Microsoft accounts** |
| Redirect URI | Single-page application: `https://assistant.bulwarkhq.com.au/taskpane.html` |

Then:

1. Copy the **Application (client) ID** into `server/wrangler.jsonc` → `MS_CLIENT_ID`.
2. **Expose an API** → Application ID URI: `api://assistant.bulwarkhq.com.au/<client ID>`.
3. **Add a scope**: name `access_as_user`, who can consent **Admins and users**, with a short display name and description for each.
4. **Authorized client applications** → add each of these with the `access_as_user` scope (these are Microsoft's Office apps):
   - `ea5a67f6-b6f3-4338-b240-c655ddc3cc8e` (all Microsoft Office endpoints)
   - `d3590ed6-52b3-4102-aeff-aad2292ab01c` (Office desktop)
   - `93d53678-613d-4013-afc1-62e9e444a0a5` (Office on the web)
   - `bc59ab01-8403-45c6-8796-ac3ef710b3e3` (Outlook on the web)
5. **API permissions**: Microsoft Graph delegated `openid`, `profile`, `email` (already there or add them).
6. **Manifest** (in the app registration): set `"accessTokenAcceptedVersion": 2` and save.
7. **Branding & properties**: add the logo, home page (`https://assistant.bulwarkhq.com.au/pricing.html`), terms and privacy URLs. To avoid "unverified" warnings for other organisations, complete **Publisher verification** (needs a Microsoft Partner Network ID; see step 6).

## 3. Stripe (about 15 minutes)

1. With a **test** secret key, create the products and prices (Essentials, Professional, Team; monthly and annual; AUD incl. GST):

```bash
STRIPE_SECRET_KEY=sk_test_xxx npm run stripe:setup
```

2. Stripe Dashboard → **Settings → Tax**: turn on Stripe Tax and add your GST registration (prices are set as GST-inclusive). If you aren't GST-registered yet, leave tax off.
3. **Settings → Billing → Customer portal**: allow cancelling (at end of period), updating payment methods, switching between the three products (monthly and annual) and changing quantity (for Team).
4. **Developers → Webhooks → Add endpoint**: `https://assistant.bulwarkhq.com.au/api/stripe/webhook` with events `checkout.session.completed`, `customer.subscription.created`, `customer.subscription.updated`, `customer.subscription.deleted`, `customer.subscription.paused` and `customer.subscription.resumed`. Copy its **signing secret**.

## 4. Secrets and first deploy

Each command asks you to paste the value (it is stored encrypted in Cloudflare, never in the repository):

```bash
npx wrangler secret put ANTHROPIC_API_KEY
```

```bash
npx wrangler secret put STRIPE_SECRET_KEY
```

```bash
npx wrangler secret put STRIPE_WEBHOOK_SECRET
```

```bash
npm run deploy
```

The custom domain `assistant.bulwarkhq.com.au` is created on first deploy. Check it: https://assistant.bulwarkhq.com.au/api/health should show `{"ok":true}`, and `/pricing.html` should load.

Use a separate **Claude API key** for this service (console.anthropic.com → API keys), with a monthly spend limit set in the Anthropic console.

## 5. Test the full journey (test mode)

1. Make the marketplace manifest (with your client ID):

```bash
node ../addin-build/make-manifest.js https://assistant.bulwarkhq.com.au/ --client-id <client ID>
```

2. Sideload `server/marketplace/manifest.xml` in Outlook (Get Add-ins → My add-ins → Add a custom add-in → Add from file).
3. Open an email → Assistant. You should be asked to sign in once, then see **Trial · 3d**.
4. Write a few replies. To test the paywall without waiting 3 days, run this, replacing `you@example.com` with your email:

```bash
npx wrangler d1 execute assistant --remote --command "UPDATE accounts SET trial_replies = 50 WHERE email = 'you@example.com'"
```

5. Write another reply → the plans sheet opens → choose a plan → pay with Stripe test card `4242 4242 4242 4242` → return to Outlook → the pill shows your plan.
6. Settings → **Manage billing** → cancel → after the period ends (or cancel immediately in the Stripe test dashboard) the pane returns to "Trial ended".
7. Switch Stripe to **live**: repeat step 3 with the live key (`sk_live_…`), set the live `STRIPE_SECRET_KEY` and `STRIPE_WEBHOOK_SECRET` secrets, and deploy again.

## 6. Microsoft marketplace (AppSource) listing

1. Join the **Microsoft AI Cloud Partner Program** (free) and create a **Partner Center** account (Marketplace offers, Office add-ins), using the business's ABN details. This also gives the MPN ID for publisher verification in step 2.7.
2. Partner Center → **Marketplace offers → New offer → Office add-in**:
   - Upload `server/marketplace/manifest.xml`.
   - Pricing: **Free**, and tick **"Additional purchase may be required"** (that's how Microsoft labels apps billed by their publisher).
   - Listing: name **Bulwark HQ Assistant**, logo `addin/icons/assistant-128.png` (Partner Center also needs a 300×300 PNG), 3 to 5 screenshots at 1366×768, a short and long description, and the links `https://assistant.bulwarkhq.com.au/privacy.html`, `/terms.html` and `/support.html`.
   - **Testing notes for the validation team**: explain that sign-in starts a free trial automatically and no payment is needed to test. Give them a Microsoft 365 test account, or say any Microsoft account works.
3. Submit. Validation usually takes 3 to 5 business days. Common rejections: broken links, the add-in failing in Outlook on the web or Mac, and purchase links shown on iOS/Android. The pane already hides all prices and upgrade buttons on iPhone, iPad and Android.

## Before going live: checklist

- [ ] **Have a lawyer review** `addin/privacy.html` and `addin/terms.html`: add the legal entity name and ABN, then remove the "Draft for legal review" line.
- [ ] Make sure the `support@bulwarkhq.com.au` and `privacy@bulwarkhq.com.au` mailboxes exist, or change the addresses in `addin/support.html`, `privacy.html` and `terms.html`.
- [ ] Set an Anthropic monthly spend limit and turn on Cloudflare usage notifications.
- [ ] Do a full journey test on Windows Outlook Classic, new Outlook, Outlook on the web and Mac.

## Running costs (estimate)

| Item | Cost |
|---|---|
| Cloudflare Workers Paid + D1 | US$5 / month (covers well beyond 10,000 users) |
| Stripe | 1.7% + A$0.30 per domestic card payment (higher for international cards) |
| Claude API | per reply; see the cost model per plan in the "Pricing & Marketplace Plan" doc. Set a monthly spend limit. |
| Microsoft Partner Center | Free |

## Day-to-day

- **Logs:** `npx wrangler tail`. Email content is never logged.
- **Look up a customer:** run this, replacing `x@y.com` with their email:

```bash
npx wrangler d1 execute assistant --remote --command "SELECT id,email,plan,status,trial_ends_at,trial_replies FROM accounts WHERE email='x@y.com'"
```

- **Extend someone's trial by a day:** run this, replacing `x@y.com` with their email:

```bash
npx wrangler d1 execute assistant --remote --command "UPDATE accounts SET trial_ends_at = trial_ends_at + 86400000 WHERE email='x@y.com'"
```

- **Change prices:** edit `PRICES` in `src/plans.js` and the pane's `PRICING` in `addin/service-client.js`, run `npm run stripe:setup` (it moves the lookup keys onto new prices; existing subscribers keep their old price until you migrate them in Stripe), then deploy.
- **Tests:** `npm test` (server) and `node test/pane.browser.mjs` (pane end-to-end). Both also run in GitHub Actions.
