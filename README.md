# Budgie: monthly budget app

An installable web app (PWA) for phone and laptop. It shows progress bars per category, auto-sorts bank SMSes, imports budgets from Excel and gives you useful insights.

```
docs/     the app itself (served by GitHub Pages)
worker/   tiny Cloudflare Worker: encrypted sync + bank-SMS inbox
```

## What it does

- **Home**: "left to spend", a safe daily amount, and a pace marker. Fixed bills such as rent and insurance count as expected once paid, so they don't make you look like you're overspending. Each category has a progress bar with a status: On track / Spending fast / Over (or Paid / Not paid yet for bills).
- **Bank SMS → transactions**: tuned for FNB, Capitec, Standard Bank, Absa, Nedbank, Discovery and TymeBank. OTPs and declined payments are ignored. Bank "reserved" holds are de-duplicated.
- **Auto-categorise**: it uses (1) merchants you've sorted before, (2) your own keywords per category, and (3) built-in knowledge of SA merchants (Checkers, Engen, Vodacom, Dis-Chem…). When it isn't sure, the transaction goes into a **"to sort" queue** with best guesses. "Remember" teaches it, and other waiting transactions from the same merchant get sorted too.
- **Excel / CSV import**: drag and drop (laptop) or choose a file (phone). It finds the name and amount columns, treats heading rows like "FIXED COSTS" as groups, skips "Total" rows, and picks out income rows. You see a preview before anything changes.
- **Insights**: where the month is heading, a comparison with the same day last month, savings rate, no-spend days, a cumulative spending chart, a 6-month chart, where the money went, the biggest changes, top places and the biggest purchases.
- **Payday months** (e.g. 25th → 24th), **Easy view** for parents (bigger text, fewer screens), light/dark mode, backup/restore, CSV export.

## Deploy (one-time, about 15 minutes, both free)

### 1. Sync Worker (Cloudflare)
```bash
cd worker
npm install
npx wrangler login
npx wrangler kv namespace create DATA     # paste the id into wrangler.toml
npx wrangler deploy                       # prints https://budgie-sync.<you>.workers.dev
```
Put that URL in `DEFAULT_SYNC_URL` at the top of `docs/app.js`. You can also paste it into Settings in the app.

### 2. App (GitHub Pages)
Push this folder to a GitHub repo, then go to Settings → Pages → Source: `main` branch, `/docs` folder.
Open `https://<you>.github.io/<repo>/` on your phone → Share → **Add to Home Screen**.

### 3. Connect devices
- In the app: Settings → **Create new sync space**. Copy the **invite link** and open it on your laptop.
- Your parents: create a *separate* sync space on Mom's phone and send the invite link to Dad's, so their budget stays theirs.
- Settings → **Set up bank SMS** has step-by-step guides for iPhone (Shortcuts automation) and Android (MacroDroid).

## Privacy
The budget is encrypted on your device (AES-256-GCM, with a key derived from your sync code) before it's uploaded, so the Worker only ever stores ciphertext. Forwarded SMS texts wait in the Worker inbox until a device pulls them, then they're deleted (or expire after 45 days). Anyone with your sync code can read your data, so share it only with people you trust.
