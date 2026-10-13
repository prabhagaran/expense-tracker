# Expense Tracker

A small personal expense tracker. The page is hosted on GitHub Pages, and the data lives in **your own private Google Sheet**.

- **No financial data is in this repo.** The page is only code. Transactions are stored in a Google Sheet in your Drive, reached through a Google Apps Script web app that answers only when a request carries your secret key.
- Works on a phone. You can install it to your home screen ("Add to Home screen").
- **Overview:** income, spending, net cash flow, savings rate and closing balance, each compared with the previous period; a 12-month cash-flow chart; spending by category against budget; top payees; recent transactions.
- **Transactions:** a sortable, paged table with search, category and type filters, multi-select to recategorise or delete in bulk, and CSV export of the filtered view.
- **Budgets:** monthly limits per category, with On track / At risk / Over status.
- **Reporting periods:** a single month, last 3/6/12 months, year to date, or all time.
- **Keyboard:** `n` new transaction, `/` search, `g o` / `g t` / `g b` to switch pages, `Esc` to close the editor.

## Adding new months

Once a month, download your Canara e-passbook as a PDF (ai1 app or net banking → e-Passbook / Statement → any date range → PDF) and drop it on **Import statement**. The PDF is read in your browser. You'll see how many transactions are new, how many are already in your sheet, and which payees still need a category; then you confirm.

- Overlapping dates are fine. Rows already in the sheet are matched on date, type and amount and skipped.
- Payees you've categorised before get the same category automatically.
- Cash spends: add them with **New transaction**.

The PDF reader is [pdf.js](https://mozilla.github.io/pdf.js/) (Apache-2.0), bundled in `vendor/`.

## One-time setup (about 5 minutes)

### 1. Create the Google Sheet + backend
1. Go to <https://sheets.new> and create a sheet. Name it something like **Expense Tracker DB**. Keep it private; don't share it.
2. In the sheet, open **Extensions → Apps Script**.
3. Delete what's in `Code.gs`, then paste in the contents of [`apps-script/Code.gs`](apps-script/Code.gs). Save.
4. Set your secret key: **Project Settings (⚙) → Script properties → Add script property**
   - Property: `API_KEY`
   - Value: a long random string (e.g. 32+ characters from a password manager). Treat it like a password.
5. Back in the editor, pick the `setup` function and click **Run**. Approve the permissions; this is your own script accessing your own sheet. This creates the `Transactions` and `Budgets` tabs.
6. **Deploy → New deployment → ⚙ → Web app**
   - Execute as: **Me**
   - Who has access: **Anyone**. This is needed so the page can reach it, and the secret key keeps everyone else out.
   - Click **Deploy** and copy the **Web app URL** (ends in `/exec`).

### 2. Connect the page
Open the site, go to **Settings**, paste the Web app URL and your secret key, then click **Save & connect**. The key is stored only in that browser. Do this once on each device you use.

### 3. (Optional) Import your history
Settings → Import & export → choose a CSV with columns:

```
id,date,type,category,counterparty,amount,note,source
```

`type` is `in` or `out`, `date` is `YYYY-MM-DD`, and `amount` is positive. Rows whose `id` already exists are skipped.

## Updating the backend
If `Code.gs` changes: paste the new version, then **Deploy → Manage deployments → ✎ → Version: New version → Deploy**. The URL stays the same.

## Security notes
- Anyone can load the page, but without the URL **and** the key they get `Unauthorized` and no data.
- If you think the key has leaked, change `API_KEY` in Script properties and re-enter it on your devices.
- `.gitignore` blocks `*.csv`, `*.xlsx` and `*.pdf` so statements can't be committed by accident.
