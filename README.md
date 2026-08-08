# 💸 Finance Tracker — Personal Expense Manager

A personal finance tracking web app built with **Flask + SQLite** — track expenses across multiple credit cards and bank accounts, import bank statements, manage cashback, and split costs by person.

![Python](https://img.shields.io/badge/python-3.11%2B-blue)
![Flask](https://img.shields.io/badge/flask-3.0-green)
![SQLite](https://img.shields.io/badge/database-SQLite-lightgrey)

---

## ✨ Features

### 🔐 Multi-User
- **Registration & Login** — Each user has their own account
- **Per-user data isolation** — Transactions, cards, people and monthly balances are all scoped to the owning account
- **Display name** — Customizable name shown on dashboard and People pages
- **Password recovery** — Set a security word in Settings → use `/forgot-password` to reset without email

### 📊 Dashboard
- **Summary cards** — Today's Spend, Today's Credits, Monthly Total, Cashback
- **Today's transactions** — Both debits (`-₹`) and credits (`+₹`) shown
- **Recent transactions** — Full paginated history with edit/delete actions
- **Cashback column** — See cashback at a glance on every row

### 📥 Statement Import (CSV + PDF)
Import a bank or credit-card statement from **Import Statement** (`/upload`) instead of typing rows by hand.

- **PDF statements** — Password-protected files supported; enter the PDF password in the form
- **Layout-aware parsing** — Rebuilds table rows from text coordinates, so multi-line descriptions rejoin correctly and debit/credit columns stay distinct
- **Balance reconciliation** — Where the statement prints a running balance, it is used to confirm each row's direction
- **CSV statements** — Auto-detects `Date` / `Description` / `Amount`, or split `Debit` / `Credit` columns in either order
- **Auto-categorize + card guess** — Each row is categorized, and rows that are clearly credit-card payments are matched to a card

Tested against real HDFC savings statements (single-page and 41-page/6-month), reconciling exactly to the statement's own summary totals.

> **Note:** HDFC bank statements are verified against real files. HDFC / SBI / ICICI **credit-card** layouts are covered by the parser and tested against synthetic fixtures matching their published formats — if a real card statement misparses, use the review screen and send the parser messages.

### 📝 Review Before Saving (`/import/review`)
With **Review before saving** ticked (the default), the parsed rows are held in a staging
area and **nothing is written to your ledger until you confirm**. On the review screen you can:

- **Edit any row** — date, description, amount, debit/credit, category, person
- **Tag people** — assign a person per row, or add a brand-new person without leaving the page
- **Bulk apply** — set a person, category or type across selected rows (or all of them) in one action; the value list is a dropdown scoped to the field you picked
- **Bank / card is fixed at upload** — it comes from the Bank / Card you chose on the upload form (or auto-detect) and isn't editable per row; change it after importing if a row needs a different source
- **Deselect rows** — untick anything you don't want; only ticked rows are imported
- **Duplicate protection** — rows matching transactions you already have are flagged and **pre-unticked**, so re-importing an overlapping statement can't silently double up your history
- **Resume later** — the batch survives navigation, so adding a person or closing the tab doesn't lose the parse
- **Discard** — throw the whole batch away without saving

Large statements are paginated at 100 rows per page; edits on one page never affect the others.
Untick **Review before saving** to skip this and import straight to the ledger.

### 💳 Multi-Card & Bank Support
Track expenses across credit cards, debit cards, and bank accounts. All payment sources are managed as "Bank / Card" — a unified list.

- **Unified label** — Transaction source shows "Bank / Card" instead of just "Card"
- **Simplified setup** — Adding a card/bank needs just a name (e.g. "SBI Cashback"). ID is auto-generated, bank field is merged into the name, type is optional
- **Per-user management** — Configure your own cards/banks via Settings page
- **Duplicate protection** — Duplicate names are blocked

| Default Card | Nickname |
|-------------|----------|
| SBI Cashback | `sbi_cb` |
| SBI PhonePe | `sbi_pp` |
| HDFC Millennia | `hdfc_mil` |
| HDFC Swiggy | `hdfc_swig` |
| BOB Eterna | `bob_eterna` |

Add, edit, or remove cards anytime from the Settings page.

### 💰 Cashback Tracking
- **Manual entry** — You decide the cashback per transaction
- **Per-card dashboard** — `/cashback` page breaks down earnings by card
- **Edit anytime** — Cashback values are fully editable

### 👥 Person / Split Tracking
- **Tag beneficiaries** — Mark each transaction as self, or for a specific person (e.g., Sister, Dad)
- **Self-transactions** — Leave person blank for your own expenses
- **People page** — Shows balances per person (auto-excludes self)
- **Add/remove people** — Manage from Settings page; a person can only be removed once settled up

### 🏷️ Auto-Categorization
- Descriptions like "Netflix", "Swiggy", "Amazon" auto-assign to categories
- Rule-based keyword matching using `categories.json`
- Categories fully customizable from Settings

### 📈 Reports & Export
- **Daily / Monthly / Card** views with grouped data
- **Monthly balances** — Record opening/closing bank balances per month and reconcile them against recorded activity
- **CSV Export** — One-click download of transactions
- **Pagination** — Browse transaction history page by page

### 📦 Backup & Restore
- **CLI tool** — `backup_restore.py` creates and restores timestamped zip archives

---

## 🚀 Quick Start

### Prerequisites
- Python 3.11+
- pip

### Installation

```bash
# Clone the repo
git clone https://github.com/prabhatia1/finance-tracker.git
cd finance-tracker

# Install dependencies
pip install -r requirements.txt

# Run the app
python app.py
```

Open **http://127.0.0.1:5000** in your browser.

### First-Run Setup

1. Register a new account on the `/register` page
   *(username: 3–20 characters, letters/numbers/underscore only)*
2. Log in with your credentials
3. Configure your cards, categories, and people from the Settings page
4. Start adding transactions — or import a statement from **Import Statement**

---

## 🗂️ Project Structure

```
finance-tracker/
├── app.py                 # Flask application — all routes & logic
├── pdf_parser.py          # PDF statement parser (layout/coordinate based)
├── cashback.py            # Cashback reference rules (labels, rates)
├── backup_restore.py      # Database backup and restore CLI
├── excel_sync.py          # Optional Excel import/export helpers
├── cards.json             # Card definitions (per-user managed in DB)
├── categories.json        # Auto-categorization keyword rules
├── people.json            # Person/beneficiary definitions
├── templates/             # Jinja2 templates (Bootstrap 5, dark theme)
├── requirements.txt       # Python dependencies
├── .gitignore             # Git ignore rules
└── README.md              # This file

# Runtime data (auto-created, git-ignored):
├── finance.db             # SQLite database
└── backups/               # Timestamped backup archives
```

---

## 🧭 Navigation

| Page | Route | Description |
|------|-------|-------------|
| Dashboard | `/` | Summary stats, today's & recent transactions |
| Add Transaction | `/add` | New expense entry with cashback & person tag |
| Import Statement | `/upload` | Import a CSV or PDF bank/card statement |
| Review Import | `/import/review` | Edit, tag and confirm parsed rows before saving |
| Edit Transaction | `/edit/<id>` | Modify existing transaction |
| All Transactions | `/transactions` | Full paginated transaction history |
| Reports | `/reports` | Daily, monthly, card-wise views |
| Cashback | `/cashback` | Per-card cashback breakdown |
| People | `/people` | Person-wise balances |
| Balances | `/balances` | Monthly opening/closing balance reconciliation |
| Settings | `/settings` | Manage cards, categories, people, password |

---

## 🔧 Configuration

### Categories (`categories.json`)
```json
{
  "categories": [
    {"name": "Electricity Bill", "keywords": ["electricity", "msedcl", "torrent"], "type": "bill"},
    {"name": "Groceries", "keywords": ["grocery", "blinkit", "zepto", "dmart"], "type": "daily"}
  ]
}
```

### Cards (`cards.json`)
```json
{
  "cards": [
    {"name": "SBI Cashback", "nickname": "sbi_cb"},
    {"name": "HDFC Millennia", "nickname": "hdfc_mil"}
  ]
}
```

> Cards and categories are also editable from the Settings page — no need to hand-edit JSON.
> Note that `categories.json` is a single shared file, so category edits apply to every account on the instance.

---

## 📄 Backup & Restore

Create a timestamped backup (writes `backups/finance_backup_<timestamp>.zip`):
```bash
python backup_restore.py backup
```

List available backups:
```bash
python backup_restore.py list
```

Restore from a backup archive:
```bash
python backup_restore.py restore finance_backup_20260709_120000.zip
```

---

## 🔒 Running It Safely

The defaults in `app.py` are tuned for local development:

```python
app.run(debug=True, host="0.0.0.0", port=5000)
```

- `debug=True` enables Werkzeug's interactive debugger, which can **execute code** on any traceback.
- `host="0.0.0.0"` binds every network interface, so anyone on the same Wi-Fi can reach it.

Fine on a trusted home network. Before exposing it anywhere else, set `debug=False`,
bind `host="127.0.0.1"`, and set `SESSION_COOKIE_SECURE = True` when serving over HTTPS.

Other things worth knowing:
- `/backup-download` ships the **entire database** (all users, password hashes, security words) and is gated on the username `pratik`.
- The `security_word` used by `/forgot-password` is stored in plaintext.
- Database backups (`finance.db.bak*`) are git-ignored — keep them out of commits, they contain credentials.

---

## 🛠️ Tech Stack

- **Backend**: Flask (Python 3)
- **Database**: SQLite (via `sqlite3` stdlib, WAL mode)
- **PDF parsing**: `pdfplumber`
- **Templating**: Jinja2 (built-in Flask)
- **Frontend**: Bootstrap 5 (dark theme), Bootstrap Icons
- **Auto-categorization**: Rule-based keyword matching

---

## 📝 License

MIT — free to use, modify, and share.

---

*Built with ❤️ by Pratik*

---

## 📋 Changelog

### 2026-08-08 — Review imports before they hit the ledger

Uploading with **Review before saving** no longer writes anything immediately. Parsed rows go to a
staging area and land on a new review screen (`/import/review`), where you can correct them and then
commit — the flow previously forced a second upload with Preview unticked, and there was no way to
tag people or fix categories without editing every transaction after the fact.

- **Edit any row before saving** — date, description, amount, debit/credit, category, person
- **Add a person mid-review** — previously this meant navigating to Settings, which would have thrown the parse away; rows are held server-side so nothing is lost
- **Bulk apply** a person, category or type across selected rows (or all of them). Both the field and its value are dropdowns, and the value list only shows options valid for the chosen field
- **Bank / card is not editable here** — it is set once on the upload form (or by auto-detect) and carried through unchanged, keeping the review table focused on what actually varies per row
- **Deselect rows** you don't want — only ticked rows are imported
- **Duplicate protection** — rows matching existing transactions are flagged and pre-unticked, so re-importing an overlapping statement can't silently double your history
- **Resumable** — an unfinished batch is offered again on the upload page
- **Discard** drops the batch without saving
- Large statements paginate at 100 rows/page; edits on one page never touch another
- Unticking **Review before saving** still imports straight to the ledger as before

Two safeguards in the new editor: an unparseable amount or date keeps the value the statement
actually contained rather than falling back to `0.00`/today, and rows with a zero amount are
skipped at commit (with a count) instead of entering the ledger.

### 2026-08-08 — Bug fixes across dashboard, import, reports and multi-user isolation

**Data correctness**
- **Dashboard "Today's Spend" was wrong** — the summary card summed only the 10 rows on the current page while its "N txns" subtitle counted the whole day. With 11+ transactions the headline figure understated the real total and changed as you paginated. Now aggregated in SQL over the full day.
- **CSV import booked income as expenses** — `amount` column detection also matched `Debit`/`Credit`, so for a `Date,Description,Credit,Debit` header every credit row imported as a debit. A ₹50,000 salary landed as a ₹50,000 expense with a success message. Debit/credit columns are now detected separately and take priority.
- **CSV rows with an unreadable date** silently defaulted to today, scattering them into the wrong month. They are now skipped with a visible warning.
- **Invalid transaction dates** were stored verbatim; `strftime()` returns NULL for them, so the row existed in the ledger but vanished from every report and monthly total. Dates are now normalized on save.

**Multi-user isolation**
- **Monthly balances leaked between accounts** — `monthly_balances` had no `user_id` and a globally-unique `month`, so every user read the same rows and editing one **overwrote another user's balances**. The table is now keyed per user; existing rows migrate to the original owner automatically.

**Accounts**
- **Registration could permanently lock you out** — signup never validated the username while login did, so registering `ab`, `my name` or `a&b` created an account that could never be signed into (password reset didn't help). Both paths now apply the same rule, and the username is no longer HTML-escaped before storage.
- Registration now relies on the UNIQUE constraint instead of check-then-insert, which could race into a 500.

**Crashes (all returned HTTP 500)**
- `?page_today=` / `?page_recent=` with an empty value — and the app generated these URLs itself, because the AJAX partials rendered without the other section's variables.
- `/reports?month=13`, `?month=0`, `?year=0` → `IllegalMonthError`.
- `amount=nan` passed validation (NaN compares False against every bound) and then failed the NOT NULL constraint at INSERT.
- `cashback=1,200` — a comma in a currency field hit a bare `float()`.
- `/settings` posts missing a field returned HTTP 400 via `request.form[...]`.

**Reliability**
- **Connection leaks** — routes returning early (notably every `POST /balances`) never closed their SQLite connection. Under WAL these hold read locks and surfaced as `database is locked` on unrelated requests. Added a `db_conn()` context manager; no leaking paths remain.
- `get_db()` re-ran the full 8-statement schema script on **every** connection (twice per dashboard load). Moved to startup, and added a busy timeout.

**Smaller fixes**
- A settled person couldn't be deleted: float residue (`5.5e-17`) failed a `!= 0` check while the message rounded to "₹0" — an unsatisfiable instruction. Now compared with a tolerance.
- A genuine `0.00` closing balance was treated as "not set", hiding real reconciliation discrepancies.
- Blank card names were accepted, creating an unnamed entry in every dropdown.
- CSV parse errors returned a string instead of a list, so `'; '.join(...)` split the message into characters (`No transactions found! C; o; u; l; d`).
- `.gitignore` didn't match `finance.db.bak_*`, so database backups containing password hashes were not ignored.

### 2026-08-07 — Statement import (PDF + CSV)
- **PDF statement import** — new `pdf_parser.py` reconstructs table rows from text coordinates rather than regex-matching flat text, which is required for HDFC statements where descriptions wrap across multiple lines and columns are not ruled
- **Password-protected PDFs** supported
- **Credit-card layouts** — handles the single amount column with trailing `D`/`C` or `Cr` markers used by HDFC/SBI/ICICI cards, and multi-token dates like `05 Aug 26`
- **Balance-based verification** — running balance confirms each row's debit/credit direction
- **Re-enabled `/upload`** with preview mode; fixed an undefined `debug_info` on the CSV preview path and a duplicated auto-categorize loop
- **Card matching corrected** — UPI handles (`@oksbi`) and IFSC codes (`HDFC0009155`) name the counterparty, not the card used; 124 of 434 rows were being mis-assigned, now only genuine card payments are
- Startup banner no longer crashes on non-UTF-8 Windows consoles

### 2026-07-09 — Card form simplified, DB column dropped, "Bank / Card" labels
- **Removed bank column from DB** — `bank` column dropped from `user_cards` table via migration. Existing cards lose only the bank name (card name, type, transactions all preserved)
- **Forgot password** — `/forgot-password` page with security-word verification. Set your word in Settings.
- **Simplified add card form** — Removed separate ID and Bank fields. Just enter a name (e.g. "SBI Cashback") — ID is auto-generated, type is optional
- **"Bank / Card" labels** — Transaction source and table headers now show "Bank / Card" instead of just "Card"
- **Duplicate name protection** — Adding a card/bank with an existing name is blocked gracefully
- **Removed PRAGMA foreign_keys=ON** — Fixes "FOREIGN KEY constraint failed" error on live site
- **Safeguard missing tables** — `get_user_cards`, `get_user_people` handle missing tables gracefully
- **Dashboard fixes** — Today's Credits shows txn count, This Month shows net (credits − debits) with +/- sign, no minus sign on Today's Spend
