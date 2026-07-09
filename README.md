# 💸 Finance Tracker — Personal Expense Manager

A personal finance tracking web app built with **Flask + SQLite** — track expenses across multiple credit cards, manage cashback, and split costs by person.

![Python](https://img.shields.io/badge/python-3.11%2B-blue)
![Flask](https://img.shields.io/badge/flask-3.0-green)
![SQLite](https://img.shields.io/badge/database-SQLite-lightgrey)

---

## ✨ Features

### 🔐 Multi-User
- **Registration & Login** — Each user has their own account
- **Per-user data isolation** — Every transaction belongs to a user; no data leaks between accounts
- **Display name** — Customizable name shown on dashboard and People pages

### 📊 Dashboard
- **Summary cards** — Today's Spend, Today's Credits, Monthly Total, Cashback
- **Today's transactions** — Both debits (`-₹`) and credits (`+₹`) shown
- **Recent transactions** — Full paginated history with edit/delete actions
- **Cashback column** — See cashback at a glance on every row

### 💳 Multi-Card Support
Manage expenses across any number of credit/debit cards. Cards configured per-user via the Settings page.

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
- **Add/remove people** — Manage from Settings page

### 🏷️ Auto-Categorization
- Descriptions like "Netflix", "Swiggy", "Amazon" auto-assign to categories
- Rule-based keyword matching using `categories.json`
- Categories fully customizable from Settings

### 📈 Reports & Export
- **Daily / Monthly / Card** views with grouped data
- **CSV Export** — One-click download of all transactions (pratik only)
- **Pagination** — Browse transaction history page by page

### 📦 Backup & Restore
- **CLI tool** — `backup_restore.py` for creating and restoring timestamps

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
2. Log in with your credentials
3. Configure your cards, categories, and people from the Settings page
4. Start adding transactions!

---

## 🗂️ Project Structure

```
finance-tracker/
├── app.py                 # Flask application — all routes & logic
├── cashback.py            # Cashback reference rules (labels, rates)
├── backup_restore.py      # Database backup and restore CLI
├── cards.json             # Card definitions (per-user managed in DB)
├── categories.json        # Auto-categorization keyword rules
├── people.json            # Person/beneficiary definitions
├── requirements.txt       # Python dependencies
├── .gitignore             # Git ignore rules
└── README.md              # This file

# Runtime data (auto-created, git-ignored):
└── finance.db             # SQLite database
```

---

## 🧭 Navigation

| Page | Route | Description |
|------|-------|-------------|
| Dashboard | `/` | Summary stats, today's & recent transactions |
| Add Transaction | `/add` | New expense entry with cashback & person tag |
| Edit Transaction | `/edit/<id>` | Modify existing transaction |
| All Transactions | `/transactions` | Full paginated transaction history |
| Reports | `/reports` | Daily, monthly, card-wise views |
| Cashback | `/cashback` | Per-card cashback breakdown |
| People | `/people` | Person-wise balances |
| Balances | `/balances` | Monthly balance report |
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

---

## 📄 Backup & Restore

Create a timestamped backup:
```bash
python backup_restore.py backup
```

Restore from a backup file:
```bash
python backup_restore.py restore backup_20260709_120000.db
```

---

## 🛠️ Tech Stack

- **Backend**: Flask (Python 3)
- **Database**: SQLite (via `sqlite3` stdlib)
- **Templating**: Jinja2 (built-in Flask)
- **Frontend**: Bootstrap 5 (dark theme), Bootstrap Icons
- **Auto-categorization**: Rule-based keyword matching

---

## 📝 License

MIT — free to use, modify, and share.

---

*Built with ❤️ by Pratik*
