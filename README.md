# 💸 Finance Tracker — Daily Transaction Manager

A personal finance tracking web app built with **Flask + SQLite** that helps you manage expenses across multiple credit cards, track cashback earnings, and sync everything to Excel for monthly reporting.

![Python](https://img.shields.io/badge/python-3.11%2B-blue)
![Flask](https://img.shields.io/badge/flask-3.0-green)
![License](https://img.shields.io/badge/license-MIT-lightgrey)

---

## ✨ Features

### 📊 Dashboard
- **Summary cards** — Total spent, balance remaining, cashback earned
- **Today's transactions** — Quick overview of current day's activity
- **Recent transactions** — Full paginated history with edit/delete actions
- **Cashback column** — See cashback at a glance on every transaction

### 💳 Multi-Card Support
Manage expenses across any number of credit/debit cards with manual cashback entry:

| Card | Nickname |
|------|----------|
| SBI Cashback | `sbi_cb` |
| SBI PhonePe | `sbi_pp` |
| HDFC Millennia | `hdfc_mil` |
| HDFC Swiggy | `hdfc_swig` |
| BOB Eterna | `bob_eterna` |

Easily configurable via `cards.json`.

### 💰 Cashback Tracking
- **Manual entry** — You decide the cashback amount per transaction, no auto-override
- **Per-card dashboard** — `/cashback` page breaks down earnings by card
- **Effective rate** — Automatic rate % calculation based on amount vs cashback
- **Edit anytime** — Cashback values are fully editable

### 📁 Excel Two-Way Sync
- **DB → Excel** — Export all transactions to monthly calendar sheets
- **Excel → DB** — Import transactions from Excel (card assignments preserved)
- **Smart sync** — Auto-detects which side has more data and syncs in the right direction
- **Sync button** — One-click sync from the dashboard
- **Excel-safe** — Warns if Excel is open (file locked)

### 🏷️ Auto-Categorization
- Descriptions like "Netflix", "Swiggy", "Amazon" auto-assign to categories
- Machine-learning-free rule-based matching (extensible in `app.py`)

### 👥 Multi-Person Support
- Tag transactions to family members or friends
- Filter by person for shared expense tracking
- Balance tracking per person

### 📈 Reports & Export
- **Daily/Monthly/Card** views with chart-ready grouped data
- **CSV Export** — One-click download of all transactions
- **Pagination** — Browse transaction history page by page

---

## 🚀 Quick Start

### Prerequisites
- Python 3.11+
- pip

### Installation

```bash
# Clone the repo
git clone https://github.com/tonystark6/finance-tracker.git
cd finance-tracker

# Install dependencies
pip install -r requirements.txt

# Run the app
python app.py
```

Open **http://127.0.0.1:5000** in your browser.

### First-Run Setup

1. The app auto-creates `finance.db` (SQLite) and `expense_tracker.xlsx` (Excel template)
2. If you have an existing Excel file, place it in the project folder — it will be detected
3. Configure your cards in `cards.json` (auto-created if missing)
4. Click **Sync Excel** on the dashboard to import existing data

---

## 🗂️ Project Structure

```
finance-tracker/
├── app.py              # Flask application — all routes & logic
├── cashback.py         # Cashback reference rules (labels, rates)
├── excel_sync.py       # Excel ↔ DB sync engine
├── cards.json          # Card definitions (names & nicknames)
├── categories.json     # Auto-categorization rules
├── requirements.txt    # Python dependencies
├── .gitignore          # Git ignore rules
└── README.md           # This file

# Runtime data files (untracked):
├── finance.db          # SQLite database
└── expense_tracker.xlsx  # Excel workbook with monthly sheets
```

---

## 🧭 Navigation

| Page | Route | Description |
|------|-------|-------------|
| Dashboard | `/` | Summary stats, today's & recent transactions |
| Add Transaction | `/add` | New expense entry with cashback |
| Edit Transaction | `/edit/<id>` | Modify existing transaction |
| Cashback Dashboard | `/cashback` | Per-card cashback breakdown |
| Reports | `/reports/<view>` | Daily, monthly, card-wise views |
| Settings | `/settings` | Manage cards & categories |
| People | `/people` | Multi-person tracking |
| Balances | `/balances` | Person-wise balances |

---

## 🔧 Configuration

### Cards (`cards.json`)
```json
{
  "cards": [
    {"name": "SBI Cashback", "nickname": "sbi_cb"},
    {"name": "HDFC Millennia", "nickname": "hdfc_mil"}
  ]
}
```

### Cashback Reference (`cashback.py`)
The `cashback.py` module contains reference labels and effective rates for display purposes. Actual cashback values are user-entered per transaction — the app never auto-calculates or overrides your input.

---

## 📄 Excel Sheet Format

The app creates month-named sheets (April, May, June...) with:
- **Row 1**: Month header
- **Row 3**: Income row (green)
- **Row 5**: Day numbers (1–31)
- **Row 6+**: Category rows with daily amounts in day columns
- **Column B**: Category names
- **Last columns**: Totals, cashback tracking

> ⚠️ Keep Excel closed when using the Sync button — open files are locked and can't be written.

---

## 🐛 Troubleshooting

| Problem | Solution |
|---------|----------|
| **"Excel could not be updated"** | Close Excel and click Sync again |
| **Cashback not showing** | Ensure you entered it on the Add/Edit form — it's manual |
| **404 on routes** | You may be running an older version; refresh the page |
| **Port 5000 in use** | Change the port in app.py or kill the other process |

---

## 📦 Backup

Run from the project directory to create a timestamped backup:
```bash
python -c "import shutil; shutil.make_archive('backup', 'zip', '.')"
```

---

## 🛠️ Tech Stack

- **Backend**: Flask (Python 3)
- **Database**: SQLite (via `sqlite3` stdlib)
- **Templating**: Jinja2 (built-in Flask)
- **Excel**: openpyxl
- **Frontend**: Bootstrap 5 (dark theme), Bootstrap Icons
- **Auto-categorization**: Rule-based keyword matching

---

## 📝 License

MIT — free to use, modify, and share.

---

*Built with ❤️ by Pratik Bhatia*
