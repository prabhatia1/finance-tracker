"""
💸 Finance Tracker — Daily Transaction Manager
Track expenses across multiple credit cards, auto-categorize bills,
and generate beautiful reports (daily, monthly, card-wise).
"""

import sqlite3
import csv
import io
import json
import os
import re
import calendar
from datetime import date, datetime, timedelta
from pathlib import Path
from io import BytesIO, StringIO

from flask import (
    Flask, render_template, request, redirect, url_for,
    flash, session, send_file, jsonify, Response
)
# ─── Optional imports (CSRF, Rate Limiting) ─────────────────────────────
try:
    from flask_wtf.csrf import CSRFProtect
    _HAS_FLASK_WTF = True
except ImportError:
    CSRFProtect = None
    _HAS_FLASK_WTF = False
    print("[warn] flask-wtf not installed — CSRF protection disabled")

try:
    from flask_limiter import Limiter
    from flask_limiter.util import get_remote_address
    _HAS_FLASK_LIMITER = True
except ImportError:
    Limiter = None
    get_remote_address = None
    _HAS_FLASK_LIMITER = False
    print("[warn] flask-limiter not installed — rate limiting disabled")

# ─── Setup ───────────────────────────────────────────────────────────────────
BASE_DIR = Path(__file__).parent
DB_PATH = BASE_DIR / "finance.db"
CARDS_PATH = BASE_DIR / "cards.json"
CATEGORIES_PATH = BASE_DIR / "categories.json"
PEOPLE_PATH = BASE_DIR / "people.json"
EXCEL_PATH = BASE_DIR / "expense_tracker.xlsx"
SECRET_KEY_PATH = BASE_DIR / ".secret_key"

# Import Excel sync
from excel_sync import (
    read_transactions_from_excel,
    add_transaction_to_excel,
    excel_sync_to_db,
    db_sync_to_excel,
    smart_sync,
    init_excel,
)

# Import cashback — labels for dashboard display
from cashback import CARD_CB_LABELS

app = Flask(__name__)
try:
    if SECRET_KEY_PATH.exists():
        app.secret_key = SECRET_KEY_PATH.read_text().strip()
    else:
        app.secret_key = os.urandom(32).hex()
        SECRET_KEY_PATH.write_text(app.secret_key)
except Exception:
    app.secret_key = os.urandom(32).hex()

# ─── CSRF Protection ────────────────────────────────────────────────────
if _HAS_FLASK_WTF:
    csrf = CSRFProtect(app)
else:
    csrf = None
    # Register dummy csrf_token for templates so they don't crash
    @app.context_processor
    def _noop_csrf():
        def csrf_token():
            return ""
        return dict(csrf_token=csrf_token)

# ─── Rate Limiting ─────────────────────────────────────────────────────
if _HAS_FLASK_LIMITER:
    limiter = Limiter(
        app=app,
        key_func=get_remote_address,
        default_limits=["200 per day", "50 per hour"],
        storage_uri="memory://"
    )
else:
    class _NoopLimiter:
        def limit(self, *a, **kw):
            def deco(f):
                return f
            return deco
    limiter = _NoopLimiter()

# ─── Session Security ─────────────────────────────────────────────────────
app.config['SESSION_COOKIE_HTTPONLY'] = True      # JS can't read cookie
app.config['SESSION_COOKIE_SAMESITE'] = 'Lax'      # blocks CSRF from external sites
app.config['SESSION_COOKIE_SECURE'] = False        # Set True for HTTPS production

# ─── Sanitization helper ──────────────────────────────────────────────────
import html as _html
def sanitize(text):
    """Strip HTML tags from user input to prevent XSS."""
    if not text:
        return ""
    return _html.escape(re.sub(r'<[^>]*>', '', text)).strip()


# ─── Password & Input Validation ────────────────────────────────────────────
class PasswordValidator:
    """Validate password strength."""
    MIN_LENGTH = 6

    @staticmethod
    def validate(password):
        errors = []
        if len(password) < PasswordValidator.MIN_LENGTH:
            errors.append(f"At least {PasswordValidator.MIN_LENGTH} characters")
        if not re.search(r'[A-Z]', password):
            errors.append("One uppercase letter")
        if not re.search(r'[a-z]', password):
            errors.append("One lowercase letter")
        if not re.search(r'\d', password):
            errors.append("One digit")
        if errors:
            return False, " • ".join(errors)
        return True, ""

    @staticmethod
    def validate_username(username):
        return bool(re.match(r'^[a-zA-Z0-9_]+$', username)) and 3 <= len(username) <= 20

    @staticmethod
    def validate_amount(amount):
        try:
            return 0.01 <= float(amount) <= 99_999_999
        except (ValueError, TypeError):
            return False

    @staticmethod
    def sanitize(text, max_length=500):
        if not isinstance(text, str):
            return ""
        return _html.escape(re.sub(r'<[^>]*>', '', text)).strip()[:max_length]

# ─── Seed Data for New Users ────────────────────────────────────────────────
def seed_new_user(user_id):
    """Add default cards, people, and sample transactions for a new user."""
    conn = get_db()
    now = date.today()

    # Default cards
    default_cards = [
        ("sbi_cb",  "SBI Cashback",     "Visa"),
        ("sbi_pp",  "SBI PhonePe",      "Visa"),
        ("hdfc_mil","HDFC Millennia",    "Visa"),
        ("hdfc_swig","HDFC Swiggy",      "Visa"),
        ("bob_eterna","BOB Eterna",      "RuPay"),
    ]
    for cid, nm, tp in default_cards:
        conn.execute(
            "INSERT OR IGNORE INTO user_cards (user_id, card_id, name, type) VALUES (?,?,?,?)",
            (user_id, cid, nm, tp))

    # Default people
    for nm in ["Sister", "Brother", "Dad"]:
        conn.execute(
            "INSERT OR IGNORE INTO user_people (user_id, name) VALUES (?,?)",
            (user_id, nm))

    # Sample transactions (current month)
    sample = [
        (f"{now.year}-{now.month:02d}-01", "Electricity Bill",  1250, "Electricity Bill",  "sbi_cb",  "",          "",  "debit"),
        (f"{now.year}-{now.month:02d}-03", "Big Basket",         845,  "Groceries",         "hdfc_mil","",          "",  "debit"),
        (f"{now.year}-{now.month:02d}-05", "Zomato Order",       320,  "Dining / Food",     "sbi_pp",  "",          "",  "debit"),
        (f"{now.year}-{now.month:02d}-07", "Petrol",            1500,  "Fuel",              "hdfc_swig","",         "",  "debit"),
        (f"{now.year}-{now.month:02d}-10", "Netflix",            499,  "Entertainment",     "hdfc_mil","Monthly",   "",  "debit"),
        (f"{now.year}-{now.month:02d}-12", "Amazon Shopping",   1200,  "Shopping",          "sbi_cb",  "Electronics","",  "debit"),
        (f"{now.year}-{now.month:02d}-15", "LIC Premium",       5000,  "LIC Premium",       "hdfc_mil","Yearly",    "",  "debit"),
        (f"{now.year}-{now.month:02d}-18", "Mobile Recharge",    299,  "Recharge / Mobile",  "sbi_pp",  "",          "", "debit"),
        (f"{now.year}-{now.month:02d}-20", "D Mart",            1560,  "Groceries",         "bob_eterna","",        "Sister","debit"),
        (f"{now.year}-{now.month:02d}-22", "Ola Cab",            250,  "Travel",            "sbi_cb",  "Office",     "",  "debit"),
        (f"{now.year}-{now.month:02d}-25", "Medicine",           340,  "Medical / Health",   "hdfc_mil","Pharmacy",  "Dad","debit"),
        (f"{now.year}-{now.month:02d}-28", "Swiggy",             180,  "Dining / Food",     "sbi_cb",  "",          "",  "debit"),
        (f"{now.year}-{now.month:02d}-01", "Salary",          75000,  "Other",             "other",   "Credit",     "",  "credit"),
    ]
    for dt, desc, amt, cat, card, notes, person, ttype in sample:
        conn.execute(
            "INSERT INTO transactions (date, description, amount, category, card_id, txn_type, notes, source, user_id, person) VALUES (?,?,?,?,?,?,?,?,?,?)",
            (dt, desc, abs(amt), cat, card, ttype, notes, "seed", user_id, person))

    conn.commit()
    conn.close()

# ─── Login helper ────────────────────────────────────────────────────────────
from functools import wraps

def login_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if "user_id" not in session:
            flash("Please log in first.", "warning")
            return redirect(url_for("login"))
        return f(*args, **kwargs)
    return decorated_function

# ─── Helpers ─────────────────────────────────────────────────────────────────

def get_db():
    conn = sqlite3.connect(str(DB_PATH))
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL")
    return conn


def get_user_cards(user_id):
    """Get a user's cards from DB."""
    conn = get_db()
    try:
        rows = conn.execute(
            "SELECT card_id AS id, name, type FROM user_cards WHERE user_id = ? ORDER BY name",
            (user_id,)
        ).fetchall()
    except Exception:
        rows = []
    conn.close()
    return [dict(r) for r in rows]


def get_user_people(user_id):
    """Get a user's people list from DB."""
    conn = get_db()
    try:
        rows = conn.execute(
            "SELECT name FROM user_people WHERE user_id = ? ORDER BY name",
            (user_id,)
        ).fetchall()
    except Exception:
        rows = []
    conn.close()
    return [dict(r) for r in rows]


def load_cards():
    """Legacy fallback: load cards for very first user."""
    conn = get_db()
    first = conn.execute("SELECT id FROM users ORDER BY id LIMIT 1").fetchone()
    conn.close()
    if first:
        return get_user_cards(first["id"])
    return []


def load_people():
    """Legacy fallback: load people for very first user."""
    conn = get_db()
    first = conn.execute("SELECT id FROM users ORDER BY id LIMIT 1").fetchone()
    conn.close()
    if first:
        return get_user_people(first["id"])
    return []


def load_categories():
    with open(CATEGORIES_PATH) as f:
        return json.load(f)["categories"]


def save_categories(categories):
    with open(CATEGORIES_PATH, "w") as f:
        json.dump({"categories": categories}, f, indent=2)

# ─── DB Init ─────────────────────────────────────────────────────────────────

def init_db():
    conn = get_db()
    conn.executescript("""
        CREATE TABLE IF NOT EXISTS transactions (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            date TEXT NOT NULL,
            description TEXT NOT NULL,
            amount REAL NOT NULL,
            category TEXT DEFAULT 'Other',
            card_id TEXT DEFAULT 'other',
            txn_type TEXT DEFAULT 'debit',
            notes TEXT DEFAULT '',
            source TEXT DEFAULT 'manual',
            created_at TEXT DEFAULT (datetime('now','localtime')),
            person TEXT DEFAULT "",
            cashback REAL DEFAULT 0,
            user_id INTEGER REFERENCES users(id)
        );

        CREATE INDEX IF NOT EXISTS idx_txn_date ON transactions(date);
        CREATE INDEX IF NOT EXISTS idx_txn_category ON transactions(category);
        CREATE INDEX IF NOT EXISTS idx_txn_card ON transactions(card_id);

        CREATE TABLE IF NOT EXISTS monthly_balances (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            month TEXT NOT NULL UNIQUE,
            start_balance REAL DEFAULT 0,
            end_balance REAL DEFAULT 0,
            updated_at TEXT DEFAULT (datetime('now','localtime'))
        );

        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT UNIQUE NOT NULL,
            password_hash TEXT NOT NULL,
            display_name TEXT NOT NULL,
            security_word TEXT DEFAULT '',
            created_at TEXT DEFAULT (datetime('now'))
        );

        CREATE TABLE IF NOT EXISTS user_cards (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER NOT NULL,
            card_id TEXT NOT NULL,
            name TEXT NOT NULL,
            type TEXT DEFAULT 'Other',
            FOREIGN KEY (user_id) REFERENCES users(id),
            UNIQUE(user_id, card_id)
        );

        CREATE TABLE IF NOT EXISTS user_people (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER NOT NULL,
            name TEXT NOT NULL,
            FOREIGN KEY (user_id) REFERENCES users(id),
            UNIQUE(user_id, name)
        );
    """)
    conn.commit()

    # Migration: drop bank column from user_cards
    try:
        conn.execute("ALTER TABLE user_cards DROP COLUMN bank")
        conn.commit()
    except Exception:
        pass  # Column already dropped or didn't exist

    # Migration: add security_word column to users
    try:
        conn.execute("ALTER TABLE users ADD COLUMN security_word TEXT DEFAULT ''")
        conn.commit()
    except Exception:
        pass  # Column already exists

    conn.close()

init_db()

# Auto-migrate: add missing columns to existing tables
for col, col_type in [("person", "TEXT DEFAULT ''"), ("user_id", "INTEGER REFERENCES users(id)")]:
    try:
        conn = get_db()
        conn.execute(f"ALTER TABLE transactions ADD COLUMN {col} {col_type}")
        conn.commit()
        conn.close()
    except Exception:
        pass  # column already exists
try:
    conn = get_db()
    conn.execute("CREATE INDEX IF NOT EXISTS idx_txn_user ON transactions(user_id)")
    conn.commit()
    conn.close()
except Exception:
    pass

# ─── Claim orphan transactions (old DB without user_id) ─────────────────
try:
    conn = get_db()
    first_user = conn.execute("SELECT id FROM users ORDER BY id LIMIT 1").fetchone()
    if first_user:
        conn.execute("PRAGMA foreign_keys = OFF")
        conn.execute("UPDATE transactions SET user_id = ? WHERE user_id IS NULL", (first_user["id"],))
        conn.commit()
        conn.execute("PRAGMA foreign_keys = ON")
        print(f"[migrate] Claimed orphan transactions → user {first_user['id']}")
    conn.close()
except Exception:
    pass

def _migrate_json_to_db():
    conn = get_db()
    card_count = conn.execute("SELECT COUNT(*) FROM user_cards").fetchone()[0]
    people_count = conn.execute("SELECT COUNT(*) FROM user_people").fetchone()[0]
    first_user = conn.execute("SELECT id FROM users ORDER BY id LIMIT 1").fetchone()
    conn.close()
    if not first_user:
        return
    uid = first_user["id"]
    migrated = False
    if card_count == 0 and os.path.exists(CARDS_PATH):
        with open(CARDS_PATH) as f:
            data = json.load(f)
        conn = get_db()
        for c in data.get("cards", []):
            conn.execute("INSERT OR IGNORE INTO user_cards (user_id, card_id, name, type) VALUES (?, ?, ?, ?)", (uid, c["id"], c["name"], c.get("type", "Other")))
        conn.commit()
        conn.close()
        migrated = True
    if people_count == 0 and os.path.exists(PEOPLE_PATH):
        with open(PEOPLE_PATH) as f:
            data = json.load(f)
        conn = get_db()
        for p in data.get("people", []):
            conn.execute("INSERT OR IGNORE INTO user_people (user_id, name) VALUES (?, ?)", (uid, p["name"]))
        conn.commit()
        conn.close()
        migrated = True
    if migrated:
        print("[migrate] Existing JSON data migrated to DB")

_migrate_json_to_db()

# ─── Lazy Startup Sync (runs on first request, not at import time) ───────────
_startup_done = False

@app.before_request
def _lazy_startup():
    """Run init_excel + smart_sync exactly once, on the first real request.
    This avoids crashes on PythonAnywhere when the WSGI worker imports the
    module before the database tables exist, and is safe even on a fresh DB."""
    global _startup_done
    if _startup_done:
        return
    _startup_done = True

    global _excel_mtime
    try:
        init_excel()
        smart_sync()
        _excel_mtime = EXCEL_PATH.stat().st_mtime if EXCEL_PATH.exists() else 0
    except Exception as _e:
        print(f"⚠️ Startup sync skipped: {_e}")

# ─── Security Headers (applied to every response) ─────────────────────────
@app.after_request
def _add_security_headers(response):
    response.headers['X-Content-Type-Options'] = 'nosniff'
    response.headers['X-Frame-Options'] = 'DENY'
    response.headers['Strict-Transport-Security'] = 'max-age=31536000; includeSubDomains'
    response.headers['Referrer-Policy'] = 'strict-origin-when-cross-origin'
    response.headers['Permissions-Policy'] = 'geolocation=(), microphone=(), camera=()'
    return response

# ─── Custom Error Pages (no stack trace leaks) ────────────────────────────
@app.errorhandler(404)
def not_found(e):
    return render_template("error.html", code=404, message="Page not found."), 404

@app.errorhandler(500)
def server_error(e):
    return render_template("error.html", code=500, message="Something went wrong."), 500

# ─── Auth Routes
from werkzeug.security import generate_password_hash, check_password_hash

@app.route("/login", methods=["GET", "POST"])
@limiter.limit("5 per minute")
def login():
    if request.method == "POST":
        username = sanitize(request.form.get("username", "")).lower()
        password = request.form.get("password", "")
        if not PasswordValidator.validate_username(username):
            flash("Invalid username format.", "danger")
            return render_template("login.html")
        conn = get_db()
        user = conn.execute(
            "SELECT id, username, password_hash, display_name FROM users WHERE username = ?",
            (username,)
        ).fetchone()
        conn.close()
        if user and check_password_hash(user["password_hash"], password):
            session.clear()
            session["user_id"] = user["id"]
            session["username"] = user["username"]
            session["display_name"] = user["display_name"]
            flash(f"Welcome back, {user['display_name']}!", "success")
            return redirect(url_for("index"))
        flash("Invalid username or password.", "danger")
    return render_template("login.html")

@app.route("/register", methods=["GET", "POST"])
@limiter.limit("3 per hour")
def register():
    if request.method == "POST":
        username = sanitize(request.form.get("username", "")).lower()
        display_name = sanitize(request.form.get("display_name", ""))
        password = request.form.get("password", "")
        confirm = request.form.get("confirm", "")
        if not username or not display_name or not password:
            flash("All fields are required.", "danger")
            return render_template("register.html")
        if password != confirm:
            flash("Passwords do not match.", "danger")
            return render_template("register.html")
        valid_pw, msg = PasswordValidator.validate(password)
        if not valid_pw:
            flash(f"Password too weak: {msg}", "danger")
            return render_template("register.html")
        conn = get_db()
        existing = conn.execute(
            "SELECT id FROM users WHERE username = ?", (username,)
        ).fetchone()
        if existing:
            conn.close()
            flash("Username already taken.", "danger")
            return render_template("register.html")
        pw_hash = generate_password_hash(password)
        cur = conn.execute(
            "INSERT INTO users (username, password_hash, display_name) VALUES (?, ?, ?)",
            (username, pw_hash, display_name)
        )
        user_id = cur.lastrowid
        conn.commit()
        conn.close()
        seed_new_user(user_id)
        # Claim orphan transactions (from old DB, no user_id) for this new user
        try:
            conn2 = get_db()
            conn2.execute("PRAGMA foreign_keys = OFF")
            conn2.execute("UPDATE transactions SET user_id = ? WHERE user_id IS NULL", (user_id,))
            conn2.commit()
            conn2.execute("PRAGMA foreign_keys = ON")
            conn2.close()
        except Exception:
            pass
        flash("Account created! Please log in.", "success")
        return redirect(url_for("login"))
    return render_template("register.html")


@app.route("/forgot-password", methods=["GET", "POST"])
@limiter.limit("3 per minute")
def forgot_password():
    """Reset password using security word verification."""
    if request.method == "POST":
        username = sanitize(request.form.get("username", "")).lower()
        security_word = sanitize(request.form.get("security_word", ""))
        new_password = request.form.get("new_password", "")
        confirm = request.form.get("confirm", "")

        if not username or not security_word or not new_password:
            flash("All fields are required.", "danger")
            return render_template("forgot_password.html", username=username)

        if new_password != confirm:
            flash("Passwords do not match.", "danger")
            return render_template("forgot_password.html", username=username)

        valid_pw, msg = PasswordValidator.validate(new_password)
        if not valid_pw:
            flash(f"Password too weak: {msg}", "danger")
            return render_template("forgot_password.html", username=username)

        conn = get_db()
        user = conn.execute(
            "SELECT id, security_word FROM users WHERE username = ?",
            (username,)
        ).fetchone()

        if not user:
            conn.close()
            flash("Username not found.", "danger")
            return render_template("forgot_password.html", username=username)

        if not user["security_word"]:
            conn.close()
            flash("This account does not have a security word set. Contact the admin.", "danger")
            return render_template("forgot_password.html", username=username)

        if user["security_word"] != security_word:
            conn.close()
            flash("Security word is incorrect.", "danger")
            return render_template("forgot_password.html", username=username)

        # Security word matches — update password
        pw_hash = generate_password_hash(new_password)
        conn.execute(
            "UPDATE users SET password_hash = ? WHERE id = ?",
            (pw_hash, user["id"])
        )
        conn.commit()
        conn.close()
        flash("✅ Password reset successfully! Please log in.", "success")
        return redirect(url_for("login"))

    return render_template("forgot_password.html")

@app.route("/logout")

def logout():
    session.clear()
    flash("Logged out.", "info")
    return redirect(url_for("login"))

# ─── Auto-Categorization Engine ──────────────────────────────────────────────

def auto_categorize(description, amount=None):
    """Match a transaction description against category keywords."""
    desc_lower = description.lower().strip()
    categories = load_categories()

    best_match = None
    best_score = 0

    for cat in categories:
        for keyword in cat["keywords"]:
            if keyword in desc_lower:
                # Longer keyword match = higher confidence
                score = len(keyword)
                if keyword == desc_lower or desc_lower.startswith(keyword):
                    score += 10  # exact/prefix boost
                if score > best_score:
                    best_score = score
                    best_match = cat["name"]

    # Special: small amounts at restaurants/food places
    if not best_match:
        for name in ["Zomato", "Swiggy", "Amazon", "Flipkart", "BigBasket"]:
            if name.lower() in desc_lower:
                return name

    return best_match or "Other"


def guess_card(description):
    """Try to figure out which card from description (bank statement context)."""
    desc_lower = description.lower()
    if "sbi" in desc_lower or "state bank" in desc_lower:
        return "sbi_cb"
    if "hdfc" in desc_lower:
        return "hdfc_mil"
    if "bob" in desc_lower or "bank of baroda" in desc_lower:
        return "bob_eterna"
    return "other"


def parse_amount(val):
    """Parse a currency string like '1,234.56' or '-1,234' or '1.234,56'."""
    if isinstance(val, (int, float)):
        return float(val)
    val = str(val).strip().replace("₹", "").replace(",", "").replace(" ", "")
    # Handle European format 1.234,56 → 1234.56
    if re.match(r'^\d+\.\d{3}[.,]\d{2}$', val):
        val = val.replace(".", "").replace(",", ".")
    val = val.replace(",", ".")
    try:
        return float(val)
    except ValueError:
        return 0.0

# ─── CSV Import Engine ────────────────────────────────────────────────────────

def parse_uploaded_csv(file_content, card_id=None):
    """
    Parse bank/credit card statement CSV.
    Returns list of dicts with keys: date, description, amount, category.
    Tries to auto-detect column layout.
    """
    reader = csv.DictReader(StringIO(file_content))
    if not reader.fieldnames:
        return [], "No columns found in CSV"

    fields_lower = [f.lower().strip() for f in reader.fieldnames]
    transactions = []
    errors = []

    # Detect column mappings
    date_cols = ["date", "txn date", "transaction date", "posting date", "value date"]
    desc_cols = ["description", "narrative", "particulars", "transaction details",
                 "remarks", "details", "merchant", "merchant name"]
    amt_cols = ["amount", "txn amount", "transaction amount", "debit", "credit",
               "dr", "cr", "charge", "value"]
    debit_cols = ["debit", "dr", "debit amount", "withdrawal"]
    credit_cols = ["credit", "cr", "credit amount", "deposit"]

    date_field = next((fn for fn in reader.fieldnames if fn.lower().strip() in date_cols), None)
    desc_field = next((fn for fn in reader.fieldnames if fn.lower().strip() in desc_cols), None)
    amt_field = next((fn for fn in reader.fieldnames if fn.lower().strip() in amt_cols), None)
    debit_field = next((fn for fn in reader.fieldnames if fn.lower().strip() in debit_cols), None)
    credit_field = next((fn for fn in reader.fieldnames if fn.lower().strip() in credit_cols), None)

    if not date_field:
        return [], "Could not find 'Date' column in CSV"
    if not amt_field and not (debit_field or credit_field):
        return [], "Could not find 'Amount' column in CSV"
    if not desc_field:
        desc_field = reader.fieldnames[1]  # fallback: second column
        errors.append("Description column not found; using second column")

    for i, row in enumerate(reader, start=2):
        try:
            raw_date = row.get(date_field, "").strip()
            raw_desc = row.get(desc_field, "").strip()
            raw_amt = row.get(amt_field, "").strip() if amt_field else "0"
            raw_debit = row.get(debit_field, "").strip() if debit_field else ""
            raw_credit = row.get(credit_field, "").strip() if credit_field else ""

            if not raw_date or (not raw_desc and not raw_amt and not raw_debit and not raw_credit):
                continue  # skip empty rows

            # Parse date - handle DD/MM/YYYY, MM/DD/YYYY, etc.
            parsed_date = None
            for fmt in ["%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%m/%d/%Y",
                        "%d/%m/%y", "%d %b %Y", "%d-%b-%Y", "%d %B %Y"]:
                try:
                    parsed_date = datetime.strptime(raw_date, fmt).strftime("%Y-%m-%d")
                    break
                except ValueError:
                    continue
            if not parsed_date:
                errors.append(f"Row {i}: Could not parse date '{raw_date}'")
                parsed_date = date.today().strftime("%Y-%m-%d")

            # Parse amount
            amount = 0.0
            txn_type = "debit"
            if raw_amt:
                amount = parse_amount(raw_amt)
            elif raw_debit:
                amount = abs(parse_amount(raw_debit))
                txn_type = "debit"
            elif raw_credit:
                amount = abs(parse_amount(raw_credit))
                txn_type = "credit"

            if amount == 0:
                continue

            # If amount is negative, flip to credit/debit accordingly
            if amount < 0:
                amount = abs(amount)
                txn_type = "credit" if txn_type == "debit" else "debit"

            if not raw_desc:
                raw_desc = f"Statement entry {i}"

            # Auto-categorize
            category = auto_categorize(raw_desc, amount)

            transactions.append({
                "date": parsed_date,
                "description": raw_desc[:200],
                "amount": round(amount, 2),
                "category": category,
                "card_id": card_id or guess_card(raw_desc),
                "txn_type": txn_type,
                "source": "upload",
            })
        except Exception as e:
            errors.append(f"Row {i}: {str(e)}")

    return transactions, errors


# ─── Report Engine ────────────────────────────────────────────────────────────

def get_daily_summary(txn_date=None, user_id=None):
    """Get today's or a specific day's transactions and totals."""
    if txn_date is None:
        txn_date = date.today().strftime("%Y-%m-%d")
    conn = get_db()
    rows = conn.execute(
        "SELECT * FROM transactions WHERE date = ? AND user_id = ? ORDER BY id", (txn_date, user_id,)
    ).fetchall()
    total = sum(r["amount"] for r in rows if r["txn_type"] == "debit")
    credits = sum(r["amount"] for r in rows if r["txn_type"] == "credit")
    conn.close()
    return {"date": txn_date, "transactions": rows, "total_debit": total, "total_credit": credits}


def get_monthly_report(year=None, month=None, user_id=None):
    """Get category-wise and card-wise breakdown for a month."""
    today = date.today()
    if year is None:
        year = today.year
    if month is None:
        month = today.month

    conn = get_db()
    rows = conn.execute(
        """SELECT * FROM transactions
           WHERE strftime('%Y', date) = ? AND strftime('%m', date) = ? AND user_id = ?
           ORDER BY date, id""",
        (str(year), f"{month:02d}", user_id,)
    ).fetchall()
    conn.close()

    # Category-wise breakdown
    cat_totals = {}
    cat_txns = {}
    for r in rows:
        cat = r["category"]
        amt = r["amount"] if r["txn_type"] == "debit" else 0
        cat_totals[cat] = cat_totals.get(cat, 0) + amt
        if cat not in cat_txns:
            cat_txns[cat] = []
        cat_txns[cat].append(dict(r))

    # Card-wise breakdown
    card_totals = {}
    card_txns = {}
    for r in rows:
        cid = r["card_id"]
        amt = r["amount"] if r["txn_type"] == "debit" else 0
        card_totals[cid] = card_totals.get(cid, 0) + amt
        if cid not in card_txns:
            card_txns[cid] = []
        card_txns[cid].append(dict(r))

    # Sort categories and cards by amount descending
    cat_totals = dict(sorted(cat_totals.items(), key=lambda x: x[1], reverse=True))
    card_totals = dict(sorted(card_totals.items(), key=lambda x: x[1], reverse=True))

    total_spend = sum(cat_totals.values())

    return {
        "year": year,
        "month": month,
        "month_name": datetime(year, month, 1).strftime("%B"),
        "total_spend": total_spend,
        "transaction_count": len(rows),
        "cat_totals": cat_totals,
        "cat_txns": cat_txns,
        "card_totals": card_totals,
        "card_txns": card_txns,
        "transactions": rows,
    }


def get_card_report(card_id=None, year=None, month=None, user_id=None):
    """Get spending breakdown for a specific card or all cards."""
    today = date.today()
    if year is None:
        year = today.year
    if month is None:
        month = today.month

    conn = get_db()
    if card_id:
        rows = conn.execute(
            """SELECT * FROM transactions
               WHERE card_id = ? AND strftime('%Y', date) = ? AND strftime('%m', date) = ? AND user_id = ?
               ORDER BY date, id""",
            (card_id, str(year), f"{month:02d}", user_id,)
        ).fetchall()
    else:
        rows = conn.execute(
            """SELECT * FROM transactions
               WHERE strftime('%Y', date) = ? AND strftime('%m', date) = ? AND user_id = ?
               ORDER BY date, id""",
            (str(year), f"{month:02d}", user_id,)
        ).fetchall()
    conn.close()

    total = sum(r["amount"] for r in rows if r["txn_type"] == "debit")
    return {"card_id": card_id or "all", "transactions": rows, "total": total}


def export_csv(transactions):
    """Generate CSV file from transaction rows."""
    output = StringIO()
    writer = csv.writer(output)
    writer.writerow(["Date", "Description", "Amount", "Category", "Card", "Type", "Notes", "Source"])
    for t in transactions:
        writer.writerow([
            t["date"], t["description"], t["amount"],
            t["category"], t["card_id"], t["txn_type"],
            t["notes"] or "", t["source"] or ""
        ])
    output.seek(0)
    return output


def export_monthly_excel(year, month, user_id=None):
    """Generate Excel report using CSV (simple, universal)."""
    conn = get_db()
    rows = conn.execute(
        """SELECT date, description, amount, category, card_id, txn_type, notes, source
           FROM transactions
           WHERE strftime('%Y', date) = ? AND strftime('%m', date) = ? AND user_id = ?
           ORDER BY date, id""",
        (str(year), f"{month:02d}", user_id,)
    ).fetchall()
    conn.close()
    return export_csv(rows)


# ─── Flask Routes ─────────────────────────────────────────────────────────────

@app.route("/")

@login_required
def index():
    """Dashboard — show daily summary + quick stats."""

    user_id = session["user_id"]
    today = date.today().strftime("%Y-%m-%d")
    sort_by = request.args.get("sort", "id")
    sort_today = request.args.get("sort_today", "id")
    per_page = 10

    conn = get_db()

    # ─── Today's Transactions (paginated) ─────────────────────────────────
    page_today = int(request.args.get("page_today", 1))
    today_total = conn.execute(
        "SELECT COUNT(*) FROM transactions WHERE date = ? AND user_id = ?", (today, user_id,)
    ).fetchone()[0]
    today_pages = max(1, (today_total + per_page - 1) // per_page)
    if page_today > today_pages:
        page_today = today_pages
    today_offset = (page_today - 1) * per_page

    if sort_today == "amount":
        today_order = "amount DESC, id DESC"
    else:
        today_order = "id DESC"
    today_rows = conn.execute(
        f"SELECT * FROM transactions WHERE date = ? AND user_id = ? ORDER BY {today_order} LIMIT ? OFFSET ?",
        (today, user_id, per_page, today_offset)
    ).fetchall()
    total_debit = sum(r["amount"] for r in today_rows if r["txn_type"] == "debit")
    total_credit = sum(r["amount"] for r in today_rows if r["txn_type"] == "credit")
    daily = {"date": today, "transactions": today_rows,
             "total_debit": total_debit, "total_credit": total_credit}

    # ─── Recent Transactions (paginated) ──────────────────────────────────
    page_recent = int(request.args.get("page_recent", 1))
    recent_total = conn.execute("SELECT COUNT(*) FROM transactions WHERE user_id = ?", (user_id,)).fetchone()[0]
    recent_pages = max(1, (recent_total + per_page - 1) // per_page)
    if page_recent > recent_pages:
        page_recent = recent_pages
    recent_offset = (page_recent - 1) * per_page

    if sort_by == "date":
        order_clause = "date DESC, id DESC"
    elif sort_by == "amount":
        order_clause = "amount DESC, id DESC"
    else:
        order_clause = "id DESC"
    recent = conn.execute(
        f"SELECT * FROM transactions WHERE user_id = ? ORDER BY {order_clause} LIMIT ? OFFSET ?",
        (user_id, per_page, recent_offset)
    ).fetchall()

    # Monthly total so far
    monthly_debit = conn.execute(
        "SELECT COALESCE(SUM(amount), 0) FROM transactions WHERE "
        "strftime('%Y', date) = ? AND strftime('%m', date) = ? AND txn_type = 'debit' AND user_id = ?",
        (str(date.today().year), f"{date.today().month:02d}", user_id,)
    ).fetchone()[0]
    monthly_credit = conn.execute(
        "SELECT COALESCE(SUM(amount), 0) FROM transactions WHERE "
        "strftime('%Y', date) = ? AND strftime('%m', date) = ? AND txn_type = 'credit' AND user_id = ?",
        (str(date.today().year), f"{date.today().month:02d}", user_id,)
    ).fetchone()[0]
    monthly_total = monthly_credit - monthly_debit

    # Today's credit count
    today_credits_count = conn.execute(
        "SELECT COUNT(*) FROM transactions WHERE date = ? AND txn_type = 'credit' AND user_id = ?",
        (today, user_id,)
    ).fetchone()[0]

    # Monthly transaction count
    monthly_txns_count = conn.execute(
        "SELECT COUNT(*) FROM transactions WHERE "
        "strftime('%Y', date) = ? AND strftime('%m', date) = ? AND user_id = ?",
        (str(date.today().year), f"{date.today().month:02d}", user_id,)
    ).fetchone()[0]
    total_cb = conn.execute(
        "SELECT COALESCE(SUM(cashback), 0) FROM transactions WHERE cashback > 0 AND user_id = ?", (user_id,)
    ).fetchone()[0]
    conn.close()

    cards = get_user_cards(user_id)
    categories = load_categories()

    # ─── AJAX partials ──────────────────────────────────────────────────────
    ajax_today = request.args.get("ajax_today") == "1"
    ajax_recent = request.args.get("ajax_recent") == "1"
    if ajax_today:
        return render_template("_today_section.html",
                             daily=daily,
                             sort_today=sort_today,
                             page_today=page_today, today_pages=today_pages, today_total=today_total, today_credits_count=today_credits_count)
    if ajax_recent:
        return render_template("_recent_section.html",
                             recent=recent,
                             sort_by=sort_by,
                             page_recent=page_recent, recent_pages=recent_pages, recent_total=recent_total)

    return render_template("index.html",
                         daily=daily,
                         recent=recent,
                         monthly_total=monthly_total,
                         monthly_txns_count=monthly_txns_count,
                         total_cb=total_cb,
                         cards=cards,
                         categories=categories,
                         sort_by=sort_by,
                         sort_today=sort_today,
                         page_today=page_today, today_pages=today_pages, today_total=today_total, today_credits_count=today_credits_count,
                         page_recent=page_recent, recent_pages=recent_pages, recent_total=recent_total)




@app.route("/add", methods=["GET", "POST"])

@login_required
def add_transaction():
    """Add a single transaction via form."""
    user_id = session["user_id"]
    display_name = session.get("display_name", "")
    cards = get_user_cards(user_id)
    categories = load_categories()
    persons = [p["name"] for p in get_user_people(user_id)]

    if request.method == "POST":
        txn_date = request.form.get("date", date.today().strftime("%Y-%m-%d"))
        description = sanitize(request.form.get("description", ""))
        amount = parse_amount(request.form.get("amount", "0"))
        category = request.form.get("category", "Other")
        card_id = request.form.get("card_id", "").strip()
        if not card_id:
            card_id = "other"
        txn_type = request.form.get("txn_type", "debit")
        notes = sanitize(request.form.get("notes", ""))
        person = sanitize(request.form.get("person", ""))

        if not description:
            flash("Description is required!", "danger")
            return render_template("add.html", cards=cards, categories=categories, persons=persons)

        if amount <= 0:
            flash("Amount must be greater than zero!", "danger")
            return render_template("add.html", cards=cards, categories=categories, persons=persons)

        if amount > 50_000_000:
            flash("Amount cannot exceed ₹5 crore!", "danger")
            return render_template("add.html", cards=cards, categories=categories, persons=persons)

        # Auto-categorize if user chose "Auto"
        if category == "Auto":
            category = auto_categorize(description, amount)

        # Manual cashback from user input (optional)
        cashback = float(request.form.get("cashback", 0) or 0)
        cashback = round(cashback, 2)

        conn = get_db()
        conn.execute(
            "INSERT INTO transactions (date, description, amount, category, card_id, txn_type, notes, source, person, cashback, user_id) "
            "VALUES (?, ?, ?, ?, ?, ?, ?, 'manual', ?, ?, ?)",
            (txn_date, description, round(amount, 2), category, card_id, txn_type, notes, person, cashback, user_id)
        )
        conn.commit()
        conn.close()

        # Sync to Excel
        add_transaction_to_excel({
            "date": txn_date,
            "description": description,
            "amount": round(amount, 2),
            "category": category,
            "card_id": card_id,
            "txn_type": txn_type,
            "notes": notes,
        })

        flash(f"✅ Transaction added: ₹{amount:,.2f} — {description}", "success")
        return redirect(url_for("index"))

    return render_template("add.html",
                         cards=cards,
                         categories=categories,
                         persons=persons,
                         today=date.today().strftime("%Y-%m-%d"))




#@app.route("/upload", methods=["GET", "POST"])
#
#@login_required
#def upload_statement():
#    """Upload a bank/credit card CSV statement."""
#
#    person = session.get("display_name", "")
#    user_id = session["user_id"]
#    cards = get_user_cards(user_id)
#
#    if request.method == "POST":
#        if "file" not in request.files:
#            flash("No file selected!", "danger")
#            return render_template("upload.html", cards=cards)
#
#        file = request.files["file"]
#        if file.filename == "":
#            flash("No file selected!", "danger")
#            return render_template("upload.html", cards=cards)
#
#        if not file.filename.endswith(".csv"):
#            flash("Please upload a CSV file!", "danger")
#            return render_template("upload.html", cards=cards)
#
#        card_id = request.form.get("card_id", "")
#        if card_id == "auto":
#            card_id = None  # Let parser guess
#
#        content = file.read().decode("utf-8-sig", errors="ignore")
#
#        # Check if it's a standard bank format or need to guess
#        transactions, errors = parse_uploaded_csv(content, card_id)
#
#        if not transactions:
#            flash(f"❌ No transactions found! Errors: {'; '.join(errors[:5])}", "danger")
#            return render_template("upload.html", cards=cards, debug_preview=content[:2000])
#
#        # Save to DB
#        conn = get_db()
#        inserted = 0
#        for txn in transactions:
#            conn.execute(
#                """INSERT INTO transactions (date, description, amount, category, card_id, txn_type, source, person)
#                   VALUES (?, ?, ?, ?, ?, ?, 'upload', ?)""",
#                (txn["date"], txn["description"][:200], txn["amount"],
#                 txn["category"], txn["card_id"], txn["txn_type"], person)
#            )
#            inserted += 1
#        conn.commit()
#        conn.close()
#
#        # Sync all to Excel
#        db_sync_to_excel()
#
#        msg = f"✅ Imported {inserted} transactions from {file.filename}"
#        if errors:
#            msg += f"\n⚠️ {len(errors)} warnings (first 3): " + "; ".join(errors[:3])
#
#        flash(msg, "success")
#        return redirect(url_for("index"))
#
#    return render_template("upload.html", cards=cards)


@app.route("/reports")

@login_required
def reports():
    """Reports page — daily, monthly, card-wise."""

    user_id = session["user_id"]
    today = date.today()

    # Daily report
    day_param = request.args.get("day", today.strftime("%Y-%m-%d"))
    daily = get_daily_summary(day_param, user_id)

    # Monthly report
    year_param = int(request.args.get("year", today.year))
    month_raw = request.args.get("month", str(today.month))
    try:
        month_param = int(month_raw)
    except ValueError:
        # Try parsing YYYY-MM format
        if "-" in month_raw:
            parts = month_raw.split("-")
            if len(parts) == 2:
                year_param = int(parts[0])
                month_param = int(parts[1])
            else:
                month_param = today.month
        else:
            month_param = today.month
    monthly = get_monthly_report(year_param, month_param, user_id)

    # Card report
    card_param = request.args.get("card_id", None)
    card_rpt = get_card_report(card_param, year_param, month_param, user_id)

    cards = get_user_cards(user_id)
    card_map = {c["id"]: c["name"] for c in cards}

    return render_template("reports.html",
                         daily=daily,
                         monthly=monthly,
                         card_rpt=card_rpt,
                         cards=cards,
                         card_map=card_map,
                         today=today,
                         selected_year=year_param,
                         selected_month=month_param,
                         selected_day=day_param,
                         selected_card=card_param,
                         datetime=datetime)


@app.route("/export")

@login_required
def export_data():
    """Export transactions as CSV."""

    user_id = session["user_id"]
    conn = get_db()
    rows = conn.execute(
        "SELECT * FROM transactions WHERE user_id = ? ORDER BY date DESC, id DESC", (user_id,)
    ).fetchall()
    conn.close()

    output = export_csv(rows)
    return Response(
        output.getvalue(),
        mimetype="text/csv",
        headers={"Content-Disposition": "attachment;filename=transactions.csv"}
    )


@app.route("/export/monthly")
@login_required
def export_monthly():
    """Export monthly report as CSV."""

    user_id = session["user_id"]
    today = date.today()
    year = int(request.args.get("year", today.year))
    month = int(request.args.get("month", today.month))
    output = export_monthly_excel(year, month, user_id)
    return Response(
        output.getvalue(),
        mimetype="text/csv",
        headers={
            "Content-Disposition":
            f"attachment;filename=report_{year}_{month:02d}.csv"
        }
    )


@app.route("/delete/<int:txn_id>", methods=["POST"])
@login_required
def delete_transaction(txn_id):
    """Delete a transaction."""
    user_id = session["user_id"]
    conn = get_db()
    conn.execute("DELETE FROM transactions WHERE id = ? AND user_id = ?", (txn_id, user_id))
    conn.commit()
    conn.close()

    # Sync to Excel
    sc = db_sync_to_excel()
    if sc == 0:
        flash("⚠️ Transaction deleted in website, but Excel could not be updated. Close Excel and Sync.", "warning")
    else:
        flash("🗑️ Transaction deleted", "info")
    return redirect(request.referrer or url_for("index"))


@app.route("/edit/<int:txn_id>", methods=["GET", "POST"])
@login_required
def edit_transaction(txn_id):
    """Edit a transaction."""
    person = session.get("display_name", "")
    user_id = session["user_id"]
    conn = get_db()
    if request.method == "POST":
        txn_date = request.form.get("date", date.today().strftime("%Y-%m-%d"))
        description = sanitize(request.form.get("description", ""))
        amount = parse_amount(request.form.get("amount", "0"))
        category = request.form.get("category", "Other")
        card_id = request.form.get("card_id", "").strip()
        if not card_id:
            card_id = "other"
        txn_type = request.form.get("txn_type", "debit")
        notes = sanitize(request.form.get("notes", ""))
        person = sanitize(request.form.get("person", ""))

        if not description:
            flash("Description is required!", "danger")
            return redirect(url_for("edit_transaction", txn_id=txn_id))

        if amount <= 0:
            flash("Amount must be greater than zero!", "danger")
            return redirect(url_for("edit_transaction", txn_id=txn_id))

        if category == "Auto":
            category = auto_categorize(description, amount)

        # Manual cashback from user input (optional)
        cashback_raw = request.form.get("cashback", "0") or "0"
        print(f"📋 EDIT FORM DATA: {dict(request.form)}")  # DEBUG
        print(f"📋 CASHBACK RAW: '{cashback_raw}'")
        cashback = float(cashback_raw)
        cashback = round(cashback, 2)

        conn.execute(
            "UPDATE transactions SET date=?, description=?, amount=?, category=?, card_id=?, txn_type=?, notes=?, person=?, cashback=? "
            "WHERE id=? AND user_id=?",
            (txn_date, description, round(amount, 2), category, card_id, txn_type, notes, person, cashback, txn_id, user_id)
        )
        conn.commit()
        conn.close()

        # Sync to Excel
        sync_count = db_sync_to_excel()
        if sync_count == 0:
            flash(f"⚠️ Transaction updated in website, but could not update Excel (file may be open). Close Excel and click Sync.", "warning")
        else:
            flash(f"✅ Transaction updated: ₹{amount:,.2f} — {description}", "success")
        return redirect(url_for("index"))

    # GET — show edit form
    txn = conn.execute(
        "SELECT id, date, description, amount, category, card_id, txn_type, notes, person, cashback FROM transactions WHERE id = ? AND user_id = ?",
        (txn_id, user_id)
    ).fetchone()
    conn.close()

    if not txn:
        flash("Transaction not found!", "danger")
        return redirect(url_for("index"))

    cards = get_user_cards(user_id)
    persons = [p["name"] for p in get_user_people(user_id)]
    return render_template("edit.html",
                         txn=txn,
                         cards=cards,
                         categories=load_categories(),
                         persons=persons,
                         today=date.today().strftime("%Y-%m-%d"))


@app.route("/api/stats")
@login_required
def api_stats():
    """JSON API for stats (handy for external use)."""

    user_id = session["user_id"]
    today = date.today()
    year = today.year
    month = today.month

    conn = get_db()

    # Daily
    day_txns = conn.execute(
        "SELECT COUNT(*) as count, COALESCE(SUM(amount), 0) as total FROM transactions "
        "WHERE date = ? AND txn_type = 'debit' AND user_id = ?",
        (today.strftime("%Y-%m-%d"), user_id,)
    ).fetchone()

    # Monthly
    month_txns = conn.execute(
        "SELECT COUNT(*) as count, COALESCE(SUM(amount), 0) as total FROM transactions "
        "WHERE strftime('%Y', date) = ? AND strftime('%m', date) = ? AND txn_type = 'debit' AND user_id = ?",
        (str(year), f"{month:02d}", user_id,)
    ).fetchone()

    # By category this month
    cat_rows = conn.execute(
        "SELECT category, COUNT(*) as count, SUM(amount) as total FROM transactions "
        "WHERE strftime('%Y', date) = ? AND strftime('%m', date) = ? AND txn_type = 'debit' AND user_id = ? "
        "GROUP BY category ORDER BY total DESC",
        (str(year), f"{month:02d}", user_id,)
    ).fetchall()

    # By card this month
    card_rows = conn.execute(
        "SELECT card_id, COUNT(*) as count, SUM(amount) as total FROM transactions "
        "WHERE strftime('%Y', date) = ? AND strftime('%m', date) = ? AND txn_type = 'debit' AND user_id = ? "
        "GROUP BY card_id ORDER BY total DESC",
        (str(year), f"{month:02d}", user_id,)
    ).fetchall()

    conn.close()

    return jsonify({
        "daily": {"count": day_txns["count"], "total": day_txns["total"]},
        "monthly": {"count": month_txns["count"], "total": month_txns["total"]},
        "by_category": [dict(c) for c in cat_rows],
        "by_card": [dict(c) for c in card_rows],
    })


@app.route("/settings", methods=["GET", "POST"])

@login_required
def settings():
    """Manage cards, categories, and people."""

    user_id = session["user_id"]
    cards = get_user_cards(user_id)
    categories = load_categories()
    people = [{"name": p["name"]} for p in get_user_people(user_id)]

    # Get balances for each person the user has transactions with (excluding self)
    conn = get_db()
    display_name = session.get("display_name", "")
    balance_rows = conn.execute("""\
        SELECT person,
               COALESCE(SUM(CASE WHEN txn_type='debit' THEN amount ELSE 0 END), 0) as total_debit,
               COALESCE(SUM(CASE WHEN txn_type='credit' THEN amount ELSE 0 END), 0) as total_credit
        FROM transactions
        WHERE user_id = ? AND person IS NOT NULL AND person != '' AND person != ?
        GROUP BY person
    """, (user_id, display_name)).fetchall()
    conn.close()
    balance_map = {}
    for r in balance_rows:
        balance_map[r["person"]] = r["total_credit"] - r["total_debit"]

    if request.method == "POST":
        action = request.form.get("action")

        if action == "add_card":
            card_name = sanitize(request.form["card_name"])
            card_id = card_name.lower().replace(" ", "_")
            card_type = sanitize(request.form.get("card_type", ""))
            conn = get_db()
            try:
                existing = conn.execute(
                    "SELECT id FROM user_cards WHERE user_id = ? AND card_id = ?",
                    (user_id, card_id)
                ).fetchone()
                if existing:
                    flash("❌ Card name already exists.", "danger")
                else:
                    conn.execute(
                        "INSERT INTO user_cards (user_id, card_id, name, type) VALUES (?, ?, ?, ?)",
                        (user_id, card_id, card_name, card_type)
                    )
                    conn.commit()
                    flash(f"✅ Card '{card_name}' added!", "success")
            except Exception as e:
                flash(f"❌ Error adding card: {e}", "danger")
            conn.close()
            cards = get_user_cards(user_id)

        elif action == "remove_card":
            cid = request.form["card_id"]
            conn = get_db()
            conn.execute("DELETE FROM user_cards WHERE user_id = ? AND card_id = ?", (user_id, cid))
            conn.commit()
            conn.close()
            cards = get_user_cards(user_id)
            flash("🗑️ Card removed", "info")

        elif action == "add_category":
            new_cat = {
                "name": sanitize(request.form["cat_name"]),
                "keywords": [sanitize(k) for k in request.form["cat_keywords"].split(",") if k.strip()],
                "type": sanitize(request.form["cat_type"]),
            }
            categories.append(new_cat)
            save_categories(categories)
            flash(f"✅ Category '{new_cat['name']}' added!", "success")

        elif action == "remove_category":
            cat_name = request.form["cat_name"]
            categories = [c for c in categories if c["name"] != cat_name]
            save_categories(categories)
            flash("🗑️ Category removed", "info")

        elif action == "add_person":
            name = sanitize(request.form["person_name"])
            if name:
                conn = get_db()
                try:
                    existing = conn.execute(
                        "SELECT id FROM user_people WHERE user_id = ? AND name = ?",
                        (user_id, name)
                    ).fetchone()
                    if existing:
                        flash("❌ Person already exists.", "danger")
                    else:
                        conn.execute(
                            "INSERT INTO user_people (user_id, name) VALUES (?, ?)",
                            (user_id, name)
                        )
                        conn.commit()
                        flash(f"✅ Person '{name}' added!", "success")
                except Exception as e:
                    flash(f"❌ Error adding person: {e}", "danger")
                conn.close()
            people = [{"name": p["name"]} for p in get_user_people(user_id)]

        elif action == "remove_person":
            name = request.form["person_name"]
            # Check if person has pending balance
            conn = get_db()
            balance_row = conn.execute("""\
                SELECT COALESCE(SUM(CASE WHEN txn_type='debit' THEN amount ELSE 0 END), 0) as total_debit,
                       COALESCE(SUM(CASE WHEN txn_type='credit' THEN amount ELSE 0 END), 0) as total_credit
                FROM transactions
                WHERE user_id = ? AND person = ?
            """, (user_id, name)).fetchone()
            conn.close()
            total_debit = balance_row["total_debit"] if balance_row else 0
            total_credit = balance_row["total_credit"] if balance_row else 0
            balance = total_credit - total_debit
            if balance != 0:
                flash(f"❌ Cannot remove '{name}' — pending amount of ₹{abs(balance):.0f} ({'you owe them' if balance > 0 else 'owes you'}). Settle up first!", "danger")
            else:
                conn = get_db()
                conn.execute("DELETE FROM user_people WHERE user_id = ? AND name = ?", (user_id, name))
                conn.commit()
                conn.close()
                flash("🗑️ Person removed", "info")
            people = [{"name": p["name"]} for p in get_user_people(user_id)]

        return redirect(request.form.get("next") or url_for("settings"))

    return render_template("settings.html", cards=cards, categories=categories, people=people, balance_map=balance_map)


@app.route("/change-password", methods=["POST"])
@login_required
def change_password():
    """Change the logged-in user's password."""
    current = request.form.get("current_password", "")
    new_pw = request.form.get("new_password", "")
    confirm = request.form.get("confirm_password", "")

    if not current or not new_pw or not confirm:
        flash("All fields are required.", "danger")
        return redirect(url_for("settings"))

    if new_pw != confirm:
        flash("New passwords don't match.", "danger")
        return redirect(url_for("settings"))

    if len(new_pw) < 4:
        flash("Password must be at least 4 characters.", "danger")
        return redirect(url_for("settings"))

    conn = get_db()
    user = conn.execute("SELECT * FROM users WHERE id = ?", (session["user_id"],)).fetchone()
    if not user or not check_password_hash(user["password_hash"], current):
        conn.close()
        flash("Current password is incorrect.", "danger")
        return redirect(url_for("settings"))

    new_hash = generate_password_hash(new_pw)
    conn.execute("UPDATE users SET password_hash = ? WHERE id = ?", (new_hash, session["user_id"]))
    conn.commit()
    conn.close()
    flash("✅ Password changed successfully!", "success")
    return redirect(url_for("settings"))


@app.route("/update-security-word", methods=["POST"])
@login_required
def update_security_word():
    """Update the logged-in user's security word."""
    current_pw = request.form.get("current_password", "")
    new_word = sanitize(request.form.get("security_word", ""))

    if not current_pw or not new_word:
        flash("Current password and security word are required.", "danger")
        return redirect(url_for("settings"))

    conn = get_db()
    user = conn.execute("SELECT * FROM users WHERE id = ?", (session["user_id"],)).fetchone()
    if not user or not check_password_hash(user["password_hash"], current_pw):
        conn.close()
        flash("Current password is incorrect.", "danger")
        return redirect(url_for("settings"))

    conn.execute("UPDATE users SET security_word = ? WHERE id = ?", (new_word, session["user_id"]))
    conn.commit()
    conn.close()
    flash("✅ Security word updated!", "success")
    return redirect(url_for("settings"))


@app.route("/cashback")

@login_required
def cashback_page():
    """Cashback dashboard — per-card, per-month breakdown."""

    user_id = session["user_id"]
    today = date.today()

    conn = get_db()
    rows = conn.execute(
        "SELECT id, date, description, amount, category, card_id, txn_type, cashback "
        "FROM transactions WHERE cashback > 0 AND user_id = ? "
        "ORDER BY date DESC, id DESC",
        (user_id,)
    ).fetchall()
    conn.close()

    # Build totals per card
    card_totals = {}
    cb_transactions = []

    for r in rows:
        cid = r["card_id"] or "other"
        if cid not in card_totals:
            card_totals[cid] = {
                "label": CARD_CB_LABELS.get(cid, ""),
                "total_cb": 0.0,
                "count": 0,
            }
        cb_val = round(r["cashback"] or 0, 2)
        card_totals[cid]["total_cb"] += cb_val
        if cb_val > 0:
            card_totals[cid]["count"] += 1
            cb_transactions.append({
                "id": r["id"],
                "date": r["date"],
                "description": r["description"],
                "amount": round(r["amount"], 2),
                "category": r["category"],
                "card_id": r["card_id"],
                "cashback": cb_val,
                "cashback_label": CARD_CB_LABELS.get(r["card_id"], ""),
            })

    # Total cashback
    total_cb = round(sum(t["total_cb"] for t in card_totals.values()), 2)

    # Sort cards by cashback descending
    card_totals = dict(sorted(card_totals.items(), key=lambda x: x[1]["total_cb"], reverse=True))

    return render_template(
        "cashback.html",
        card_totals=card_totals,
        total_cb=total_cb,
        cb_transactions=cb_transactions,
        today=today,
    )


@app.route("/open-excel")
@login_required
def open_excel():
    """Download the Excel file (instead of opening on server)."""
    if session.get("username") != "pratik":
        flash("⛔ Excel download is only available for the primary user.", "danger")
        return redirect(request.referrer or url_for("index"))
    try:
        return send_file(
            EXCEL_PATH,
            as_attachment=True,
            download_name="expense_tracker.xlsx",
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        flash(f"❌ Could not serve Excel file: {e}", "danger")
        return redirect(request.referrer or url_for("index"))


@app.route("/backup-download")

@login_required
def backup_download():
    """Download a ZIP of database + categories (Pratik only)."""
    if session.get("username") != "pratik":
        flash("⛔ Backup download is only available for the primary user.", "danger")
        return redirect(request.referrer or url_for("index"))

    import io, zipfile
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for path, arcname in [(DB_PATH, "finance.db"), (CATEGORIES_PATH, "categories.json")]:
            if path.exists():
                z.write(path, arcname)
    buf.seek(0)
    from datetime import datetime
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    return send_file(
        buf,
        as_attachment=True,
        download_name=f"finance_backup_{ts}.zip",
        mimetype="application/zip"
    )


@app.route("/balances", methods=["GET", "POST"])

@login_required
def balances():
    """View and edit monthly start/end bank balances."""

    user_id = session["user_id"]
    conn = get_db()
    year = date.today().year

    if request.method == "POST":
        month = request.form.get("month", "")
        field = request.form.get("field", "")
        value = request.form.get("value", "0")
        try:
            value = float(value) if value else 0
        except ValueError:
            value = 0
        if month and field in ("start_balance", "end_balance"):
            conn.execute(f"""
                INSERT INTO monthly_balances (month, {field})
                VALUES (?, ?)
                ON CONFLICT(month) DO UPDATE SET {field}=excluded.{field}, updated_at=datetime('now','localtime')
            """, (month, value))
            conn.commit()
            flash("✅ Balance updated!", "success")
        return redirect(url_for("balances"))

    # Load balances from DB
    rows = conn.execute("SELECT month, start_balance, end_balance FROM monthly_balances ORDER BY month").fetchall()
    balance_map = {r["month"]: {"start": r["start_balance"], "end": r["end_balance"]} for r in rows}

    conn.close()

    # Build month data with computed net change
    months_data = []
    monthly_totals = get_monthly_totals(user_id)

    for m in range(1, 13):
        month_key = f"{year}-{m:02d}"
        month_name = date(year, m, 1).strftime("%B")
        start_val = balance_map.get(month_key, {}).get("start", 0)
        end_val = balance_map.get(month_key, {}).get("end", 0)
        totals = monthly_totals.get(month_key, {"debit": 0, "credit": 0})
        net_change = totals["credit"] - totals["debit"]
        expected_end = start_val + net_change
        diff = end_val - expected_end if end_val else 0

        months_data.append({
            "key": month_key,
            "name": month_name,
            "start_balance": start_val,
            "end_balance": end_val,
            "total_debit": totals["debit"],
            "total_credit": totals["credit"],
            "net_change": net_change,
            "expected_end": round(expected_end, 2),
            "difference": round(diff, 2),
        })

    return render_template("balances.html", months=months_data, year=year)


def get_monthly_totals(user_id=None):
    """Get total debit/credit per month from transactions."""
    conn = get_db()
    rows = conn.execute("""\
        SELECT 
            substr(date, 1, 7) as month,
            SUM(CASE WHEN txn_type='debit' THEN amount ELSE 0 END) as total_debit,
            SUM(CASE WHEN txn_type='credit' THEN amount ELSE 0 END) as total_credit
        FROM transactions
        WHERE user_id = ?
        GROUP BY month
        ORDER BY month
    """, (user_id,)).fetchall()
    conn.close()
    return {r["month"]: {"debit": r["total_debit"], "credit": r["total_credit"]} for r in rows}


@app.route("/people")

@login_required
def people():
    """Track money lent to / borrowed from friends and family."""

    user_id = session["user_id"]
    display_name = session.get("display_name", "")
    conn = get_db()
    filter_person = request.args.get("person", "").strip()

    # Get all unique persons + their balances for this user
    people_query = """\
        SELECT 
            person,
            SUM(CASE WHEN txn_type='debit' THEN amount ELSE 0 END) as given,
            SUM(CASE WHEN txn_type='credit' THEN amount ELSE 0 END) as received,
            COUNT(*) as count
        FROM transactions 
        WHERE user_id = ? AND person != '' AND person IS NOT NULL AND person != ?
        GROUP BY person
        ORDER BY person
    """
    people_rows = conn.execute(people_query, (user_id, display_name)).fetchall()

    # Get transactions — filtered by person if selected
    if filter_person:
        recent = conn.execute("""\
            SELECT id, date, description, amount, category, card_id, txn_type, person, notes
            FROM transactions 
            WHERE user_id = ? AND person = ?
            ORDER BY date DESC, id DESC
            LIMIT 100
        """, (user_id, filter_person)).fetchall()
    else:
        recent = conn.execute("""\
            SELECT id, date, description, amount, category, card_id, txn_type, person, notes
            FROM transactions 
            WHERE user_id = ? AND person != '' AND person IS NOT NULL AND person != ?
            ORDER BY date DESC, id DESC
            LIMIT 50
        """, (user_id, display_name)).fetchall()

    conn.close()

    return render_template("people.html", people=people_rows, recent=recent,
                         filter_person=filter_person, all_people=load_people())




#@app.route("/sync")
#
#@login_required
#def manual_sync():
#    \"\"\"Sync Excel → Database (picks up new entries from Excel).\"\"\"
#
#    person = session.get(\"display_name\", \"\")
#    direction, count = smart_sync()
#
#    # Also import balances from Excel
#    try:
#        from excel_sync import _read_balances_from_excel, EXCEL_PATH as _excel_path
#        import openpyxl
#        if _excel_path.exists():
#            wb = openpyxl.load_workbook(str(_excel_path), data_only=True)
#            excel_balances = _read_balances_from_excel(wb)
#            wb.close()
#            conn = get_db()
#            for m_key, vals in excel_balances.items():
#                if vals[\"start\"] or vals[\"end\"]:
#                    conn.execute(\"\"\"
#                        INSERT INTO monthly_balances (month, start_balance, end_balance)
#                        VALUES (?, ?, ?)
#                        ON CONFLICT(month) DO UPDATE SET
#                            start_balance=excluded.start_balance,
#                            end_balance=excluded.end_balance,
#                            updated_at=datetime('now','localtime')
#                    \"\"\", (m_key, vals[\"start\"], vals[\"end\"]))
#            conn.commit()
#            conn.close()
#    except Exception as e:
#        print(f\"⚠️ Balance import error: {e}\")
#
#    if direction == \"excel_to_db\":
#        flash(f\"📥 Synced: Excel → Website ({count} transactions imported)\", \"info\")
#    elif direction == \"db_to_excel\":
#        flash(f\"📤 Synced: Website → Excel ({count} transactions exported)\", \"info\")
#    else:
#        flash(f\"✅ Already in sync\", \"success\")
#    return redirect(request.referrer or url_for(\"index\"))


# ─── All Transactions Page ─────────────────────────────────────────────────────
@app.route("/transactions")

@login_required
def all_transactions():
    """Paginated, sortable, filterable list of all transactions."""

    user_id = session["user_id"]
    page = int(request.args.get("page", 1))
    per_page = 50
    sort_by = request.args.get("sort", "date")
    sort_dir = request.args.get("dir", "desc")

    # Filters
    filter_cat = request.args.get("category", "")
    filter_card = request.args.get("card_id", "")
    filter_type = request.args.get("txn_type", "")
    filter_search = request.args.get("q", "").strip()
    filter_from = request.args.get("from", "")
    filter_to = request.args.get("to", "")

    conn = get_db()
    where_clauses = ["user_id = ?"]
    params = [user_id]

    if filter_cat:
        where_clauses.append("category = ?")
        params.append(filter_cat)
    if filter_card:
        where_clauses.append("card_id = ?")
        params.append(filter_card)
    if filter_type in ("debit", "credit"):
        where_clauses.append("txn_type = ?")
        params.append(filter_type)
    if filter_search:
        where_clauses.append("(description LIKE ? OR notes LIKE ?)")
        params.extend([f"%{filter_search}%", f"%{filter_search}%"])
    if filter_from:
        where_clauses.append("date >= ?")
        params.append(filter_from)
    if filter_to:
        where_clauses.append("date <= ?")
        params.append(filter_to)

    where_sql = " WHERE " + " AND ".join(where_clauses)

    allowed_sorts = {"date": "date", "amount": "amount", "description": "description", "id": "id"}
    sort_col = allowed_sorts.get(sort_by, "date")
    direction = "DESC" if sort_dir == "desc" else "ASC"
    order_sql = f" ORDER BY {sort_col} {direction}, id DESC"

    total = conn.execute(f"SELECT COUNT(*) FROM transactions{where_sql}", params).fetchone()[0]
    pages = max(1, (total + per_page - 1) // per_page)
    if page > pages:
        page = pages
    offset = (page - 1) * per_page

    rows = conn.execute(
        f"SELECT * FROM transactions{where_sql}{order_sql} LIMIT ? OFFSET ?",
        params + [per_page, offset]
    ).fetchall()
    conn.close()

    cards = get_user_cards(user_id)
    categories = load_categories()

    is_ajax = request.args.get("ajax") == "1"

    if is_ajax:
        table_html = render_template("_transactions_table.html",
            transactions=rows, page=page, pages=pages, total=total,
            sort_by=sort_by, sort_dir=sort_dir,
            filter_cat=filter_cat, filter_card=filter_card,
            filter_type=filter_type, filter_search=filter_search,
            filter_from=filter_from, filter_to=filter_to)
        pagination_html = render_template("_transactions_pagination.html",
            page=page, pages=pages, total=total,
            sort_by=sort_by, sort_dir=sort_dir,
            filter_cat=filter_cat, filter_card=filter_card,
            filter_type=filter_type, filter_search=filter_search,
            filter_from=filter_from, filter_to=filter_to)
        return jsonify(table_html=table_html, pagination_html=pagination_html)

    return render_template("all_transactions.html",
                         transactions=rows,
                         page=page, pages=pages, total=total,
                         sort_by=sort_by, sort_dir=sort_dir,
                         filter_cat=filter_cat, filter_card=filter_card,
                         filter_type=filter_type, filter_search=filter_search,
                         filter_from=filter_from, filter_to=filter_to,
                         cards=cards, categories=categories,
                         per_page=per_page)


# ─── Main ─────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    print("=" * 60)
    print("💰 FINANCE TRACKER — Daily Transaction Manager")
    print("=" * 60)
    print(f"📁 Database: {DB_PATH}")
    print(f"🔗 Open: http://127.0.0.1:5000")
    print("=" * 60)
    app.run(debug=True, host="0.0.0.0", port=5000)
