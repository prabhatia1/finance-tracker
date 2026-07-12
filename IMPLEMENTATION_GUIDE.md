# 🚀 Finance Tracker - Implementation Guide

## Phase 1: Critical Security Fixes (MUST DO FIRST)

### Step 1: Update Dependencies

```bash
# Update your requirements.txt with new packages
pip install -r requirements.txt
```

**File: `requirements.txt`**
```txt
flask>=3.0.0
openpyxl>=3.1.0
werkzeug>=2.3.0
Flask-WTF>=1.1.0
Flask-Limiter>=3.5.0
python-dotenv>=1.0.0
```

---

### Step 2: Replace Password Hashing

**OLD CODE (INSECURE):**
```python
import hashlib

def hash_password(password):
    return hashlib.sha256(password.encode()).hexdigest()

def check_password(hash, password):
    return hash == hashlib.sha256(password.encode()).hexdigest()
```

**NEW CODE (SECURE):**
```python
from werkzeug.security import generate_password_hash, check_password_hash

def hash_password(password):
    return generate_password_hash(password, method='pbkdf2:sha256')

def check_password(password_hash, password):
    return check_password_hash(password_hash, password)
```

**UPDATE in app.py around line ~600 (register route):**

```python
@app.route("/register", methods=["GET", "POST"])
@limiter.limit("3 per hour")  # ADD RATE LIMIT
def register():
    if request.method == "POST":
        username = sanitize(request.form.get("username", "")).strip()
        password = request.form.get("password", "")
        password_confirm = request.form.get("password_confirm", "")
        display_name = sanitize(request.form.get("display_name", "")).strip()
        
        # ✅ NEW: Validate username format
        if not username or len(username) < 3 or len(username) > 20:
            flash("Username must be 3-20 characters", "error")
            return render_template("register.html")
        
        # ✅ NEW: Check passwords match
        if password != password_confirm:
            flash("Passwords do not match", "error")
            return render_template("register.html")
        
        # ✅ NEW: Validate password strength
        if len(password) < 8:
            flash("Password must be at least 8 characters", "error")
            return render_template("register.html")
        
        if not any(c.isupper() for c in password):
            flash("Password must contain uppercase letter", "error")
            return render_template("register.html")
        
        if not any(c.isdigit() for c in password):
            flash("Password must contain digit", "error")
            return render_template("register.html")
        
        conn = get_db()
        try:
            # ✅ CHANGED: Use new hashing method
            password_hash = generate_password_hash(password, method='pbkdf2:sha256')
            
            conn.execute(
                "INSERT INTO users (username, password_hash, display_name) VALUES (?,?,?)",
                (username, password_hash, display_name)
            )
            conn.commit()
            flash("Account created! Please log in.", "success")
            return redirect(url_for("login"))
        except sqlite3.IntegrityError:
            flash("Username already exists", "error")
        finally:
            conn.close()
    
    return render_template("register.html")
```

**UPDATE in app.py login route:**

```python
@app.route("/login", methods=["GET", "POST"])
@limiter.limit("5 per minute")  # ADD RATE LIMIT
def login():
    if request.method == "POST":
        username = sanitize(request.form.get("username", "")).strip()
        password = request.form.get("password", "")
        
        if not username or not password:
            flash("Username and password required", "error")
            return render_template("login.html")
        
        conn = get_db()
        user = conn.execute(
            "SELECT id, display_name, password_hash FROM users WHERE username = ?",
            (username,)
        ).fetchone()
        conn.close()
        
        # ✅ CHANGED: Use new verification method
        if user and check_password_hash(user["password_hash"], password):
            session["user_id"] = user["id"]
            session["display_name"] = user["display_name"]
            session.permanent = True  # ✅ NEW: Make session persistent
            flash("Logged in successfully!", "success")
            return redirect(url_for("dashboard"))
        else:
            flash("Invalid username or password", "error")
    
    return render_template("login.html")
```

---

### Step 3: Add CSRF Protection

**At the top of app.py:**

```python
from flask_wtf.csrf import CSRFProtect

# Initialize CSRF protection
csrf = CSRFProtect(app)

# Configure CSRF
app.config['WTF_CSRF_ENABLED'] = True
app.config['WTF_CSRF_TIME_LIMIT'] = None
```

**In every HTML form template (add this line):**

```html
<form method="POST">
    {{ csrf_token() }}  <!-- ← ADD THIS LINE -->
    <!-- form fields... -->
</form>
```

**Example:**
```html
<form method="POST" action="{{ url_for('add_transaction') }}" class="needs-validation" novalidate>
    {{ csrf_token() }}
    
    <div class="mb-3">
        <label for="date" class="form-label">Date</label>
        <input type="date" class="form-control" id="date" name="date" required>
    </div>
    
    <button type="submit" class="btn btn-primary">Add Transaction</button>
</form>
```

---

### Step 4: Add Rate Limiting

**At the top of app.py:**

```python
from flask_limiter import Limiter
from flask_limiter.util import get_remote_address

limiter = Limiter(
    app=app,
    key_func=get_remote_address,
    default_limits=["200 per day", "50 per hour"],
    storage_uri="memory://"  # Use Redis in production
)

app.config['RATELIMIT_STORAGE_URL'] = 'memory://'
```

**Apply to auth routes:**

```python
@app.route("/login", methods=["GET", "POST"])
@limiter.limit("5 per minute")  # ← ADD THIS
def login():
    # ...

@app.route("/register", methods=["GET", "POST"])
@limiter.limit("3 per hour")  # ← ADD THIS
def register():
    # ...

@app.route("/forgot-password", methods=["GET", "POST"])
@limiter.limit("3 per hour")  # ← ADD THIS
def forgot_password():
    # ...
```

---

### Step 5: Add Security Headers

**At the top of app.py, after app initialization:**

```python
@app.after_request
def set_security_headers(response):
    """Add security headers to all responses."""
    response.headers['X-Frame-Options'] = 'SAMEORIGIN'
    response.headers['X-Content-Type-Options'] = 'nosniff'
    response.headers['X-XSS-Protection'] = '1; mode=block'
    response.headers['Referrer-Policy'] = 'strict-origin-when-cross-origin'
    response.headers['Permissions-Policy'] = 'geolocation=(), microphone=(), camera=()'
    
    # Content Security Policy
    response.headers['Content-Security-Policy'] = (
        "default-src 'self'; "
        "script-src 'self' https://cdn.jsdelivr.net; "
        "style-src 'self' 'unsafe-inline' https://cdn.jsdelivr.net https://fonts.googleapis.com; "
        "font-src 'self' https://fonts.gstatic.com https://cdn.jsdelivr.net; "
        "img-src 'self' data: https:; "
        "connect-src 'self';"
    )
    return response
```

---

### Step 6: Session Security Configuration

**At the top of app.py, after app initialization:**

```python
from datetime import timedelta

# Session security
app.config['SESSION_COOKIE_SECURE'] = True       # HTTPS only
app.config['SESSION_COOKIE_HTTPONLY'] = True     # No JavaScript access
app.config['SESSION_COOKIE_SAMESITE'] = 'Lax'    # CSRF protection
app.config['PERMANENT_SESSION_LIFETIME'] = timedelta(hours=24)
app.config['SESSION_REFRESH_EACH_REQUEST'] = True

@app.before_request
def make_session_permanent():
    session.permanent = True
```

---

### Step 7: Migrate Existing Passwords

**Create file: `migrate_passwords.py`**

```python
"""
⚠️  IMPORTANT: Run this script ONCE to migrate existing password hashes
to the new secure werkzeug format.

Run: python migrate_passwords.py
"""

import sqlite3
from werkzeug.security import generate_password_hash

def migrate_passwords():
    conn = sqlite3.connect('finance.db')
    
    print("Starting password migration...")
    
    users = conn.execute("SELECT id, username, password_hash FROM users").fetchall()
    
    migrated = 0
    already_new = 0
    
    for user_id, username, old_hash in users:
        # Check if already using new format
        if old_hash and old_hash.startswith('pbkdf2:'):
            already_new += 1
            continue
        
        # For existing users with old hashes, force password reset
        # Create a temporary hash they can reset
        temp_password = f"TempPass{user_id}!@#"
        new_hash = generate_password_hash(temp_password, method='pbkdf2:sha256')
        
        conn.execute(
            "UPDATE users SET password_hash = ? WHERE id = ?",
            (new_hash, user_id)
        )
        
        print(f"  ✓ Migrated user: {username} (ID: {user_id})")
        migrated += 1
    
    conn.commit()
    conn.close()
    
    print(f"\n✅ Migration complete!")
    print(f"   - Migrated: {migrated} users")
    print(f"   - Already new format: {already_new} users")
    print(f"   - Total: {migrated + already_new} users")
    print("\n⚠️  Users will need to reset their passwords with: Forgot Password")

if __name__ == '__main__':
    migrate_passwords()
```

**Run the migration:**
```bash
python migrate_passwords.py
```

---

## Phase 2: UI/UX Improvements

### Step 1: Replace base.html

Copy the improved `IMPROVED_BASE.html` to your `templates/base.html`:

```bash
cp IMPROVED_BASE.html templates/base.html
```

**Key improvements in new base.html:**
- ✅ Better color palette with WCAG contrast compliance
- ✅ Improved mobile responsiveness
- ✅ Enhanced form styling with validation feedback
- ✅ Better button styling with hover effects
- ✅ Improved table styling
- ✅ Smooth animations and transitions
- ✅ Safe area insets for notched devices
- ✅ Auto-dismissing alert messages
- ✅ Improved accessibility

### Step 2: Add Date Picker

**Update add.html transaction form:**

```html
<form method="POST" class="needs-validation" novalidate>
    {{ csrf_token() }}
    
    <div class="mb-3">
        <label for="date" class="form-label">
            <i class="bi bi-calendar-event"></i> Date
        </label>
        <input 
            type="date" 
            class="form-control" 
            id="date" 
            name="date" 
            value="{{ today }}"
            required
        >
        <div class="invalid-feedback">
            Please select a valid date (not in future)
        </div>
    </div>
    
    <div class="mb-3">
        <label for="amount" class="form-label">
            <i class="bi bi-cash-coin"></i> Amount (₹)
        </label>
        <input 
            type="number" 
            class="form-control" 
            id="amount" 
            name="amount" 
            step="0.01"
            min="0.01"
            max="10000000"
            placeholder="0.00"
            required
        >
        <div class="invalid-feedback">
            Please enter a valid amount (0.01 - 10,000,000)
        </div>
    </div>
    
    <!-- ... more fields ... -->
    
    <button type="submit" class="btn btn-primary btn-lg w-100">
        <i class="bi bi-plus-circle"></i> Add Transaction
    </button>
</form>

<script>
    // Set date to today
    document.getElementById('date').valueAsDate = new Date();
    
    // Validate date is not in future
    document.getElementById('date').addEventListener('change', function() {
        const selectedDate = new Date(this.value);
        const today = new Date();
        today.setHours(0,0,0,0);
        
        if (selectedDate > today) {
            this.classList.add('is-invalid');
            this.setCustomValidity('Date cannot be in the future');
        } else {
            this.classList.remove('is-invalid');
            this.setCustomValidity('');
        }
    });
</script>
```

### Step 3: Add Input Validation Function

**Create file: `static/js/form-validation.js`**

```javascript
/**
 * Form Validation Helper
 * Provides real-time validation feedback
 */

class FormValidator {
    static validateAmount(value) {
        const amount = parseFloat(value);
        return !isNaN(amount) && amount > 0 && amount <= 10000000;
    }
    
    static validateDate(value) {
        const date = new Date(value);
        const today = new Date();
        today.setHours(0, 0, 0, 0);
        return date <= today;
    }
    
    static validateEmail(value) {
        return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(value);
    }
    
    static validatePassword(value) {
        return value.length >= 8 &&
               /[A-Z]/.test(value) &&
               /[a-z]/.test(value) &&
               /\d/.test(value) &&
               /[!@#$%^&*]/.test(value);
    }
    
    static getPasswordStrength(password) {
        let strength = 0;
        if (password.length >= 8) strength++;
        if (/[A-Z]/.test(password)) strength++;
        if (/[a-z]/.test(password)) strength++;
        if (/\d/.test(password)) strength++;
        if (/[!@#$%^&*]/.test(password)) strength++;
        return ['Weak', 'Fair', 'Good', 'Strong', 'Very Strong'][strength] || 'Weak';
    }
    
    static setupValidation() {
        // Auto-validate on input
        document.querySelectorAll('[data-validate]').forEach(field => {
            field.addEventListener('blur', function() {
                const type = this.dataset.validate;
                const isValid = FormValidator['validate' + 
                    type.charAt(0).toUpperCase() + type.slice(1)](this.value);
                
                if (isValid) {
                    this.classList.remove('is-invalid');
                    this.classList.add('is-valid');
                } else {
                    this.classList.remove('is-valid');
                    this.classList.add('is-invalid');
                }
            });
        });
    }
}

// Initialize on page load
document.addEventListener('DOMContentLoaded', FormValidator.setupValidation);
```

**Use in templates:**
```html
<script src="{{ url_for('static', filename='js/form-validation.js') }}"></script>

<input type="email" data-validate="email" class="form-control">
<input type="number" data-validate="amount" class="form-control">
<input type="date" data-validate="date" class="form-control">
<input type="password" data-validate="password" class="form-control">
```

---

### Step 4: Add Search & Filter

**Update all_transactions.html:**

```html
<div class="card mb-3">
    <div class="card-body">
        <form method="GET" class="row g-2" id="filter-form">
            <div class="col-12 col-md-3">
                <input 
                    type="date" 
                    name="date_from" 
                    class="form-control" 
                    placeholder="From date"
                    value="{{ request.args.get('date_from', '') }}"
                >
            </div>
            <div class="col-12 col-md-3">
                <input 
                    type="date" 
                    name="date_to" 
                    class="form-control" 
                    placeholder="To date"
                    value="{{ request.args.get('date_to', '') }}"
                >
            </div>
            <div class="col-12 col-md-3">
                <select name="category" class="form-select">
                    <option value="">All Categories</option>
                    {% for cat in categories %}
                    <option value="{{ cat.name }}" 
                        {% if request.args.get('category') == cat.name %}selected{% endif %}>
                        {{ cat.name }}
                    </option>
                    {% endfor %}
                </select>
            </div>
            <div class="col-12 col-md-3">
                <button type="submit" class="btn btn-primary w-100">
                    <i class="bi bi-search"></i> Filter
                </button>
            </div>
        </form>
    </div>
</div>
```

**Backend route update:**

```python
@app.route("/transactions", methods=["GET"])
@login_required
def all_transactions():
    user_id = session["user_id"]
    
    # Get filters
    date_from = request.args.get('date_from', '')
    date_to = request.args.get('date_to', '')
    category = request.args.get('category', '')
    search = request.args.get('search', '').strip()
    
    conn = get_db()
    query = "SELECT * FROM transactions WHERE user_id = ? AND txn_type = 'debit'"
    params = [user_id]
    
    # Apply filters
    if date_from:
        query += " AND date >= ?"
        params.append(date_from)
    
    if date_to:
        query += " AND date <= ?"
        params.append(date_to)
    
    if category:
        query += " AND category = ?"
        params.append(category)
    
    if search:
        query += " AND (description LIKE ? OR notes LIKE ?)"
        params.extend([f'%{search}%', f'%{search}%'])
    
    query += " ORDER BY date DESC LIMIT 100"
    
    transactions = conn.execute(query, params).fetchall()
    categories = [dict(c) for c in conn.execute(
        "SELECT DISTINCT category FROM transactions WHERE user_id = ?",
        (user_id,)
    ).fetchall()]
    conn.close()
    
    return render_template(
        'all_transactions.html',
        transactions=[dict(t) for t in transactions],
        categories=categories
    )
```

---

## Testing Checklist

### Security Testing
- [ ] Try old passwords (should fail)
- [ ] Test rate limiting on login (5 attempts in 1 minute)
- [ ] Test CSRF protection (remove token, should fail)
- [ ] Test XSS prevention (try HTML in description)
- [ ] Test SQL injection prevention (try SQL in search)
- [ ] Verify security headers in browser DevTools

### UI/UX Testing
- [ ] Test on mobile devices
- [ ] Test form validation (required fields)
- [ ] Test date picker on different browsers
- [ ] Test filter functionality
- [ ] Test accessibility (Tab key navigation)
- [ ] Test responsive design (resize browser)

### Functional Testing
- [ ] Login/Register with new passwords
- [ ] Add transaction with date picker
- [ ] Edit existing transaction
- [ ] Filter transactions by date range
- [ ] Search transactions by description
- [ ] Generate reports
- [ ] Export to Excel

---

## Deployment Checklist

### Before Going Live
- [ ] Run password migration script
- [ ] Test on staging environment
- [ ] Backup existing database
- [ ] Update `.gitignore` (add logs, secrets)
- [ ] Set secure `SECRET_KEY` in environment variable
- [ ] Enable HTTPS in production
- [ ] Configure proper session storage (Redis)

### Production Configuration

**Create file: `.env`**
```
FLASK_ENV=production
SECRET_KEY=your-secure-random-key-here
DATABASE_URL=sqlite:///finance.db
WTF_CSRF_ENABLED=true
SESSION_COOKIE_SECURE=true
RATELIMIT_STORAGE_URL=redis://localhost:6379
```

**In app.py:**
```python
from dotenv import load_dotenv
import os

load_dotenv()

app.config['SECRET_KEY'] = os.getenv('SECRET_KEY')
app.config['ENV'] = os.getenv('FLASK_ENV', 'production')
```

---

## Getting Help

- **Security Questions**: See `SECURITY_IMPROVEMENTS.py`
- **UI Improvements**: See `IMPROVED_BASE.html`
- **Full Analysis**: See `IMPROVEMENTS.md`
- **GitHub Issues**: Report bugs on GitHub

---

## Next Steps

1. ✅ Complete Phase 1 (Security) - PRIORITY
2. ✅ Test thoroughly on staging
3. ✅ Deploy to production
4. ⏳ Phase 2 (UI/UX) - Schedule for next sprint
5. ⏳ Phase 3 (Code refactoring) - Plan for future

**Need help? Check the implementation files or create an issue on GitHub!**
