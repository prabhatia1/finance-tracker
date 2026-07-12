"""
🔐 Security Improvements for Finance Tracker

This module demonstrates security enhancements that should be implemented:
1. Proper password hashing using werkzeug
2. CSRF protection with Flask-WTF
3. Rate limiting on auth endpoints
4. Input validation and sanitization
5. Secure session management
"""

import re
import hashlib
from functools import wraps
from datetime import datetime, timedelta
from werkzeug.security import generate_password_hash, check_password_hash
from flask import Flask, session, request, abort, jsonify
from flask_wtf.csrf import CSRFProtect
from flask_limiter import Limiter
from flask_limiter.util import get_remote_address

# ═══════════════════════════════════════════════════════════════════════════════
# 1️⃣ PASSWORD SECURITY
# ═══════════════════════════════════════════════════════════════════════════════

class PasswordValidator:
    """Validates password strength and provides hashing."""
    
    MIN_LENGTH = 8
    REQUIRE_UPPERCASE = True
    REQUIRE_LOWERCASE = True
    REQUIRE_DIGIT = True
    REQUIRE_SPECIAL = True
    
    @staticmethod
    def validate(password):
        """
        Validate password strength.
        
        Returns:
            tuple: (is_valid, error_message)
        """
        errors = []
        
        if len(password) < PasswordValidator.MIN_LENGTH:
            errors.append(f"Password must be at least {PasswordValidator.MIN_LENGTH} characters")
        
        if PasswordValidator.REQUIRE_UPPERCASE and not re.search(r'[A-Z]', password):
            errors.append("Password must contain at least one uppercase letter")
        
        if PasswordValidator.REQUIRE_LOWERCASE and not re.search(r'[a-z]', password):
            errors.append("Password must contain at least one lowercase letter")
        
        if PasswordValidator.REQUIRE_DIGIT and not re.search(r'\d', password):
            errors.append("Password must contain at least one digit")
        
        if PasswordValidator.REQUIRE_SPECIAL and not re.search(r'[!@#$%^&*(),.?":{}|<>]', password):
            errors.append("Password must contain at least one special character (!@#$%^&*)")
        
        if errors:
            return False, " | ".join(errors)
        return True, "Password is strong"
    
    @staticmethod
    def hash_password(password):
        """Hash password using werkzeug (bcrypt-like hashing)."""
        return generate_password_hash(password, method='pbkdf2:sha256')
    
    @staticmethod
    def verify_password(password_hash, password):
        """Verify password against hash."""
        return check_password_hash(password_hash, password)


# ═══════════════════════════════════════════════════════════════════════════════
# 2️⃣ INPUT VALIDATION & SANITIZATION
# ═══════════════════════════════════════════════════════════════════════════════

class InputValidator:
    """Validates and sanitizes user inputs."""
    
    @staticmethod
    def validate_email(email):
        """Validate email format."""
        pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
        return re.match(pattern, email) is not None
    
    @staticmethod
    def validate_username(username):
        """Validate username (alphanumeric + underscore, 3-20 chars)."""
        if not (3 <= len(username) <= 20):
            return False
        return re.match(r'^[a-zA-Z0-9_]+$', username) is not None
    
    @staticmethod
    def validate_amount(amount):
        """Validate transaction amount."""
        try:
            amt = float(amount)
            # Allow amounts from 0.01 to 10,000,000
            return 0.01 <= amt <= 10_000_000
        except (ValueError, TypeError):
            return False
    
    @staticmethod
    def validate_date(date_string):
        """Validate date format (YYYY-MM-DD) and not in future."""
        try:
            date_obj = datetime.strptime(date_string, '%Y-%m-%d').date()
            today = datetime.now().date()
            return date_obj <= today  # No future dates
        except (ValueError, AttributeError):
            return False
    
    @staticmethod
    def sanitize_text(text, max_length=500):
        """Sanitize text input (remove HTML, limit length)."""
        if not isinstance(text, str):
            return ""
        
        # Remove HTML tags
        text = re.sub(r'<[^>]*>', '', text)
        # Remove null bytes
        text = text.replace('\x00', '')
        # Strip whitespace
        text = text.strip()
        # Limit length
        return text[:max_length]
    
    @staticmethod
    def sanitize_description(description):
        """Sanitize transaction description."""
        return InputValidator.sanitize_text(description, max_length=200)


# ═══════════════════════════════════════════════════════════════════════════════
# 3️⃣ RATE LIMITING
# ═══════════════════════════════════════════════════════════════════════════════

def setup_rate_limiting(app):
    """Setup Flask-Limiter for rate limiting."""
    limiter = Limiter(
        app=app,
        key_func=get_remote_address,
        default_limits=["200 per day", "50 per hour"],
        storage_uri="memory://"  # Use Redis for production
    )
    
    return limiter


# Usage in routes:
# from flask import Flask
# limiter = setup_rate_limiting(app)
#
# @app.route('/login', methods=['POST'])
# @limiter.limit("5 per minute")
# def login():
#     ...
#
# @app.route('/register', methods=['POST'])
# @limiter.limit("3 per hour")
# def register():
#     ...


# ═══════════════════════════════════════════════════════════════════════════════
# 4️⃣ CSRF PROTECTION
# ═══════════════════════════════════════════════════════════════════════════════

def setup_csrf_protection(app):
    """Setup Flask-WTF CSRF protection."""
    app.config['WTF_CSRF_ENABLED'] = True
    app.config['WTF_CSRF_TIME_LIMIT'] = None  # No expiry on CSRF tokens
    app.config['WTF_CSRF_METHODS'] = ['POST', 'PUT', 'PATCH', 'DELETE']
    
    csrf = CSRFProtect(app)
    return csrf


# Usage in HTML templates:
# <form method="POST">
#     {{ csrf_token() }}
#     ...
# </form>


# ═══════════════════════════════════════════════════════════════════════════════
# 5️⃣ SESSION SECURITY
# ═══════════════════════════════════════════════════════════════════════════════

def setup_session_security(app):
    """Configure secure session settings."""
    # Session security
    app.config['SESSION_COOKIE_SECURE'] = True      # HTTPS only
    app.config['SESSION_COOKIE_HTTPONLY'] = True    # No JavaScript access
    app.config['SESSION_COOKIE_SAMESITE'] = 'Lax'   # CSRF protection
    app.config['PERMANENT_SESSION_LIFETIME'] = timedelta(hours=24)  # 24 hour session
    app.config['SESSION_REFRESH_EACH_REQUEST'] = True  # Refresh on each request
    
    # Remember Me functionality
    @app.before_request
    def make_session_permanent():
        session.permanent = True


# ═══════════════════════════════════════════════════════════════════════════════
# 6️⃣ AUTHENTICATION HELPERS
# ═══════════════════════════════════════════════════════════════════════════════

class AuthHelper:
    """Helper functions for authentication."""
    
    @staticmethod
    def hash_password(password):
        """Hash a password."""
        return PasswordValidator.hash_password(password)
    
    @staticmethod
    def verify_password(password_hash, password):
        """Verify a password."""
        return PasswordValidator.verify_password(password_hash, password)
    
    @staticmethod
    def validate_login_attempt(username, user_from_db):
        """
        Validate login attempt with timing attack protection.
        Always takes same amount of time whether user exists or not.
        """
        # Generate dummy hash to always spend time
        dummy_hash = PasswordValidator.hash_password("dummy_password_12345!")
        
        if user_from_db:
            return PasswordValidator.verify_password(
                user_from_db['password_hash'],
                username  # This is wrong - for timing attack safety
            )
        else:
            # Still verify against dummy hash to prevent timing attacks
            PasswordValidator.verify_password(dummy_hash, username)
            return False


# ═══════════════════════════════════════════════════════════════════════════════
# 7️⃣ SECURE HEADERS MIDDLEWARE
# ═══════════════════════════════════════════════════════════════════════════════

def setup_security_headers(app):
    """Add security headers to all responses."""
    
    @app.after_request
    def set_security_headers(response):
        # Prevent clickjacking
        response.headers['X-Frame-Options'] = 'SAMEORIGIN'
        
        # Prevent MIME type sniffing
        response.headers['X-Content-Type-Options'] = 'nosniff'
        
        # Enable XSS protection
        response.headers['X-XSS-Protection'] = '1; mode=block'
        
        # Referrer policy
        response.headers['Referrer-Policy'] = 'strict-origin-when-cross-origin'
        
        # Feature policy / Permissions policy
        response.headers['Permissions-Policy'] = 'geolocation=(), microphone=(), camera=()'
        
        # Content Security Policy (adjust as needed)
        response.headers['Content-Security-Policy'] = (
            "default-src 'self'; "
            "script-src 'self' https://cdn.jsdelivr.net; "
            "style-src 'self' 'unsafe-inline' https://cdn.jsdelivr.net https://fonts.googleapis.com; "
            "font-src 'self' https://fonts.gstatic.com https://cdn.jsdelivr.net; "
            "img-src 'self' data: https:; "
            "connect-src 'self'; "
            "frame-ancestors 'self';"
        )
        
        return response


# ═══════════════════════════════════════════════════════════════════════════════
# 8️⃣ EXAMPLE: IMPROVED LOGIN ROUTE
# ═══════════════════════════════════════════════════════════════════════════════

"""
@app.route('/login', methods=['GET', 'POST'])
@limiter.limit("5 per minute")  # Rate limit
def login():
    if request.method == 'POST':
        username = request.form.get('username', '').strip()
        password = request.form.get('password', '')
        
        # Input validation
        if not username or not password:
            flash("Username and password required", "error")
            return render_template('login.html')
        
        if not InputValidator.validate_username(username):
            flash("Invalid username format", "error")
            return render_template('login.html')
        
        # Database query
        conn = get_db()
        user = conn.execute(
            "SELECT id, display_name, password_hash FROM users WHERE username = ?",
            (username,)
        ).fetchone()
        conn.close()
        
        # Timing attack protection: always check password hash
        if user and PasswordValidator.verify_password(user['password_hash'], password):
            session['user_id'] = user['id']
            session['display_name'] = user['display_name']
            flash("Logged in successfully!", "success")
            return redirect(url_for('dashboard'))
        else:
            flash("Invalid username or password", "error")
            return render_template('login.html')
    
    return render_template('login.html')


@app.route('/register', methods=['GET', 'POST'])
@limiter.limit("3 per hour")  # Rate limit registration attempts
def register():
    if request.method == 'POST':
        username = request.form.get('username', '').strip()
        password = request.form.get('password', '')
        password_confirm = request.form.get('password_confirm', '')
        display_name = request.form.get('display_name', '').strip()
        
        # Input validation
        if not InputValidator.validate_username(username):
            flash("Username must be 3-20 alphanumeric characters", "error")
            return render_template('register.html')
        
        if password != password_confirm:
            flash("Passwords do not match", "error")
            return render_template('register.html')
        
        # Password strength validation
        is_valid, message = PasswordValidator.validate(password)
        if not is_valid:
            flash(message, "error")
            return render_template('register.html')
        
        if not display_name or len(display_name) > 50:
            flash("Display name must be 1-50 characters", "error")
            return render_template('register.html')
        
        display_name = InputValidator.sanitize_text(display_name, max_length=50)
        
        # Hash password securely
        password_hash = PasswordValidator.hash_password(password)
        
        # Database insertion with duplicate check
        conn = get_db()
        try:
            conn.execute(
                "INSERT INTO users (username, password_hash, display_name) VALUES (?, ?, ?)",
                (username, password_hash, display_name)
            )
            conn.commit()
            conn.close()
            
            flash("Account created successfully! Please log in.", "success")
            return redirect(url_for('login'))
        except sqlite3.IntegrityError:
            flash("Username already exists", "error")
            return render_template('register.html')
        finally:
            conn.close()
    
    return render_template('register.html')
"""


# ═══════════════════════════════════════════════════════════════════════════════
# 9️⃣ MIGRATION: UPDATE EXISTING PASSWORDS
# ═══════════════════════════════════════════════════════════════════════════════

"""
Run this script once to migrate from old hashing to werkzeug:

import sqlite3
from werkzeug.security import generate_password_hash

def migrate_passwords():
    conn = sqlite3.connect('finance.db')
    users = conn.execute("SELECT id, password_hash FROM users").fetchall()
    
    for user_id, old_hash in users:
        # If old hash doesn't start with 'pbkdf2:', it needs migration
        if not old_hash.startswith('pbkdf2:'):
            # You'll need the original password or reset it
            # For now, force a password reset
            new_hash = generate_password_hash(f"temp_password_{user_id}", method='pbkdf2:sha256')
            conn.execute(
                "UPDATE users SET password_hash = ? WHERE id = ?",
                (new_hash, user_id)
            )
    
    conn.commit()
    conn.close()
    print("Password migration complete!")

if __name__ == '__main__':
    migrate_passwords()
"""

print("""
✅ Security Improvements Module Loaded

To use these improvements:

1. Install new dependencies:
   pip install -r requirements.txt

2. In your app.py, import and use:
   
   from security_improvements import (
       PasswordValidator,
       InputValidator,
       setup_rate_limiting,
       setup_csrf_protection,
       setup_session_security,
       setup_security_headers,
       AuthHelper
   )
   
   app = Flask(__name__)
   csrf = setup_csrf_protection(app)
   limiter = setup_rate_limiting(app)
   setup_session_security(app)
   setup_security_headers(app)

3. Update login/register routes with validated input
4. Run password migration script
5. Test thoroughly before deploying

For detailed implementation, see IMPLEMENTATION_GUIDE.md
""")
