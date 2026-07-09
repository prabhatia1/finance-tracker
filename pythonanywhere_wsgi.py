"""
PythonAnywhere WSGI file.

HOW TO SET UP:
1. Log in to https://www.pythonanywhere.com/
2. Go to Web → Add new web app → Manual config → Python 3.10
3. Open the WSGI configuration file
4. Replace its contents with this file
5. Go to Web → Virtualenv → create a virtualenv
6. Open a Bash console and run:
     cd ~/finance-tracker
     pip install -r requirements.txt
7. Reload the web app
"""
import sys
import os

# ── Path to your project (adjust if cloned elsewhere) ──────────────────────
PROJECT_DIR = os.path.dirname(os.path.abspath(__file__))
if PROJECT_DIR not in sys.path:
    sys.path.insert(0, PROJECT_DIR)

# ── Start the Flask app ────────────────────────────────────────────────────
os.environ["FLASK_ENV"] = "production"
from app import app as application  # noqa: E402
