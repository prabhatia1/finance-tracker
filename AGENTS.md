# AGENTS.md

## Project overview
- This repository contains a Flask-based personal finance tracker with SQLite-backed storage.
- The main application entry point is app.py.
- User data is stored in finance.db and related JSON files such as cards.json, categories.json, and people.json.

## Working conventions
- Preserve the existing multi-user model. Any database change should keep user-specific data isolated by user_id.
- Prefer minimal, targeted changes that match the existing Flask route and template patterns.
- Keep the app safe for personal use: do not expose secrets, credentials, or raw database contents in responses or logs.
- Avoid changing password handling, backup behavior, or security-related flows without explicit instruction.
- When adding features, update the relevant templates and README notes if behavior changes.

## Verification expectations
- Before claiming success, run the relevant checks. For Python changes, verify with:
  - python -m compileall app.py backup_restore.py cashback.py create_monthly_excel.py excel_sync.py pdf_parser.py
- If a change touches the web app, run the app locally when practical and confirm the affected route behaves correctly.

## Repository-specific guidance
- The app uses Flask, SQLite, Jinja templates, and optional CSV/PDF import helpers.
- Keep parsing and import logic robust and backward-compatible.
- Be careful with backup and restore scripts because they can handle sensitive financial data.
