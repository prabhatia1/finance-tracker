# Copilot instructions for finance-tracker

## Project context
- This is a Flask + SQLite personal finance application.
- The primary workflow is centered around app.py and the templates directory.
- Data is user-scoped, so preserve multi-user isolation.

## Coding guidance
- Follow the existing style in the repository: simple Flask routes, SQLite helpers, and Jinja templates.
- Keep changes focused and compatible with the existing database schema.
- Avoid introducing new dependencies unless necessary.
- If you update import, parsing, reporting, or backup functionality, verify the relevant behavior locally.

## Safety and privacy
- Do not expose or print secrets, password hashes, or backup archives.
- Treat finance data as sensitive and avoid logging full contents of transactions or user records.
