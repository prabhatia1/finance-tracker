"""Exercise account creation and dashboard queries against DATABASE_URL."""

import re
import sys
import uuid
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))


def main():
    from app import app, get_db

    username = "smoke_" + uuid.uuid4().hex[:12]
    client = app.test_client()
    response = client.get("/register")
    match = re.search(
        r'name="csrf_token" value="([^"]+)"',
        response.get_data(as_text=True),
    )
    if response.status_code != 200 or not match:
        raise RuntimeError("Registration page did not render a CSRF token.")

    try:
        response = client.post(
            "/register",
            data={
                "csrf_token": match.group(1),
                "username": username,
                "display_name": "Smoke Test",
                "security_word": "smoke-test",
                "password": "Temporary-Smoke-123",
                "confirm": "Temporary-Smoke-123",
            },
            follow_redirects=True,
        )
        body = response.get_data(as_text=True)
        if response.status_code != 200 or "Welcome, Smoke Test!" not in body:
            raise RuntimeError("Registration or authenticated dashboard request failed.")
        print("PostgreSQL registration and dashboard smoke test passed.")
    finally:
        conn = get_db()
        try:
            user = conn.execute("SELECT id FROM users WHERE username = ?", (username,)).fetchone()
            if user:
                user_id = user["id"]
                for table in ("import_staging", "transactions", "monthly_balances", "user_cards", "user_people"):
                    conn.execute(f"DELETE FROM {table} WHERE user_id = ?", (user_id,))
                conn.execute("DELETE FROM users WHERE id = ?", (user_id,))
                conn.commit()
        finally:
            conn.close()


if __name__ == "__main__":
    main()
