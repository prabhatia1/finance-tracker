"""Regression checks using disposable data, never the local finance database."""

import json
import io
import os
import sqlite3
import tempfile
import unittest
import zipfile
from datetime import date
from pathlib import Path
from unittest.mock import patch

from database import HybridRow


_test_directory = tempfile.TemporaryDirectory(prefix="finance-regressions-")
_test_root = Path(_test_directory.name)
_environment = {
    "SQLITE_DB_PATH": str(_test_root / "startup.db"),
    "SECRET_KEY": "isolated-regression-tests",
}
with patch.dict(os.environ, _environment):
    for _key in (
        "DATABASE_URL", "POSTGRES_URL", "POSTGRES_URL_NON_POOLING",
        "DATABASE_URL_UNPOOLED", "VERCEL",
    ):
        os.environ.pop(_key, None)
    import app as tracker


def tearDownModule():
    _test_directory.cleanup()


class AppRegressionTests(unittest.TestCase):
    def setUp(self):
        self.directory = tempfile.TemporaryDirectory(dir=_test_root)
        self.addCleanup(self.directory.cleanup)
        root = Path(self.directory.name)
        self.defaults = [{"name": "Other", "keywords": [], "type": "other"}]
        categories_path = root / "categories.json"
        categories_path.write_text(json.dumps({"categories": self.defaults}), encoding="utf-8")
        for name, value in (
            ("DB_PATH", root / "finance.db"),
            ("CATEGORIES_PATH", categories_path),
        ):
            patcher = patch.object(tracker, name, value)
            patcher.start()
            self.addCleanup(patcher.stop)
        tracker.app.config.update(TESTING=True, WTF_CSRF_ENABLED=False)
        if tracker._HAS_FLASK_LIMITER:
            patcher = patch.object(tracker.limiter, "enabled", False)
            patcher.start()
            self.addCleanup(patcher.stop)
        tracker.init_db()
        with tracker.db_conn(commit=True) as conn:
            for user_id in (1, 2):
                conn.execute(
                    "INSERT INTO users (id, username, password_hash, display_name) VALUES (?, ?, ?, ?)",
                    (user_id, f"regression_{user_id}", "unused", f"User {user_id}"),
                )
        self.client = self.client_for(1)

    def client_for(self, user_id):
        client = tracker.app.test_client()
        with client.session_transaction() as session:
            session.update(user_id=user_id, username=f"regression_{user_id}", display_name=f"User {user_id}")
        return client

    def stage(self):
        return tracker.stage_import(1, [{
            "date": date.today().isoformat(), "description": "Coffee", "amount": 12.5,
            "category": "Other", "card_id": "other", "txn_type": "debit",
        }])

    def test_dashboard_handles_postgres_aggregate_rows(self):
        with tracker.db_conn(commit=True) as conn:
            for txn_type, amount in (("debit", 12.5), ("credit", 4)):
                conn.execute(
                    "INSERT INTO transactions (date, description, amount, txn_type, user_id) VALUES (?, ?, ?, ?, ?)",
                    (date.today().isoformat(), "Dashboard sample", amount, txn_type, 1),
                )

        original_get_db = tracker.get_db

        def postgres_style_rows():
            conn = original_get_db()

            def row_factory(cursor, values):
                # PostgreSQL labels both unaliased aggregate expressions "coalesce".
                names = [
                    "coalesce" if col[0].upper().startswith("COALESCE(") else col[0]
                    for col in cursor.description
                ]
                return HybridRow(names, values)

            conn.row_factory = row_factory
            return conn

        with patch.object(tracker, "get_db", postgres_style_rows):
            response = self.client.get("/")
        self.assertEqual(response.status_code, 200)
        body = response.get_data(as_text=True)
        self.assertIn("₹12", body)
        self.assertIn("+₹4", body)

    def test_category_changes_are_isolated_by_user(self):
        response = self.client.post("/settings", data={
            "action": "add_category", "cat_name": "Travel", "cat_keywords": "RAILWAY", "cat_type": "transport",
        })
        self.assertEqual(response.status_code, 302)
        own_page = self.client.get("/settings").get_data(as_text=True)
        other_page = self.client_for(2).get("/settings").get_data(as_text=True)
        self.assertIn("Travel", own_page)
        self.assertNotIn("Travel", other_page)
        self.assertEqual(json.loads(tracker.CATEGORIES_PATH.read_text(encoding="utf-8"))["categories"], self.defaults)

    def test_custom_keywords_match_regardless_of_case(self):
        self.client.post("/settings", data={
            "action": "add_category", "cat_name": "Travel", "cat_keywords": "RAILWAY", "cat_type": "transport",
        })
        with tracker.app.test_request_context("/"):
            tracker.session["user_id"] = 1
            self.assertEqual(tracker.auto_categorize("Railway ticket"), "Travel")

    def test_csv_debit_column_takes_precedence_over_type_hint(self):
        rows, errors = tracker.parse_uploaded_csv(
            "Date,Description,Debit,Credit,Dr/Cr\n25/11/2026,Coffee,12.50,,CR\n"
        )
        self.assertFalse(errors)
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0]["txn_type"], "debit")

    def test_import_cleanup_failure_rolls_back_transactions(self):
        batch = self.stage()
        with tracker.db_conn(commit=True) as conn:
            conn.execute("""
                CREATE TRIGGER fail_staging_delete BEFORE DELETE ON import_staging
                BEGIN SELECT RAISE(ABORT, 'synthetic cleanup failure'); END
            """)
        with self.assertRaises(sqlite3.IntegrityError):
            self.client.post("/import/review", data={"action": "commit", "batch_id": batch})
        with tracker.db_conn() as conn:
            self.assertEqual(conn.execute("SELECT COUNT(*) FROM transactions").fetchone()[0], 0)
            self.assertEqual(conn.execute("SELECT COUNT(*) FROM import_staging").fetchone()[0], 1)

    def test_backup_is_account_scoped_and_portable(self):
        with tracker.db_conn(commit=True) as conn:
            for user_id, description in ((1, "User one"), (2, "User two")):
                conn.execute(
                    "INSERT INTO transactions (date, description, amount, txn_type, user_id) VALUES (?, ?, ?, ?, ?)",
                    (date.today().isoformat(), description, 5, "debit", user_id),
                )
        response = self.client.get("/backup-download")
        self.assertEqual(response.status_code, 200)
        with zipfile.ZipFile(io.BytesIO(response.data)) as archive:
            self.assertNotIn("finance.db", archive.namelist())
            payload = json.loads(archive.read("finance_tracker_account.json"))
        descriptions = {row["description"] for row in payload["transactions"]}
        self.assertEqual(descriptions, {"User one"})
        self.assertNotIn("password_hash", payload["account"])

    def test_post_action_redirect_rejects_external_urls(self):
        with tracker.app.test_request_context("/settings", base_url="https://finance.example"):
            fallback = "/settings"
            self.assertEqual(tracker.safe_local_redirect("//evil.example/path", fallback), fallback)
            self.assertEqual(tracker.safe_local_redirect("https://evil.example/path", fallback), fallback)
            self.assertEqual(tracker.safe_local_redirect("https://finance.example/add?x=1", fallback), "/add?x=1")
            self.assertEqual(tracker.safe_local_redirect("/add?x=1", fallback), "/add?x=1")

    def test_password_validator_allows_weak_passwords(self):
        for password in ("x", "123", "password", ""):
            self.assertTrue(tracker.PasswordValidator.validate(password)[0], password)

    def test_registration_accepts_one_character_password(self):
        response = self.client.post("/register", data={
            "username": "weakuser",
            "display_name": "Weak User",
            "security_word": "remember",
            "password": "x",
            "confirm": "x",
        })
        self.assertEqual(response.status_code, 302)
        with tracker.db_conn() as conn:
            user = conn.execute(
                "SELECT password_hash FROM users WHERE username = ?", ("weakuser",)
            ).fetchone()
        self.assertTrue(tracker.check_password_hash(user["password_hash"], "x"))

    def test_password_reset_accepts_one_character_password(self):
        with tracker.db_conn(commit=True) as conn:
            conn.execute(
                "UPDATE users SET security_word = ? WHERE id = ?", ("remember", 1)
            )
        response = self.client.post("/forgot-password", data={
            "username": "regression_1",
            "security_word": "remember",
            "new_password": "x",
            "confirm": "x",
        })
        self.assertEqual(response.status_code, 302)
        with tracker.db_conn() as conn:
            password_hash = conn.execute(
                "SELECT password_hash FROM users WHERE id = ?", (1,)
            ).fetchone()["password_hash"]
        self.assertTrue(tracker.check_password_hash(password_hash, "x"))

    def test_password_change_accepts_one_character_password(self):
        old_hash = tracker.generate_password_hash("old-password")
        with tracker.db_conn(commit=True) as conn:
            conn.execute(
                "UPDATE users SET password_hash = ? WHERE id = ?", (old_hash, 1)
            )
        response = self.client.post("/change-password", data={
            "current_password": "old-password",
            "new_password": "x",
            "confirm_password": "x",
        })
        self.assertEqual(response.status_code, 302)
        with tracker.db_conn() as conn:
            password_hash = conn.execute(
                "SELECT password_hash FROM users WHERE id = ?", (1,)
            ).fetchone()["password_hash"]
        self.assertTrue(tracker.check_password_hash(password_hash, "x"))

    def test_favicon_route_is_available(self):
        response = self.client.get("/favicon.ico")
        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.mimetype, "image/svg+xml")


if __name__ == "__main__":
    unittest.main()
