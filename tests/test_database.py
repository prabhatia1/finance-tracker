import unittest

from database import HybridRow, adapt_sql


class DatabaseAdapterTests(unittest.TestCase):
    def test_month_filters_and_placeholders(self):
        sql = "SELECT id FROM transactions WHERE strftime('%Y', date) = ? AND strftime('%m', date) = ?"
        self.assertEqual(
            adapt_sql(sql),
            "SELECT id FROM transactions WHERE SUBSTRING(date FROM 1 FOR 4) = %s "
            "AND SUBSTRING(date FROM 6 FOR 2) = %s",
        )

    def test_ignored_insert_becomes_conflict_safe_postgres_insert(self):
        sql = "INSERT OR IGNORE INTO user_people (user_id, name) VALUES (?, ?)"
        self.assertEqual(
            adapt_sql(sql),
            "INSERT INTO user_people (user_id, name) VALUES (%s, %s) ON CONFLICT DO NOTHING",
        )

    def test_datetime_and_literal_percent(self):
        sql = "SELECT datetime('now','localtime') WHERE label LIKE '%tax%' AND id = ?"
        self.assertEqual(
            adapt_sql(sql),
            "SELECT LOCALTIMESTAMP WHERE label LIKE '%%tax%%' AND id = %s",
        )

    def test_row_supports_name_and_position_access(self):
        row = HybridRow(("id", "name"), (7, "sample"))
        self.assertEqual(row[0], 7)
        self.assertEqual(row[1], "sample")
        self.assertEqual(row["id"], 7)
        self.assertEqual(dict(row), {"id": 7, "name": "sample"})


if __name__ == "__main__":
    unittest.main()
