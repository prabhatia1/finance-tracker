"""Database connection helpers for local SQLite and hosted PostgreSQL."""

import os
import re
import sqlite3

DATABASE_URL = (
    os.environ.get("DATABASE_URL")
    or os.environ.get("POSTGRES_URL")
    or os.environ.get("POSTGRES_URL_NON_POOLING")
    or os.environ.get("DATABASE_URL_UNPOOLED")
)
USE_POSTGRES = bool(DATABASE_URL)

try:
    import psycopg
except ImportError:
    psycopg = None

if psycopg:
    INTEGRITY_ERRORS = (sqlite3.IntegrityError, psycopg.IntegrityError)
else:
    INTEGRITY_ERRORS = (sqlite3.IntegrityError,)


class HybridRow(dict):
    """Mapping row that also supports SQLite-style integer indexing."""

    def __init__(self, names, values):
        values = tuple(values)
        super().__init__(zip(names, values))
        self._values = values

    def __getitem__(self, key):
        if isinstance(key, int):
            return self._values[key]
        return super().__getitem__(key)


def _hybrid_row_factory(cursor):
    names = [column.name for column in cursor.description or ()]
    return lambda values: HybridRow(names, values)


def adapt_sql(query):
    """Translate the small set of SQLite SQL syntax used by the app."""
    query = re.sub(
        r"strftime\('%Y',\s*date\)", "SUBSTRING(date FROM 1 FOR 4)", query, flags=re.I
    )
    query = re.sub(
        r"strftime\('%m',\s*date\)", "SUBSTRING(date FROM 6 FOR 2)", query, flags=re.I
    )
    query = re.sub(
        r"datetime\(\s*'now'\s*(?:,\s*'localtime'\s*)?\)",
        "LOCALTIMESTAMP",
        query,
        flags=re.I,
    )

    if re.match(r"\s*INSERT\s+OR\s+IGNORE\s+INTO\b", query, re.I):
        query = re.sub(r"INSERT\s+OR\s+IGNORE\s+INTO", "INSERT INTO", query, count=1, flags=re.I)
        query = query.rstrip().rstrip(";") + " ON CONFLICT DO NOTHING"

    placeholder = "\x00SQL_PLACEHOLDER\x00"
    query = query.replace("?", placeholder).replace("%", "%%")
    return query.replace(placeholder, "%s")


class PostgresConnection:
    def __init__(self, connection):
        self._connection = connection

    def execute(self, query, params=()):
        return self._connection.execute(adapt_sql(query), params)

    def executescript(self, script):
        for statement in script.split(";"):
            if statement.strip():
                self.execute(statement)

    def commit(self):
        self._connection.commit()

    def close(self):
        self._connection.close()


def connect_database(sqlite_path):
    if not USE_POSTGRES:
        return sqlite3.connect(str(sqlite_path), timeout=15)
    if psycopg is None:
        raise RuntimeError("Install psycopg to use DATABASE_URL with PostgreSQL.")
    raw_connection = psycopg.connect(
        DATABASE_URL,
        connect_timeout=10,
        row_factory=_hybrid_row_factory,
    )
    return PostgresConnection(raw_connection)
