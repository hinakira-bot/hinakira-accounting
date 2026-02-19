"""
Database initialization and connection helpers for AI Accounting Tool.
Supports PostgreSQL (production) and SQLite (local development).
Set DATABASE_URL env var to use PostgreSQL; otherwise SQLite is used.
"""
import os

# --- Database mode detection ---
DATABASE_URL = os.environ.get('DATABASE_URL')
USE_PG = bool(DATABASE_URL)

# SQLite path (used only when DATABASE_URL is not set)
if os.environ.get('VERCEL'):
    DB_PATH = '/tmp/accounting.db'
else:
    DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'data', 'accounting.db')


def P(sql):
    """Convert SQLite-style ? placeholders to PostgreSQL %s when needed."""
    if USE_PG:
        return sql.replace('?', '%s')
    return sql


# --- Default Account Master (青色申告 個人事業主向け) ---
DEFAULT_ACCOUNTS = [
    ("100", "現金", "資産", "不課税", 100),
    ("101", "小口現金", "資産", "不課税", 101),
    ("110", "普通預金", "資産", "不課税", 110),
    ("120", "売掛金", "資産", "不課税", 120),
    ("121", "未収入金", "資産", "不課税", 121),
    ("130", "棚卸資産", "資産", "不課税", 130),
    ("190", "事業主貸", "資産", "不課税", 190),
    ("200", "買掛金", "負債", "不課税", 200),
    ("201", "未払金", "負債", "不課税", 201),
    ("210", "借入金", "負債", "不課税", 210),
    ("220", "預り金", "負債", "不課税", 220),
    ("290", "事業主借", "負債", "不課税", 290),
    ("300", "資本金", "純資産", "不課税", 300),
    ("301", "元入金", "純資産", "不課税", 301),
    ("400", "売上高", "収益", "10%", 400),
    ("410", "雑収入", "収益", "10%", 410),
    ("500", "仕入高", "費用", "10%", 500),
    ("510", "役員報酬", "費用", "不課税", 510),
    ("511", "給料手当", "費用", "不課税", 511),
    ("520", "外注工賃", "費用", "10%", 520),
    ("530", "旅費交通費", "費用", "10%", 530),
    ("531", "通信費", "費用", "10%", 531),
    ("540", "広告宣伝費", "費用", "10%", 540),
    ("541", "接待交際費", "費用", "10%", 541),
    ("550", "消耗品費", "費用", "10%", 550),
    ("551", "会議費", "費用", "10%", 551),
    ("560", "水道光熱費", "費用", "10%", 560),
    ("570", "地代家賃", "費用", "非課税", 570),
    ("580", "修繕費", "費用", "10%", 580),
    ("590", "支払手数料", "費用", "10%", 590),
    ("600", "租税公課", "費用", "不課税", 600),
    ("610", "新聞図書費", "費用", "10%", 610),
    ("620", "保険料", "費用", "非課税", 620),
    ("630", "減価償却費", "費用", "不課税", 630),
    ("900", "雑費", "費用", "10%", 900),
]

# --- Schema (PostgreSQL) ---
PG_SCHEMA_STATEMENTS = [
    """CREATE TABLE IF NOT EXISTS users (
        id          SERIAL PRIMARY KEY,
        email       TEXT UNIQUE NOT NULL,
        name        TEXT DEFAULT '',
        picture     TEXT DEFAULT '',
        created_at  TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        last_login  TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    )""",
    """CREATE TABLE IF NOT EXISTS accounts_master (
        id              SERIAL PRIMARY KEY,
        code            TEXT NOT NULL UNIQUE,
        name            TEXT NOT NULL UNIQUE,
        account_type    TEXT NOT NULL,
        tax_default     TEXT DEFAULT '10%',
        is_active       INTEGER DEFAULT 1,
        display_order   INTEGER DEFAULT 0,
        created_at      TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        updated_at      TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    )""",
    """CREATE TABLE IF NOT EXISTS journal_entries (
        id                  SERIAL PRIMARY KEY,
        user_id             INTEGER NOT NULL DEFAULT 0,
        entry_date          TEXT NOT NULL,
        debit_account_id    INTEGER NOT NULL,
        credit_account_id   INTEGER NOT NULL,
        amount              INTEGER NOT NULL,
        tax_classification  TEXT NOT NULL DEFAULT '10%',
        tax_amount          INTEGER DEFAULT 0,
        counterparty        TEXT DEFAULT '',
        memo                TEXT DEFAULT '',
        evidence_url        TEXT DEFAULT '',
        source              TEXT DEFAULT 'manual',
        is_deleted          INTEGER DEFAULT 0,
        created_at          TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        updated_at          TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY (debit_account_id) REFERENCES accounts_master(id),
        FOREIGN KEY (credit_account_id) REFERENCES accounts_master(id)
    )""",
    "CREATE INDEX IF NOT EXISTS idx_journal_entry_date ON journal_entries(entry_date)",
    "CREATE INDEX IF NOT EXISTS idx_journal_debit ON journal_entries(debit_account_id)",
    "CREATE INDEX IF NOT EXISTS idx_journal_credit ON journal_entries(credit_account_id)",
    "CREATE INDEX IF NOT EXISTS idx_journal_deleted ON journal_entries(is_deleted)",
    "CREATE INDEX IF NOT EXISTS idx_journal_user ON journal_entries(user_id)",
    """CREATE TABLE IF NOT EXISTS opening_balances (
        id          SERIAL PRIMARY KEY,
        user_id     INTEGER NOT NULL DEFAULT 0,
        fiscal_year TEXT NOT NULL,
        account_id  INTEGER NOT NULL,
        amount      INTEGER NOT NULL DEFAULT 0,
        note        TEXT DEFAULT '',
        FOREIGN KEY (account_id) REFERENCES accounts_master(id)
    )""",
    """CREATE TABLE IF NOT EXISTS counterparties (
        id              SERIAL PRIMARY KEY,
        user_id         INTEGER NOT NULL DEFAULT 0,
        name            TEXT NOT NULL,
        code            TEXT DEFAULT '',
        contact_info    TEXT DEFAULT '',
        notes           TEXT DEFAULT '',
        is_active       INTEGER DEFAULT 1,
        created_at      TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        updated_at      TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    )""",
    "CREATE INDEX IF NOT EXISTS idx_counterparties_name ON counterparties(name)",
    "CREATE INDEX IF NOT EXISTS idx_counterparties_user ON counterparties(user_id)",
    """CREATE TABLE IF NOT EXISTS user_settings (
        user_id     INTEGER NOT NULL DEFAULT 0,
        key         TEXT NOT NULL,
        value       TEXT NOT NULL,
        PRIMARY KEY (user_id, key)
    )""",
    """CREATE TABLE IF NOT EXISTS import_history (
        id              SERIAL PRIMARY KEY,
        user_id         INTEGER NOT NULL DEFAULT 0,
        filename        TEXT NOT NULL,
        file_hash       TEXT NOT NULL,
        source_name     TEXT NOT NULL DEFAULT '',
        row_count       INTEGER DEFAULT 0,
        imported_count  INTEGER DEFAULT 0,
        date_range_start TEXT DEFAULT '',
        date_range_end   TEXT DEFAULT '',
        created_at      TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    )""",
    "CREATE INDEX IF NOT EXISTS idx_import_history_user ON import_history(user_id)",
    "CREATE INDEX IF NOT EXISTS idx_import_history_hash ON import_history(file_hash)",
    """CREATE TABLE IF NOT EXISTS statement_sources (
        id              SERIAL PRIMARY KEY,
        user_id         INTEGER NOT NULL DEFAULT 0,
        source_name     TEXT NOT NULL,
        default_debit   TEXT DEFAULT '',
        default_credit  TEXT DEFAULT '',
        created_at      TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        updated_at      TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    )""",
    "CREATE UNIQUE INDEX IF NOT EXISTS idx_statement_sources_unique ON statement_sources(user_id, source_name)",
]

# --- Schema (SQLite — kept for local dev) ---
SQLITE_SCHEMA_SQL = """
CREATE TABLE IF NOT EXISTS users (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    email       TEXT UNIQUE NOT NULL,
    name        TEXT DEFAULT '',
    picture     TEXT DEFAULT '',
    created_at  TEXT DEFAULT (datetime('now')),
    last_login  TEXT DEFAULT (datetime('now'))
);
CREATE TABLE IF NOT EXISTS accounts_master (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    code            TEXT NOT NULL UNIQUE,
    name            TEXT NOT NULL UNIQUE,
    account_type    TEXT NOT NULL,
    tax_default     TEXT DEFAULT '10%',
    is_active       INTEGER DEFAULT 1,
    display_order   INTEGER DEFAULT 0,
    created_at      TEXT DEFAULT (datetime('now')),
    updated_at      TEXT DEFAULT (datetime('now'))
);
CREATE TABLE IF NOT EXISTS journal_entries (
    id                  INTEGER PRIMARY KEY AUTOINCREMENT,
    user_id             INTEGER NOT NULL DEFAULT 0,
    entry_date          TEXT NOT NULL,
    debit_account_id    INTEGER NOT NULL,
    credit_account_id   INTEGER NOT NULL,
    amount              INTEGER NOT NULL,
    tax_classification  TEXT NOT NULL DEFAULT '10%',
    tax_amount          INTEGER DEFAULT 0,
    counterparty        TEXT DEFAULT '',
    memo                TEXT DEFAULT '',
    evidence_url        TEXT DEFAULT '',
    source              TEXT DEFAULT 'manual',
    is_deleted          INTEGER DEFAULT 0,
    created_at          TEXT DEFAULT (datetime('now')),
    updated_at          TEXT DEFAULT (datetime('now')),
    FOREIGN KEY (debit_account_id) REFERENCES accounts_master(id),
    FOREIGN KEY (credit_account_id) REFERENCES accounts_master(id)
);
CREATE INDEX IF NOT EXISTS idx_journal_entry_date ON journal_entries(entry_date);
CREATE INDEX IF NOT EXISTS idx_journal_debit ON journal_entries(debit_account_id);
CREATE INDEX IF NOT EXISTS idx_journal_credit ON journal_entries(credit_account_id);
CREATE INDEX IF NOT EXISTS idx_journal_deleted ON journal_entries(is_deleted);
CREATE INDEX IF NOT EXISTS idx_journal_user ON journal_entries(user_id);
CREATE TABLE IF NOT EXISTS opening_balances (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    user_id     INTEGER NOT NULL DEFAULT 0,
    fiscal_year TEXT NOT NULL,
    account_id  INTEGER NOT NULL,
    amount      INTEGER NOT NULL DEFAULT 0,
    note        TEXT DEFAULT '',
    FOREIGN KEY (account_id) REFERENCES accounts_master(id)
);
CREATE TABLE IF NOT EXISTS counterparties (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    user_id         INTEGER NOT NULL DEFAULT 0,
    name            TEXT NOT NULL,
    code            TEXT DEFAULT '',
    contact_info    TEXT DEFAULT '',
    notes           TEXT DEFAULT '',
    is_active       INTEGER DEFAULT 1,
    created_at      TEXT DEFAULT (datetime('now')),
    updated_at      TEXT DEFAULT (datetime('now'))
);
CREATE INDEX IF NOT EXISTS idx_counterparties_name ON counterparties(name);
CREATE INDEX IF NOT EXISTS idx_counterparties_user ON counterparties(user_id);
CREATE TABLE IF NOT EXISTS user_settings (
    user_id     INTEGER NOT NULL DEFAULT 0,
    key         TEXT NOT NULL,
    value       TEXT NOT NULL,
    PRIMARY KEY (user_id, key)
);
CREATE TABLE IF NOT EXISTS import_history (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    user_id         INTEGER NOT NULL DEFAULT 0,
    filename        TEXT NOT NULL,
    file_hash       TEXT NOT NULL,
    source_name     TEXT NOT NULL DEFAULT '',
    row_count       INTEGER DEFAULT 0,
    imported_count  INTEGER DEFAULT 0,
    date_range_start TEXT DEFAULT '',
    date_range_end   TEXT DEFAULT '',
    created_at      TEXT DEFAULT (datetime('now'))
);
CREATE INDEX IF NOT EXISTS idx_import_history_user ON import_history(user_id);
CREATE INDEX IF NOT EXISTS idx_import_history_hash ON import_history(file_hash);
CREATE TABLE IF NOT EXISTS statement_sources (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    user_id         INTEGER NOT NULL DEFAULT 0,
    source_name     TEXT NOT NULL,
    default_debit   TEXT DEFAULT '',
    default_credit  TEXT DEFAULT '',
    created_at      TEXT DEFAULT (datetime('now')),
    updated_at      TEXT DEFAULT (datetime('now'))
);
CREATE UNIQUE INDEX IF NOT EXISTS idx_statement_sources_unique ON statement_sources(user_id, source_name);
"""


class PgConnectionWrapper:
    """Wrap psycopg2 connection so that conn.execute() works like SQLite.
    Returns a cursor that supports fetchone()/fetchall() directly."""
    def __init__(self, raw_conn):
        self._conn = raw_conn

    def execute(self, sql, params=None):
        cur = self._conn.cursor()
        cur.execute(sql, params)
        return cur

    def commit(self):
        self._conn.commit()

    def rollback(self):
        self._conn.rollback()

    def close(self):
        self._conn.close()

    def cursor(self):
        return self._conn.cursor()

    @property
    def autocommit(self):
        return self._conn.autocommit

    @autocommit.setter
    def autocommit(self, value):
        self._conn.autocommit = value


def get_db():
    """Get a database connection. Uses PostgreSQL if DATABASE_URL is set, else SQLite."""
    if USE_PG:
        import psycopg2
        import psycopg2.extras
        raw = psycopg2.connect(DATABASE_URL)
        raw.autocommit = False
        # Use RealDictCursor so rows are dicts (like sqlite3.Row)
        raw.cursor_factory = psycopg2.extras.RealDictCursor
        return PgConnectionWrapper(raw)
    else:
        import sqlite3
        os.makedirs(os.path.dirname(DB_PATH), exist_ok=True)
        conn = sqlite3.connect(DB_PATH)
        conn.row_factory = sqlite3.Row
        conn.execute("PRAGMA journal_mode=WAL")
        conn.execute("PRAGMA foreign_keys=ON")
        return conn


def _table_exists(conn, table_name):
    """Check if a table exists (works for both PG and SQLite)."""
    if USE_PG:
        cur = conn.cursor()
        cur.execute(
            "SELECT 1 FROM information_schema.tables WHERE table_name = %s",
            (table_name,)
        )
        return cur.fetchone() is not None
    else:
        return conn.execute(
            "SELECT name FROM sqlite_master WHERE type='table' AND name=?",
            (table_name,)
        ).fetchone() is not None


def _column_exists(conn, table_name, column_name):
    """Check if a column exists in a table (works for both PG and SQLite)."""
    if USE_PG:
        cur = conn.cursor()
        cur.execute(
            "SELECT 1 FROM information_schema.columns WHERE table_name = %s AND column_name = %s",
            (table_name, column_name)
        )
        return cur.fetchone() is not None
    else:
        cols = [r['name'] for r in conn.execute(f"PRAGMA table_info({table_name})").fetchall()]
        return column_name in cols


def migrate_db():
    """Run schema migrations for multi-user support."""
    conn = get_db()
    try:
        cur = conn.cursor()

        if not _table_exists(conn, 'users'):
            if USE_PG:
                cur.execute("""CREATE TABLE IF NOT EXISTS users (
                    id SERIAL PRIMARY KEY, email TEXT UNIQUE NOT NULL,
                    name TEXT DEFAULT '', picture TEXT DEFAULT '',
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    last_login TIMESTAMP DEFAULT CURRENT_TIMESTAMP)""")
            else:
                cur.execute("""CREATE TABLE IF NOT EXISTS users (
                    id INTEGER PRIMARY KEY AUTOINCREMENT, email TEXT UNIQUE NOT NULL,
                    name TEXT DEFAULT '', picture TEXT DEFAULT '',
                    created_at TEXT DEFAULT (datetime('now')),
                    last_login TEXT DEFAULT (datetime('now')))""")
            print("Migration: Created users table.")

        # Add user_id columns if missing
        for tbl in ['journal_entries', 'counterparties', 'opening_balances']:
            if _table_exists(conn, tbl) and not _column_exists(conn, tbl, 'user_id'):
                cur.execute(f"ALTER TABLE {tbl} ADD COLUMN user_id INTEGER NOT NULL DEFAULT 0")
                if tbl != 'opening_balances':
                    cur.execute(f"CREATE INDEX IF NOT EXISTS idx_{tbl}_user ON {tbl}(user_id)")
                print(f"Migration: Added user_id to {tbl}.")

        if not _table_exists(conn, 'user_settings'):
            cur.execute("""CREATE TABLE IF NOT EXISTS user_settings (
                user_id INTEGER NOT NULL DEFAULT 0, key TEXT NOT NULL,
                value TEXT NOT NULL, PRIMARY KEY (user_id, key))""")
            if _table_exists(conn, 'settings'):
                if USE_PG:
                    cur.execute("""INSERT INTO user_settings (user_id, key, value)
                        SELECT 0, key, value FROM settings ON CONFLICT DO NOTHING""")
                else:
                    cur.execute("""INSERT OR IGNORE INTO user_settings (user_id, key, value)
                        SELECT 0, key, value FROM settings""")
            print("Migration: Created user_settings table.")

        if not _table_exists(conn, 'import_history'):
            if USE_PG:
                cur.execute("""CREATE TABLE IF NOT EXISTS import_history (
                    id SERIAL PRIMARY KEY, user_id INTEGER NOT NULL DEFAULT 0,
                    filename TEXT NOT NULL, file_hash TEXT NOT NULL,
                    source_name TEXT NOT NULL DEFAULT '', row_count INTEGER DEFAULT 0,
                    imported_count INTEGER DEFAULT 0, date_range_start TEXT DEFAULT '',
                    date_range_end TEXT DEFAULT '',
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP)""")
            else:
                cur.execute("""CREATE TABLE IF NOT EXISTS import_history (
                    id INTEGER PRIMARY KEY AUTOINCREMENT, user_id INTEGER NOT NULL DEFAULT 0,
                    filename TEXT NOT NULL, file_hash TEXT NOT NULL,
                    source_name TEXT NOT NULL DEFAULT '', row_count INTEGER DEFAULT 0,
                    imported_count INTEGER DEFAULT 0, date_range_start TEXT DEFAULT '',
                    date_range_end TEXT DEFAULT '',
                    created_at TEXT DEFAULT (datetime('now')))""")
            cur.execute("CREATE INDEX IF NOT EXISTS idx_import_history_user ON import_history(user_id)")
            cur.execute("CREATE INDEX IF NOT EXISTS idx_import_history_hash ON import_history(file_hash)")
            print("Migration: Created import_history table.")

        if not _table_exists(conn, 'statement_sources'):
            if USE_PG:
                cur.execute("""CREATE TABLE IF NOT EXISTS statement_sources (
                    id SERIAL PRIMARY KEY, user_id INTEGER NOT NULL DEFAULT 0,
                    source_name TEXT NOT NULL, default_debit TEXT DEFAULT '',
                    default_credit TEXT DEFAULT '',
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP)""")
            else:
                cur.execute("""CREATE TABLE IF NOT EXISTS statement_sources (
                    id INTEGER PRIMARY KEY AUTOINCREMENT, user_id INTEGER NOT NULL DEFAULT 0,
                    source_name TEXT NOT NULL, default_debit TEXT DEFAULT '',
                    default_credit TEXT DEFAULT '',
                    created_at TEXT DEFAULT (datetime('now')),
                    updated_at TEXT DEFAULT (datetime('now')))""")
            cur.execute("CREATE UNIQUE INDEX IF NOT EXISTS idx_statement_sources_unique ON statement_sources(user_id, source_name)")
            print("Migration: Created statement_sources table.")

        conn.commit()
        print("Migration completed successfully.")
    except Exception as e:
        print(f"Migration error: {e}")
        conn.rollback()
    finally:
        conn.close()


def init_db():
    """Create tables and seed account master data if empty."""
    conn = get_db()
    try:
        cur = conn.cursor()

        if USE_PG:
            for stmt in PG_SCHEMA_STATEMENTS:
                try:
                    cur.execute(stmt)
                except Exception as e:
                    print(f"Schema note: {e}")
            print("PostgreSQL schema initialized.", flush=True)
        else:
            try:
                conn.executescript(SQLITE_SCHEMA_SQL)
            except Exception as e:
                print(f"Schema note (will be fixed by migration): {e}")

        # Seed accounts only if table is empty
        cur.execute("SELECT COUNT(*) AS cnt FROM accounts_master")
        row = cur.fetchone()
        count = row['cnt'] if USE_PG else row[0]
        if count == 0:
            if USE_PG:
                for acc in DEFAULT_ACCOUNTS:
                    cur.execute(
                        "INSERT INTO accounts_master (code, name, account_type, tax_default, display_order) "
                        "VALUES (%s, %s, %s, %s, %s)", acc
                    )
            else:
                conn.executemany(
                    "INSERT INTO accounts_master (code, name, account_type, tax_default, display_order) VALUES (?, ?, ?, ?, ?)",
                    DEFAULT_ACCOUNTS
                )
            print(f"Seeded {len(DEFAULT_ACCOUNTS)} accounts into accounts_master.")

        conn.commit()
        db_type = "PostgreSQL" if USE_PG else "SQLite"
        print(f"Database initialized successfully. ({db_type})", flush=True)
    except Exception as e:
        print(f"Database init error: {e}")
        conn.rollback()
    finally:
        conn.close()

    migrate_db()


if __name__ == "__main__":
    init_db()
    if USE_PG:
        print(f"Database: PostgreSQL ({DATABASE_URL[:30]}...)")
    else:
        print(f"Database: SQLite ({DB_PATH})")
