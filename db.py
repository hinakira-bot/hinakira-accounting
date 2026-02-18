"""
Database initialization and connection helpers for AI Accounting Tool.
Uses SQLite as the primary data store.
"""
import sqlite3
import os

DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'data', 'accounting.db')

# --- Default Account Master (青色申告 個人事業主向け) ---
DEFAULT_ACCOUNTS = [
    # (code, name, account_type, tax_default, display_order)
    # 資産 (1xx)
    ("100", "現金", "資産", "不課税", 100),
    ("101", "小口現金", "資産", "不課税", 101),
    ("110", "普通預金", "資産", "不課税", 110),
    ("120", "売掛金", "資産", "不課税", 120),
    ("121", "未収入金", "資産", "不課税", 121),
    ("130", "棚卸資産", "資産", "不課税", 130),
    ("190", "事業主貸", "資産", "不課税", 190),
    # 負債 (2xx)
    ("200", "買掛金", "負債", "不課税", 200),
    ("201", "未払金", "負債", "不課税", 201),
    ("210", "借入金", "負債", "不課税", 210),
    ("220", "預り金", "負債", "不課税", 220),
    ("290", "事業主借", "負債", "不課税", 290),
    # 純資産 (3xx)
    ("300", "資本金", "純資産", "不課税", 300),
    ("301", "元入金", "純資産", "不課税", 301),
    # 収益 (4xx)
    ("400", "売上高", "収益", "10%", 400),
    ("410", "雑収入", "収益", "10%", 410),
    # 費用 (5xx-8xx)
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

SCHEMA_SQL = """
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

CREATE TABLE IF NOT EXISTS opening_balances (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    fiscal_year TEXT NOT NULL,
    account_id  INTEGER NOT NULL,
    amount      INTEGER NOT NULL DEFAULT 0,
    note        TEXT DEFAULT '',
    FOREIGN KEY (account_id) REFERENCES accounts_master(id),
    UNIQUE(fiscal_year, account_id)
);

CREATE TABLE IF NOT EXISTS counterparties (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    name            TEXT NOT NULL,
    code            TEXT DEFAULT '',
    contact_info    TEXT DEFAULT '',
    notes           TEXT DEFAULT '',
    is_active       INTEGER DEFAULT 1,
    created_at      TEXT DEFAULT (datetime('now')),
    updated_at      TEXT DEFAULT (datetime('now'))
);
CREATE INDEX IF NOT EXISTS idx_counterparties_name ON counterparties(name);

CREATE TABLE IF NOT EXISTS settings (
    key     TEXT PRIMARY KEY,
    value   TEXT NOT NULL
);
"""


def get_db():
    """Get a database connection with row_factory set."""
    os.makedirs(os.path.dirname(DB_PATH), exist_ok=True)
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA foreign_keys=ON")
    return conn


def init_db():
    """Create tables and seed account master data if empty."""
    conn = get_db()
    try:
        conn.executescript(SCHEMA_SQL)

        # Seed accounts only if table is empty
        count = conn.execute("SELECT COUNT(*) FROM accounts_master").fetchone()[0]
        if count == 0:
            conn.executemany(
                "INSERT INTO accounts_master (code, name, account_type, tax_default, display_order) VALUES (?, ?, ?, ?, ?)",
                DEFAULT_ACCOUNTS
            )
            print(f"Seeded {len(DEFAULT_ACCOUNTS)} accounts into accounts_master.")

        # Seed default settings
        defaults = [
            ("fiscal_year_start", "04-01"),
            ("business_name", ""),
        ]
        for key, value in defaults:
            conn.execute(
                "INSERT OR IGNORE INTO settings (key, value) VALUES (?, ?)",
                (key, value)
            )

        # Migrate existing counterparty strings into new counterparties table
        cp_count = conn.execute("SELECT COUNT(*) FROM counterparties").fetchone()[0]
        if cp_count == 0:
            conn.execute("""
                INSERT OR IGNORE INTO counterparties (name)
                SELECT DISTINCT counterparty FROM journal_entries
                WHERE is_deleted = 0 AND counterparty != ''
            """)

        conn.commit()
        print("Database initialized successfully.")
    except Exception as e:
        print(f"Database init error: {e}")
        conn.rollback()
    finally:
        conn.close()


if __name__ == "__main__":
    init_db()
    print(f"Database created at: {DB_PATH}")
