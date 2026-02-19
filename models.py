"""
Data access layer for AI Accounting Tool.
CRUD operations for journal entries, accounts, trial balance, etc.
All user-scoped functions require user_id parameter.
"""
import math
from db import get_db

# --- Tax Calculation ---
TAX_RATES = {
    '10%': 10,
    '8%': 8,
    '非課税': 0,
    '不課税': 0,
}


def calculate_tax_amount(amount: int, classification: str) -> int:
    """Calculate consumption tax from tax-inclusive amount (税込経理方式)."""
    rate = TAX_RATES.get(classification, 0)
    if rate == 0:
        return 0
    return amount * rate // (100 + rate)


# --- Users ---
def get_or_create_user(email: str, name: str = '', picture: str = '') -> dict:
    """Find user by email or create new one. Returns user dict with id."""
    conn = get_db()
    try:
        row = conn.execute("SELECT id, email, name, picture FROM users WHERE email = ?", (email,)).fetchone()
        if row:
            # Update last_login
            conn.execute("UPDATE users SET last_login = datetime('now'), name = ?, picture = ? WHERE id = ?",
                         (name or row['name'], picture or row['picture'], row['id']))
            conn.commit()
            user = dict(row)
            # Migrate legacy data (user_id=0) to first user
            _migrate_legacy_data(conn, user['id'])
            return user
        else:
            cursor = conn.execute(
                "INSERT INTO users (email, name, picture) VALUES (?, ?, ?)",
                (email, name, picture)
            )
            conn.commit()
            new_id = cursor.lastrowid
            user = {'id': new_id, 'email': email, 'name': name, 'picture': picture}
            # Migrate legacy data to first registered user
            _migrate_legacy_data(conn, new_id)
            return user
    finally:
        conn.close()


def _migrate_legacy_data(conn, user_id: int):
    """Migrate user_id=0 data to the given user (first-come gets legacy data)."""
    # Check if there's any legacy data
    legacy_count = conn.execute("SELECT COUNT(*) FROM journal_entries WHERE user_id = 0").fetchone()[0]
    if legacy_count == 0:
        return

    # Only migrate if no other user has claimed this data yet
    any_claimed = conn.execute("SELECT COUNT(*) FROM journal_entries WHERE user_id != 0").fetchone()[0]
    if any_claimed > 0:
        return  # Another user already has data, don't re-migrate

    conn.execute("UPDATE journal_entries SET user_id = ? WHERE user_id = 0", (user_id,))
    conn.execute("UPDATE opening_balances SET user_id = ? WHERE user_id = 0", (user_id,))
    conn.execute("UPDATE counterparties SET user_id = ? WHERE user_id = 0", (user_id,))
    conn.execute("UPDATE user_settings SET user_id = ? WHERE user_id = 0", (user_id,))
    conn.commit()
    print(f"Migrated legacy data to user {user_id}")


# --- Accounts (Global, no user_id) ---
def get_accounts(active_only=True):
    """Get all accounts from master."""
    conn = get_db()
    try:
        sql = "SELECT id, code, name, account_type, tax_default, display_order FROM accounts_master"
        if active_only:
            sql += " WHERE is_active = 1"
        sql += " ORDER BY display_order, code"
        rows = conn.execute(sql).fetchall()
        return [dict(r) for r in rows]
    finally:
        conn.close()


def create_account(code: str, name: str, account_type: str, tax_default: str = '10%') -> dict:
    """Create a new account in accounts_master."""
    conn = get_db()
    try:
        display_order = int(code) if code.isdigit() else 0
        conn.execute(
            "INSERT INTO accounts_master (code, name, account_type, tax_default, display_order) VALUES (?, ?, ?, ?, ?)",
            (code, name, account_type, tax_default, display_order)
        )
        conn.commit()
        row = conn.execute(
            "SELECT id, code, name, account_type, tax_default, display_order FROM accounts_master WHERE code = ?",
            (code,)
        ).fetchone()
        return dict(row) if row else {}
    finally:
        conn.close()


def delete_account(account_id: int) -> bool:
    """Soft-delete an account (set is_active = 0). Prevents deletion if used in journal entries."""
    conn = get_db()
    try:
        used = conn.execute(
            "SELECT COUNT(*) FROM journal_entries WHERE (debit_account_id = ? OR credit_account_id = ?) AND is_deleted = 0",
            (account_id, account_id)
        ).fetchone()[0]
        if used > 0:
            return False
        conn.execute(
            "UPDATE accounts_master SET is_active = 0, updated_at = datetime('now') WHERE id = ?",
            (account_id,)
        )
        conn.commit()
        return True
    finally:
        conn.close()


def get_account_by_name(name: str):
    """Get account by name."""
    conn = get_db()
    try:
        row = conn.execute(
            "SELECT id, code, name, account_type, tax_default FROM accounts_master WHERE name = ? AND is_active = 1",
            (name,)
        ).fetchone()
        return dict(row) if row else None
    finally:
        conn.close()


def get_account_by_id(account_id: int):
    """Get account by ID."""
    conn = get_db()
    try:
        row = conn.execute(
            "SELECT id, code, name, account_type, tax_default FROM accounts_master WHERE id = ?",
            (account_id,)
        ).fetchone()
        return dict(row) if row else None
    finally:
        conn.close()


# --- Journal Entries (user-scoped) ---
def create_journal_entry(entry: dict, user_id: int = 0) -> int:
    """Create a single journal entry. Returns the new entry ID."""
    conn = get_db()
    try:
        amount = int(entry.get('amount', 0))
        tax_class = entry.get('tax_classification', '10%')
        tax_amount = calculate_tax_amount(amount, tax_class)

        debit_id = _resolve_account_id(conn, entry.get('debit_account_id'), entry.get('debit_account'))
        credit_id = _resolve_account_id(conn, entry.get('credit_account_id'), entry.get('credit_account'))

        if not debit_id or not credit_id:
            raise ValueError(f"Invalid account: debit={entry.get('debit_account', entry.get('debit_account_id'))}, credit={entry.get('credit_account', entry.get('credit_account_id'))}")

        entry_date = entry.get('entry_date', entry.get('date', ''))
        memo = entry.get('memo', '')
        import datetime
        current_year = datetime.date.today().year
        if entry_date:
            try:
                parsed_date = datetime.date.fromisoformat(entry_date)
                if parsed_date.year < current_year:
                    original_date = entry_date
                    entry_date = f"{current_year}-01-01"
                    memo = f"[実際日付:{original_date}] {memo}".strip()
            except (ValueError, TypeError):
                pass

        cursor = conn.execute(
            """INSERT INTO journal_entries
            (user_id, entry_date, debit_account_id, credit_account_id, amount, tax_classification, tax_amount,
             counterparty, memo, evidence_url, source)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
            (
                user_id,
                entry_date,
                debit_id,
                credit_id,
                amount,
                tax_class,
                tax_amount,
                entry.get('counterparty', ''),
                memo,
                entry.get('evidence_url', ''),
                entry.get('source', 'manual'),
            )
        )
        conn.commit()
        return cursor.lastrowid
    except Exception as e:
        conn.rollback()
        raise e
    finally:
        conn.close()


def create_journal_entries_batch(entries: list, user_id: int = 0) -> list:
    """Create multiple journal entries. Returns list of new IDs."""
    ids = []
    for entry in entries:
        try:
            new_id = create_journal_entry(entry, user_id)
            ids.append(new_id)
        except Exception as e:
            print(f"Error creating entry: {e}, data: {entry}")
            ids.append(None)
    return ids


def get_journal_entries(filters: dict = None, user_id: int = 0) -> dict:
    """Get journal entries with filters and pagination (user-scoped)."""
    if filters is None:
        filters = {}

    conn = get_db()
    try:
        conditions = ["je.is_deleted = 0", "je.user_id = ?"]
        params = [user_id]

        if filters.get('start_date'):
            conditions.append("je.entry_date >= ?")
            params.append(filters['start_date'])
        if filters.get('end_date'):
            conditions.append("je.entry_date <= ?")
            params.append(filters['end_date'])
        if filters.get('account_id'):
            conditions.append("(je.debit_account_id = ? OR je.credit_account_id = ?)")
            params.extend([filters['account_id'], filters['account_id']])
        if filters.get('counterparty'):
            conditions.append("je.counterparty LIKE ?")
            params.append(f"%{filters['counterparty']}%")
        if filters.get('memo'):
            conditions.append("je.memo LIKE ?")
            params.append(f"%{filters['memo']}%")
        if filters.get('amount_min'):
            conditions.append("je.amount >= ?")
            params.append(int(filters['amount_min']))
        if filters.get('amount_max'):
            conditions.append("je.amount <= ?")
            params.append(int(filters['amount_max']))

        where = " AND ".join(conditions)

        count_sql = f"SELECT COUNT(*) FROM journal_entries je WHERE {where}"
        total = conn.execute(count_sql, params).fetchone()[0]

        page = max(1, int(filters.get('page', 1)))
        per_page = min(100, max(1, int(filters.get('per_page', 20))))
        offset = (page - 1) * per_page

        sql = f"""
        SELECT
            je.id, je.entry_date, je.amount, je.tax_classification, je.tax_amount,
            je.counterparty, je.memo, je.evidence_url, je.source,
            je.debit_account_id, je.credit_account_id,
            da.name AS debit_account, da.code AS debit_code,
            ca.name AS credit_account, ca.code AS credit_code
        FROM journal_entries je
        JOIN accounts_master da ON je.debit_account_id = da.id
        JOIN accounts_master ca ON je.credit_account_id = ca.id
        WHERE {where}
        ORDER BY je.entry_date DESC, je.id DESC
        LIMIT ? OFFSET ?
        """
        params.extend([per_page, offset])
        rows = conn.execute(sql, params).fetchall()

        return {
            "entries": [dict(r) for r in rows],
            "total": total,
            "page": page,
            "per_page": per_page,
        }
    finally:
        conn.close()


def get_recent_entries(limit=5, user_id: int = 0) -> list:
    """Get the most recent journal entries (user-scoped)."""
    conn = get_db()
    try:
        rows = conn.execute("""
        SELECT
            je.id, je.entry_date, je.amount, je.tax_classification, je.tax_amount,
            je.counterparty, je.memo, je.evidence_url, je.source,
            da.name AS debit_account, ca.name AS credit_account
        FROM journal_entries je
        JOIN accounts_master da ON je.debit_account_id = da.id
        JOIN accounts_master ca ON je.credit_account_id = ca.id
        WHERE je.is_deleted = 0 AND je.user_id = ?
        ORDER BY je.id DESC
        LIMIT ?
        """, (user_id, limit)).fetchall()
        return [dict(r) for r in rows]
    finally:
        conn.close()


def update_journal_entry(entry_id: int, entry: dict, user_id: int = 0) -> bool:
    """Update a journal entry (user-scoped ownership check)."""
    conn = get_db()
    try:
        amount = int(entry.get('amount', 0))
        tax_class = entry.get('tax_classification', '10%')
        tax_amount = calculate_tax_amount(amount, tax_class)

        debit_id = _resolve_account_id(conn, entry.get('debit_account_id'), entry.get('debit_account'))
        credit_id = _resolve_account_id(conn, entry.get('credit_account_id'), entry.get('credit_account'))

        if not debit_id or not credit_id:
            return False

        conn.execute(
            """UPDATE journal_entries SET
                entry_date = ?, debit_account_id = ?, credit_account_id = ?,
                amount = ?, tax_classification = ?, tax_amount = ?,
                counterparty = ?, memo = ?, updated_at = datetime('now')
            WHERE id = ? AND is_deleted = 0 AND user_id = ?""",
            (
                entry.get('entry_date', entry.get('date', '')),
                debit_id, credit_id, amount, tax_class, tax_amount,
                entry.get('counterparty', ''), entry.get('memo', ''),
                entry_id, user_id,
            )
        )
        conn.commit()
        return True
    except Exception as e:
        conn.rollback()
        print(f"Update error: {e}")
        return False
    finally:
        conn.close()


def delete_journal_entry(entry_id: int, user_id: int = 0) -> bool:
    """Soft-delete a journal entry (user-scoped ownership check)."""
    conn = get_db()
    try:
        conn.execute(
            "UPDATE journal_entries SET is_deleted = 1, updated_at = datetime('now') WHERE id = ? AND user_id = ?",
            (entry_id, user_id)
        )
        conn.commit()
        return True
    except Exception as e:
        conn.rollback()
        print(f"Delete error: {e}")
        return False
    finally:
        conn.close()


# --- Trial Balance (user-scoped) ---
def get_trial_balance(start_date=None, end_date=None, user_id: int = 0) -> list:
    """Get trial balance grouped by account (user-scoped)."""
    conn = get_db()
    try:
        accounts = conn.execute(
            "SELECT id, code, name, account_type, display_order FROM accounts_master WHERE is_active = 1 ORDER BY display_order, code"
        ).fetchall()

        date_conditions = ["je.is_deleted = 0", "je.user_id = ?"]
        date_params = [user_id]
        if start_date:
            date_conditions.append("je.entry_date >= ?")
            date_params.append(start_date)
        if end_date:
            date_conditions.append("je.entry_date <= ?")
            date_params.append(end_date)
        date_where = " AND ".join(date_conditions)

        debit_sql = f"""
        SELECT debit_account_id AS account_id, SUM(amount) AS total
        FROM journal_entries je WHERE {date_where}
        GROUP BY debit_account_id
        """
        debit_totals = {row['account_id']: row['total'] for row in conn.execute(debit_sql, date_params).fetchall()}

        credit_sql = f"""
        SELECT credit_account_id AS account_id, SUM(amount) AS total
        FROM journal_entries je WHERE {date_where}
        GROUP BY credit_account_id
        """
        credit_totals = {row['account_id']: row['total'] for row in conn.execute(credit_sql, date_params).fetchall()}

        fiscal_year = start_date[:4] if start_date else "2025"
        opening_sql = "SELECT account_id, amount FROM opening_balances WHERE fiscal_year = ? AND user_id = ?"
        openings = {row['account_id']: row['amount'] for row in conn.execute(opening_sql, (fiscal_year, user_id)).fetchall()}

        cf_debit = {}
        cf_credit = {}
        if start_date:
            cf_cond = "je.is_deleted = 0 AND je.user_id = ? AND je.entry_date < ?"
            cf_params = [user_id, start_date]
            fy_start = fiscal_year + "-01-01"
            if start_date > fy_start:
                cf_cond += " AND je.entry_date >= ?"
                cf_params.append(fy_start)

            cf_debit_sql = f"SELECT debit_account_id AS account_id, SUM(amount) AS total FROM journal_entries je WHERE {cf_cond} GROUP BY debit_account_id"
            cf_debit = {row['account_id']: row['total'] for row in conn.execute(cf_debit_sql, cf_params).fetchall()}

            cf_credit_sql = f"SELECT credit_account_id AS account_id, SUM(amount) AS total FROM journal_entries je WHERE {cf_cond} GROUP BY credit_account_id"
            cf_credit = {row['account_id']: row['total'] for row in conn.execute(cf_credit_sql, cf_params).fetchall()}

        result = []
        for acc in accounts:
            aid = acc['id']
            atype = acc['account_type']
            opening = openings.get(aid, 0)
            debit = debit_totals.get(aid, 0)
            credit = credit_totals.get(aid, 0)

            cf_d = cf_debit.get(aid, 0)
            cf_c = cf_credit.get(aid, 0)
            if atype in ('資産', '費用'):
                carry_forward = opening + cf_d - cf_c
                closing = carry_forward + debit - credit
            else:
                carry_forward = opening + cf_c - cf_d
                closing = carry_forward + credit - debit

            result.append({
                "account_id": aid,
                "code": acc['code'],
                "name": acc['name'],
                "account_type": atype,
                "opening_balance": opening,
                "carry_forward": carry_forward,
                "debit_total": debit,
                "credit_total": credit,
                "closing_balance": closing,
            })

        return result
    finally:
        conn.close()


# --- Counterparties (user-scoped) ---
def get_counterparty_names(user_id: int = 0) -> list:
    """Get counterparty names for autocomplete (user-scoped)."""
    conn = get_db()
    try:
        rows = conn.execute("""
        SELECT name FROM counterparties WHERE is_active = 1 AND user_id = ?
        UNION
        SELECT DISTINCT counterparty AS name FROM journal_entries WHERE is_deleted = 0 AND counterparty != '' AND user_id = ?
        ORDER BY name
        """, (user_id, user_id)).fetchall()
        return [r['name'] for r in rows]
    finally:
        conn.close()


def get_counterparties_list(user_id: int = 0) -> list:
    """Get all counterparties with full details (user-scoped)."""
    conn = get_db()
    try:
        rows = conn.execute(
            "SELECT id, name, code, contact_info, notes, is_active FROM counterparties WHERE is_active = 1 AND user_id = ? ORDER BY name",
            (user_id,)
        ).fetchall()
        return [dict(r) for r in rows]
    finally:
        conn.close()


def create_counterparty(data: dict, user_id: int = 0) -> int:
    """Create a new counterparty (user-scoped)."""
    conn = get_db()
    try:
        cursor = conn.execute(
            "INSERT INTO counterparties (user_id, name, code, contact_info, notes) VALUES (?, ?, ?, ?, ?)",
            (user_id, data.get('name', ''), data.get('code', ''), data.get('contact_info', ''), data.get('notes', ''))
        )
        conn.commit()
        return cursor.lastrowid
    except Exception as e:
        conn.rollback()
        raise e
    finally:
        conn.close()


def update_counterparty(cp_id: int, data: dict, user_id: int = 0) -> bool:
    """Update a counterparty (user-scoped ownership check)."""
    conn = get_db()
    try:
        conn.execute(
            "UPDATE counterparties SET name=?, code=?, contact_info=?, notes=?, updated_at=datetime('now') WHERE id=? AND user_id=?",
            (data.get('name', ''), data.get('code', ''), data.get('contact_info', ''), data.get('notes', ''), cp_id, user_id)
        )
        conn.commit()
        return True
    except Exception:
        conn.rollback()
        return False
    finally:
        conn.close()


def delete_counterparty(cp_id: int, user_id: int = 0) -> bool:
    """Soft-delete a counterparty (user-scoped ownership check)."""
    conn = get_db()
    try:
        conn.execute("UPDATE counterparties SET is_active=0, updated_at=datetime('now') WHERE id=? AND user_id=?", (cp_id, user_id))
        conn.commit()
        return True
    except Exception:
        conn.rollback()
        return False
    finally:
        conn.close()


# --- Opening Balances (user-scoped) ---
def get_opening_balances(fiscal_year: str, user_id: int = 0) -> list:
    """Get opening balances for a fiscal year (user-scoped)."""
    conn = get_db()
    try:
        rows = conn.execute("""
        SELECT am.id AS account_id, am.code, am.name, am.account_type,
               COALESCE(ob.amount, 0) AS amount, COALESCE(ob.note, '') AS note
        FROM accounts_master am
        LEFT JOIN opening_balances ob ON am.id = ob.account_id AND ob.fiscal_year = ? AND ob.user_id = ?
        WHERE am.is_active = 1
        ORDER BY am.display_order, am.code
        """, (fiscal_year, user_id)).fetchall()
        return [dict(r) for r in rows]
    finally:
        conn.close()


def save_opening_balances(fiscal_year: str, balances: list, user_id: int = 0) -> bool:
    """Bulk upsert opening balances for a fiscal year (user-scoped)."""
    conn = get_db()
    try:
        for b in balances:
            # Check if exists
            existing = conn.execute(
                "SELECT id FROM opening_balances WHERE user_id = ? AND fiscal_year = ? AND account_id = ?",
                (user_id, fiscal_year, b['account_id'])
            ).fetchone()
            if existing:
                conn.execute(
                    "UPDATE opening_balances SET amount = ?, note = ? WHERE id = ?",
                    (int(b.get('amount', 0)), b.get('note', ''), existing['id'])
                )
            else:
                conn.execute(
                    "INSERT INTO opening_balances (user_id, fiscal_year, account_id, amount, note) VALUES (?, ?, ?, ?, ?)",
                    (user_id, fiscal_year, b['account_id'], int(b.get('amount', 0)), b.get('note', ''))
                )
        conn.commit()
        return True
    except Exception:
        conn.rollback()
        return False
    finally:
        conn.close()


# --- General Ledger (user-scoped) ---
def get_ledger_entries(account_id: int, start_date=None, end_date=None, user_id: int = 0) -> dict:
    """Get journal entries for a specific account with running balance (user-scoped)."""
    conn = get_db()
    try:
        account = conn.execute(
            "SELECT id, code, name, account_type FROM accounts_master WHERE id = ?", (account_id,)
        ).fetchone()
        if not account:
            return {"error": "Account not found"}
        account = dict(account)
        atype = account['account_type']

        conditions = ["je.is_deleted = 0", "je.user_id = ?", "(je.debit_account_id = ? OR je.credit_account_id = ?)"]
        params = [user_id, account_id, account_id]
        if start_date:
            conditions.append("je.entry_date >= ?")
            params.append(start_date)
        if end_date:
            conditions.append("je.entry_date <= ?")
            params.append(end_date)
        where = " AND ".join(conditions)

        rows = conn.execute(f"""
        SELECT je.id, je.entry_date, je.amount, je.tax_classification,
               je.counterparty, je.memo,
               je.debit_account_id, je.credit_account_id,
               da.name AS debit_account, ca.name AS credit_account
        FROM journal_entries je
        JOIN accounts_master da ON je.debit_account_id = da.id
        JOIN accounts_master ca ON je.credit_account_id = ca.id
        WHERE {where}
        ORDER BY je.entry_date ASC, je.id ASC
        """, params).fetchall()

        fiscal_year = start_date[:4] if start_date else str(__import__('datetime').date.today().year)
        ob_row = conn.execute(
            "SELECT amount FROM opening_balances WHERE fiscal_year = ? AND account_id = ? AND user_id = ?",
            (fiscal_year, account_id, user_id)
        ).fetchone()
        opening = ob_row['amount'] if ob_row else 0

        entries = []
        balance = opening
        for r in rows:
            r = dict(r)
            debit_amount = r['amount'] if r['debit_account_id'] == account_id else 0
            credit_amount = r['amount'] if r['credit_account_id'] == account_id else 0
            if atype in ('資産', '費用'):
                balance += debit_amount - credit_amount
            else:
                balance += credit_amount - debit_amount
            r['debit_amount'] = debit_amount
            r['credit_amount'] = credit_amount
            r['balance'] = balance
            if r['debit_account_id'] == account_id:
                r['counter_account'] = r['credit_account']
            else:
                r['counter_account'] = r['debit_account']
            entries.append(r)

        return {"account": account, "opening_balance": opening, "entries": entries}
    finally:
        conn.close()


# --- Backup (user-scoped) ---
def get_full_backup(user_id: int = 0) -> dict:
    """Export user's data as a JSON structure."""
    conn = get_db()
    try:
        result = {}
        # Accounts master is global
        rows = conn.execute("SELECT * FROM accounts_master").fetchall()
        result['accounts_master'] = [dict(r) for r in rows]

        # User-scoped tables
        rows = conn.execute("SELECT * FROM journal_entries WHERE user_id = ?", (user_id,)).fetchall()
        result['journal_entries'] = [dict(r) for r in rows]

        rows = conn.execute("SELECT * FROM opening_balances WHERE user_id = ?", (user_id,)).fetchall()
        result['opening_balances'] = [dict(r) for r in rows]

        rows = conn.execute("SELECT * FROM counterparties WHERE user_id = ?", (user_id,)).fetchall()
        result['counterparties'] = [dict(r) for r in rows]

        rows = conn.execute("SELECT key, value FROM user_settings WHERE user_id = ?", (user_id,)).fetchall()
        result['settings'] = [{'key': r['key'], 'value': r['value']} for r in rows]

        return result
    finally:
        conn.close()


def restore_from_backup(data: dict, user_id: int = 0) -> dict:
    """Restore database from a JSON backup (user-scoped)."""
    conn = get_db()
    try:
        conn.execute("PRAGMA foreign_keys=OFF")
        summary = {}

        # Restore accounts_master (global)
        acct_rows = data.get('accounts_master', [])
        if acct_rows:
            conn.execute("DELETE FROM accounts_master")
            columns = list(acct_rows[0].keys())
            placeholders = ', '.join(['?' for _ in columns])
            col_names = ', '.join(columns)
            inserted = 0
            for row in acct_rows:
                try:
                    values = [row.get(col) for col in columns]
                    conn.execute(f"INSERT INTO accounts_master ({col_names}) VALUES ({placeholders})", values)
                    inserted += 1
                except Exception as e:
                    print(f"Restore skip accounts_master: {e}")
            summary['accounts_master'] = inserted

        # Clear user's data only
        conn.execute("DELETE FROM journal_entries WHERE user_id = ?", (user_id,))
        conn.execute("DELETE FROM opening_balances WHERE user_id = ?", (user_id,))
        conn.execute("DELETE FROM counterparties WHERE user_id = ?", (user_id,))
        conn.execute("DELETE FROM user_settings WHERE user_id = ?", (user_id,))

        # Restore journal_entries
        je_rows = data.get('journal_entries', [])
        inserted = 0
        for row in je_rows:
            try:
                row_user_id = user_id  # Force to current user
                conn.execute(
                    """INSERT INTO journal_entries (user_id, entry_date, debit_account_id, credit_account_id,
                       amount, tax_classification, tax_amount, counterparty, memo, evidence_url, source, is_deleted, created_at, updated_at)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                    (row_user_id, row.get('entry_date'), row.get('debit_account_id'), row.get('credit_account_id'),
                     row.get('amount', 0), row.get('tax_classification', '10%'), row.get('tax_amount', 0),
                     row.get('counterparty', ''), row.get('memo', ''), row.get('evidence_url', ''),
                     row.get('source', 'manual'), row.get('is_deleted', 0),
                     row.get('created_at', ''), row.get('updated_at', ''))
                )
                inserted += 1
            except Exception as e:
                print(f"Restore skip journal_entries: {e}")
        summary['journal_entries'] = inserted

        # Restore opening_balances
        ob_rows = data.get('opening_balances', [])
        inserted = 0
        for row in ob_rows:
            try:
                conn.execute(
                    "INSERT INTO opening_balances (user_id, fiscal_year, account_id, amount, note) VALUES (?, ?, ?, ?, ?)",
                    (user_id, row.get('fiscal_year'), row.get('account_id'), row.get('amount', 0), row.get('note', ''))
                )
                inserted += 1
            except Exception as e:
                print(f"Restore skip opening_balances: {e}")
        summary['opening_balances'] = inserted

        # Restore counterparties
        cp_rows = data.get('counterparties', [])
        inserted = 0
        for row in cp_rows:
            try:
                conn.execute(
                    "INSERT INTO counterparties (user_id, name, code, contact_info, notes, is_active) VALUES (?, ?, ?, ?, ?, ?)",
                    (user_id, row.get('name', ''), row.get('code', ''), row.get('contact_info', ''), row.get('notes', ''), row.get('is_active', 1))
                )
                inserted += 1
            except Exception as e:
                print(f"Restore skip counterparties: {e}")
        summary['counterparties'] = inserted

        # Restore settings
        settings_rows = data.get('settings', [])
        inserted = 0
        for row in settings_rows:
            try:
                key = row.get('key', '')
                value = row.get('value', '')
                if key:
                    conn.execute(
                        "INSERT OR REPLACE INTO user_settings (user_id, key, value) VALUES (?, ?, ?)",
                        (user_id, key, value)
                    )
                    inserted += 1
            except Exception as e:
                print(f"Restore skip settings: {e}")
        summary['settings'] = inserted

        conn.execute("PRAGMA foreign_keys=ON")
        conn.commit()
        return summary
    except Exception as e:
        conn.rollback()
        raise e
    finally:
        conn.close()


# --- Export (user-scoped) ---
def get_journal_export(start_date=None, end_date=None, user_id: int = 0) -> list:
    """Get all journal entries for export (user-scoped)."""
    conn = get_db()
    try:
        conditions = ["je.is_deleted = 0", "je.user_id = ?"]
        params = [user_id]
        if start_date:
            conditions.append("je.entry_date >= ?")
            params.append(start_date)
        if end_date:
            conditions.append("je.entry_date <= ?")
            params.append(end_date)
        where = " AND ".join(conditions)
        rows = conn.execute(f"""
        SELECT je.entry_date, da.code AS debit_code, da.name AS debit_account,
               ca.code AS credit_code, ca.name AS credit_account,
               je.amount, je.tax_classification, je.tax_amount,
               je.counterparty, je.memo
        FROM journal_entries je
        JOIN accounts_master da ON je.debit_account_id = da.id
        JOIN accounts_master ca ON je.credit_account_id = ca.id
        WHERE {where}
        ORDER BY je.entry_date ASC, je.id ASC
        """, params).fetchall()
        return [dict(r) for r in rows]
    finally:
        conn.close()


# --- History for AI (user-scoped) ---
def get_accounting_history(limit=200, user_id: int = 0) -> list:
    """Get recent journal entries for AI context (user-scoped)."""
    conn = get_db()
    try:
        rows = conn.execute("""
        SELECT je.counterparty, je.memo, da.name AS account
        FROM journal_entries je
        JOIN accounts_master da ON je.debit_account_id = da.id
        WHERE je.is_deleted = 0 AND je.counterparty != '' AND je.user_id = ?
        ORDER BY je.id DESC LIMIT ?
        """, (user_id, limit)).fetchall()
        return [dict(r) for r in rows]
    finally:
        conn.close()


def get_existing_entry_keys(user_id: int = 0) -> set:
    """Get set of 'date_amount_counterparty' keys for duplicate detection (user-scoped)."""
    conn = get_db()
    try:
        rows = conn.execute(
            "SELECT entry_date, amount, counterparty FROM journal_entries WHERE is_deleted = 0 AND user_id = ?",
            (user_id,)
        ).fetchall()
        return {f"{r['entry_date']}_{r['amount']}_{r['counterparty']}" for r in rows}
    finally:
        conn.close()


# --- Helpers ---
def _resolve_account_id(conn, account_id=None, account_name=None) -> int:
    """Resolve account ID from either ID or name."""
    if account_id:
        return int(account_id)
    if account_name:
        row = conn.execute(
            "SELECT id FROM accounts_master WHERE name = ? AND is_active = 1",
            (account_name.strip(),)
        ).fetchone()
        if row:
            return row['id']
        row = conn.execute(
            "SELECT id FROM accounts_master WHERE name = '雑費' AND is_active = 1"
        ).fetchone()
        return row['id'] if row else None
    return None


# ============================================================
#  Import History & Statement Sources
# ============================================================

def get_import_by_hash(file_hash: str, user_id: int = 0):
    """Check if a file with this hash has already been imported."""
    conn = get_db()
    try:
        row = conn.execute(
            "SELECT * FROM import_history WHERE file_hash = ? AND user_id = ? ORDER BY created_at DESC LIMIT 1",
            (file_hash, user_id)
        ).fetchone()
        return dict(row) if row else None
    except Exception:
        return None
    finally:
        conn.close()


def create_import_record(user_id: int, filename: str, file_hash: str,
                          source_name: str, row_count: int, imported_count: int,
                          date_range_start: str = '', date_range_end: str = '') -> int:
    """Record a completed import."""
    conn = get_db()
    try:
        cursor = conn.execute(
            """INSERT INTO import_history
            (user_id, filename, file_hash, source_name, row_count, imported_count,
             date_range_start, date_range_end)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)""",
            (user_id, filename, file_hash, source_name, row_count,
             imported_count, date_range_start, date_range_end)
        )
        conn.commit()
        return cursor.lastrowid
    except Exception as e:
        print(f"Import record error: {e}")
        conn.rollback()
        return 0
    finally:
        conn.close()


def get_import_history(user_id: int = 0, limit: int = 20) -> list:
    """Get recent import history."""
    conn = get_db()
    try:
        rows = conn.execute(
            "SELECT * FROM import_history WHERE user_id = ? ORDER BY created_at DESC LIMIT ?",
            (user_id, limit)
        ).fetchall()
        return [dict(r) for r in rows]
    except Exception:
        return []
    finally:
        conn.close()


def get_statement_source(source_name: str, user_id: int = 0):
    """Get saved source mapping for a specific source."""
    conn = get_db()
    try:
        row = conn.execute(
            "SELECT * FROM statement_sources WHERE source_name = ? AND user_id = ?",
            (source_name, user_id)
        ).fetchone()
        return dict(row) if row else None
    except Exception:
        return None
    finally:
        conn.close()


def upsert_statement_source(user_id: int, source_name: str,
                             default_debit: str = '', default_credit: str = ''):
    """Save or update statement source mapping."""
    conn = get_db()
    try:
        conn.execute(
            """INSERT INTO statement_sources (user_id, source_name, default_debit, default_credit)
            VALUES (?, ?, ?, ?)
            ON CONFLICT(user_id, source_name)
            DO UPDATE SET default_debit = ?, default_credit = ?, updated_at = datetime('now')""",
            (user_id, source_name, default_debit, default_credit,
             default_debit, default_credit)
        )
        conn.commit()
    except Exception as e:
        print(f"Statement source upsert error: {e}")
        conn.rollback()
    finally:
        conn.close()
