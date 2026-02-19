"""
Data access layer for AI Accounting Tool.
CRUD operations for journal entries, accounts, trial balance, etc.
All user-scoped functions require user_id parameter.
"""
import math
import secrets
import string
from db import get_db, P, USE_PG

# --- DB dialect helpers ---
_NOW = "CURRENT_TIMESTAMP" if USE_PG else "datetime('now')"

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
        row = conn.execute(P("SELECT id, email, name, picture FROM users WHERE email = ?"), (email,)).fetchone()
        if row:
            # Update last_login
            conn.execute(P(f"UPDATE users SET last_login = {_NOW}, name = ?, picture = ? WHERE id = ?"),
                         (name or row['name'], picture or row['picture'], row['id']))
            conn.commit()
            user = dict(row)
            # Migrate legacy data (user_id=0) to first user
            _migrate_legacy_data(conn, user['id'])
            return user
        else:
            if USE_PG:
                cur = conn.cursor()
                cur.execute(P("INSERT INTO users (email, name, picture) VALUES (?, ?, ?) RETURNING id"),
                            (email, name, picture))
                new_id = cur.fetchone()['id']
            else:
                cur = conn.execute("INSERT INTO users (email, name, picture) VALUES (?, ?, ?)",
                                   (email, name, picture))
                new_id = cur.lastrowid
            conn.commit()
            user = {'id': new_id, 'email': email, 'name': name, 'picture': picture}
            # Migrate legacy data to first registered user
            _migrate_legacy_data(conn, new_id)
            return user
    finally:
        conn.close()


def _migrate_legacy_data(conn, user_id: int):
    """Migrate user_id=0 data to the given user (first-come gets legacy data)."""
    # Check if there's any legacy data
    row = conn.execute("SELECT COUNT(*) AS cnt FROM journal_entries WHERE user_id = 0").fetchone()
    legacy_count = row['cnt'] if USE_PG else row[0]
    if legacy_count == 0:
        return

    # Only migrate if no other user has claimed this data yet
    row2 = conn.execute("SELECT COUNT(*) AS cnt FROM journal_entries WHERE user_id != 0").fetchone()
    any_claimed = row2['cnt'] if USE_PG else row2[0]
    if any_claimed > 0:
        return  # Another user already has data, don't re-migrate

    conn.execute(P("UPDATE journal_entries SET user_id = ? WHERE user_id = 0"), (user_id,))
    conn.execute(P("UPDATE opening_balances SET user_id = ? WHERE user_id = 0"), (user_id,))
    conn.execute(P("UPDATE counterparties SET user_id = ? WHERE user_id = 0"), (user_id,))
    conn.execute(P("UPDATE user_settings SET user_id = ? WHERE user_id = 0"), (user_id,))
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
            P("INSERT INTO accounts_master (code, name, account_type, tax_default, display_order) VALUES (?, ?, ?, ?, ?)"),
            (code, name, account_type, tax_default, display_order)
        )
        conn.commit()
        row = conn.execute(
            P("SELECT id, code, name, account_type, tax_default, display_order FROM accounts_master WHERE code = ?"),
            (code,)
        ).fetchone()
        return dict(row) if row else {}
    finally:
        conn.close()


def delete_account(account_id: int) -> bool:
    """Soft-delete an account (set is_active = 0). Prevents deletion if used in journal entries."""
    conn = get_db()
    try:
        row = conn.execute(
            P("SELECT COUNT(*) AS cnt FROM journal_entries WHERE (debit_account_id = ? OR credit_account_id = ?) AND is_deleted = 0"),
            (account_id, account_id)
        ).fetchone()
        used = row['cnt'] if USE_PG else row[0]
        if used > 0:
            return False
        conn.execute(
            P(f"UPDATE accounts_master SET is_active = 0, updated_at = {_NOW} WHERE id = ?"),
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
            P("SELECT id, code, name, account_type, tax_default FROM accounts_master WHERE name = ? AND is_active = 1"),
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
            P("SELECT id, code, name, account_type, tax_default FROM accounts_master WHERE id = ?"),
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

        _params = (
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
        _sql = """INSERT INTO journal_entries
            (user_id, entry_date, debit_account_id, credit_account_id, amount, tax_classification, tax_amount,
             counterparty, memo, evidence_url, source)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"""
        if USE_PG:
            cur = conn.cursor()
            cur.execute(P(_sql + " RETURNING id"), _params)
            new_id = cur.fetchone()['id']
        else:
            cur = conn.execute(_sql, _params)
            new_id = cur.lastrowid
        conn.commit()
        return new_id
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

        count_sql = f"SELECT COUNT(*) AS cnt FROM journal_entries je WHERE {where}"
        row = conn.execute(P(count_sql), params).fetchone()
        total = row['cnt'] if USE_PG else row[0]

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
        rows = conn.execute(P(sql), params).fetchall()

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
        rows = conn.execute(P("""
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
        """), (user_id, limit)).fetchall()
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
            P(f"""UPDATE journal_entries SET
                entry_date = ?, debit_account_id = ?, credit_account_id = ?,
                amount = ?, tax_classification = ?, tax_amount = ?,
                counterparty = ?, memo = ?, updated_at = {_NOW}
            WHERE id = ? AND is_deleted = 0 AND user_id = ?"""),
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
            P(f"UPDATE journal_entries SET is_deleted = 1, updated_at = {_NOW} WHERE id = ? AND user_id = ?"),
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
        debit_totals = {row['account_id']: row['total'] for row in conn.execute(P(debit_sql), date_params).fetchall()}

        credit_sql = f"""
        SELECT credit_account_id AS account_id, SUM(amount) AS total
        FROM journal_entries je WHERE {date_where}
        GROUP BY credit_account_id
        """
        credit_totals = {row['account_id']: row['total'] for row in conn.execute(P(credit_sql), date_params).fetchall()}

        fiscal_year = start_date[:4] if start_date else "2025"
        opening_sql = P("SELECT account_id, amount FROM opening_balances WHERE fiscal_year = ? AND user_id = ?")
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
            cf_debit = {row['account_id']: row['total'] for row in conn.execute(P(cf_debit_sql), cf_params).fetchall()}

            cf_credit_sql = f"SELECT credit_account_id AS account_id, SUM(amount) AS total FROM journal_entries je WHERE {cf_cond} GROUP BY credit_account_id"
            cf_credit = {row['account_id']: row['total'] for row in conn.execute(P(cf_credit_sql), cf_params).fetchall()}

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
        rows = conn.execute(P("""
        SELECT name FROM counterparties WHERE is_active = 1 AND user_id = ?
        UNION
        SELECT DISTINCT counterparty AS name FROM journal_entries WHERE is_deleted = 0 AND counterparty != '' AND user_id = ?
        ORDER BY name
        """), (user_id, user_id)).fetchall()
        return [r['name'] for r in rows]
    finally:
        conn.close()


def get_counterparties_list(user_id: int = 0) -> list:
    """Get all counterparties with full details (user-scoped)."""
    conn = get_db()
    try:
        rows = conn.execute(
            P("SELECT id, name, code, contact_info, notes, is_active FROM counterparties WHERE is_active = 1 AND user_id = ? ORDER BY name"),
            (user_id,)
        ).fetchall()
        return [dict(r) for r in rows]
    finally:
        conn.close()


def create_counterparty(data: dict, user_id: int = 0) -> int:
    """Create a new counterparty (user-scoped)."""
    conn = get_db()
    try:
        _sql = "INSERT INTO counterparties (user_id, name, code, contact_info, notes) VALUES (?, ?, ?, ?, ?)"
        _params = (user_id, data.get('name', ''), data.get('code', ''), data.get('contact_info', ''), data.get('notes', ''))
        if USE_PG:
            cur = conn.cursor()
            cur.execute(P(_sql + " RETURNING id"), _params)
            new_id = cur.fetchone()['id']
        else:
            cur = conn.execute(_sql, _params)
            new_id = cur.lastrowid
        conn.commit()
        return new_id
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
            P(f"UPDATE counterparties SET name=?, code=?, contact_info=?, notes=?, updated_at={_NOW} WHERE id=? AND user_id=?"),
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
        conn.execute(P(f"UPDATE counterparties SET is_active=0, updated_at={_NOW} WHERE id=? AND user_id=?"), (cp_id, user_id))
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
        rows = conn.execute(P("""
        SELECT am.id AS account_id, am.code, am.name, am.account_type,
               COALESCE(ob.amount, 0) AS amount, COALESCE(ob.note, '') AS note
        FROM accounts_master am
        LEFT JOIN opening_balances ob ON am.id = ob.account_id AND ob.fiscal_year = ? AND ob.user_id = ?
        WHERE am.is_active = 1
        ORDER BY am.display_order, am.code
        """), (fiscal_year, user_id)).fetchall()
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
                P("SELECT id FROM opening_balances WHERE user_id = ? AND fiscal_year = ? AND account_id = ?"),
                (user_id, fiscal_year, b['account_id'])
            ).fetchone()
            if existing:
                conn.execute(
                    P("UPDATE opening_balances SET amount = ?, note = ? WHERE id = ?"),
                    (int(b.get('amount', 0)), b.get('note', ''), existing['id'])
                )
            else:
                conn.execute(
                    P("INSERT INTO opening_balances (user_id, fiscal_year, account_id, amount, note) VALUES (?, ?, ?, ?, ?)"),
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
            P("SELECT id, code, name, account_type FROM accounts_master WHERE id = ?"), (account_id,)
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

        rows = conn.execute(P(f"""
        SELECT je.id, je.entry_date, je.amount, je.tax_classification,
               je.counterparty, je.memo,
               je.debit_account_id, je.credit_account_id,
               da.name AS debit_account, ca.name AS credit_account
        FROM journal_entries je
        JOIN accounts_master da ON je.debit_account_id = da.id
        JOIN accounts_master ca ON je.credit_account_id = ca.id
        WHERE {where}
        ORDER BY je.entry_date ASC, je.id ASC
        """), params).fetchall()

        fiscal_year = start_date[:4] if start_date else str(__import__('datetime').date.today().year)
        ob_row = conn.execute(
            P("SELECT amount FROM opening_balances WHERE fiscal_year = ? AND account_id = ? AND user_id = ?"),
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
        rows = conn.execute(P("SELECT * FROM journal_entries WHERE user_id = ?"), (user_id,)).fetchall()
        result['journal_entries'] = [dict(r) for r in rows]

        rows = conn.execute(P("SELECT * FROM opening_balances WHERE user_id = ?"), (user_id,)).fetchall()
        result['opening_balances'] = [dict(r) for r in rows]

        rows = conn.execute(P("SELECT * FROM counterparties WHERE user_id = ?"), (user_id,)).fetchall()
        result['counterparties'] = [dict(r) for r in rows]

        rows = conn.execute(P("SELECT key, value FROM user_settings WHERE user_id = ?"), (user_id,)).fetchall()
        result['settings'] = [{'key': r['key'], 'value': r['value']} for r in rows]

        return result
    finally:
        conn.close()


def restore_from_backup(data: dict, user_id: int = 0) -> dict:
    """Restore database from a JSON backup (user-scoped)."""
    conn = get_db()
    try:
        if not USE_PG:
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
                    conn.execute(P(f"INSERT INTO accounts_master ({col_names}) VALUES ({placeholders})"), values)
                    inserted += 1
                except Exception as e:
                    print(f"Restore skip accounts_master: {e}")
            summary['accounts_master'] = inserted

        # Clear user's data only
        conn.execute(P("DELETE FROM journal_entries WHERE user_id = ?"), (user_id,))
        conn.execute(P("DELETE FROM opening_balances WHERE user_id = ?"), (user_id,))
        conn.execute(P("DELETE FROM counterparties WHERE user_id = ?"), (user_id,))
        conn.execute(P("DELETE FROM user_settings WHERE user_id = ?"), (user_id,))

        # Restore journal_entries
        je_rows = data.get('journal_entries', [])
        inserted = 0
        for row in je_rows:
            try:
                row_user_id = user_id  # Force to current user
                conn.execute(
                    P("""INSERT INTO journal_entries (user_id, entry_date, debit_account_id, credit_account_id,
                       amount, tax_classification, tax_amount, counterparty, memo, evidence_url, source, is_deleted, created_at, updated_at)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"""),
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
                    P("INSERT INTO opening_balances (user_id, fiscal_year, account_id, amount, note) VALUES (?, ?, ?, ?, ?)"),
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
                    P("INSERT INTO counterparties (user_id, name, code, contact_info, notes, is_active) VALUES (?, ?, ?, ?, ?, ?)"),
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
                    if USE_PG:
                        conn.execute(
                            "INSERT INTO user_settings (user_id, key, value) VALUES (%s, %s, %s) ON CONFLICT (user_id, key) DO UPDATE SET value = EXCLUDED.value",
                            (user_id, key, value)
                        )
                    else:
                        conn.execute(
                            "INSERT OR REPLACE INTO user_settings (user_id, key, value) VALUES (?, ?, ?)",
                            (user_id, key, value)
                        )
                    inserted += 1
            except Exception as e:
                print(f"Restore skip settings: {e}")
        summary['settings'] = inserted

        if not USE_PG:
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
        rows = conn.execute(P(f"""
        SELECT je.entry_date, da.code AS debit_code, da.name AS debit_account,
               ca.code AS credit_code, ca.name AS credit_account,
               je.amount, je.tax_classification, je.tax_amount,
               je.counterparty, je.memo
        FROM journal_entries je
        JOIN accounts_master da ON je.debit_account_id = da.id
        JOIN accounts_master ca ON je.credit_account_id = ca.id
        WHERE {where}
        ORDER BY je.entry_date ASC, je.id ASC
        """), params).fetchall()
        return [dict(r) for r in rows]
    finally:
        conn.close()


# --- History for AI (user-scoped) ---
def get_accounting_history(limit=200, user_id: int = 0) -> list:
    """Get recent journal entries for AI context (user-scoped)."""
    conn = get_db()
    try:
        rows = conn.execute(P("""
        SELECT je.counterparty, je.memo, da.name AS account
        FROM journal_entries je
        JOIN accounts_master da ON je.debit_account_id = da.id
        WHERE je.is_deleted = 0 AND je.counterparty != '' AND je.user_id = ?
        ORDER BY je.id DESC LIMIT ?
        """), (user_id, limit)).fetchall()
        return [dict(r) for r in rows]
    finally:
        conn.close()


def get_existing_entry_keys(user_id: int = 0) -> set:
    """Get set of 'date_amount_counterparty' keys for duplicate detection (user-scoped)."""
    conn = get_db()
    try:
        rows = conn.execute(
            P("SELECT entry_date, amount, counterparty FROM journal_entries WHERE is_deleted = 0 AND user_id = ?"),
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
            P("SELECT id FROM accounts_master WHERE name = ? AND is_active = 1"),
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
#  Blue Return Tax Form (青色申告決算書) — Monthly Summary
# ============================================================

def get_monthly_summary(fiscal_year: str, user_id: int = 0) -> dict:
    """Get monthly revenue/expense breakdown for 青色申告決算書 page 2.
    Returns dict with:
      monthly: list of 12 dicts {month, revenue, purchases, expenses_by_account}
      annual:  totals for the year
    """
    conn = get_db()
    try:
        fy_start = f"{fiscal_year}-01-01"
        fy_end = f"{fiscal_year}-12-31"

        if USE_PG:
            month_expr = "to_char(je.entry_date::date, 'MM')"
        else:
            month_expr = "strftime('%m', je.entry_date)"

        # Monthly revenue (売上) by month
        revenue_sql = P(f"""
        SELECT {month_expr} AS month,
               SUM(je.amount) AS total
        FROM journal_entries je
        JOIN accounts_master am ON je.credit_account_id = am.id
        WHERE je.is_deleted = 0 AND je.user_id = ?
          AND je.entry_date >= ? AND je.entry_date <= ?
          AND am.account_type = '収益'
        GROUP BY {month_expr}
        """)
        rev_rows = conn.execute(revenue_sql, (user_id, fy_start, fy_end)).fetchall()
        rev_by_month = {int(r['month']): r['total'] for r in rev_rows}

        # Monthly purchases (仕入) by month -- code 500
        purchase_sql = P(f"""
        SELECT {month_expr} AS month,
               SUM(je.amount) AS total
        FROM journal_entries je
        JOIN accounts_master am ON je.debit_account_id = am.id
        WHERE je.is_deleted = 0 AND je.user_id = ?
          AND je.entry_date >= ? AND je.entry_date <= ?
          AND am.code = '500'
        GROUP BY {month_expr}
        """)
        pur_rows = conn.execute(purchase_sql, (user_id, fy_start, fy_end)).fetchall()
        pur_by_month = {int(r['month']): r['total'] for r in pur_rows}

        # Monthly expense totals by account (for page 1 detail)
        expense_sql = P(f"""
        SELECT am.code, am.name,
               {month_expr} AS month,
               SUM(je.amount) AS total
        FROM journal_entries je
        JOIN accounts_master am ON je.debit_account_id = am.id
        WHERE je.is_deleted = 0 AND je.user_id = ?
          AND je.entry_date >= ? AND je.entry_date <= ?
          AND am.account_type = '費用'
        GROUP BY am.code, am.name, {month_expr}
        """)
        exp_rows = conn.execute(expense_sql, (user_id, fy_start, fy_end)).fetchall()

        # Build per-account monthly breakdown
        expense_accounts = {}
        for r in exp_rows:
            code = r['code']
            if code not in expense_accounts:
                expense_accounts[code] = {'code': code, 'name': r['name'], 'months': {}, 'total': 0}
            m = int(r['month'])
            expense_accounts[code]['months'][m] = r['total']
            expense_accounts[code]['total'] += r['total']

        # Build monthly array
        monthly = []
        annual_revenue = 0
        annual_purchase = 0
        for m in range(1, 13):
            rev = rev_by_month.get(m, 0)
            pur = pur_by_month.get(m, 0)
            annual_revenue += rev
            annual_purchase += pur
            monthly.append({
                'month': m,
                'revenue': rev,
                'purchases': pur,
            })

        # Sort expense accounts by code
        sorted_expenses = sorted(expense_accounts.values(), key=lambda x: x['code'])
        annual_expense = sum(ea['total'] for ea in sorted_expenses)

        return {
            'fiscal_year': fiscal_year,
            'monthly': monthly,
            'expense_accounts': sorted_expenses,
            'annual_revenue': annual_revenue,
            'annual_purchase': annual_purchase,
            'annual_expense': annual_expense,
        }
    finally:
        conn.close()


# ============================================================
#  Import History & Statement Sources
# ============================================================

def get_import_by_hash(file_hash: str, user_id: int = 0):
    """Check if a file with this hash has already been imported."""
    conn = get_db()
    try:
        row = conn.execute(
            P("SELECT * FROM import_history WHERE file_hash = ? AND user_id = ? ORDER BY created_at DESC LIMIT 1"),
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
        _sql = """INSERT INTO import_history
            (user_id, filename, file_hash, source_name, row_count, imported_count,
             date_range_start, date_range_end)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)"""
        _params = (user_id, filename, file_hash, source_name, row_count,
             imported_count, date_range_start, date_range_end)
        if USE_PG:
            cur = conn.cursor()
            cur.execute(P(_sql + " RETURNING id"), _params)
            new_id = cur.fetchone()['id']
        else:
            cur = conn.execute(_sql, _params)
            new_id = cur.lastrowid
        conn.commit()
        return new_id
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
            P("SELECT * FROM import_history WHERE user_id = ? ORDER BY created_at DESC LIMIT ?"),
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
            P("SELECT * FROM statement_sources WHERE source_name = ? AND user_id = ?"),
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
            P(f"""INSERT INTO statement_sources (user_id, source_name, default_debit, default_credit)
            VALUES (?, ?, ?, ?)
            ON CONFLICT(user_id, source_name)
            DO UPDATE SET default_debit = ?, default_credit = ?, updated_at = {_NOW}"""),
            (user_id, source_name, default_debit, default_credit,
             default_debit, default_credit)
        )
        conn.commit()
    except Exception as e:
        print(f"Statement source upsert error: {e}")
        conn.rollback()
    finally:
        conn.close()


# ============================
#  License Key Management
# ============================

def generate_license_key() -> str:
    """Generate a HINA-XXXX-XXXX-XXXX format license key."""
    chars = string.ascii_uppercase + string.digits
    segments = [''.join(secrets.choice(chars) for _ in range(4)) for _ in range(3)]
    return f"HINA-{segments[0]}-{segments[1]}-{segments[2]}"


def create_license_key(created_by: str = '', notes: str = '') -> dict:
    """Create a single license key and store in DB."""
    conn = get_db()
    try:
        key = generate_license_key()
        if USE_PG:
            row = conn.execute(
                "INSERT INTO license_keys (license_key, created_by, notes) "
                "VALUES (%s, %s, %s) RETURNING id",
                (key, created_by, notes)
            ).fetchone()
            key_id = row['id']
        else:
            cur = conn.execute(
                "INSERT INTO license_keys (license_key, created_by, notes) VALUES (?, ?, ?)",
                (key, created_by, notes)
            )
            key_id = cur.lastrowid
        conn.commit()
        return {"id": key_id, "license_key": key}
    except Exception as e:
        conn.rollback()
        return {"error": str(e)}
    finally:
        conn.close()


def create_license_keys_batch(count: int, created_by: str = '', notes: str = '') -> list:
    """Generate multiple license keys at once."""
    results = []
    for _ in range(min(count, 100)):
        result = create_license_key(created_by, notes)
        results.append(result)
    return results


def activate_license(license_key_str: str, user_id: int, user_email: str) -> dict:
    """Activate a license key for a user."""
    conn = get_db()
    try:
        # 1. Find the key
        row = conn.execute(
            P("SELECT id, is_revoked FROM license_keys WHERE license_key = ?"),
            (license_key_str,)
        ).fetchone()
        if not row:
            return {"error": "無効なライセンスキーです"}
        key_id = row['id'] if USE_PG else row[0]
        is_revoked = row['is_revoked'] if USE_PG else row[1]
        if is_revoked:
            return {"error": "このライセンスキーは無効化されています"}

        # 2. Check if key is already used by another user
        existing = conn.execute(
            P("SELECT user_id FROM license_activations WHERE license_key_id = ? AND is_active = 1"),
            (key_id,)
        ).fetchone()
        if existing:
            existing_uid = existing['user_id'] if USE_PG else existing[0]
            if existing_uid != user_id:
                return {"error": "このライセンスキーは既に使用されています"}
            else:
                return {"status": "success", "message": "既に有効化済みです"}

        # 3. Check if user already has an active license
        user_lic = conn.execute(
            P("SELECT id FROM license_activations WHERE user_id = ? AND is_active = 1"),
            (user_id,)
        ).fetchone()
        if user_lic:
            return {"error": "既にライセンスが有効化されています"}

        # 4. Activate
        if USE_PG:
            conn.execute(
                "INSERT INTO license_activations (license_key_id, user_id, user_email) "
                "VALUES (%s, %s, %s)",
                (key_id, user_id, user_email)
            )
        else:
            conn.execute(
                "INSERT INTO license_activations (license_key_id, user_id, user_email) "
                "VALUES (?, ?, ?)",
                (key_id, user_id, user_email)
            )
        conn.commit()
        return {"status": "success", "message": "ライセンスが有効化されました"}
    except Exception as e:
        conn.rollback()
        return {"error": str(e)}
    finally:
        conn.close()


def check_user_license(user_id: int) -> bool:
    """Check if a user has an active license."""
    conn = get_db()
    try:
        row = conn.execute(
            P("""SELECT la.id FROM license_activations la
                JOIN license_keys lk ON la.license_key_id = lk.id
                WHERE la.user_id = ? AND la.is_active = 1 AND lk.is_revoked = 0"""),
            (user_id,)
        ).fetchone()
        return row is not None
    finally:
        conn.close()


def get_all_license_keys() -> list:
    """Get all license keys with activation info (admin)."""
    conn = get_db()
    try:
        rows = conn.execute(
            P("""SELECT lk.id, lk.license_key, lk.created_at, lk.created_by,
                    lk.is_revoked, lk.notes,
                    la.user_email, la.activated_at, la.is_active
                FROM license_keys lk
                LEFT JOIN license_activations la ON lk.id = la.license_key_id
                ORDER BY lk.id DESC""")
        ).fetchall()
        result = []
        for r in rows:
            if USE_PG:
                result.append(dict(r))
            else:
                result.append({
                    'id': r[0], 'license_key': r[1], 'created_at': r[2],
                    'created_by': r[3], 'is_revoked': r[4], 'notes': r[5],
                    'user_email': r[6], 'activated_at': r[7], 'is_active': r[8]
                })
        return result
    finally:
        conn.close()


def revoke_license(license_key_id: int) -> dict:
    """Revoke a license key (admin)."""
    _now = "CURRENT_TIMESTAMP" if USE_PG else "datetime('now')"
    conn = get_db()
    try:
        conn.execute(
            P(f"UPDATE license_keys SET is_revoked = 1, revoked_at = {_now} WHERE id = ?"),
            (license_key_id,)
        )
        conn.execute(
            P(f"UPDATE license_activations SET is_active = 0, deactivated_at = {_now} WHERE license_key_id = ?"),
            (license_key_id,)
        )
        conn.commit()
        return {"status": "success"}
    except Exception as e:
        conn.rollback()
        return {"error": str(e)}
    finally:
        conn.close()


# =====================================================
#  Fixed Assets (固定資産台帳)
# =====================================================

def create_fixed_asset(data: dict, user_id: int = 0) -> int:
    """Create a new fixed asset record."""
    conn = get_db()
    try:
        cur = conn.execute(
            P("""INSERT INTO fixed_assets
                (user_id, asset_name, acquisition_date, useful_life,
                 acquisition_cost, depreciation_method, notes)
                VALUES (?, ?, ?, ?, ?, ?, ?)"""),
            (user_id, data['asset_name'], data['acquisition_date'],
             int(data['useful_life']), int(data['acquisition_cost']),
             data.get('depreciation_method', '定額法'),
             data.get('notes', ''))
        )
        conn.commit()
        new_id = cur.lastrowid if not USE_PG else None
        if USE_PG:
            row = conn.execute("SELECT lastval()").fetchone()
            new_id = row['lastval']
        return new_id
    except Exception as e:
        conn.rollback()
        raise e
    finally:
        conn.close()


def get_fixed_assets(user_id: int = 0) -> list:
    """Get all active fixed assets for a user."""
    conn = get_db()
    try:
        rows = conn.execute(
            P("SELECT * FROM fixed_assets WHERE user_id = ? AND is_deleted = 0 ORDER BY acquisition_date DESC"),
            (user_id,)
        ).fetchall()
        return [dict(r) for r in rows]
    finally:
        conn.close()


def update_fixed_asset(asset_id: int, data: dict, user_id: int = 0) -> bool:
    """Update a fixed asset."""
    _now = "CURRENT_TIMESTAMP" if USE_PG else "datetime('now')"
    conn = get_db()
    try:
        conn.execute(
            P(f"""UPDATE fixed_assets SET
                asset_name = ?, acquisition_date = ?, useful_life = ?,
                acquisition_cost = ?, depreciation_method = ?, notes = ?,
                updated_at = {_now}
                WHERE id = ? AND user_id = ? AND is_deleted = 0"""),
            (data['asset_name'], data['acquisition_date'],
             int(data['useful_life']), int(data['acquisition_cost']),
             data.get('depreciation_method', '定額法'),
             data.get('notes', ''),
             asset_id, user_id)
        )
        conn.commit()
        return True
    except Exception as e:
        conn.rollback()
        raise e
    finally:
        conn.close()


def delete_fixed_asset(asset_id: int, user_id: int = 0) -> bool:
    """Soft-delete a fixed asset."""
    conn = get_db()
    try:
        conn.execute(
            P("UPDATE fixed_assets SET is_deleted = 1 WHERE id = ? AND user_id = ?"),
            (asset_id, user_id)
        )
        conn.commit()
        return True
    except Exception as e:
        conn.rollback()
        raise e
    finally:
        conn.close()


def dispose_fixed_asset(asset_id: int, data: dict, user_id: int = 0) -> bool:
    """Record disposal (売却 or 除却) of a fixed asset.
    data: {disposal_type: '売却'|'除却', disposal_date: 'YYYY-MM-DD', disposal_price: int}
    """
    _now = "CURRENT_TIMESTAMP" if USE_PG else "datetime('now')"
    conn = get_db()
    try:
        conn.execute(
            P(f"""UPDATE fixed_assets SET
                disposal_type = ?, disposal_date = ?, disposal_price = ?,
                updated_at = {_now}
                WHERE id = ? AND user_id = ? AND is_deleted = 0"""),
            (data['disposal_type'], data['disposal_date'],
             int(data.get('disposal_price', 0)),
             asset_id, user_id)
        )
        conn.commit()
        return True
    except Exception as e:
        conn.rollback()
        raise e
    finally:
        conn.close()


def cancel_disposal(asset_id: int, user_id: int = 0) -> bool:
    """Cancel a disposal (undo 売却/除却)."""
    _now = "CURRENT_TIMESTAMP" if USE_PG else "datetime('now')"
    conn = get_db()
    try:
        conn.execute(
            P(f"""UPDATE fixed_assets SET
                disposal_type = '', disposal_date = '', disposal_price = 0,
                updated_at = {_now}
                WHERE id = ? AND user_id = ? AND is_deleted = 0"""),
            (asset_id, user_id)
        )
        conn.commit()
        return True
    except Exception as e:
        conn.rollback()
        raise e
    finally:
        conn.close()


def calculate_depreciation(user_id: int = 0, fiscal_year: str = '2025') -> list:
    """Calculate depreciation schedule for all active fixed assets.

    定額法 (Straight-Line Method):
      年間償却額 = (取得原価 - 1) ÷ 耐用年数  (端数切捨て)
      初年度 = 年間償却額 × 使用月数 ÷ 12 (取得月〜12月, 端数切捨て)
      最終年度: 備忘価額1円になるまで

    売却/除却: 処分月までの月割り償却費を計上し、
      除却 → 償却後の帳簿価額が除却損
      売却 → 売却額との差額が売却損益（譲渡所得）

    Returns list of dicts with per-asset depreciation details for the fiscal year.
    """
    assets = get_fixed_assets(user_id)
    fy = int(fiscal_year)
    result = []

    for asset in assets:
        cost = asset['acquisition_cost']
        life = asset['useful_life']
        acq_date = asset['acquisition_date']  # YYYY-MM-DD
        method = asset.get('depreciation_method', '定額法')
        disposal_type = asset.get('disposal_type', '') or ''
        disposal_date = asset.get('disposal_date', '') or ''
        disposal_price = int(asset.get('disposal_price', 0) or 0)

        # Parse acquisition date
        parts = acq_date.split('-')
        acq_year = int(parts[0])
        acq_month = int(parts[1]) if len(parts) > 1 else 1

        # Parse disposal date if exists
        disp_year = 0
        disp_month = 12
        if disposal_date and disposal_type:
            disp_parts = disposal_date.split('-')
            disp_year = int(disp_parts[0])
            disp_month = int(disp_parts[1]) if len(disp_parts) > 1 else 12

        # Skip assets acquired after fiscal year end
        if acq_year > fy:
            continue

        # Skip assets disposed before this fiscal year
        if disp_year and disp_year < fy:
            continue

        # 定額法 calculation
        depreciable = cost - 1  # 備忘価額1円
        annual_amount = depreciable // life  # 年間償却額

        # Calculate cumulative depreciation up to previous fiscal year
        cumulative_before = 0
        for yr in range(acq_year, fy):
            if yr == acq_year:
                months_used = 12 - acq_month + 1
                yr_dep = annual_amount * months_used // 12
            else:
                yr_dep = annual_amount
            if cumulative_before + yr_dep >= depreciable:
                yr_dep = depreciable - cumulative_before
            cumulative_before += yr_dep
            if cumulative_before >= depreciable:
                break

        opening_bv = cost - cumulative_before
        remark = asset.get('notes', '')

        # Already fully depreciated (before disposal check)?
        if cumulative_before >= depreciable and not (disp_year == fy and disposal_type):
            result.append({
                'id': asset['id'],
                'asset_name': asset['asset_name'],
                'acquisition_date': acq_date,
                'acquisition_cost': cost,
                'useful_life': life,
                'depreciation_method': method,
                'notes': remark,
                'opening_book_value': 1,
                'depreciation_amount': 0,
                'closing_book_value': 1,
                'annual_rate': round(1 / life, 4) if life else 0,
                'disposal_type': '',
                'disposal_remark': '',
            })
            continue

        # --- Handle disposal in this fiscal year ---
        if disp_year == fy and disposal_type:
            # Step 1: Calculate depreciation up to disposal month (月割り)
            if fy == acq_year:
                # Acquired and disposed in the same year
                months_used = disp_month - acq_month + 1
                if months_used < 1:
                    months_used = 1
            else:
                # 1月〜処分月 (e.g. June disposal = 6 months)
                months_used = disp_month
            this_year_dep = annual_amount * months_used // 12
            remaining = depreciable - cumulative_before
            if this_year_dep > remaining:
                this_year_dep = remaining
            # Book value after prorated depreciation
            bv_after_dep = opening_bv - this_year_dep

            if disposal_type == '除却':
                # 除却: normal depreciation + remaining BV is 除却損
                # closing_bv = 0 (asset removed entirely)
                retirement_loss = bv_after_dep  # 帳簿価額そのものが除却損（備忘価額含む）
                if bv_after_dep <= 1:
                    disp_remark = '除却（償却済）'
                else:
                    disp_remark = f'除却（固定資産除却損 {bv_after_dep:,}円）'
                result.append({
                    'id': asset['id'],
                    'asset_name': asset['asset_name'],
                    'acquisition_date': acq_date,
                    'acquisition_cost': cost,
                    'useful_life': life,
                    'depreciation_method': method,
                    'notes': remark,
                    'opening_book_value': opening_bv,
                    'depreciation_amount': this_year_dep,
                    'closing_book_value': 0,
                    'annual_rate': round(1 / life, 4) if life else 0,
                    'disposal_type': '除却',
                    'disposal_remark': disp_remark,
                })
            elif disposal_type == '売却':
                # 売却: normal depreciation + gain/loss vs sale price
                # gain/loss = disposal_price - bv_after_dep
                gain_loss = disposal_price - bv_after_dep
                if gain_loss >= 0:
                    gl_text = f'売却益 {gain_loss:,}円'
                else:
                    gl_text = f'売却損 {abs(gain_loss):,}円'
                result.append({
                    'id': asset['id'],
                    'asset_name': asset['asset_name'],
                    'acquisition_date': acq_date,
                    'acquisition_cost': cost,
                    'useful_life': life,
                    'depreciation_method': method,
                    'notes': remark,
                    'opening_book_value': opening_bv,
                    'depreciation_amount': this_year_dep,
                    'closing_book_value': 0,
                    'annual_rate': round(1 / life, 4) if life else 0,
                    'disposal_type': '売却',
                    'disposal_remark': f'売却（売却額 {disposal_price:,}円 / {gl_text}）',
                })
            continue

        # --- Normal depreciation ---
        if fy == acq_year:
            months_used = 12 - acq_month + 1
            this_year_dep = annual_amount * months_used // 12
        else:
            this_year_dep = annual_amount

        remaining = depreciable - cumulative_before
        if this_year_dep > remaining:
            this_year_dep = remaining

        closing_bv = cost - cumulative_before - this_year_dep

        result.append({
            'id': asset['id'],
            'asset_name': asset['asset_name'],
            'acquisition_date': acq_date,
            'acquisition_cost': cost,
            'useful_life': life,
            'depreciation_method': method,
            'notes': remark,
            'opening_book_value': opening_bv,
            'depreciation_amount': this_year_dep,
            'closing_book_value': closing_bv,
            'annual_rate': round(1 / life, 4) if life else 0,
            'disposal_type': '',
            'disposal_remark': '',
        })

    return result
