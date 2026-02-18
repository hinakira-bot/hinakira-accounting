"""
Data access layer for AI Accounting Tool.
CRUD operations for journal entries, accounts, trial balance, etc.
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


# --- Accounts ---
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


# --- Journal Entries ---
def create_journal_entry(entry: dict) -> int:
    """Create a single journal entry. Returns the new entry ID."""
    conn = get_db()
    try:
        amount = int(entry.get('amount', 0))
        tax_class = entry.get('tax_classification', '10%')
        tax_amount = calculate_tax_amount(amount, tax_class)

        # Resolve account IDs from names if needed
        debit_id = _resolve_account_id(conn, entry.get('debit_account_id'), entry.get('debit_account'))
        credit_id = _resolve_account_id(conn, entry.get('credit_account_id'), entry.get('credit_account'))

        if not debit_id or not credit_id:
            raise ValueError(f"Invalid account: debit={entry.get('debit_account', entry.get('debit_account_id'))}, credit={entry.get('credit_account', entry.get('credit_account_id'))}")

        cursor = conn.execute(
            """INSERT INTO journal_entries
            (entry_date, debit_account_id, credit_account_id, amount, tax_classification, tax_amount,
             counterparty, memo, evidence_url, source)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
            (
                entry.get('entry_date', entry.get('date', '')),
                debit_id,
                credit_id,
                amount,
                tax_class,
                tax_amount,
                entry.get('counterparty', ''),
                entry.get('memo', ''),
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


def create_journal_entries_batch(entries: list) -> list:
    """Create multiple journal entries. Returns list of new IDs."""
    ids = []
    for entry in entries:
        try:
            new_id = create_journal_entry(entry)
            ids.append(new_id)
        except Exception as e:
            print(f"Error creating entry: {e}, data: {entry}")
            ids.append(None)
    return ids


def get_journal_entries(filters: dict = None) -> dict:
    """
    Get journal entries with filters and pagination.
    filters: start_date, end_date, account_id, counterparty, memo, page, per_page
    Returns: { entries: [...], total: N, page: N, per_page: N }
    """
    if filters is None:
        filters = {}

    conn = get_db()
    try:
        conditions = ["je.is_deleted = 0"]
        params = []

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

        where = " AND ".join(conditions)

        # Count total
        count_sql = f"SELECT COUNT(*) FROM journal_entries je WHERE {where}"
        total = conn.execute(count_sql, params).fetchone()[0]

        # Pagination
        page = max(1, int(filters.get('page', 1)))
        per_page = min(100, max(1, int(filters.get('per_page', 20))))
        offset = (page - 1) * per_page

        # Fetch with joins
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


def get_recent_entries(limit=5) -> list:
    """Get the most recent journal entries."""
    conn = get_db()
    try:
        rows = conn.execute("""
        SELECT
            je.id, je.entry_date, je.amount, je.tax_classification, je.tax_amount,
            je.counterparty, je.memo, je.source,
            da.name AS debit_account, ca.name AS credit_account
        FROM journal_entries je
        JOIN accounts_master da ON je.debit_account_id = da.id
        JOIN accounts_master ca ON je.credit_account_id = ca.id
        WHERE je.is_deleted = 0
        ORDER BY je.id DESC
        LIMIT ?
        """, (limit,)).fetchall()
        return [dict(r) for r in rows]
    finally:
        conn.close()


def update_journal_entry(entry_id: int, entry: dict) -> bool:
    """Update a journal entry."""
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
            WHERE id = ? AND is_deleted = 0""",
            (
                entry.get('entry_date', entry.get('date', '')),
                debit_id, credit_id, amount, tax_class, tax_amount,
                entry.get('counterparty', ''), entry.get('memo', ''),
                entry_id,
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


def delete_journal_entry(entry_id: int) -> bool:
    """Soft-delete a journal entry."""
    conn = get_db()
    try:
        conn.execute(
            "UPDATE journal_entries SET is_deleted = 1, updated_at = datetime('now') WHERE id = ?",
            (entry_id,)
        )
        conn.commit()
        return True
    except Exception as e:
        conn.rollback()
        print(f"Delete error: {e}")
        return False
    finally:
        conn.close()


# --- Trial Balance ---
def get_trial_balance(start_date=None, end_date=None) -> list:
    """
    Get trial balance grouped by account.
    Returns list of dicts with: account_id, code, name, account_type,
    opening_balance, debit_total, credit_total, closing_balance
    """
    conn = get_db()
    try:
        # Get all active accounts
        accounts = conn.execute(
            "SELECT id, code, name, account_type, display_order FROM accounts_master WHERE is_active = 1 ORDER BY display_order, code"
        ).fetchall()

        # Build date conditions for journal entries
        date_conditions = ["je.is_deleted = 0"]
        date_params = []
        if start_date:
            date_conditions.append("je.entry_date >= ?")
            date_params.append(start_date)
        if end_date:
            date_conditions.append("je.entry_date <= ?")
            date_params.append(end_date)
        date_where = " AND ".join(date_conditions)

        # Debit totals per account
        debit_sql = f"""
        SELECT debit_account_id AS account_id, SUM(amount) AS total
        FROM journal_entries je WHERE {date_where}
        GROUP BY debit_account_id
        """
        debit_totals = {row['account_id']: row['total'] for row in conn.execute(debit_sql, date_params).fetchall()}

        # Credit totals per account
        credit_sql = f"""
        SELECT credit_account_id AS account_id, SUM(amount) AS total
        FROM journal_entries je WHERE {date_where}
        GROUP BY credit_account_id
        """
        credit_totals = {row['account_id']: row['total'] for row in conn.execute(credit_sql, date_params).fetchall()}

        # Opening balances (fiscal year derived from start_date or current year)
        fiscal_year = start_date[:4] if start_date else "2025"
        opening_sql = "SELECT account_id, amount FROM opening_balances WHERE fiscal_year = ?"
        openings = {row['account_id']: row['amount'] for row in conn.execute(opening_sql, (fiscal_year,)).fetchall()}

        # Build result
        result = []
        for acc in accounts:
            aid = acc['id']
            atype = acc['account_type']
            opening = openings.get(aid, 0)
            debit = debit_totals.get(aid, 0)
            credit = credit_totals.get(aid, 0)

            # Closing balance calculation depends on account type
            # 資産・費用: Opening + Debit - Credit
            # 負債・純資産・収益: Opening + Credit - Debit
            if atype in ('資産', '費用'):
                closing = opening + debit - credit
            else:
                closing = opening + credit - debit

            # Only include accounts with activity or opening balance
            if opening != 0 or debit != 0 or credit != 0:
                result.append({
                    "account_id": aid,
                    "code": acc['code'],
                    "name": acc['name'],
                    "account_type": atype,
                    "opening_balance": opening,
                    "debit_total": debit,
                    "credit_total": credit,
                    "closing_balance": closing,
                })

        return result
    finally:
        conn.close()


# --- Counterparties ---
def get_counterparties(limit=200) -> list:
    """Get distinct counterparty names from recent journal entries."""
    conn = get_db()
    try:
        rows = conn.execute("""
        SELECT DISTINCT counterparty FROM journal_entries
        WHERE is_deleted = 0 AND counterparty != ''
        ORDER BY counterparty
        LIMIT ?
        """, (limit,)).fetchall()
        return [r['counterparty'] for r in rows]
    finally:
        conn.close()


# --- History for AI (replaces Sheets-based history) ---
def get_accounting_history(limit=200) -> list:
    """Get recent journal entries for AI context (counterparty + memo + debit account)."""
    conn = get_db()
    try:
        rows = conn.execute("""
        SELECT je.counterparty, je.memo, da.name AS account
        FROM journal_entries je
        JOIN accounts_master da ON je.debit_account_id = da.id
        WHERE je.is_deleted = 0 AND je.counterparty != ''
        ORDER BY je.id DESC LIMIT ?
        """, (limit,)).fetchall()
        return [dict(r) for r in rows]
    finally:
        conn.close()


def get_existing_entry_keys() -> set:
    """Get set of 'date_amount_counterparty' keys for duplicate detection."""
    conn = get_db()
    try:
        rows = conn.execute(
            "SELECT entry_date, amount, counterparty FROM journal_entries WHERE is_deleted = 0"
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
        # Fallback: try to find 雑費 as default
        row = conn.execute(
            "SELECT id FROM accounts_master WHERE name = '雑費' AND is_active = 1"
        ).fetchone()
        return row['id'] if row else None
    return None
