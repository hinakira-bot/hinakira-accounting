"""
Statement Parser — 銀行/クレジットカードCSV明細の自動判定・構造化パース
Standard library only (csv, hashlib, io, re, datetime). No heavy imports.
"""
import csv
import hashlib
import io
import re
from datetime import datetime, date


# ============================================================
#  Format Registry — 対応する銀行/カードのCSVフォーマット定義
# ============================================================

FORMAT_REGISTRY = [
    # --- 銀行 (Banks) ---
    {
        'name': 'みずほ銀行',
        'type': 'bank',
        'keywords': ['お取引日', 'お引出金額', 'お預入金額'],
        'date_col': 0,
        'description_col': 1,
        'withdrawal_col': 2,
        'deposit_col': 3,
        'balance_col': 4,
        'skip_rows': 1,
        'date_formats': ['%Y/%m/%d', '%Y-%m-%d'],
        'default_debit': '普通預金',
        'default_credit': '普通預金',
    },
    {
        'name': '三菱UFJ銀行',
        'type': 'bank',
        'keywords': ['日付', 'お支払い金額', 'お預かり金額'],
        'date_col': 0,
        'description_col': 1,
        'withdrawal_col': 2,
        'deposit_col': 3,
        'balance_col': None,
        'skip_rows': 1,
        'date_formats': ['%Y/%m/%d', '%Y-%m-%d'],
        'default_debit': '普通預金',
        'default_credit': '普通預金',
    },
    {
        'name': '三井住友銀行',
        'type': 'bank',
        'keywords': ['年月日', 'お引出し', 'お預入れ'],
        'date_col': 0,
        'description_col': 1,
        'withdrawal_col': 2,
        'deposit_col': 3,
        'balance_col': 4,
        'skip_rows': 1,
        'date_formats': ['%Y/%m/%d', '%Y-%m-%d'],
        'default_debit': '普通預金',
        'default_credit': '普通預金',
    },
    {
        'name': 'ゆうちょ銀行',
        'type': 'bank',
        'keywords': ['取引日', '受入金額', '払出金額', '詳細'],
        'date_col': 0,
        'description_col': 4,  # 詳細１（取引種別: カード/振込/利子等）
        'detail_col': 5,       # 詳細２（取引先名・半角カナ）
        'deposit_col': 2,      # 受入金額
        'withdrawal_col': 3,   # 払出金額
        'balance_col': 6,      # 現在（貸付）高
        'skip_rows': 8,        # 行1-7: 情報ヘッダー, 行8: データヘッダー
        'date_formats': ['%Y%m%d', '%Y/%m/%d'],
        'default_debit': '普通預金',
        'default_credit': '普通預金',
    },
    {
        'name': '楽天銀行',
        'type': 'bank',
        'keywords': ['取引日', '入出金'],
        'date_col': 0,
        'description_col': 1,
        'amount_col': 2,  # signed: negative=withdrawal, positive=deposit
        'balance_col': None,
        'skip_rows': 1,
        'date_formats': ['%Y%m%d', '%Y/%m/%d'],
        'default_debit': '普通預金',
        'default_credit': '普通預金',
    },
    {
        'name': '住信SBIネット銀行',
        'type': 'bank',
        'keywords': ['日付', '内容', '出金金額', '入金金額'],
        'date_col': 0,
        'description_col': 1,
        'withdrawal_col': 2,
        'deposit_col': 3,
        'balance_col': 4,
        'skip_rows': 1,
        'date_formats': ['%Y/%m/%d', '%Y-%m-%d'],
        'default_debit': '普通預金',
        'default_credit': '普通預金',
    },
    {
        'name': 'PayPay銀行',
        'type': 'bank',
        'keywords': ['日付', 'お支払金額', 'お預り金額'],
        'date_col': 0,
        'description_col': 1,
        'withdrawal_col': 2,
        'deposit_col': 3,
        'balance_col': None,
        'skip_rows': 1,
        'date_formats': ['%Y/%m/%d', '%Y-%m-%d'],
        'default_debit': '普通預金',
        'default_credit': '普通預金',
    },

    # --- クレジットカード (Credit Cards) ---
    {
        'name': '楽天カード',
        'type': 'card',
        'keywords': ['利用日', '利用店名・商品名'],
        'date_col': 0,
        'description_col': 1,
        'amount_col': -1,  # last numeric column (varies by format)
        'skip_rows': 1,
        'date_formats': ['%Y/%m/%d', '%Y-%m-%d'],
        'default_debit': '',
        'default_credit': '未払金',
    },
    {
        'name': '三井住友カード',
        'type': 'card',
        'keywords': ['ご利用日', 'ご利用先'],
        'date_col': 0,
        'description_col': 1,
        'amount_col': -1,
        'skip_rows': 1,
        'date_formats': ['%Y/%m/%d', '%Y-%m-%d'],
        'default_debit': '',
        'default_credit': '未払金',
    },
    {
        'name': 'JCBカード',
        'type': 'card',
        'keywords': ['ご利用年月日', 'ご利用先'],
        'date_col': 0,
        'description_col': 1,
        'amount_col': -1,
        'skip_rows': 1,
        'date_formats': ['%Y/%m/%d', '%Y-%m-%d'],
        'default_debit': '',
        'default_credit': '未払金',
    },
    {
        'name': 'Amazonカード',
        'type': 'card',
        'keywords': ['ご利用日', 'ご利用金額'],
        'date_col': 0,
        'description_col': 1,
        'amount_col': -1,
        'skip_rows': 1,
        'date_formats': ['%Y/%m/%d', '%Y-%m-%d'],
        'default_debit': '',
        'default_credit': '未払金',
    },
    {
        'name': 'アメリカン・エキスプレス',
        'type': 'card',
        'keywords': ['日付', 'ご利用先', '金額'],
        'date_col': 0,
        'description_col': 1,
        'amount_col': -1,
        'skip_rows': 1,
        'date_formats': ['%Y/%m/%d', '%m/%d/%Y', '%Y-%m-%d'],
        'default_debit': '',
        'default_credit': '未払金',
    },
]


# ============================================================
#  Utility Functions
# ============================================================

def compute_file_hash(file_bytes: bytes) -> str:
    """SHA-256 hash of file content for duplicate detection."""
    return hashlib.sha256(file_bytes).hexdigest()


def read_csv_with_encoding(file_bytes: bytes) -> str:
    """Try multiple encodings to decode CSV bytes."""
    for enc in ['cp932', 'shift_jis', 'utf-8-sig', 'utf-8']:
        try:
            text = file_bytes.decode(enc)
            # Quick validation: should have commas or tabs
            if ',' in text or '\t' in text:
                return text
        except (UnicodeDecodeError, ValueError):
            continue
    # Last resort
    return file_bytes.decode('utf-8', errors='replace')


def _clean_amount(s: str) -> int:
    """Parse amount string: remove commas, yen signs, spaces. Return int or 0."""
    if not s:
        return 0
    s = s.strip()
    s = re.sub(r'[¥￥,、\s　]', '', s)
    if not s or s == '-' or s == '—':
        return 0
    try:
        return int(float(s))
    except (ValueError, TypeError):
        return 0


def _parse_date(s: str, formats: list) -> str:
    """Try multiple date formats, return YYYY-MM-DD or empty string."""
    s = s.strip()
    if not s:
        return ''
    # Remove common noise
    s = re.sub(r'[年月]', '/', s).replace('日', '').strip()
    for fmt in formats:
        try:
            dt = datetime.strptime(s, fmt)
            return dt.strftime('%Y-%m-%d')
        except ValueError:
            continue
    # Try ISO format as last resort
    try:
        dt = datetime.fromisoformat(s)
        return dt.strftime('%Y-%m-%d')
    except (ValueError, TypeError):
        return ''


def _extract_counterparty(description: str) -> str:
    """Extract counterparty name from bank/card description."""
    if not description:
        return ''
    desc = description.strip()

    # Bank transfer patterns
    patterns = [
        r'^振込\s*(.+)',
        r'^ﾌﾘｺﾐ\s*(.+)',
        r'^カード\s*(.+)',
        r'^ｶｰﾄﾞ\s*(.+)',
        r'^デビット\s*(.+)',
        r'^ﾃﾞﾋﾞｯﾄ\s*(.+)',
        r'^引落\s*(.+)',
        r'^ﾋｷｵﾄｼ\s*(.+)',
        r'^口座振替\s*(.+)',
        r'^ｺｳｻﾞﾌﾘｶｴ\s*(.+)',
        r'^給与\s*(.+)',
        r'^ｷﾕｳﾖ\s*(.+)',
    ]
    for pat in patterns:
        m = re.match(pat, desc)
        if m:
            return m.group(1).strip()

    return desc


# ============================================================
#  Core Functions
# ============================================================

def detect_format(csv_text: str) -> dict | None:
    """
    Auto-detect CSV format by scanning first 5 lines for keywords.
    Returns matching format dict, or None if no match.
    """
    # Get first 5 lines
    lines = csv_text.strip().split('\n')[:5]
    header_text = '\n'.join(lines)

    best_match = None
    best_score = 0

    for fmt in FORMAT_REGISTRY:
        score = 0
        for kw in fmt['keywords']:
            if kw in header_text:
                score += 1
        # Require ALL keywords to match
        if score == len(fmt['keywords']) and score > best_score:
            best_match = fmt
            best_score = score

    return best_match


def parse_statement(csv_text: str, fmt: dict) -> list:
    """
    Parse CSV rows using format definition.
    Returns list of dicts with: date, description, amount, direction, balance, raw_row
    No row limit.
    """
    results = []
    reader = csv.reader(io.StringIO(csv_text))
    rows = list(reader)

    skip = fmt.get('skip_rows', 1)
    date_formats = fmt.get('date_formats', ['%Y/%m/%d'])

    for i, row in enumerate(rows):
        if i < skip:
            continue
        if not row or all(not cell.strip() for cell in row):
            continue

        # Extract date
        date_col = fmt.get('date_col', 0)
        if date_col >= len(row):
            continue
        entry_date = _parse_date(row[date_col], date_formats)
        if not entry_date:
            continue  # Skip rows without valid date

        # Extract description
        desc_col = fmt.get('description_col', 1)
        description = row[desc_col].strip() if desc_col < len(row) else ''

        # Extract amount and direction
        amount = 0
        direction = 'withdrawal'

        if 'amount_col' in fmt:
            # Single amount column (signed or card)
            col = fmt['amount_col']
            if col == -1:
                # Find last numeric column
                for j in range(len(row) - 1, -1, -1):
                    val = _clean_amount(row[j])
                    if val != 0:
                        amount = abs(val)
                        direction = 'deposit' if val > 0 else 'withdrawal'
                        break
            else:
                val = _clean_amount(row[col] if col < len(row) else '')
                amount = abs(val)
                direction = 'deposit' if val > 0 else 'withdrawal'

            # For credit cards, all entries are expenses (withdrawals)
            if fmt.get('type') == 'card':
                direction = 'withdrawal'
        else:
            # Separate withdrawal/deposit columns
            w_col = fmt.get('withdrawal_col')
            d_col = fmt.get('deposit_col')
            withdrawal = _clean_amount(row[w_col] if w_col is not None and w_col < len(row) else '')
            deposit = _clean_amount(row[d_col] if d_col is not None and d_col < len(row) else '')

            if withdrawal > 0:
                amount = withdrawal
                direction = 'withdrawal'
            elif deposit > 0:
                amount = deposit
                direction = 'deposit'
            else:
                continue  # Skip zero-amount rows

        if amount == 0:
            continue

        # Extract balance if available
        bal_col = fmt.get('balance_col')
        balance = _clean_amount(row[bal_col] if bal_col is not None and bal_col < len(row) else '') if bal_col else None

        # Extract detail column (e.g. ゆうちょ銀行の詳細2 = 取引先名)
        detail_col = fmt.get('detail_col')
        detail = row[detail_col].strip() if detail_col is not None and detail_col < len(row) else ''

        results.append({
            'date': entry_date,
            'description': description,
            'detail': detail,
            'amount': amount,
            'direction': direction,
            'balance': balance,
            'raw_row': ','.join(row),
        })

    return results


def build_journal_candidates(
    parsed_rows: list,
    source_type: str,
    source_name: str,
    default_debit: str = '',
    default_credit: str = '',
) -> list:
    """
    Convert parsed rows into journal entry candidates.
    Output format matches existing scan results for seamless integration.

    Bank withdrawals: debit=TBD(expense), credit=普通預金
    Bank deposits:    debit=普通預金, credit=TBD(revenue)
    Card purchases:   debit=TBD(expense), credit=未払金
    """
    # 銀行固有取引のキーワード（取引先が銀行自体の場合）
    BANK_SELF_KEYWORDS = ['利息', '利子', '受取利子', '手数料', '料\u3000金', '料 金', '税金', '源泉', '記帳']
    # 口座種別（取引先名ではないので除外）
    ACCOUNT_TYPE_WORDS = ['普通預金', '通常貯金', '当座預金', '貯蓄預金', '定期預金', '普通貯金', '通常貯蓄']

    entries = []

    for row in parsed_rows:
        detail = row.get('detail', '')
        description = row['description']

        # 詳細列が口座種別の場合は取引先名として使わない
        effective_detail = detail if detail and detail not in ACCOUNT_TYPE_WORDS else ''

        # 詳細列に有効な取引先名がある場合は使用（ゆうちょ等）
        if effective_detail:
            counterparty = effective_detail
            memo = f"{description} {effective_detail}".strip()
        else:
            counterparty = _extract_counterparty(description)
            memo = description if description != counterparty else ''

        # 預金明細で「カード」かつ取引先空 → ATM引出/預入
        if source_type == 'bank' and description in ('カード', 'ｶｰﾄﾞ') and not effective_detail:
            counterparty = source_name  # 例: 'ゆうちょ銀行'
            if row['direction'] == 'withdrawal':
                memo = '預金引出'
            else:
                memo = '預金預入'

        # 預金明細で銀行固有キーワードに該当 → 取引先を銀行名に設定
        if source_type == 'bank':
            combined = f"{description} {detail}".strip()
            if any(kw in combined for kw in BANK_SELF_KEYWORDS):
                counterparty = source_name  # 例: 'ゆうちょ銀行'
                memo = description   # 取引種別をmemoに保存

        if source_type == 'card':
            # Credit card: all are expenses
            debit = ''  # To be predicted by AI
            credit = default_credit or '未払金'
        elif source_type == 'bank':
            if row['direction'] == 'withdrawal':
                # Bank withdrawal (出金) = expense
                debit = ''  # To be predicted by AI
                credit = default_debit or '普通預金'  # default_debit is the bank account name
            else:
                # Bank deposit (入金) = revenue
                debit = default_debit or '普通預金'
                credit = ''  # To be predicted by AI
        else:
            debit = ''
            credit = ''

        entries.append({
            'date': row['date'],
            'debit_account': debit,
            'credit_account': credit,
            'amount': row['amount'],
            'tax_classification': '10%',  # Default, AI will refine
            'counterparty': counterparty,
            'memo': memo,
            'source': 'csv_import',
            'source_name': source_name,
        })

    return entries


def parse_csv_smart(file_bytes: bytes) -> dict:
    """
    All-in-one function: detect format, parse, build candidates.
    Returns dict with:
      detected: bool
      source_type: str
      source_name: str
      entries: list
      file_hash: str
      total_rows: int
      default_debit: str
      default_credit: str
    """
    file_hash = compute_file_hash(file_bytes)
    csv_text = read_csv_with_encoding(file_bytes)
    fmt = detect_format(csv_text)

    if fmt:
        parsed = parse_statement(csv_text, fmt)
        candidates = build_journal_candidates(
            parsed,
            source_type=fmt['type'],
            source_name=fmt['name'],
            default_debit=fmt.get('default_debit', ''),
            default_credit=fmt.get('default_credit', ''),
        )
        return {
            'detected': True,
            'source_type': fmt['type'],
            'source_name': fmt['name'],
            'entries': candidates,
            'file_hash': file_hash,
            'total_rows': len(parsed),
            'default_debit': fmt.get('default_debit', ''),
            'default_credit': fmt.get('default_credit', ''),
        }
    else:
        return {
            'detected': False,
            'source_type': '',
            'source_name': '',
            'entries': [],
            'file_hash': file_hash,
            'total_rows': 0,
            'default_debit': '',
            'default_credit': '',
        }
