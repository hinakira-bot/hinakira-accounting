"""
Hinakira会計 - Flask Backend
Main routing and API endpoints.
Multi-user support via Google OAuth user identification.
"""
import os
import io
import json
import time
import requests as http_requests
from flask import Flask, request, jsonify, send_from_directory, g
from flask_cors import CORS
from dotenv import load_dotenv

import db
import models

load_dotenv()

print("=== Hinakira Accounting Starting ===", flush=True)
print(f"Python version: {__import__('sys').version}", flush=True)
print(f"Working directory: {os.getcwd()}", flush=True)

# In-memory cache: access_token -> {user: dict, expires: timestamp}
# Avoids calling Google UserInfo API on every request
_token_cache = {}
_TOKEN_CACHE_TTL = 300  # 5 minutes

app = Flask(__name__, static_folder='.', static_url_path='/static')
CORS(app)

# Initialize database on startup (includes migration for existing DBs)
try:
    print(f"DB_PATH will be: {db.DB_PATH}", flush=True)
    db.init_db()
    print("Database initialized successfully", flush=True)
except Exception as e:
    print(f"CRITICAL: Database initialization failed: {e}", flush=True)
    import traceback
    traceback.print_exc()


# ============================
#  Health Check (no auth needed)
# ============================
@app.route('/healthz')
def healthz():
    """Health check endpoint for Render."""
    return jsonify({"status": "ok"}), 200


# ============================
#  Authentication Middleware
# ============================
def get_current_user():
    """Extract user from Authorization header (Bearer token).
    Calls Google UserInfo API to get email, then finds/creates user in DB.
    Uses in-memory cache to avoid calling Google API on every request.
    Caches result in flask.g for the duration of the request."""
    if hasattr(g, 'user') and g.user:
        return g.user

    auth_header = request.headers.get('Authorization', '')
    if not auth_header.startswith('Bearer '):
        return None

    access_token = auth_header[7:]
    if not access_token:
        return None

    now = time.time()

    # Check in-memory token cache first
    cached = _token_cache.get(access_token)
    if cached and cached['expires'] > now:
        g.user = cached['user']
        return cached['user']

    try:
        # Call Google UserInfo API (only when not cached)
        resp = http_requests.get(
            'https://www.googleapis.com/oauth2/v3/userinfo',
            headers={'Authorization': f'Bearer {access_token}'},
            timeout=10
        )
        if resp.status_code != 200:
            # Clean expired cache entry if exists
            _token_cache.pop(access_token, None)
            print(f"Auth: Google UserInfo returned {resp.status_code}: {resp.text[:200]}")
            return None

        info = resp.json()
        email = info.get('email', '')
        if not email:
            return None

        user = models.get_or_create_user(
            email=email,
            name=info.get('name', ''),
            picture=info.get('picture', '')
        )

        # Cache the result
        print(f"Auth: User authenticated: {email} (id={user['id']})")
        _token_cache[access_token] = {
            'user': user,
            'expires': now + _TOKEN_CACHE_TTL
        }

        # Clean old entries periodically (keep cache small)
        if len(_token_cache) > 100:
            expired = [k for k, v in _token_cache.items() if v['expires'] < now]
            for k in expired:
                del _token_cache[k]

        g.user = user
        return user
    except Exception as e:
        print(f"Auth error: {e}")
        return None


@app.before_request
def require_auth():
    """Require authentication for API endpoints (except static files and drive APIs)."""
    path = request.path

    # Static files don't need auth
    if not path.startswith('/api/'):
        return None

    # Drive APIs use access_token in body (handled internally)
    if path.startswith('/api/drive/'):
        return None

    # Accounts API is global (no user-scoping needed, but still require login)
    # All other APIs need user identification
    user = get_current_user()
    if not user:
        return jsonify({"error": "認証が必要です。再ログインしてください。"}), 401


def get_user_id():
    """Get current user's ID. Must be called after require_auth middleware."""
    user = get_current_user()
    return user['id'] if user else 0


# --- Lazy imports for heavy Google libraries (reduce startup memory) ---
_lazy_cache = {}

def _get_credentials():
    if 'Credentials' not in _lazy_cache:
        from google.oauth2.credentials import Credentials
        _lazy_cache['Credentials'] = Credentials
    return _lazy_cache['Credentials']

def _get_drive_build():
    if 'build' not in _lazy_cache:
        from googleapiclient.discovery import build
        _lazy_cache['build'] = build
    return _lazy_cache['build']

def _get_media_upload():
    if 'MediaIoBaseUpload' not in _lazy_cache:
        from googleapiclient.http import MediaIoBaseUpload
        _lazy_cache['MediaIoBaseUpload'] = MediaIoBaseUpload
    return _lazy_cache['MediaIoBaseUpload']

def _get_ai_service():
    if 'ai_service' not in _lazy_cache:
        import ai_service
        _lazy_cache['ai_service'] = ai_service
    return _lazy_cache['ai_service']

def _get_genai():
    if 'genai' not in _lazy_cache:
        import google.generativeai as genai
        _lazy_cache['genai'] = genai
    return _lazy_cache['genai']


# --- Google Drive Upload Helper ---
def upload_to_drive(file_bytes, filename, mime_type, access_token):
    """Upload evidence file to Google Drive."""
    try:
        if not access_token:
            return ""
        creds = _get_credentials()(token=access_token)
        service = _get_drive_build()('drive', 'v3', credentials=creds)

        folder_name = "Accounting_Evidence"
        results = service.files().list(
            q=f"mimeType='application/vnd.google-apps.folder' and name='{folder_name}' and trashed=false",
            spaces='drive'
        ).execute()
        items = results.get('files', [])

        if not items:
            folder_metadata = {'name': folder_name, 'mimeType': 'application/vnd.google-apps.folder'}
            file = service.files().create(body=folder_metadata, fields='id').execute()
            folder_id = file.get('id')
        else:
            folder_id = items[0]['id']

        file_metadata = {'name': filename, 'parents': [folder_id]}
        media = _get_media_upload()(io.BytesIO(file_bytes), mimetype=mime_type, resumable=True)
        file = service.files().create(body=file_metadata, media_body=media, fields='id, webViewLink').execute()
        return file.get('webViewLink')
    except Exception as e:
        print(f"Drive Upload Error: {e}")
        return ""


def get_or_create_folder(service, folder_name, parent_id=None):
    """Get or create a Google Drive folder. Returns folder_id."""
    q = f"mimeType='application/vnd.google-apps.folder' and name='{folder_name}' and trashed=false"
    if parent_id:
        q += f" and '{parent_id}' in parents"
    results = service.files().list(q=q, spaces='drive', fields='files(id)').execute()
    items = results.get('files', [])
    if items:
        return items[0]['id']
    meta = {'name': folder_name, 'mimeType': 'application/vnd.google-apps.folder'}
    if parent_id:
        meta['parents'] = [parent_id]
    f = service.files().create(body=meta, fields='id').execute()
    return f.get('id')


def get_all_folder_ids(service, folder_name, parent_ids=None):
    """Get ALL folder IDs matching name under any of the parent_ids."""
    all_ids = []
    if parent_ids:
        for pid in parent_ids:
            q = f"mimeType='application/vnd.google-apps.folder' and name='{folder_name}' and trashed=false and '{pid}' in parents"
            results = service.files().list(q=q, spaces='drive', fields='files(id)').execute()
            all_ids.extend([f['id'] for f in results.get('files', [])])
    else:
        q = f"mimeType='application/vnd.google-apps.folder' and name='{folder_name}' and trashed=false"
        results = service.files().list(q=q, spaces='drive', fields='files(id)').execute()
        all_ids = [f['id'] for f in results.get('files', [])]
    return all_ids


def upload_to_drive_processed(file_bytes, filename, mime_type, access_token):
    """Upload evidence file to Accounting_Evidence/processed/ on Google Drive."""
    try:
        if not access_token:
            return ""
        creds = _get_credentials()(token=access_token)
        service = _get_drive_build()('drive', 'v3', credentials=creds)
        root_id = get_or_create_folder(service, "Accounting_Evidence")
        processed_id = get_or_create_folder(service, "processed", root_id)

        file_metadata = {'name': filename, 'parents': [processed_id]}
        media = _get_media_upload()(io.BytesIO(file_bytes), mimetype=mime_type, resumable=True)
        file = service.files().create(body=file_metadata, media_body=media, fields='id, webViewLink').execute()
        return file.get('webViewLink')
    except Exception as e:
        print(f"Drive Upload Error: {e}")
        return ""


# ============================
#  Drive Inbox API
# ============================
@app.route('/api/drive/init', methods=['POST'])
def api_drive_init():
    """Initialize Drive folder structure on login."""
    data = request.json or {}
    access_token = data.get('access_token', '')
    if not access_token:
        return jsonify({"error": "Googleログインが必要です"}), 401

    try:
        creds = _get_credentials()(token=access_token)
        service = _get_drive_build()('drive', 'v3', credentials=creds)
        root_id = get_or_create_folder(service, "Accounting_Evidence")
        get_or_create_folder(service, "inbox", root_id)
        get_or_create_folder(service, "processed", root_id)
        return jsonify({"status": "success", "folder_id": root_id})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route('/api/drive/inbox', methods=['POST'])
def api_drive_inbox_list():
    """List unprocessed files in Accounting_Evidence/inbox/."""
    data = request.json or {}
    access_token = data.get('access_token', '')
    if not access_token:
        return jsonify({"error": "Googleログインが必要です"}), 401

    import sys
    try:
        creds = _get_credentials()(token=access_token)
        service = _get_drive_build()('drive', 'v3', credentials=creds)

        root_ids = get_all_folder_ids(service, "Accounting_Evidence")
        if not root_ids:
            return jsonify({"files": [], "inbox_folder_ids": []})

        inbox_ids = get_all_folder_ids(service, "inbox", root_ids)
        if not inbox_ids:
            return jsonify({"files": [], "inbox_folder_ids": []})

        print(f"[Drive Inbox] root_ids={root_ids}, inbox_ids={inbox_ids}", file=sys.stderr, flush=True)

        all_files = []
        for iid in inbox_ids:
            q = f"'{iid}' in parents and trashed=false"
            results = service.files().list(
                q=q,
                spaces='drive',
                fields='files(id, name, mimeType, createdTime, size)',
                orderBy='createdTime desc',
                pageSize=50
            ).execute()
            all_files.extend(results.get('files', []))

        print(f"[Drive Inbox] Found {len(all_files)} files: {[f['name'] for f in all_files]}", file=sys.stderr, flush=True)
        return jsonify({"files": all_files, "inbox_folder_ids": inbox_ids})
    except Exception as e:
        print(f"[Drive Inbox] Error: {e}", file=sys.stderr, flush=True)
        return jsonify({"error": str(e)}), 500


@app.route('/api/drive/inbox/analyze', methods=['POST'])
def api_drive_inbox_analyze():
    """Download files from Drive inbox and analyze with AI."""
    data = request.json or {}
    access_token = data.get('access_token', '')
    gemini_api_key = data.get('gemini_api_key', '')
    file_ids = data.get('file_ids', [])

    if not access_token:
        return jsonify({"error": "Googleログインが必要です"}), 401
    if not gemini_api_key:
        return jsonify({"error": "Gemini APIキーが設定されていません"}), 401
    if not file_ids:
        return jsonify({"error": "ファイルが指定されていません"}), 400

    # Get user_id from Authorization header if available
    uid = get_user_id()

    _get_ai_service().configure_gemini(gemini_api_key)
    history = models.get_accounting_history(user_id=uid)
    existing = models.get_existing_entry_keys(user_id=uid)

    try:
        creds = _get_credentials()(token=access_token)
        service = _get_drive_build()('drive', 'v3', credentials=creds)
        from googleapiclient.http import MediaIoBaseDownload

        results = []
        for fid in file_ids:
            meta = service.files().get(fileId=fid, fields='name, mimeType, webViewLink').execute()
            fname = meta.get('name', '')
            mime = meta.get('mimeType', '')
            web_link = meta.get('webViewLink', '')

            req = service.files().get_media(fileId=fid)
            buf = io.BytesIO()
            dl = MediaIoBaseDownload(buf, req)
            done = False
            while not done:
                _, done = dl.next_chunk()
            buf.seek(0)
            fbytes = buf.read()

            if fname.lower().endswith('.csv'):
                res = _get_ai_service().analyze_csv(fbytes, history)
            else:
                res = _get_ai_service().analyze_document(fbytes, mime, history)

            for item in res:
                item['evidence_url'] = web_link
                item['drive_file_id'] = fid
                key = f"{item.get('date')}_{str(item.get('amount'))}_{item.get('counterparty')}"
                item['is_duplicate'] = key in existing
            results.extend(res)

        return jsonify(results)
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route('/api/drive/inbox/move', methods=['POST'])
def api_drive_inbox_move():
    """Move files from inbox to processed after journal registration."""
    data = request.json or {}
    access_token = data.get('access_token', '')
    file_ids = data.get('file_ids', [])

    if not access_token:
        return jsonify({"error": "Googleログインが必要です"}), 401
    if not file_ids:
        return jsonify({"status": "success", "moved": 0})

    try:
        creds = _get_credentials()(token=access_token)
        service = _get_drive_build()('drive', 'v3', credentials=creds)
        root_id = get_or_create_folder(service, "Accounting_Evidence")
        processed_id = get_or_create_folder(service, "processed", root_id)

        moved = 0
        for fid in file_ids:
            try:
                f = service.files().get(fileId=fid, fields='parents').execute()
                current_parents = ','.join(f.get('parents', []))
                service.files().update(
                    fileId=fid,
                    addParents=processed_id,
                    removeParents=current_parents,
                    fields='id'
                ).execute()
                moved += 1
            except Exception:
                pass

        return jsonify({"status": "success", "moved": moved})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


# ============================
#  User Info API
# ============================
@app.route('/api/me', methods=['GET'])
def api_me():
    """Get current user info."""
    user = get_current_user()
    if not user:
        return jsonify({"error": "Not authenticated"}), 401
    return jsonify(user)


# ============================
#  Accounts API (Global)
# ============================
@app.route('/api/accounts', methods=['GET'])
def api_accounts():
    """Get all active accounts."""
    accounts = models.get_accounts()
    return jsonify({"accounts": accounts})


@app.route('/api/accounts', methods=['POST'])
def api_accounts_create():
    """Create a new account."""
    data = request.get_json()
    code = (data.get('code') or '').strip()
    name = (data.get('name') or '').strip()
    account_type = (data.get('account_type') or '').strip()
    tax_default = (data.get('tax_default') or '10%').strip()
    if not code or not name or not account_type:
        return jsonify({"status": "error", "error": "コード・科目名・区分は必須です"}), 400
    try:
        account = models.create_account(code, name, account_type, tax_default)
        return jsonify({"status": "success", "account": account})
    except Exception as e:
        return jsonify({"status": "error", "error": str(e)}), 400


@app.route('/api/accounts/<int:account_id>', methods=['DELETE'])
def api_accounts_delete(account_id):
    """Soft-delete an account."""
    ok = models.delete_account(account_id)
    if ok:
        return jsonify({"status": "success"})
    else:
        return jsonify({"status": "error", "error": "この科目は仕訳で使用中のため削除できません"}), 400


# ============================
#  Journal Entries API (user-scoped)
# ============================
@app.route('/api/journal', methods=['GET'])
def api_journal_list():
    """List journal entries with filters and pagination."""
    uid = get_user_id()
    filters = {
        'start_date': request.args.get('start_date'),
        'end_date': request.args.get('end_date'),
        'account_id': request.args.get('account_id'),
        'counterparty': request.args.get('counterparty'),
        'memo': request.args.get('memo'),
        'amount_min': request.args.get('amount_min'),
        'amount_max': request.args.get('amount_max'),
        'page': request.args.get('page', 1),
        'per_page': request.args.get('per_page', 20),
    }
    filters = {k: v for k, v in filters.items() if v is not None}
    result = models.get_journal_entries(filters, user_id=uid)
    return jsonify(result)


@app.route('/api/journal/recent', methods=['GET'])
def api_journal_recent():
    """Get recent journal entries."""
    uid = get_user_id()
    limit = request.args.get('limit', 5, type=int)
    entries = models.get_recent_entries(limit, user_id=uid)
    return jsonify({"entries": entries})


@app.route('/api/journal', methods=['POST'])
def api_journal_create():
    """Create one or more journal entries."""
    print(f"[Journal POST] Request received", flush=True)
    uid = get_user_id()
    print(f"[Journal POST] user_id={uid}", flush=True)
    data = request.json
    if not data:
        return jsonify({"error": "No data provided"}), 400

    if isinstance(data, list):
        entries = data
    elif isinstance(data, dict) and 'entries' in data:
        entries = data['entries']
    else:
        entries = [data]

    print(f"[Journal POST] Creating {len(entries)} entries", flush=True)
    try:
        ids = models.create_journal_entries_batch(entries, user_id=uid)
        print(f"[Journal POST] Created ids={ids}", flush=True)
        successful = [i for i in ids if i is not None]
        return jsonify({
            "status": "success",
            "created": len(successful),
            "ids": successful,
        })
    except Exception as e:
        print(f"Journal create error (user_id={uid}): {e}")
        return jsonify({"error": str(e)}), 500


@app.route('/api/journal/<int:entry_id>', methods=['PUT'])
def api_journal_update(entry_id):
    """Update a journal entry."""
    uid = get_user_id()
    data = request.json
    if not data:
        return jsonify({"error": "No data provided"}), 400

    success = models.update_journal_entry(entry_id, data, user_id=uid)
    if success:
        return jsonify({"status": "success"})
    else:
        return jsonify({"error": "Update failed"}), 400


@app.route('/api/journal/<int:entry_id>', methods=['DELETE'])
def api_journal_delete(entry_id):
    """Soft-delete a journal entry."""
    uid = get_user_id()
    success = models.delete_journal_entry(entry_id, user_id=uid)
    if success:
        return jsonify({"status": "success"})
    else:
        return jsonify({"error": "Delete failed"}), 400


# ============================
#  Trial Balance API (user-scoped)
# ============================
@app.route('/api/trial-balance', methods=['GET'])
def api_trial_balance():
    """Get trial balance by account for a date range."""
    uid = get_user_id()
    start_date = request.args.get('start_date')
    end_date = request.args.get('end_date')
    result = models.get_trial_balance(start_date, end_date, user_id=uid)
    return jsonify({"balances": result})


# ============================
#  AI Analysis API (user-scoped)
# ============================
@app.route('/api/analyze', methods=['POST'])
def api_analyze():
    """Upload and analyze files with AI."""
    uid = get_user_id()
    if 'files' not in request.files:
        return jsonify({"error": "No files"}), 400

    files = request.files.getlist('files')
    gemini_api_key = request.form.get('gemini_api_key')
    access_token = request.form.get('access_token', '')

    if not gemini_api_key:
        return jsonify({"error": "Missing Gemini API key"}), 401

    _get_ai_service().configure_gemini(gemini_api_key)

    history = models.get_accounting_history(user_id=uid)
    existing = models.get_existing_entry_keys(user_id=uid)

    results = []
    for file in files:
        fname = file.filename.lower()
        ftype = file.content_type
        fbytes = file.read()

        ev_url = ""
        if access_token:
            ev_url = upload_to_drive_processed(fbytes, file.filename, ftype, access_token)

        if fname.endswith('.csv'):
            res = _get_ai_service().analyze_csv(fbytes, history)
        else:
            res = _get_ai_service().analyze_document(fbytes, ftype, history)

        for item in res:
            item['evidence_url'] = ev_url
            key = f"{item.get('date')}_{str(item.get('amount'))}_{item.get('counterparty')}"
            item['is_duplicate'] = key in existing
        results.extend(res)

    return jsonify(results)


@app.route('/api/predict', methods=['POST'])
def api_predict():
    """AI-powered account prediction."""
    uid = get_user_id()
    data = request.json.get('data', [])
    gemini_api_key = request.json.get('gemini_api_key')

    if not data or not gemini_api_key:
        return jsonify({"error": "Missing data or API key"}), 400

    history = models.get_accounting_history(user_id=uid)
    accounts = models.get_accounts()
    valid_account_names = [a['name'] for a in accounts]

    try:
        predictions = _get_ai_service().predict_accounts(data, history, valid_account_names, gemini_api_key)
        return jsonify(predictions)
    except Exception as e:
        return jsonify({"error": str(e)}), 500


# ============================
#  AI Chat API (user-scoped)
# ============================
@app.route('/api/chat', methods=['POST'])
def api_chat():
    """AI chat endpoint for accounting consultation."""
    data = request.json or {}
    message = data.get('message', '').strip()
    history = data.get('history', [])
    gemini_api_key = data.get('gemini_api_key', '')

    if not message:
        return jsonify({"error": "メッセージが空です"}), 400
    if not gemini_api_key:
        return jsonify({"error": "Gemini APIキーが設定されていません"}), 401

    _get_ai_service().configure_gemini(gemini_api_key)

    account_list = models.get_accounts()
    account_names = ", ".join([f"{a['name']}({a['code']})" for a in account_list[:40]])

    system_prompt = f"""あなたは「Hinakira会計」の会計AIアシスタントです。
以下の役割で、ユーザーの質問に日本語で丁寧に答えてください。

【あなたの役割】
- 日本の個人事業主向け会計・税務の相談対応
- このツール（Hinakira会計）の使い方の案内
- 仕訳の考え方、勘定科目の選び方のアドバイス
- 消費税区分（課税仕入、課税売上、非課税、不課税）の判定支援
- 確定申告・青色申告に関する一般的な質問への回答

【このツールの機能】
- 仕訳入力: 手動入力 + AI自動判定（摘要を入力→借方・貸方・税区分をAI推定）
- 証憑読み取り: レシート・請求書をアップロード → AIが仕訳を自動作成、Google Driveに証憑保存
- 仕訳帳: 登録済み仕訳の一覧・検索・編集・削除
- 勘定科目残高: 貸借対照表(B/S)・損益計算書(P/L)を月次/年次で表示
- 取引先管理: 取引先の登録・管理
- 勘定科目管理: 科目の追加・削除
- 期首残高: 期首残高の設定
- バックアップ: JSON/Google Drive バックアップ・復元
- アウトプット: 仕訳帳・総勘定元帳・B/S・P/LのCSV/PDF出力

【登録されている勘定科目】
{account_names}

【注意事項】
- 税理士資格に基づく個別の税務判断や申告代行はできません
- 一般的な会計知識に基づいたアドバイスを提供してください
- 回答は簡潔にわかりやすくしてください
"""

    try:
        model = _get_genai().GenerativeModel('gemini-2.5-flash')

        contents = []
        contents.append({"role": "user", "parts": [{"text": system_prompt + "\n\n（ここからユーザーとの会話開始）"}]})
        contents.append({"role": "model", "parts": [{"text": "はい、Hinakira会計のAIアシスタントです。会計処理やツールの使い方について、お気軽にご質問ください。"}]})

        for h in history[-10:]:
            role = "user" if h.get('role') == 'user' else "model"
            contents.append({"role": role, "parts": [{"text": h.get('text', '')}]})

        contents.append({"role": "user", "parts": [{"text": message}]})

        response = model.generate_content(contents)
        return jsonify({"reply": response.text})
    except Exception as e:
        print(f"Chat Error: {e}")
        return jsonify({"error": f"AI応答に失敗しました: {str(e)}"}), 500


# ============================
#  Settings API (user-scoped)
# ============================
@app.route('/api/settings', methods=['GET'])
def api_settings_get():
    """Get user's settings."""
    uid = get_user_id()
    conn = db.get_db()
    try:
        rows = conn.execute("SELECT key, value FROM user_settings WHERE user_id = ?", (uid,)).fetchall()
        return jsonify({r['key']: r['value'] for r in rows})
    finally:
        conn.close()


@app.route('/api/settings', methods=['POST'])
def api_settings_update():
    """Update user's settings."""
    uid = get_user_id()
    data = request.json
    if not data:
        return jsonify({"error": "No data"}), 400

    conn = db.get_db()
    try:
        for key, value in data.items():
            conn.execute(
                "INSERT OR REPLACE INTO user_settings (user_id, key, value) VALUES (?, ?, ?)",
                (uid, key, str(value))
            )
        conn.commit()
        return jsonify({"status": "success"})
    except Exception as e:
        conn.rollback()
        return jsonify({"error": str(e)}), 500
    finally:
        conn.close()


# ============================
#  Counterparties API (user-scoped)
# ============================
@app.route('/api/counterparties', methods=['GET'])
def api_counterparties():
    """Get counterparty names for autocomplete."""
    uid = get_user_id()
    counterparties = models.get_counterparty_names(user_id=uid)
    return jsonify({"counterparties": counterparties})


@app.route('/api/counterparties/list', methods=['GET'])
def api_counterparties_list():
    """Get all counterparties with full details."""
    uid = get_user_id()
    items = models.get_counterparties_list(user_id=uid)
    return jsonify({"counterparties": items})


@app.route('/api/counterparties', methods=['POST'])
def api_counterparty_create():
    """Create a new counterparty."""
    uid = get_user_id()
    data = request.json
    if not data or not data.get('name'):
        return jsonify({"error": "Name is required"}), 400
    new_id = models.create_counterparty(data, user_id=uid)
    return jsonify({"status": "success", "id": new_id})


@app.route('/api/counterparties/<int:cp_id>', methods=['PUT'])
def api_counterparty_update(cp_id):
    """Update a counterparty."""
    uid = get_user_id()
    data = request.json
    if not data:
        return jsonify({"error": "No data"}), 400
    success = models.update_counterparty(cp_id, data, user_id=uid)
    return jsonify({"status": "success"}) if success else (jsonify({"error": "Failed"}), 400)


@app.route('/api/counterparties/<int:cp_id>', methods=['DELETE'])
def api_counterparty_delete(cp_id):
    """Soft-delete a counterparty."""
    uid = get_user_id()
    success = models.delete_counterparty(cp_id, user_id=uid)
    return jsonify({"status": "success"}) if success else (jsonify({"error": "Failed"}), 400)


# ============================
#  Opening Balances API (user-scoped)
# ============================
@app.route('/api/opening-balances', methods=['GET'])
def api_opening_balances_get():
    """Get opening balances for a fiscal year."""
    uid = get_user_id()
    fiscal_year = request.args.get('fiscal_year', str(__import__('datetime').date.today().year))
    balances = models.get_opening_balances(fiscal_year, user_id=uid)
    return jsonify({"balances": balances, "fiscal_year": fiscal_year})


@app.route('/api/opening-balances', methods=['POST'])
def api_opening_balances_save():
    """Bulk save opening balances."""
    uid = get_user_id()
    data = request.json
    if not data:
        return jsonify({"error": "No data"}), 400
    fiscal_year = data.get('fiscal_year', str(__import__('datetime').date.today().year))
    balances = data.get('balances', [])
    success = models.save_opening_balances(fiscal_year, balances, user_id=uid)
    return jsonify({"status": "success"}) if success else (jsonify({"error": "Save failed"}), 500)


# ============================
#  General Ledger API (user-scoped)
# ============================
@app.route('/api/ledger/<int:account_id>', methods=['GET'])
def api_ledger(account_id):
    """Get per-account entries with running balance."""
    uid = get_user_id()
    start_date = request.args.get('start_date')
    end_date = request.args.get('end_date')
    result = models.get_ledger_entries(account_id, start_date, end_date, user_id=uid)
    return jsonify(result)


# ============================
#  Backup API (user-scoped)
# ============================
@app.route('/api/backup/download', methods=['GET'])
def api_backup_download():
    """Download database backup as JSON."""
    uid = get_user_id()
    data = models.get_full_backup(user_id=uid)
    return jsonify(data)


@app.route('/api/backup/drive', methods=['POST'])
def api_backup_to_drive():
    """Upload JSON backup to Google Drive."""
    uid = get_user_id()
    data = request.json or {}
    access_token = data.get('access_token', '')
    if not access_token:
        return jsonify({"error": "Googleログインが必要です"}), 401

    try:
        backup_data = models.get_full_backup(user_id=uid)
        backup_json = json.dumps(backup_data, ensure_ascii=False, indent=2)
        backup_bytes = backup_json.encode('utf-8')

        import datetime
        filename = f"hinakira_backup_{datetime.date.today().isoformat()}.json"
        link = upload_to_drive(backup_bytes, filename, 'application/json', access_token)
        if link:
            return jsonify({"status": "success", "link": link, "filename": filename})
        else:
            return jsonify({"error": "Driveへのアップロードに失敗しました"}), 500
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route('/api/backup/drive/list', methods=['POST'])
def api_backup_drive_list():
    """List backup files from Google Drive."""
    data = request.json or {}
    access_token = data.get('access_token', '')
    if not access_token:
        return jsonify({"error": "Googleログインが必要です"}), 401

    try:
        creds = _get_credentials()(token=access_token)
        service = _get_drive_build()('drive', 'v3', credentials=creds)

        folder_name = "Accounting_Evidence"
        results = service.files().list(
            q=f"mimeType='application/vnd.google-apps.folder' and name='{folder_name}' and trashed=false",
            spaces='drive'
        ).execute()
        items = results.get('files', [])
        if not items:
            return jsonify({"files": []})

        folder_id = items[0]['id']

        results = service.files().list(
            q=f"'{folder_id}' in parents and name contains 'hinakira_backup' and mimeType='application/json' and trashed=false",
            spaces='drive',
            fields='files(id, name, createdTime, size)',
            orderBy='createdTime desc',
            pageSize=20
        ).execute()
        files = results.get('files', [])
        return jsonify({"files": files})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route('/api/backup/drive/restore', methods=['POST'])
def api_backup_drive_restore():
    """Download and restore a backup file from Google Drive."""
    uid = get_user_id()
    data = request.json or {}
    access_token = data.get('access_token', '')
    file_id = data.get('file_id', '')
    if not access_token:
        return jsonify({"error": "Googleログインが必要です"}), 401
    if not file_id:
        return jsonify({"error": "ファイルが指定されていません"}), 400

    try:
        creds = _get_credentials()(token=access_token)
        service = _get_drive_build()('drive', 'v3', credentials=creds)

        from googleapiclient.http import MediaIoBaseDownload
        request_dl = service.files().get_media(fileId=file_id)
        buffer = io.BytesIO()
        downloader = MediaIoBaseDownload(buffer, request_dl)
        done = False
        while not done:
            _, done = downloader.next_chunk()

        buffer.seek(0)
        content = buffer.read().decode('utf-8')
        backup_data = json.loads(content)

        expected_tables = ['accounts_master', 'journal_entries', 'opening_balances', 'counterparties', 'settings']
        if not any(t in backup_data for t in expected_tables):
            return jsonify({"error": "有効なバックアップデータが見つかりません"}), 400

        summary = models.restore_from_backup(backup_data, user_id=uid)
        return jsonify({"status": "success", "summary": summary})
    except Exception as e:
        return jsonify({"error": f"復元に失敗しました: {str(e)}"}), 500


@app.route('/api/backup/restore', methods=['POST'])
def api_backup_restore():
    """Restore database from JSON backup file."""
    uid = get_user_id()
    if 'file' not in request.files:
        return jsonify({"error": "ファイルが選択されていません"}), 400

    file = request.files['file']
    if not file.filename:
        return jsonify({"error": "ファイルが選択されていません"}), 400

    if not file.filename.lower().endswith('.json'):
        return jsonify({"error": "JSONファイルのみ対応しています"}), 400

    try:
        content = file.read().decode('utf-8')
        data = json.loads(content)
    except (UnicodeDecodeError, json.JSONDecodeError) as e:
        return jsonify({"error": f"JSONファイルの読み込みに失敗しました: {str(e)}"}), 400

    expected_tables = ['accounts_master', 'journal_entries', 'opening_balances', 'counterparties', 'settings']
    has_any = any(t in data for t in expected_tables)
    if not has_any:
        return jsonify({"error": "有効なバックアップデータが見つかりません"}), 400

    try:
        summary = models.restore_from_backup(data, user_id=uid)
        return jsonify({"status": "success", "summary": summary})
    except Exception as e:
        return jsonify({"error": f"復元に失敗しました: {str(e)}"}), 500


# ============================
#  Export API (user-scoped)
# ============================
@app.route('/api/export/journal', methods=['GET'])
def api_export_journal():
    """Export journal entries as CSV or JSON."""
    uid = get_user_id()
    start_date = request.args.get('start_date')
    end_date = request.args.get('end_date')
    fmt = request.args.get('format', 'json')
    entries = models.get_journal_export(start_date, end_date, user_id=uid)

    if fmt == 'csv':
        import csv
        output = io.StringIO()
        if entries:
            writer = csv.DictWriter(output, fieldnames=list(entries[0].keys()))
            writer.writeheader()
            writer.writerows(entries)
        from flask import Response
        return Response(output.getvalue(), mimetype='text/csv',
                       headers={"Content-Disposition": "attachment; filename=journal_export.csv"})
    return jsonify({"entries": entries})


@app.route('/api/export/trial-balance', methods=['GET'])
def api_export_trial_balance():
    """Export trial balance data."""
    uid = get_user_id()
    start_date = request.args.get('start_date')
    end_date = request.args.get('end_date')
    balances = models.get_trial_balance(start_date, end_date, user_id=uid)
    return jsonify({"balances": balances})


@app.route('/api/export/ledger', methods=['GET'])
def api_export_ledger():
    """Export general ledger for all accounts."""
    uid = get_user_id()
    start_date = request.args.get('start_date')
    end_date = request.args.get('end_date')
    account_list = models.get_accounts()
    result = []
    for acct in account_list:
        ledger = models.get_ledger_entries(acct['id'], start_date, end_date, user_id=uid)
        if ledger.get('entries'):
            result.append({
                'account_code': acct['code'],
                'account_name': acct['name'],
                'account_type': acct['account_type'],
                'opening_balance': ledger.get('opening_balance', 0),
                'entries': ledger.get('entries', [])
            })
    return jsonify({"accounts": result})


# ============================
#  Static File Routes (MUST be last to avoid intercepting API routes)
# ============================
@app.route('/')
def index():
    return send_from_directory('.', 'index.html')


@app.route('/<path:path>')
def static_files(path):
    return send_from_directory('.', path)


# ============================
#  Main
# ============================
if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5001))
    app.run(host='0.0.0.0', port=port)
