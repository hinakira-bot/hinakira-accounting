"""
Hinakira会計 - Flask Backend
Main routing and API endpoints.
"""
import os
import io
import json
from flask import Flask, request, jsonify, send_from_directory
from flask_cors import CORS
from dotenv import load_dotenv
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload

import db
import models
import ai_service

load_dotenv()

app = Flask(__name__, static_folder='.', static_url_path='/static')
CORS(app)

# Initialize database on startup
db.init_db()


# --- Google Drive Upload Helper ---
def upload_to_drive(file_bytes, filename, mime_type, access_token):
    """Upload evidence file to Google Drive."""
    try:
        if not access_token:
            return ""
        creds = Credentials(token=access_token)
        service = build('drive', 'v3', credentials=creds)

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
        media = MediaIoBaseUpload(io.BytesIO(file_bytes), mimetype=mime_type, resumable=True)
        file = service.files().create(body=file_metadata, media_body=media, fields='id, webViewLink').execute()
        return file.get('webViewLink')
    except Exception as e:
        print(f"Drive Upload Error: {e}")
        return ""


# ============================
#  Accounts API
# ============================
@app.route('/api/accounts', methods=['GET'])
def api_accounts():
    """Get all active accounts."""
    accounts = models.get_accounts()
    return jsonify({"accounts": accounts})


# ============================
#  Journal Entries API (CRUD)
# ============================
@app.route('/api/journal', methods=['GET'])
def api_journal_list():
    """List journal entries with filters and pagination."""
    filters = {
        'start_date': request.args.get('start_date'),
        'end_date': request.args.get('end_date'),
        'account_id': request.args.get('account_id'),
        'counterparty': request.args.get('counterparty'),
        'memo': request.args.get('memo'),
        'page': request.args.get('page', 1),
        'per_page': request.args.get('per_page', 20),
    }
    # Remove None values
    filters = {k: v for k, v in filters.items() if v is not None}
    result = models.get_journal_entries(filters)
    return jsonify(result)


@app.route('/api/journal/recent', methods=['GET'])
def api_journal_recent():
    """Get recent journal entries (for display below input form)."""
    limit = request.args.get('limit', 5, type=int)
    entries = models.get_recent_entries(limit)
    return jsonify({"entries": entries})


@app.route('/api/journal', methods=['POST'])
def api_journal_create():
    """Create one or more journal entries."""
    data = request.json
    if not data:
        return jsonify({"error": "No data provided"}), 400

    # Support both single entry and batch
    if isinstance(data, list):
        entries = data
    elif isinstance(data, dict) and 'entries' in data:
        entries = data['entries']
    else:
        entries = [data]

    ids = models.create_journal_entries_batch(entries)
    successful = [i for i in ids if i is not None]
    return jsonify({
        "status": "success",
        "created": len(successful),
        "ids": successful,
    })


@app.route('/api/journal/<int:entry_id>', methods=['PUT'])
def api_journal_update(entry_id):
    """Update a journal entry."""
    data = request.json
    if not data:
        return jsonify({"error": "No data provided"}), 400

    success = models.update_journal_entry(entry_id, data)
    if success:
        return jsonify({"status": "success"})
    else:
        return jsonify({"error": "Update failed"}), 400


@app.route('/api/journal/<int:entry_id>', methods=['DELETE'])
def api_journal_delete(entry_id):
    """Soft-delete a journal entry."""
    success = models.delete_journal_entry(entry_id)
    if success:
        return jsonify({"status": "success"})
    else:
        return jsonify({"error": "Delete failed"}), 400


# ============================
#  Trial Balance API
# ============================
@app.route('/api/trial-balance', methods=['GET'])
def api_trial_balance():
    """Get trial balance by account for a date range."""
    start_date = request.args.get('start_date')
    end_date = request.args.get('end_date')
    result = models.get_trial_balance(start_date, end_date)
    return jsonify({"balances": result})


# ============================
#  AI Analysis API
# ============================
@app.route('/api/analyze', methods=['POST'])
def api_analyze():
    """Upload and analyze files with AI."""
    if 'files' not in request.files:
        return jsonify({"error": "No files"}), 400

    files = request.files.getlist('files')
    gemini_api_key = request.form.get('gemini_api_key')
    access_token = request.form.get('access_token', '')

    if not gemini_api_key:
        return jsonify({"error": "Missing Gemini API key"}), 401

    ai_service.configure_gemini(gemini_api_key)

    # Get history from SQLite (instead of Sheets)
    history = models.get_accounting_history()
    existing = models.get_existing_entry_keys()

    results = []
    for file in files:
        fname = file.filename.lower()
        ftype = file.content_type
        fbytes = file.read()

        # Upload evidence to Drive (optional)
        ev_url = ""
        if access_token:
            ev_url = upload_to_drive(fbytes, file.filename, ftype, access_token)

        # AI Analysis
        if fname.endswith('.csv'):
            res = ai_service.analyze_csv(fbytes, history)
        else:
            res = ai_service.analyze_document(fbytes, ftype, history)

        for item in res:
            item['evidence_url'] = ev_url
            key = f"{item.get('date')}_{str(item.get('amount'))}_{item.get('counterparty')}"
            item['is_duplicate'] = key in existing
        results.extend(res)

    return jsonify(results)


@app.route('/api/predict', methods=['POST'])
def api_predict():
    """AI-powered account prediction."""
    data = request.json.get('data', [])
    gemini_api_key = request.json.get('gemini_api_key')

    if not data or not gemini_api_key:
        return jsonify({"error": "Missing data or API key"}), 400

    history = models.get_accounting_history()
    accounts = models.get_accounts()
    valid_account_names = [a['name'] for a in accounts]

    try:
        predictions = ai_service.predict_accounts(data, history, valid_account_names, gemini_api_key)
        return jsonify(predictions)
    except Exception as e:
        return jsonify({"error": str(e)}), 500


# ============================
#  Settings API
# ============================
@app.route('/api/settings', methods=['GET'])
def api_settings_get():
    """Get all settings."""
    conn = db.get_db()
    try:
        rows = conn.execute("SELECT key, value FROM settings").fetchall()
        return jsonify({r['key']: r['value'] for r in rows})
    finally:
        conn.close()


@app.route('/api/settings', methods=['POST'])
def api_settings_update():
    """Update settings."""
    data = request.json
    if not data:
        return jsonify({"error": "No data"}), 400

    conn = db.get_db()
    try:
        for key, value in data.items():
            conn.execute(
                "INSERT OR REPLACE INTO settings (key, value) VALUES (?, ?)",
                (key, str(value))
            )
        conn.commit()
        return jsonify({"status": "success"})
    except Exception as e:
        conn.rollback()
        return jsonify({"error": str(e)}), 500
    finally:
        conn.close()


# ============================
#  Counterparties API (CRUD)
# ============================
@app.route('/api/counterparties', methods=['GET'])
def api_counterparties():
    """Get counterparty names for autocomplete."""
    counterparties = models.get_counterparty_names()
    return jsonify({"counterparties": counterparties})


@app.route('/api/counterparties/list', methods=['GET'])
def api_counterparties_list():
    """Get all counterparties with full details."""
    items = models.get_counterparties_list()
    return jsonify({"counterparties": items})


@app.route('/api/counterparties', methods=['POST'])
def api_counterparty_create():
    """Create a new counterparty."""
    data = request.json
    if not data or not data.get('name'):
        return jsonify({"error": "Name is required"}), 400
    new_id = models.create_counterparty(data)
    return jsonify({"status": "success", "id": new_id})


@app.route('/api/counterparties/<int:cp_id>', methods=['PUT'])
def api_counterparty_update(cp_id):
    """Update a counterparty."""
    data = request.json
    if not data:
        return jsonify({"error": "No data"}), 400
    success = models.update_counterparty(cp_id, data)
    return jsonify({"status": "success"}) if success else (jsonify({"error": "Failed"}), 400)


@app.route('/api/counterparties/<int:cp_id>', methods=['DELETE'])
def api_counterparty_delete(cp_id):
    """Soft-delete a counterparty."""
    success = models.delete_counterparty(cp_id)
    return jsonify({"status": "success"}) if success else (jsonify({"error": "Failed"}), 400)


# ============================
#  Opening Balances API
# ============================
@app.route('/api/opening-balances', methods=['GET'])
def api_opening_balances_get():
    """Get opening balances for a fiscal year."""
    fiscal_year = request.args.get('fiscal_year', str(__import__('datetime').date.today().year))
    balances = models.get_opening_balances(fiscal_year)
    return jsonify({"balances": balances, "fiscal_year": fiscal_year})


@app.route('/api/opening-balances', methods=['POST'])
def api_opening_balances_save():
    """Bulk save opening balances."""
    data = request.json
    if not data:
        return jsonify({"error": "No data"}), 400
    fiscal_year = data.get('fiscal_year', str(__import__('datetime').date.today().year))
    balances = data.get('balances', [])
    success = models.save_opening_balances(fiscal_year, balances)
    return jsonify({"status": "success"}) if success else (jsonify({"error": "Save failed"}), 500)


# ============================
#  General Ledger API
# ============================
@app.route('/api/ledger/<int:account_id>', methods=['GET'])
def api_ledger(account_id):
    """Get per-account entries with running balance."""
    start_date = request.args.get('start_date')
    end_date = request.args.get('end_date')
    result = models.get_ledger_entries(account_id, start_date, end_date)
    return jsonify(result)


# ============================
#  Backup API
# ============================
@app.route('/api/backup/download', methods=['GET'])
def api_backup_download():
    """Download database backup."""
    format_type = request.args.get('format', 'json')
    if format_type == 'sqlite':
        return send_from_directory('data', 'accounting.db', as_attachment=True,
                                   download_name='accounting_backup.db')
    else:
        data = models.get_full_backup()
        return jsonify(data)


# ============================
#  Export API
# ============================
@app.route('/api/export/journal', methods=['GET'])
def api_export_journal():
    """Export journal entries as CSV or JSON."""
    start_date = request.args.get('start_date')
    end_date = request.args.get('end_date')
    fmt = request.args.get('format', 'json')
    entries = models.get_journal_export(start_date, end_date)

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
    start_date = request.args.get('start_date')
    end_date = request.args.get('end_date')
    balances = models.get_trial_balance(start_date, end_date)
    return jsonify({"balances": balances})


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
