import os
import json
import base64
from flask import Flask, request, jsonify, send_from_directory
from flask_cors import CORS
import google.generativeai as genai
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from dotenv import load_dotenv
from PIL import Image
import io

load_dotenv()

app = Flask(__name__, static_folder='.')
CORS(app)

# --- Configuration ---
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")
SPREADSHEET_ID = os.getenv("SPREADSHEET_ID")
GOOGLE_CREDENTIALS_FILE = "credentials.json" # Service Account JSON

if GEMINI_API_KEY:
    genai.configure(api_key=GEMINI_API_KEY)

# --- Google Sheets Logic (History Only) ---
def get_accounting_history():
    print("Fetching accounting history from sheets...")
    if not os.path.exists(GOOGLE_CREDENTIALS_FILE):
        return {}
    try:
        scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
        creds = ServiceAccountCredentials.from_json_keyfile_name(GOOGLE_CREDENTIALS_FILE, scope)
        client = gspread.authorize(creds)
        sh = client.open_by_key(SPREADSHEET_ID)
        sheet = sh.worksheet("仕訳明細")
        data = sheet.get_all_values()
        
        # 履歴の辞書作成 { 取引先: 借方勘定 }
        history = {}
        if len(data) > 1:
            for row in data[1:]:
                if len(row) >= 5:
                    counterparty = row[4].strip()
                    debit = row[1].strip()
                    if counterparty and debit:
                        history[counterparty] = debit
        return history
    except Exception as e:
        print(f"Error fetching history: {e}")
        return {}

import csv
import io

# --- CSV Logic ---
def analyze_csv(csv_bytes, history={}):
    print("Analyzing CSV statement...")
    try:
        text = csv_bytes.decode('shift_jis', errors='replace') # 日本のカード会社はShift-JISが多い
        if '確定日' not in text and '利用日' not in text and ',' not in text:
             text = csv_bytes.decode('utf-8', errors='replace')

        f = io.StringIO(text)
        reader = csv.reader(f)
        rows = list(reader)
        
        # CSVの内容をテキスト化してAIに整形させる
        csv_text = "\n".join([",".join(row) for row in rows[:50]]) # 最初の50行程度
        
        model = genai.GenerativeModel('gemini-2.0-flash-exp')
        history_str = json.dumps(history, ensure_ascii=False, indent=2) if history else "なし"
        
        prompt = f"""
        あなたは優秀な会計士です。以下のクレジットカード明細（CSV形式）から、仕訳データを抽出してください。
        
        履歴データ（優先）:
        {history_str}

        JSON形式（配列）で出力してください：
        [
          {{
            "date": "YYYY-MM-DD",
            "debit_account": "借方勘定科目",
            "credit_account": "貸方勘定科目",
            "amount": 数値,
            "counterparty": "取引先名",
            "memo": "カード利用明細"
          }}
        ]
        
        明細データ:
        {csv_text}
        """
        
        response = model.generate_content(prompt)
        content = response.text.strip()
        if "```json" in content:
            content = content.split("```json")[-1].split("```")[0].strip()
        return json.loads(content)
    except Exception as e:
        print(f"Error in analyze_csv: {e}")
        return []

# --- Google Sheets Logic (Duplicate Check) ---
def get_existing_entries():
    print("Fetching existing entries for duplicate check...")
    if not os.path.exists(GOOGLE_CREDENTIALS_FILE):
        return set()
    try:
        scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
        creds = ServiceAccountCredentials.from_json_keyfile_name(GOOGLE_CREDENTIALS_FILE, scope)
        client = gspread.authorize(creds)
        sh = client.open_by_key(SPREADSHEET_ID)
        sheet = sh.worksheet("仕訳明細")
        data = sheet.get_all_values()
        
        # (日付, 金額, 取引先) のセットを作成
        existing = set()
        if len(data) > 1:
            for row in data[1:]:
                if len(row) >= 5:
                    date = row[0].strip()
                    amount = str(row[3]).strip()
                    counterparty = row[4].strip()
                    existing.add(f"{date}_{amount}_{counterparty}")
        return existing
    except Exception as e:
        print(f"Error fetching existing entries: {e}")
        return set()

# --- AI Logic ---
def analyze_document(file_bytes, mime_type, history={}):
    print(f"Analyzing {mime_type} with history context...")
    models_to_try = ['gemini-2.0-flash-exp', 'gemini-1.5-flash']
    
    history_str = json.dumps(history, ensure_ascii=False, indent=2) if history else "なし"

    for model_name in models_to_try:
        try:
            model = genai.GenerativeModel(model_name)
            
            prompt = f"""
            あなたは日本の税務・会計士です。
            渡されたドキュメント（レシート、領収書、またはクレジットカードの利用明細書）から仕訳を作成してください。

            履歴（優先）:
            {history_str}

            JSON形式（配列）で出力してください。他の説明は不要です。
            [
              {{
                "date": "YYYY-MM-DD",
                "debit_account": "借方勘定科目",
                "credit_account": "貸方勘定科目",
                "amount": 数値,
                "counterparty": "取引先名",
                "memo": "詳細・摘要"
              }}
            ]
            """
            
            if mime_type.startswith('image/'):
                content_part = Image.open(io.BytesIO(file_bytes))
            else:
                # PDFなどのドキュメント
                content_part = {
                    "mime_type": mime_type,
                    "data": file_bytes
                }
                
            response = model.generate_content([prompt, content_part])
            
            content = response.text.strip()
            if "```json" in content:
                content = content.split("```json")[-1].split("```")[0].strip()
            elif "```" in content:
                content = content.split("```")[-1].split("```")[0].strip()
            
            start_idx = content.find("[")
            end_idx = content.rfind("]")
            if start_idx != -1 and end_idx != -1:
                content = content[start_idx:end_idx+1]
                
            return json.loads(content)
        except Exception as e:
            print(f"Error with {model_name}: {e}")
            continue
    return []

# --- Google Sheets Logic (Save Only) ---
def save_to_sheets(data):
    print(f"Saving {len(data)} items to sheets...")
    if not os.path.exists(GOOGLE_CREDENTIALS_FILE):
        print("Credentials file not found.")
        return False
        
    try:
        scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
        creds = ServiceAccountCredentials.from_json_keyfile_name(GOOGLE_CREDENTIALS_FILE, scope)
        client = gspread.authorize(creds)
        
        sh = client.open_by_key(SPREADSHEET_ID)
        
        # --- Sheet 1: 仕訳明細 ---
        try:
            sheet1 = sh.get_worksheet(0)
            if sheet1.title != "仕訳明細":
                sheet1.update_title("仕訳明細")
        except:
            sheet1 = sh.add_worksheet(title="仕訳明細", rows="1000", cols="6")

        # 1行目がヘッダー設定
        headers = ["日にち", "借方", "貸方", "金額", "取引先", "摘要（内容）"]
        existing_values = sheet1.get_all_values()
        
        if not existing_values:
            sheet1.append_row(headers)
            sheet1.freeze(rows=1)
        elif existing_values[0] != headers:
            if not any(existing_values[0]):
                 sheet1.update('A1', [headers])
                 sheet1.freeze(rows=1)
            else:
                sheet1.insert_row(headers, index=1)
                sheet1.freeze(rows=1)

        # --- Sheet 2: 損益計算書 (P&L) ---
        try:
            sheet2 = sh.worksheet("損益計算書")
        except gspread.exceptions.WorksheetNotFound:
            sheet2 = sh.add_worksheet(title="損益計算書", rows="100", cols="10")
            # 損益計算書のレイアウト作成
            sheet2.update('A1', [["勘定科目別 集計表 (提出用参考)"]])
            sheet2.update('A3', [["勘定科目", "合計金額"]])
            # QUERY関数を使って、Sheet1(仕訳明細)の「借方」と「金額」を自動集計する
            # B列: 借方勘定, D列: 金額
            formula = "=QUERY('仕訳明細'!A:F, \"select B, sum(D) where B != '' and B != '借方' group by B label sum(D) ''\", 1)"
            sheet2.update('A4', [[formula]], value_input_option='USER_ENTERED')
            sheet2.freeze(rows=3)

        # データの書き込み
        rows = []
        for item in data:
            rows.append([
                str(item.get('date', '')),
                str(item.get('debit_account', '')),
                str(item.get('credit_account', '')),
                item.get('amount', 0),
                str(item.get('counterparty', '')),
                str(item.get('memo', ''))
            ])
        
        sheet1.append_rows(rows, value_input_option='USER_ENTERED')
        print("Successfully saved!")
        return True
    except Exception as e:
        print(f"Detailed Error saving to sheets: {str(e)}")
        return False

# --- Routes ---
@app.route('/')
def index():
    return send_from_directory('.', 'index.html')

@app.route('/<path:path>')
def static_files(path):
    return send_from_directory('.', path)

@app.route('/api/analyze', methods=['POST'])
def api_analyze():
    if 'files' not in request.files:
        return jsonify({"error": "No files uploaded"}), 400
    
    # 履歴と既存エントリを取得
    history = get_accounting_history()
    existing_entries = get_existing_entries()
    
    files = request.files.getlist('files')
    all_results = []
    
    for file in files:
        filename = file.filename.lower()
        mime_type = file.content_type
        file_bytes = file.read()
        
        if filename.endswith('.csv'):
            results = analyze_csv(file_bytes, history)
        else:
            results = analyze_document(file_bytes, mime_type, history)
            
        # 重複チェックのタグ付け
        for item in results:
            key = f"{item.get('date', '')}_{item.get('amount', '')}_{item.get('counterparty', '')}"
            if key in existing_entries:
                item['is_duplicate'] = True
            else:
                item['is_duplicate'] = False
            
        all_results.extend(results)
    
    return jsonify(all_results)

@app.route('/api/save', methods=['POST'])
def api_save():
    data = request.json
    success = save_to_sheets(data)
    if success:
        return jsonify({"message": "Success"})
    else:
        return jsonify({"error": "Failed to save to sheets"}), 500

if __name__ == '__main__':
    app.run(debug=True, port=5001)
