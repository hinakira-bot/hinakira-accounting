import os
import json
import base64
from flask import Flask, request, jsonify, send_from_directory
from flask_cors import CORS
import google.generativeai as genai
import gspread
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload
from dotenv import load_dotenv
from PIL import Image
import io
import csv

load_dotenv()

app = Flask(__name__, static_folder='.')
CORS(app)

# --- Standard Japanese Account (Masta) ---
DEFAULT_ACCOUNTS = [
    ["勘定科目", "タイプ"],
    ["現金", "資産"],
    ["小口現金", "資産"],
    ["普通預金", "資産"],
    ["売掛金", "資産"],
    ["未収入金", "資産"],
    ["棚卸資産", "資産"],
    ["買掛金", "負債"],
    ["未払金", "負債"],
    ["借入金", "負債"],
    ["預り金", "負債"],
    ["資本金", "純資産"],
    ["元入金", "純資産"],
    ["売上高", "収益"],
    ["雑収入", "収益"],
    ["仕入高", "費用"],
    ["役員報酬", "費用"],
    ["給料手当", "費用"],
    ["外注工賃", "費用"],
    ["旅費交通費", "費用"],
    ["通信費", "費用"],
    ["広告宣伝費", "費用"],
    ["接待交際費", "費用"],
    ["消耗品費", "費用"],
    ["会議費", "費用"],
    ["水道光熱費", "費用"],
    ["地代家賃", "費用"],
    ["修繕費", "費用"],
    ["支払手数料", "費用"],
    ["租税公課", "費用"],
    ["新聞図書費", "費用"],
    ["雑費", "費用"],
    ["事業主貸", "資産"], # For Sole Proprietor
    ["事業主借", "負債"], # For Sole Proprietor
]

# --- Helpers ---
def get_gspread_client(access_token):
    try:
        if not access_token:
            return None
        creds = Credentials(token=access_token)
        return gspread.authorize(creds)
    except Exception as e:
        print(f"Auth Error: {e}")
        return None

def upload_to_drive(file_bytes, filename, mime_type, access_token):
    try:
        creds = Credentials(token=access_token)
        service = build('drive', 'v3', credentials=creds)
        
        # Check specific folder exists, if not create 'Accounting_Evidence_2025'
        folder_name = "Accounting_Evidence"
        results = service.files().list(q=f"mimeType='application/vnd.google-apps.folder' and name='{folder_name}' and trashed=false", spaces='drive').execute()
        items = results.get('files', [])
        
        if not items:
            folder_metadata = {
                'name': folder_name,
                'mimeType': 'application/vnd.google-apps.folder'
            }
            file = service.files().create(body=folder_metadata, fields='id').execute()
            folder_id = file.get('id')
        else:
            folder_id = items[0]['id']
            
        file_metadata = {
            'name': filename,
            'parents': [folder_id]
        }
        
        media = MediaIoBaseUpload(io.BytesIO(file_bytes), mimetype=mime_type, resumable=True)
        file = service.files().create(body=file_metadata, media_body=media, fields='id, webViewLink').execute()
        
        return file.get('webViewLink')
    except Exception as e:
        print(f"Drive Upload Error: {e}")
        return ""

# --- Google Sheets Logic (History Only) ---
def get_accounting_history(spreadsheet_id, access_token):
    print("Fetching accounting history (Semantic Reference) from sheets...")
    try:
        client = get_gspread_client(access_token)
        if not client:
             return []

        sh = client.open_by_key(spreadsheet_id)
        sheet = sh.worksheet("仕訳明細")
        data = sheet.get_all_values()
        
        history = []
        if len(data) > 1:
            recent_rows = data[1:][-200:]
            for row in recent_rows:
                if len(row) >= 6:
                    history.append({
                        "counterparty": row[4].strip(),
                        "memo": row[5].strip(),
                        "account": row[1].strip() # Debit account usually signifies expense type
                    })
        return history
    except Exception as e:
        print(f"Error fetching history: {e}")
        return []

# --- Standardized Prompts ---
def get_analysis_prompt(history_str, input_text_or_type):
    return f"""
    あなたは日本の税務・会計士です。
    入力されたデータ（{input_text_or_type}）から会計仕訳を作成してください。
    
    【重要ルール】
    1. **日付**: YYYY-MM-DD形式。不明な場合は今日の日付。
    2. **借方**: 費用科目（旅費交通費、消耗品費など）または資産増加（普通預金など）。
    3. **貸方**: 
       - クレジットカード/電子マネー/後払い → 「**未払金**」
       - 銀行振込/引落 → 「普通預金」
       - 現金払い → 「現金」
    4. **金額**: 税込金額。
    5. **消費税**: 
        - 10%標準対象 → 税込金額から計算。
        - 8%軽減税率（飲食料品など） → 8%で計算。
        - 非課税/不課税（給料、保険料、税金支払い、預金移動、借入返済など） → 消費税額は **0**。
    6. **摘要 (Memo)**:
        - **最重要**: 過去の履歴（下記）に類似の取引先がある場合、その「摘要」を極力踏襲または参考にしてください。
        - 履歴がない場合、単なる科目名ではなく「具体的な内容」を推測して記載してください（例：「会議費」ではなく「〇〇プロジェクト打ち合わせ」など）。

    【過去の学習データ (取引先: 摘要 => 科目)】
    {history_str}
    
    以下のJSON形式（配列）で出力してください。Markdownは不要です。
    [
      {{
        "date": "YYYY-MM-DD",
        "debit_account": "借方勘定科目",
        "credit_account": "貸方勘定科目",
        "amount": 税込金額(数値),
        "tax_amount": 消費税額(数値・推定),
        "counterparty": "取引先名",
        "memo": "摘要(履歴を優先)"
      }}
    ]
    """

# --- CSV Logic ---
def analyze_csv(csv_bytes, history=[]):
    print("Analyzing CSV...")
    try:
        text = csv_bytes.decode('shift_jis', errors='replace') 
        if '確定日' not in text and '利用日' not in text and ',' not in text:
             text = csv_bytes.decode('utf-8', errors='replace')

        f = io.StringIO(text)
        reader = csv.reader(f)
        rows = list(reader)
        csv_text = "\n".join([",".join(row) for row in rows[:60]]) # Limit context
        
        model = genai.GenerativeModel('gemini-2.0-flash-exp')
        # History format: Counterparty: Memo => Account
        history_str = "\n".join([f"- {h['counterparty']}: {h['memo']} => {h['account']}" for h in history[:50]])
        
        prompt = get_analysis_prompt(history_str, "CSV明細（クレジットカードまたは銀行）") + f"\n\nデータ:\n{csv_text}"
        
        response = model.generate_content(prompt)
        content = CleanJSON(response.text)
        return json.loads(content)
    except Exception as e:
        print(f"Error in analyze_csv: {e}")
        return []

# --- AI Logic (Images/PDF) ---
def analyze_document(file_bytes, mime_type, history=[]):
    print(f"Analyzing {mime_type}...")
    models = ['gemini-2.0-flash-exp', 'gemini-1.5-flash']
    history_str = "\n".join([f"- {h['counterparty']}: {h['memo']} => {h['account']}" for h in history[:50]])

    for model_name in models:
        try:
            model = genai.GenerativeModel(model_name)
            prompt = get_analysis_prompt(history_str, "領収書/請求書画像")
            
            if mime_type.startswith('image/'):
                content_part = Image.open(io.BytesIO(file_bytes))
            else:
                content_part = {"mime_type": mime_type, "data": file_bytes}
                
            response = model.generate_content([prompt, content_part])
            content = CleanJSON(response.text)
            return json.loads(content)
        except Exception as e:
            print(f"Error with {model_name}: {e}")
            continue
    return []

def CleanJSON(text):
    text = text.strip()
    if "```json" in text:
        text = text.split("```json")[-1].split("```")[0].strip()
    elif "```" in text:
        text = text.split("```")[-1].split("```")[0].strip()
    # Attempt to extract array if wrapped
    s = text.find('[')
    e = text.rfind(']')
    if s != -1 and e != -1:
        text = text[s:e+1]
    return text

def get_existing_entries(spreadsheet_id, access_token):
    try:
        client = get_gspread_client(access_token)
        if not client: return set()
        sh = client.open_by_key(spreadsheet_id)
        existing = set()
        for sname in ["仕訳明細", "仕訳明細（手入力）"]:
            try:
                sheet = sh.worksheet(sname)
                # Check column A(Date), D(Amount), E(Counterparty)
                # Indexes: 0, 3, 4
                data = sheet.get_all_values()
                if len(data) > 1:
                    for row in data[1:]:
                        if len(row) >= 5:
                            existing.add(f"{row[0]}_{row[3]}_{row[4]}")
            except: pass
        return existing
    except: return set()

# --- Google Sheets Logic (Save & Structure) ---
def save_to_sheets(data, spreadsheet_id, access_token):
    print(f"Saving {len(data)} items to sheets...")
    try:
        client = get_gspread_client(access_token)
        if not client: return False
        sh = client.open_by_key(spreadsheet_id)
        
        # 1. Ensure "勘定科目マスタ" exists
        try:
            sh.worksheet("勘定科目マスタ")
        except:
            ws_master = sh.add_worksheet(title="勘定科目マスタ", rows="100", cols="2")
            ws_master.update("A1", DEFAULT_ACCOUNTS)
            ws_master.hide() # Hide it to keep UI clean
            
        # 2. Ensure "期首残高" exists
        try:
            sh.worksheet("期首残高")
        except:
            ws_open = sh.add_worksheet(title="期首残高", rows="50", cols="3")
            ws_open.update("A1", [["勘定科目", "期首残高金額", "備考"]])
            # Default helpful rows
            ws_open.update("A2", [["現金", "0", ""], ["普通預金", "0", ""], ["資本金", "0", ""], ["元入金", "0", ""]])

        # 3. Setup Detail Sheets
        headers = ["日にち", "借方", "貸方", "金額", "取引先", "摘要", "消費税額", "証憑URL"]
        
        # Helper to init sheet with dropdowns
        def init_detail_sheet(name):
            try:
                ws = sh.worksheet(name)
            except:
                ws = sh.add_worksheet(title=name, rows="1000", cols="8")
            
            vals = ws.get_all_values()
            if not vals:
                ws.append_row(headers)
                ws.freeze(rows=1)
                # Apply Dropdown Logic (Data Validation)
                # Gspread new version: set_data_validation_for_cell_range
                # Range B2:B1000 and C2:C1000 -> Rule from '勘定科目マスタ'!A2:A
                try:
                    rule = gspread.utils.ValidationCondition(
                        "ONE_OF_RANGE",
                        ["=勘定科目マスタ!$A$2:$A$100"]
                    )
                    ws.set_data_validation("B2:B1000", rule)
                    ws.set_data_validation("C2:C1000", rule)
                except Exception as ex:
                    print(f"Validation Error (Ignored): {ex}")
            return ws

        sheet_auto = init_detail_sheet("仕訳明細")
        sheet_manual = init_detail_sheet("仕訳明細（手入力）")

        # 4. Save Data
        rows = []
        for item in data:
            rows.append([
                str(item.get('date', '')),
                str(item.get('debit_account', '')),
                str(item.get('credit_account', '')),
                item.get('amount', 0),
                str(item.get('counterparty', '')),
                str(item.get('memo', '')),
                item.get('tax_amount', 0),
                item.get('evidence_url', '') # New column
            ])
        if rows:
            sheet_auto.append_rows(rows, value_input_option='USER_ENTERED')

        # 5. Advanced PL & BS Formulas
        # P/L: Exclude Assets/Liabilities known in Master
        # Actually easier to just exclude standard list literals in Query for robustness without script complications
        
        # P/L Sheet
        try:
            sheet_pl = sh.worksheet("損益計算書")
        except:
            sheet_pl = sh.add_worksheet(title="損益計算書", rows="100", cols="10")
            
        sheet_pl.clear()
        sheet_pl.update("A1", [["損益計算書 (P/L)"]])
        sheet_pl.update("A3", [["【経費 (借方)】", "金額", "", "【売上 (貸方)】", "金額"]])
        
        # Exclude B/S items from P/L
        # Regex for matches must be 'A|B|C'
        bs_items_list = "'現金|小口現金|普通預金|売掛金|未収入金|棚卸資産|買掛金|未払金|借入金|預り金|資本金|元入金|事業主貸|事業主借'"
        
        # Debit (Expenses)
        # Query: Select Col2, Sum(Col4) Where Not Col2 matches BS_Items
        # Range: {'仕訳明細'!A2:G; '仕訳明細（手入力）'!A2:G}
        f_debit = f"=QUERY({{'仕訳明細'!A2:G; '仕訳明細（手入力）'!A2:G}}, \"select Col2, sum(Col4) where Col2 is not null and not Col2 matches {bs_items_list} group by Col2 label sum(Col4) ''\", 0)"
        sheet_pl.update_acell("A4", f_debit)
        
        # Credit (Revenue)
        f_credit = f"=QUERY({{'仕訳明細'!A2:G; '仕訳明細（手入力）'!A2:G}}, \"select Col3, sum(Col4) where Col3 is not null and not Col3 matches {bs_items_list} group by Col3 label sum(Col4) ''\", 0)"
        sheet_pl.update_acell("D4", f_credit)

        # B/S Sheet
        try:
            sheet_bs = sh.worksheet("貸借対照表")
        except:
            sheet_bs = sh.add_worksheet(title="貸借対照表", rows="100", cols="6")
        
        sheet_bs.clear()
        sheet_bs.update("A1", [["貸借対照表 (B/S) - 資産・負債残高"]])
        sheet_bs.update("A2", [["※期首残高 ＋ (借方合計 - 貸方合計) で算出"]])
        sheet_bs.update("A4", [["科目", "現在残高"]])
        
        # B/S Calculation is tricky in pure sheets without a helper table.
        # We will use a robust Query that UNIONs Opening Balance + Debits + Credits(inverted).
        # But for reliability, let's keep it simple: List Opening Balances and append movement.
        # Or better: Create a 'General Ledger' sheet that does the math per account.
        
        # 6. General Ledger (総勘定元帳)
        try:
            ws_gl = sh.worksheet("総勘定元帳")
        except:
            ws_gl = sh.add_worksheet(title="総勘定元帳", rows="1000", cols="8")
            ws_gl.update("A1", [["科目選択:", "現金", "←プルダウンで選択"]])
            # Creating dropdown for cell B1 using Master
            rule_gl = gspread.utils.ValidationCondition("ONE_OF_RANGE", ["=勘定科目マスタ!$A$2:$A$100"])
            ws_gl.set_data_validation("B1", rule_gl)
            
            ws_gl.update("A3", [["日付", "借方", "貸方", "金額", "摘要", "借/貸判定", "残高"]])
            # Formula to filter data for selected account
            # This is a complex formula for the user to see, but very effective.
            # =QUERY({'仕訳明細'!A2:H; '仕訳明細（手入力）'!A2:H}, "select Col1, Col2, Col3, Col4, Col6 where Col2 = '"&B1&"' or Col3 = '"&B1&"' order by Col1", 0)
            ws_gl.update_acell("A4", "=IF(B1=\"\",\"科目をB1で選択してください\", QUERY({'仕訳明細'!A2:H; '仕訳明細（手入力）'!A2:H}, \"select Col1, Col2, Col3, Col4, Col6 where Col2 = '\"&B1&\"' or Col3 = '\"&B1&\"' order by Col1 asc\", 0))")
            
        return True
    except Exception as e:
        print(f"Save Error: {e}")
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
    if 'files' not in request.files: return jsonify({"error": "No files"}), 400
    
    api_key = request.form.get('gemini_api_key')
    spreadsheet_id = request.form.get('spreadsheet_id')
    access_token = request.form.get('access_token')

    if not api_key or not spreadsheet_id or not access_token:
        return jsonify({"error": "Missing config"}), 401
    
    genai.configure(api_key=api_key)
    history = get_accounting_history(spreadsheet_id, access_token)
    existing = get_existing_entries(spreadsheet_id, access_token)
    
    files = request.files.getlist('files')
    results = []
    
    for file in files:
        fname = file.filename.lower()
        ftype = file.content_type
        fbytes = file.read()
        
        # Evidence Upload
        ev_url = upload_to_drive(fbytes, file.filename, ftype, access_token)
        
        if fname.endswith('.csv'):
            res = analyze_csv(fbytes, history)
        else:
            res = analyze_document(fbytes, ftype, history)
            
        for item in res:
            item['evidence_url'] = ev_url
            key = f"{item.get('date')}_{item.get('amount')}_{item.get('counterparty')}"
            item['is_duplicate'] = key in existing
        results.extend(res)
        
    return jsonify(results)

@app.route('/api/save', methods=['POST'])
def api_save():
    d = request.json
    return jsonify({"message": "Saved"}) if save_to_sheets(d.get('data'), d.get('spreadsheet_id'), d.get('access_token')) else (jsonify({"error": "Save failed"}), 500)

if __name__ == '__main__':
    app.run(debug=True, port=5001)


