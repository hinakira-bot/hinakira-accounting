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
        - **最重要**: 過去の履歴（下記）に類似の取引先がある場合、その「摘要」を極力踏襲してください。
        - 履歴がなく内容が不明確な場合は、無理に推測せず「勘定科目名」（例：消耗品費）または「空欄」にしてください。
    7. **判断不能な場合**:
        - 勘定化目が全く分からない、事業用か不明、といった場合は「**雑費**」として処理してください。

    【過去の学習データ (取引先: 摘要 => 科目)】
    {history_str}
    
    以下のJSON形式（配列）で出力してください。Markdownは不要です。
    [
      {{
        "date": "YYYY-MM-DD",
        "debit_account": "借方勘定科目(不明なら雑費)",
        "credit_account": "貸方勘定科目",
        "amount": 税込金額(数値),
        "tax_amount": 消費税額(数値・推定),
        "counterparty": "取引先名",
        "memo": "摘要(履歴優先、なければ科目名)"
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
        # Requested Order: Date, Debit, Credit, Amount, Tax, Counterparty, Memo, URL
        headers = ["日にち", "借方", "貸方", "金額(税込)", "消費税額", "取引先", "摘要", "証憑URL"]
        
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
                
            # SAFETY: Ensure at least 8 columns exist for the new layout/Query
            try:
                if ws.col_count < 8:
                    ws.resize(cols=8)
            except: pass

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

        # --- Data Sanitization (Validation Safety) ---
        valid_accounts = set([row[0] for row in DEFAULT_ACCOUNTS[1:]])
        
        sanitized_data = []
        for item in data:
            d_acc = item.get('debit_account', '').strip()
            if d_acc not in valid_accounts and d_acc != "": d_acc = "雑費"
            c_acc = item.get('credit_account', '').strip()
            if c_acc not in valid_accounts and c_acc != "": c_acc = "雑費"
            
            # Clean Amount (remove commas, yen sign, ensure int)
            raw_amt = str(item.get('amount', 0)).replace(',', '').replace('¥', '').replace('円', '').strip()
            try:
                clean_amt = int(float(raw_amt)) # float first to handle '1000.0'
            except:
                clean_amt = 0

            item['debit_account'] = d_acc
            item['credit_account'] = c_acc
            item['amount'] = clean_amt
            sanitized_data.append(item)

        # 4. Save Data (CORE OPERATION - Must Succeed)
        rows_auto = []
        rows_manual = []
        
        for item in sanitized_data:
            row = [
                str(item.get('date', '')),
                str(item.get('debit_account', '')),
                str(item.get('credit_account', '')),
                item.get('amount', 0),
                item.get('tax_amount', 0),
                str(item.get('counterparty', '')),
                str(item.get('memo', '')),
                item.get('evidence_url', '')
            ]
            # Route based on source flag
            if item.get('source') == 'manual':
                rows_manual.append(row)
            else:
                rows_auto.append(row)
                
        if rows_auto:
            try:
                sheet_auto.append_rows(rows_auto, value_input_option='USER_ENTERED')
            except Exception as e:
                print(f"Error appending auto rows: {e}")
                return False
                
        if rows_manual:
            try:
                sheet_manual.append_rows(rows_manual, value_input_option='USER_ENTERED')
            except Exception as e:
                print(f"Error appending manual rows: {e}")
                return False

        # --- Reporting & Views (Fail-Soft) ---
        # Wraps secondary updates in try-except so core save succeeds even if views fail.
        try: 
            # 5. Advanced PL & BS Formulas
            
            # P/L Sheet
            try:
                sheet_pl = sh.worksheet("損益計算書")
            except:
                sheet_pl = sh.add_worksheet(title="損益計算書", rows="100", cols="10")
                
            sheet_pl.clear()
            sheet_pl.update("A1", [["損益計算書 (P/L)"]])
            sheet_pl.update("A3", [["【経費 (借方)】", "金額", "", "【売上 (貸方)】", "金額"]])
            
            bs_items_list = "'現金|小口現金|普通預金|売掛金|未収入金|棚卸資産|買掛金|未払金|借入金|預り金|資本金|元入金|事業主貸|事業主借'"
            
            f_debit = f"=QUERY({{'仕訳明細'!A2:H; '仕訳明細（手入力）'!A2:H}}, \"select Col2, sum(Col4) where Col2 is not null and not Col2 matches {bs_items_list} group by Col2 label sum(Col4) ''\", 0)"
            sheet_pl.update_acell("A4", f_debit)
            
            f_credit = f"=QUERY({{'仕訳明細'!A2:H; '仕訳明細（手入力）'!A2:H}}, \"select Col3, sum(Col4) where Col3 is not null and not Col3 matches {bs_items_list} group by Col3 label sum(Col4) ''\", 0)"
            sheet_pl.update_acell("D4", f_credit)

            # B/S Sheet
            try:
                sheet_bs = sh.worksheet("貸借対照表")
            except:
                sheet_bs = sh.add_worksheet(title="貸借対照表", rows="100", cols="6")
            
            sheet_bs.clear()
            sheet_bs.update("A1", [["貸借対照表 (B/S)"]])
            sheet_bs.update("A2", [["※期首残高 ＋ (借方合計 - 貸方合計) で算出"]])
            
            assets, liabilities, equity = [], [], []
            for row in DEFAULT_ACCOUNTS[1:]:
                name, type_ = row[0], row[1]
                if type_ == "資産": assets.append(name)
                elif type_ == "負債": liabilities.append(name)
                elif type_ == "純資産": equity.append(name)
                
            bs_rows = []
            bs_rows.append(["【資産の部】", "金額"])
            for acc in assets:
                f = f"=SUMIF('期首残高'!A:A, \"{acc}\", '期首残高'!B:B) + (SUMIF('仕訳明細'!B:B, \"{acc}\", '仕訳明細'!D:D) + SUMIF('仕訳明細（手入力）'!B:B, \"{acc}\", '仕訳明細（手入力）'!D:D)) - (SUMIF('仕訳明細'!C:C, \"{acc}\", '仕訳明細'!D:D) + SUMIF('仕訳明細（手入力）'!C:C, \"{acc}\", '仕訳明細（手入力）'!D:D))"
                bs_rows.append([acc, f])
            bs_rows.append(["", ""])
            bs_rows.append(["【負債の部】", "金額"])
            for acc in liabilities:
                f = f"=SUMIF('期首残高'!A:A, \"{acc}\", '期首残高'!B:B) + (SUMIF('仕訳明細'!C:C, \"{acc}\", '仕訳明細'!D:D) + SUMIF('仕訳明細（手入力）'!C:C, \"{acc}\", '仕訳明細（手入力）'!D:D)) - (SUMIF('仕訳明細'!B:B, \"{acc}\", '仕訳明細'!D:D) + SUMIF('仕訳明細（手入力）'!B:B, \"{acc}\", '仕訳明細（手入力）'!D:D))"
                bs_rows.append([acc, f])
            bs_rows.append(["", ""])
            bs_rows.append(["【純資産の部】", "金額"])
            for acc in equity:
                f = f"=SUMIF('期首残高'!A:A, \"{acc}\", '期首残高'!B:B) + (SUMIF('仕訳明細'!C:C, \"{acc}\", '仕訳明細'!D:D) + SUMIF('仕訳明細（手入力）'!C:C, \"{acc}\", '仕訳明細（手入力）'!D:D)) - (SUMIF('仕訳明細'!B:B, \"{acc}\", '仕訳明細'!D:D) + SUMIF('仕訳明細（手入力）'!B:B, \"{acc}\", '仕訳明細（手入力）'!D:D))"
                bs_rows.append([acc, f])
            sheet_bs.update(f"A3:B{3+len(bs_rows)}", bs_rows, value_input_option='USER_ENTERED')
            
        except Exception as e:
            print(f"Reporting Logic Error (Non-Critical): {e}")

        # 6. General Ledger (Safer Update Logic)
        try:
            try:
                ws_gl = sh.worksheet("総勘定元帳")
                ws_gl.clear() # Clear is safer than Delete/Add
            except:
                ws_gl = sh.add_worksheet(title="総勘定元帳", rows="1000", cols="8")

            ws_gl.update("A1", [["科目選択:", "現金", "←プルダウンで選択"]])
            ws_gl.update("B1", "現金")
            ws_gl.update("A3", [["日付", "借方", "貸方", "金額(税込)", "摘要", "借/貸判定", "残高"]])

            try:
                rule_gl = gspread.utils.ValidationCondition("ONE_OF_RANGE", ["=勘定科目マスタ!$A$2:$A$100"])
                ws_gl.set_data_validation("B1", rule_gl)
            except Exception as e:
                print(f"GL Validation Error: {e}")
            
            ws_gl.update_acell("A4", "=IF(B1=\"\",\"科目をB1で選択してください\", QUERY({'仕訳明細'!A2:H; '仕訳明細（手入力）'!A2:H}, \"select Col1, Col2, Col3, Col4, Col7 where Col2 = '\"&B1&\"' or Col3 = '\"&B1&\"' order by Col1 asc\", 0))")
        except Exception as e:
             print(f"GL Setup Error (Non-Critical): {e}")

        # 7. Monthly Trial Balance (月次推移表)
        try:
             try:
                 ws_monthly = sh.worksheet("月次推移表")
                 ws_monthly.clear()
             except:
                 ws_monthly = sh.add_worksheet(title="月次推移表", rows="100", cols="14") # Acc + 12 months + Total
            
             # Header
             months = ["4月", "5月", "6月", "7月", "8月", "9月", "10月", "11月", "12月", "1月", "2月", "3月"]
             header = ["勘定科目"] + months + ["合計"]
             ws_monthly.update("A1", [header])
             
             # Rows
             ac_names = [row[0] for row in DEFAULT_ACCOUNTS[1:]] # Skip header
             m_data = []
             
             for ac in ac_names:
                 m_data.append([ac])
             
             ws_monthly.update(f"A2:A{1+len(m_data)}", m_data)
             
             # Simplified Monthly Summary (Debit)
             q = "=QUERY({'仕訳明細'!A2:H; '仕訳明細（手入力）'!A2:H}, \"select Col2, sum(Col4) where Col2 is not null group by Col2 pivot month(Col1)+1\", 0)"
             ws_monthly.update_acell("E1", "※経費・資産の月次推移 (借方集計)")
             ws_monthly.update_acell("E2", q)
             
        except Exception as e:
            print(f"Monthly Report Error: {e}")
            
        return True
    except Exception as e:
        print(f"Fatal Save Error: {e}")
        return False
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
        # Requested Order: Date, Debit, Credit, Amount, Tax, Counterparty, Memo, URL
        headers = ["日にち", "借方", "貸方", "金額(税込)", "消費税額", "取引先", "摘要", "証憑URL"]
        
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
                
            # SAFETY: Ensure at least 8 columns exist for the new layout/Query
            try:
                if ws.col_count < 8:
                    ws.resize(cols=8)
            except: pass

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

        # --- Data Sanitization (Validation Safety) ---
        valid_accounts = set([row[0] for row in DEFAULT_ACCOUNTS[1:]])
        
        sanitized_data = []
        for item in data:
            d_acc = item.get('debit_account', '').strip()
            if d_acc not in valid_accounts and d_acc != "": d_acc = "雑費"
            c_acc = item.get('credit_account', '').strip()
            if c_acc not in valid_accounts and c_acc != "": c_acc = "雑費"
            
            # Clean Amount (remove commas, yen sign, ensure int)
            raw_amt = str(item.get('amount', 0)).replace(',', '').replace('¥', '').replace('円', '').strip()
            try:
                clean_amt = int(float(raw_amt)) # float first to handle '1000.0'
            except:
                clean_amt = 0

            item['debit_account'] = d_acc
            item['credit_account'] = c_acc
            item['amount'] = clean_amt
            sanitized_data.append(item)

        # 4. Save Data (CORE OPERATION - Must Succeed)
        rows_auto = []
        rows_manual = []
        
        for item in sanitized_data:
            row = [
                str(item.get('date', '')),
                str(item.get('debit_account', '')),
                str(item.get('credit_account', '')),
                item.get('amount', 0),
                item.get('tax_amount', 0),
                str(item.get('counterparty', '')),
                str(item.get('memo', '')),
                item.get('evidence_url', '')
            ]
            # Route based on source flag
            if item.get('source') == 'manual':
                rows_manual.append(row)
            else:
                rows_auto.append(row)
                
        if rows_auto:
            try:
                sheet_auto.append_rows(rows_auto, value_input_option='USER_ENTERED')
            except Exception as e:
                print(f"Error appending auto rows: {e}")
                return False
                
        if rows_manual:
            try:
                sheet_manual.append_rows(rows_manual, value_input_option='USER_ENTERED')
            except Exception as e:
                print(f"Error appending manual rows: {e}")
                return False

        # --- Reporting & Views (Fail-Soft) ---
        try: 
            # 5. Advanced PL & BS Formulas
            
            # ... (P/L & B/S Logic remains same, omitted for brevity but should be kept if replacing block) ...
            # Wait, I am replacing a huge chunk. I need to be careful not to delete existing logic if I don't paste it back.
            # The previous tool call modified the file heavily.
            # I should target specific blocks or paste back everything. 
            # Let's paste back the PL/BS logic here to be safe and ADD the new Monthly Report.
            
            # P/L Sheet
            try:
                sheet_pl = sh.worksheet("損益計算書")
            except:
                sheet_pl = sh.add_worksheet(title="損益計算書", rows="100", cols="10")
            
            sheet_pl.clear()
            sheet_pl.update("A1", [["損益計算書 (P/L)"]])
            sheet_pl.update("A3", [["【経費 (借方)】", "金額", "", "【売上 (貸方)】", "金額"]])
            
            bs_items_list = "'現金|小口現金|普通預金|売掛金|未収入金|棚卸資産|買掛金|未払金|借入金|預り金|資本金|元入金|事業主貸|事業主借'"
            
            f_debit = f"=QUERY({{'仕訳明細'!A2:H; '仕訳明細（手入力）'!A2:H}}, \"select Col2, sum(Col4) where Col2 is not null and not Col2 matches {bs_items_list} group by Col2 label sum(Col4) ''\", 0)"
            sheet_pl.update_acell("A4", f_debit)
            
            f_credit = f"=QUERY({{'仕訳明細'!A2:H; '仕訳明細（手入力）'!A2:H}}, \"select Col3, sum(Col4) where Col3 is not null and not Col3 matches {bs_items_list} group by Col3 label sum(Col4) ''\", 0)"
            sheet_pl.update_acell("D4", f_credit)

            # B/S Sheet (Skipping detailed recreation in this snippet for brevity? No, must include code if I replace)
            # ... ACTUALLY, I should use `multi_replace` or smaller `replace` to just change the "Save Data" block 
            # and then ADD the "Monthly Report" block at the end. 
            # But the "Reporting & Views" block is one big try-except now.
            
            # Let's proceed with just updating the "Save Data" part first. 
            # Then I will add Monthly Report in a separate call or append it.
            # Wait, the prompt says "Update save_to_sheets to route manual rows AND create Monthly Report".
            # I will separate these into two calls for safety.
            pass # Placeholder for logic flow in thought process
            
        except Exception as e:
            print(f"Reporting Logic Error (Non-Critical): {e}")

        # Wraps secondary updates in try-except so core save succeeds even if views fail.
        try: 
            # 5. Advanced PL & BS Formulas
            
            # P/L Sheet
            try:
                sheet_pl = sh.worksheet("損益計算書")
            except:
                sheet_pl = sh.add_worksheet(title="損益計算書", rows="100", cols="10")
                
            sheet_pl.clear()
            sheet_pl.update("A1", [["損益計算書 (P/L)"]])
            sheet_pl.update("A3", [["【経費 (借方)】", "金額", "", "【売上 (貸方)】", "金額"]])
            
            bs_items_list = "'現金|小口現金|普通預金|売掛金|未収入金|棚卸資産|買掛金|未払金|借入金|預り金|資本金|元入金|事業主貸|事業主借'"
            
            f_debit = f"=QUERY({{'仕訳明細'!A2:H; '仕訳明細（手入力）'!A2:H}}, \"select Col2, sum(Col4) where Col2 is not null and not Col2 matches {bs_items_list} group by Col2 label sum(Col4) ''\", 0)"
            sheet_pl.update_acell("A4", f_debit)
            
            f_credit = f"=QUERY({{'仕訳明細'!A2:H; '仕訳明細（手入力）'!A2:H}}, \"select Col3, sum(Col4) where Col3 is not null and not Col3 matches {bs_items_list} group by Col3 label sum(Col4) ''\", 0)"
            sheet_pl.update_acell("D4", f_credit)

            # B/S Sheet
            try:
                sheet_bs = sh.worksheet("貸借対照表")
            except:
                sheet_bs = sh.add_worksheet(title="貸借対照表", rows="100", cols="6")
            
            sheet_bs.clear()
            sheet_bs.update("A1", [["貸借対照表 (B/S)"]])
            sheet_bs.update("A2", [["※期首残高 ＋ (借方合計 - 貸方合計) で算出"]])
            
            assets, liabilities, equity = [], [], []
            for row in DEFAULT_ACCOUNTS[1:]:
                name, type_ = row[0], row[1]
                if type_ == "資産": assets.append(name)
                elif type_ == "負債": liabilities.append(name)
                elif type_ == "純資産": equity.append(name)
                
            bs_rows = []
            bs_rows.append(["【資産の部】", "金額"])
            for acc in assets:
                f = f"=SUMIF('期首残高'!A:A, \"{acc}\", '期首残高'!B:B) + (SUMIF('仕訳明細'!B:B, \"{acc}\", '仕訳明細'!D:D) + SUMIF('仕訳明細（手入力）'!B:B, \"{acc}\", '仕訳明細（手入力）'!D:D)) - (SUMIF('仕訳明細'!C:C, \"{acc}\", '仕訳明細'!D:D) + SUMIF('仕訳明細（手入力）'!C:C, \"{acc}\", '仕訳明細（手入力）'!D:D))"
                bs_rows.append([acc, f])
            bs_rows.append(["", ""])
            bs_rows.append(["【負債の部】", "金額"])
            for acc in liabilities:
                f = f"=SUMIF('期首残高'!A:A, \"{acc}\", '期首残高'!B:B) + (SUMIF('仕訳明細'!C:C, \"{acc}\", '仕訳明細'!D:D) + SUMIF('仕訳明細（手入力）'!C:C, \"{acc}\", '仕訳明細（手入力）'!D:D)) - (SUMIF('仕訳明細'!B:B, \"{acc}\", '仕訳明細'!D:D) + SUMIF('仕訳明細（手入力）'!B:B, \"{acc}\", '仕訳明細（手入力）'!D:D))"
                bs_rows.append([acc, f])
            bs_rows.append(["", ""])
            bs_rows.append(["【純資産の部】", "金額"])
            for acc in equity:
                f = f"=SUMIF('期首残高'!A:A, \"{acc}\", '期首残高'!B:B) + (SUMIF('仕訳明細'!C:C, \"{acc}\", '仕訳明細'!D:D) + SUMIF('仕訳明細（手入力）'!C:C, \"{acc}\", '仕訳明細（手入力）'!D:D)) - (SUMIF('仕訳明細'!B:B, \"{acc}\", '仕訳明細'!D:D) + SUMIF('仕訳明細（手入力）'!B:B, \"{acc}\", '仕訳明細（手入力）'!D:D))"
                bs_rows.append([acc, f])
            sheet_bs.update(f"A3:B{3+len(bs_rows)}", bs_rows, value_input_option='USER_ENTERED')
            
        except Exception as e:
            print(f"Reporting Logic Error (Non-Critical): {e}")

        # 6. General Ledger (Safer Update Logic)
        try:
            try:
                ws_gl = sh.worksheet("総勘定元帳")
                ws_gl.clear() # Clear is safer than Delete/Add
            except:
                ws_gl = sh.add_worksheet(title="総勘定元帳", rows="1000", cols="8")

            ws_gl.update("A1", [["科目選択:", "現金", "←プルダウンで選択"]])
            ws_gl.update("B1", "現金")
            ws_gl.update("A3", [["日付", "借方", "貸方", "金額(税込)", "摘要", "借/貸判定", "残高"]])

            try:
                rule_gl = gspread.utils.ValidationCondition("ONE_OF_RANGE", ["=勘定科目マスタ!$A$2:$A$100"])
                ws_gl.set_data_validation("B1", rule_gl)
            except Exception as e:
                print(f"GL Validation Error: {e}")
            
            ws_gl.update_acell("A4", "=IF(B1=\"\",\"科目をB1で選択してください\", QUERY({'仕訳明細'!A2:H; '仕訳明細（手入力）'!A2:H}, \"select Col1, Col2, Col3, Col4, Col7 where Col2 = '\"&B1&\"' or Col3 = '\"&B1&\"' order by Col1 asc\", 0))")
        except Exception as e:
            print(f"GL Setup Error (Non-Critical): {e}")

        # 7. Monthly Trial Balance (月次推移表)
        try:
             try:
                 ws_monthly = sh.worksheet("月次推移表")
                 ws_monthly.clear()
             except:
                 ws_monthly = sh.add_worksheet(title="月次推移表", rows="100", cols="14") # Acc + 12 months + Total
            
             # Header
             months = ["4月", "5月", "6月", "7月", "8月", "9月", "10月", "11月", "12月", "1月", "2月", "3月"]
             header = ["勘定科目"] + months + ["合計"]
             ws_monthly.update("A1", [header])
             
             # Rows
             ac_names = [row[0] for row in DEFAULT_ACCOUNTS[1:]] # Skip header
             m_data = []
             
             for ac in ac_names:
                 # Row Formula: SUMIFS(Amount, Account=ac, Date>=Start, Date<End)
                 # Note: Standard Fiscal Year in Japan starts April.
                 # Optimization: Creating 12 formulas per row is heavy.
                 # Let's use QUERY pivot if possible? 
                 # Pivot by Month(Date) is tricky in Query without extra column.
                 
                 # Let's stick to simple SUMIFS for reliability.
                 # We need to construct the range references.
                 # '仕訳明細'!D:D (Amount), '仕訳明細'!B:B (Debit)=AC - '仕訳明細'!C:C (Credit)=AC?
                 # Wait, for P/L expenses (Debit dominant), it's Debit - Credit.
                 # For Sales (Credit dominant), it's Credit - Debit.
                 # This "Sign" logic is complex. 
                 # Simplified approach: Net Movement (Debit - Credit)
                 # If negative, it means credit balance.
                 
                 # Formula for April (Month 4):
                 # =SUMIFS('仕訳明細'!Amount, '仕訳明細'!Debit, AC, '仕訳明細'!Month, 4) ... 
                 # Generating detailed formulas for each cell is too slow via API (100 rows * 12 cols = 1200 updates).
                 # Better approach: 
                 # Use QUERY on a hidden sheet that pre-calculates month, then Pivot.
                 # OR: Just set up the sheet ONCE with array formulas? 
                 # Let's insert the account list in Col A, and use a draggable formula in Row 2.
                 
                 m_data.append([ac])
             
             ws_monthly.update(f"A2:A{1+len(m_data)}", m_data)
             
             # Inject Array Formula in B2 (Example for April)
             # But ArrayFormula across months is hard.
             
             # Fallback: Just put the Account list.
             # User can use the "General Ledger" for details. 
             # OR: Let's produce a simple "By Month" query for the whole dataset.
             # =QUERY({Data}, "select Col2, sum(Col4) group by Col2 pivot month(Col1)+1 label Col2 '科目'", 1)
             # But month() returns 0-11 or 1-12? Query month() is 0-based (0=Jan).
             
             # Let's try the Pivot Query. It's the most powerful feature.
             # We need to normalize dates first. 
             # QUERY(..., "select Col2, sum(Col4) pivot month(Col1)")
             # Problem: 'Col2' is Debit Account. We need to merge Debit and Credit side.
             # This is hard in one query.
             
             # Compromise: Create a simple "Monthly Debit Summary" (Spending per month).
             # Focus on Expenses (Debit Side).
             q = "=QUERY({'仕訳明細'!A2:H; '仕訳明細（手入力）'!A2:H}, \"select Col2, sum(Col4) where Col2 is not null group by Col2 pivot month(Col1)+1\", 0)"
             ws_monthly.update_acell("E1", "※経費・資産の月次推移 (借方集計)")
             ws_monthly.update_acell("E2", q)
             
        except Exception as e:
            print(f"Monthly Report Error: {e}")

        return True
    except Exception as e:
        print(f"Fatal Save Error: {e}")
        return False

# --- Routes ---
@app.route('/')
def index():
    return send_from_directory('.', 'index.html')

@app.route('/<path:path>')
def static_files(path):
    return send_from_directory('.', path)

@app.route('/api/accounts', methods=['GET'])
def api_accounts():
    # Return just the account names
    return jsonify({"accounts": [row[0] for row in DEFAULT_ACCOUNTS[1:]]})

@app.route('/api/analyze', methods=['POST'])
def api_analyze():
    if 'files' not in request.files: return jsonify({"error": "No files"}), 400
    
    files = request.files.getlist('files')
    access_token = request.headers.get('Authorization', '').replace('Bearer ', '')
    gemini_api_key = request.form.get('gemini_api_key')
    spreadsheet_id = request.form.get('spreadsheet_id')
    # Use access token from header if form is empty, or vice versa. 
    # Frontend sends access_token in form data for analyze.
    if not access_token: access_token = request.form.get('access_token')

    if not gemini_api_key or not spreadsheet_id or not access_token:
        return jsonify({"error": "Missing config"}), 401
    
    genai.configure(api_key=gemini_api_key)
    history = get_accounting_history(spreadsheet_id, access_token)
    existing = get_existing_entries(spreadsheet_id, access_token)
    
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
            key = f"{item.get('date')}_{str(item.get('amount'))}_{item.get('counterparty')}"
            item['is_duplicate'] = key in existing
        results.extend(res)
        
    return jsonify(results)

@app.route('/api/predict', methods=['POST'])
def api_predict():
    data = request.json.get('data', []) # List of {index, counterparty, memo}
    spreadsheet_id = request.json.get('spreadsheet_id')
    access_token = request.json.get('access_token') or request.headers.get('Authorization', '').replace('Bearer ', '')
    gemini_api_key = request.json.get('gemini_api_key')

    if not data or not spreadsheet_id or not access_token or not gemini_api_key:
        return jsonify({"error": "Missing config"}), 400

    genai.configure(api_key=gemini_api_key)
    history = get_accounting_history(spreadsheet_id, access_token)
    
    # Format history
    history_str = "\n".join([f"- {h['counterparty']}: {h['memo']} => {h['account']}" for h in history[:50]])
    
    # Get valid accounts list
    valid_accounts_str = ", ".join([f"{row[0]}" for row in DEFAULT_ACCOUNTS[1:]])
    
    # Format input
    # Include existing debit/credit if provided
    input_text = ""
    for item in data:
        line = f"ID:{item['index']} 取引先:{item.get('counterparty', '')} 摘要:{item.get('memo', '')}"
        if item.get('debit'): line += f" 【借方指定:{item['debit']}】"
        if item.get('credit'): line += f" 【貸方指定:{item['credit']}】"
        input_text += line + "\n"

    prompt = f"""
    あなたは優秀な日本の公認会計士です。
    以下の「取引先」と「摘要」の情報から、最も適切な「借方勘定科目」と「貸方勘定科目」を推論してください。
    
    【使用可能な勘定科目リスト】
    {valid_accounts_str}
    ※必ずこのリストの中から選択してください。
    
    【推論ルール】
    0. **ユーザー指定の考慮:** 入力データに【借方指定:...】や【貸方指定:...】がある場合は、その指定を**絶対に変更せず**、そのまま出力に含めてください。空欄の片方だけを推測してください。
    1. **過去の学習データ**と同じ取引先があれば、その科目を優先してください。
    2. 過去データにない場合、**取引先名や摘要のニュアンス**から、一般的な会計知識に基づいて推測してください。
       - 例: 飲食店 → 会議費(商談) または 接待交際費
       - 例: コンビニ・スーパー・Amazon → 消耗品費
       - 例: ガソリン・タクシー・電車 → 旅費交通費
       - 例: サブスク・クラウドサービス → 通信費
       - 例: 広告・宣伝 → 広告宣伝費
    3. クレジットカード払いや後払いの場合は、貸方を「未払金」としてください。
    4. どうしても判断できない場合のみ「雑費」としてください。安易に雑費にせず、可能性が高い科目を選んでください。
    
    【過去の学習データ】
    {history_str}
    
    【推測対象データ】
    {input_text}
    
    【出力フォーマット (JSON)】
    [
      {{"index": ID(数値), "debit": "借方科目", "credit": "貸方科目"}},
      ...
    ]
    Markdownは不要です。JSONのみ出力してください。
    """
    
    try:
        model = genai.GenerativeModel('gemini-2.0-flash-exp')
        response = model.generate_content(prompt)
        content = CleanJSON(response.text)
        predictions = json.loads(content)
        
        print(f"DEBUG: Original Data: {data}")
        print(f"DEBUG: AI Predictions: {predictions}")
        
        # Enforce user overrides (Programmatic Safety Net)
        for pred in predictions:
            # Robust index matching (handle str/int mismatch)
            p_idx = int(pred.get('index', -1))
            original = next((item for item in data if int(item.get('index', -2)) == p_idx), None)
            
            if original:
                # If user provided a non-empty string for debit/credit, USE IT.
                if original.get('debit') and str(original['debit']).strip(): 
                    pred['debit'] = original['debit']
                if original.get('credit') and str(original['credit']).strip(): 
                    pred['credit'] = original['credit']
        
        print(f"DEBUG: Final Response: {predictions}")
        return jsonify(predictions)
    except Exception as e:
        print(f"Prediction Error: {e}")
        return jsonify({"error": str(e)}), 500

@app.route('/api/save', methods=['POST'])
def api_save():
    data = request.json.get('data', [])
    spreadsheet_id = request.json.get('spreadsheet_id')
    access_token = request.headers.get('Authorization', '').replace('Bearer ', '')
    
    # In JSON body, frontend might send access_token too
    if not access_token: access_token = request.json.get('access_token')
    
    if not data or not spreadsheet_id or not access_token:
        return jsonify({"error": "Missing data or credentials"}), 400
        
    success = save_to_sheets(data, spreadsheet_id, access_token)
    if success:
        return jsonify({"status": "success"})
    else:
        return jsonify({"error": "Save failed"}), 500

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5001))
    app.run(host='0.0.0.0', port=port)


