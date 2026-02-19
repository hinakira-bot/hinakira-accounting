"""
AI analysis service using Gemini for receipt/CSV/document analysis.
Extracted from main.py for modularity.
All heavy imports (google.generativeai, PIL) are lazy-loaded inside functions
to reduce startup memory on Render (512MB Starter plan).
"""
import json
import io
import csv

# genai and PIL are imported lazily inside functions that need them
_genai = None


def _get_genai():
    """Lazy-load google.generativeai."""
    global _genai
    if _genai is None:
        import google.generativeai as genai
        _genai = genai
    return _genai


def configure_gemini(api_key: str):
    """Configure Gemini API with the given key."""
    _get_genai().configure(api_key=api_key)


def get_analysis_prompt(history_str: str, input_text_or_type: str) -> str:
    """Generate the standardized AI analysis prompt."""
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
    5. **消費税区分 (tax_classification)**:
        - 飲食料品（テイクアウト・食料品購入） → 「8%」（軽減税率）
        - それ以外の課税取引 → 「10%」（標準税率）
        - 給与・社会保険料・税金・保険料・地代（土地） → 「非課税」
        - 預金移動・借入返済・事業主貸/借 → 「不課税」
    6. **摘要 (Memo)**:
        - **最重要**: 過去の履歴（下記）に類似の取引先がある場合、その「摘要」を極力踏襲してください。
        - 履歴がなく内容が不明確な場合は、「勘定科目名」（例：消耗品費）または空欄にしてください。
    7. **判断不能な場合**:
        - 勘定科目が全く分からない場合は「**雑費**」として処理してください。

    【過去の学習データ (取引先: 摘要 => 科目)】
    {history_str}

    以下のJSON形式（配列）で出力してください。Markdownは不要です。
    [
      {{
        "date": "YYYY-MM-DD",
        "debit_account": "借方勘定科目(不明なら雑費)",
        "credit_account": "貸方勘定科目",
        "amount": 税込金額(数値),
        "tax_classification": "10%" or "8%" or "非課税" or "不課税",
        "counterparty": "取引先名",
        "memo": "摘要(履歴優先、なければ科目名)"
      }}
    ]
    """


def analyze_csv(csv_bytes: bytes, history: list = None) -> list:
    """Analyze CSV data (credit card / bank statement) using Gemini AI."""
    if history is None:
        history = []
    print("Analyzing CSV...")
    try:
        genai = _get_genai()
        text = csv_bytes.decode('shift_jis', errors='replace')
        if '確定日' not in text and '利用日' not in text and ',' not in text:
            text = csv_bytes.decode('utf-8', errors='replace')

        f = io.StringIO(text)
        reader = csv.reader(f)
        rows = list(reader)
        csv_text = "\n".join([",".join(row) for row in rows[:60]])

        model = genai.GenerativeModel('gemini-2.5-flash')
        history_str = "\n".join([f"- {h['counterparty']}: {h['memo']} => {h['account']}" for h in history[:50]])

        prompt = get_analysis_prompt(history_str, "CSV明細（クレジットカードまたは銀行）") + f"\n\nデータ:\n{csv_text}"

        response = model.generate_content(prompt)
        content = clean_json(response.text)
        return json.loads(content)
    except Exception as e:
        print(f"Error in analyze_csv: {e}")
        return []


def analyze_document(file_bytes: bytes, mime_type: str, history: list = None) -> list:
    """Analyze document (image/PDF) using Gemini AI."""
    if history is None:
        history = []
    print(f"Analyzing {mime_type}...")
    genai = _get_genai()
    from PIL import Image
    models = ['gemini-2.5-flash', 'gemini-2.0-flash']
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
            content = clean_json(response.text)
            return json.loads(content)
        except Exception as e:
            print(f"Error with {model_name}: {e}")
            continue
    return []


def predict_accounts(data: list, history: list, valid_accounts: list, gemini_api_key: str) -> list:
    """Predict debit/credit accounts, tax category, and tax rate for entries using AI."""
    configure_gemini(gemini_api_key)
    genai = _get_genai()

    history_str = "\n".join([f"- {h['counterparty']}: {h['memo']} => {h['account']}" for h in history[:50]])
    valid_accounts_str = ", ".join(valid_accounts)

    input_text = ""
    for item in data:
        line = f"ID:{item['index']} 取引先:{item.get('counterparty', '')} 摘要:{item.get('memo', '')} 金額:{item.get('amount', '')}"
        if item.get('debit'):
            line += f" 【借方(固定):{item['debit']}】"
        if item.get('credit'):
            line += f" 【貸方(固定):{item['credit']}】"
        if item.get('tax_category'):
            line += f" 【課税区分(固定):{item['tax_category']}】"
        if item.get('tax_rate'):
            line += f" 【税率(固定):{item['tax_rate']}】"
        input_text += line + "\n"

    prompt = f"""
    あなたは優秀な日本の公認会計士です。
    以下の「取引先」と「摘要」の情報から、最も適切な仕訳情報を推論してください。

    【使用可能な勘定科目リスト】
    {valid_accounts_str}

    【重要：固定指定の厳守】
    入力データに【借方(固定):...】【貸方(固定):...】【課税区分(固定):...】【税率(固定):...】と書かれている場合は、
    ユーザーが既に決定した確定事項です。**絶対に**その値を変更せず、そのまま出力に含めてください。
    空欄になっている項目のみを推論して埋めてください。

    【推論ルール】
    1. **過去の学習データ**と同じ取引先があれば、その科目を優先してください。
    2. 一般的な会計知識に基づいて推測してください。
    3. クレジットカード払いや後払いの場合は、貸方を「未払金」とします（ただし固定指定がある場合は指定を優先）。

    【課税区分の判定ルール】
    - 通常の経費（消耗品、交通費、外注費等の購入）→「課税仕入」
    - 売上や収入 →「課税売上」
    - 給与・社会保険料・保険料・地代（土地） →「非課税」
    - 預金移動・借入返済・事業主貸/借・租税公課 →「不課税」

    【税率の判定ルール】
    - 飲食料品（テイクアウト・食料品購入）→「8%」（軽減税率）
    - それ以外の課税取引 →「10%」（標準税率）
    - 非課税・不課税 →「0%」

    【過去の学習データ】
    {history_str}

    【推測対象データ】
    {input_text}

    【出力フォーマット (JSON)】
    [
      {{"index": ID(数値), "debit": "借方科目", "credit": "貸方科目", "tax_category": "課税仕入 or 課税売上 or 非課税 or 不課税", "tax_rate": "10% or 8% or 0%"}},
      ...
    ]
    """

    try:
        model = genai.GenerativeModel('gemini-2.5-flash')
        response = model.generate_content(prompt)
        content = clean_json(response.text)
        predictions = json.loads(content)

        # Enforce user overrides
        for pred in predictions:
            p_idx = int(pred.get('index', -1))
            original = next((item for item in data if int(item.get('index', -2)) == p_idx), None)
            if original:
                if original.get('debit') and str(original['debit']).strip():
                    pred['debit'] = original['debit']
                if original.get('credit') and str(original['credit']).strip():
                    pred['credit'] = original['credit']
                if original.get('tax_category') and str(original['tax_category']).strip():
                    pred['tax_category'] = original['tax_category']
                if original.get('tax_rate') and str(original['tax_rate']).strip():
                    pred['tax_rate'] = original['tax_rate']

        return predictions
    except Exception as e:
        print(f"Prediction Error: {e}")
        raise e


def estimate_useful_life(asset_name: str, user_id: int = 0) -> dict:
    """Use Gemini AI to estimate useful life based on Japanese tax law (耐用年数表)."""
    genai = _get_genai()
    prompt = f"""
あなたは日本の税務の専門家です。
以下の固定資産の名称から、国税庁の「減価償却資産の耐用年数等に関する省令」に基づいて法定耐用年数を判定してください。

【資産名】{asset_name}

【主要な法定耐用年数（参考）】
■ 器具備品
- パソコン・電子計算機（サーバー除く）→ 4年
- サーバー用電子計算機 → 5年
- コピー機・複合機・プリンター → 5年
- 電話設備・ファクシミリ → 6年
- エアコン・空調設備 → 6年
- 冷蔵庫・冷凍庫（電気式） → 6年
- テレビ・モニター → 5年
- カメラ（デジカメ含む） → 5年
- 事務机・事務いす（金属製） → 15年
- 事務机・事務いす（金属製以外） → 8年
- 応接セット（接客用） → 5年（金属以外は8年）
- 看板・ネオンサイン → 3年
- 陳列ケース・棚（金属製） → 6年
- 金庫 → 20年
- ベッド（金属製） → 15年

■ 車両運搬具
- 普通自動車（総排気量0.66L超） → 6年
- 軽自動車（総排気量0.66L以下） → 4年
- 二輪自動車（バイク） → 3年
- 自転車 → 2年
- 運送用トラック → 4〜5年

■ 機械装置
- 食料品製造設備 → 10年
- 金属加工機械 → 10年
- 印刷設備 → 10年
- 一般的な製造設備 → 7〜15年

■ 建物
- 木造・合成樹脂造 → 22年（事務所用）/ 20年（店舗用）
- 鉄骨鉄筋コンクリート造（SRC）→ 50年（事務所用）
- 鉄筋コンクリート造（RC） → 47年（事務所用）
- 鉄骨造（骨格材3mm以下） → 22年 / (3mm超4mm以下) → 30年 / (4mm超) → 38年
- 建物附属設備（電気・給排水・ガス） → 15年

■ 無形固定資産
- ソフトウェア（自社利用・研究開発以外） → 5年
- ソフトウェア（市場販売目的） → 3年
- 特許権 → 8年 / 商標権 → 10年

■ 工具
- 測定工具・検査工具 → 5年
- 治具・金型 → 3年

【重要ルール】
1. まずGoogle検索で「{asset_name} 耐用年数」「{asset_name} 減価償却 法定耐用年数」を検索し、国税庁や税理士サイトの情報を確認してください
2. 検索結果と上記の表を照合して、最も正確な耐用年数を判定してください
3. 上記の表に該当するものはその年数を使用
4. 該当しない場合は最も近い資産分類で判定
5. 少額減価償却資産（10万円未満）は一括経費計上が可能だが、耐用年数の判定は行う
6. 中古資産の場合は新品の耐用年数を回答（中古計算はユーザー側で行う）

以下のJSON形式で回答してください。Markdownは不要です。
{{
  "useful_life": 法定耐用年数(数値),
  "asset_category": "大分類（器具備品、車両運搬具、機械装置、建物、無形固定資産等）",
  "detail_category": "細目（電子計算機、事務机いす等の具体名）",
  "reasoning": "判定理由（1行で簡潔に）"
}}
"""
    try:
        model = genai.GenerativeModel('gemini-2.5-flash')
        # Use Google Search grounding for accurate depreciation data
        # Try google_search first (Gemini 2.0+), fall back to google_search_retrieval
        response = None
        for search_tool in ['google_search_retrieval', None]:
            try:
                if search_tool:
                    response = model.generate_content(prompt, tools=search_tool)
                else:
                    response = model.generate_content(prompt)
                break
            except Exception as tool_err:
                print(f"Search grounding attempt ({search_tool}): {tool_err}")
                continue
        if response is None:
            response = model.generate_content(prompt)
        text = response.text.strip()
        # Clean JSON
        if "```json" in text:
            text = text.split("```json")[-1].split("```")[0].strip()
        elif "```" in text:
            text = text.split("```")[1].split("```")[0].strip()
        s = text.find('{')
        e = text.rfind('}')
        if s != -1 and e != -1:
            text = text[s:e + 1]
        result = json.loads(text)
        return result
    except Exception as e:
        print(f"AI useful life estimation error: {e}")
        return {"useful_life": 4, "asset_category": "不明", "reasoning": "AI判定エラー。デフォルト4年を設定。"}


def clean_json(text: str) -> str:
    """Clean AI response to extract valid JSON."""
    text = text.strip()
    if "```json" in text:
        text = text.split("```json")[-1].split("```")[0].strip()
    elif "```" in text:
        text = text.split("```")[-1].split("```")[0].strip()
    s = text.find('[')
    e = text.rfind(']')
    if s != -1 and e != -1:
        text = text[s:e + 1]
    return text
