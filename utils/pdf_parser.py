"""
台新銀行信用卡 PDF 帳單解析器
支援：民國年日期、加密 PDF、word 座標解析、多行消費明細

實際 PDF 版面（由 pdfplumber extract_words 得知）：
  x ≈  75 : 消費日     (115/03/12)
  x ≈ 117 : 入帳起息日 (115/03/12)
  x ≈ 160 : 消費明細
  x ≈ 304~347 : 新臺幣金額（正數或負數，含逗號）
  x ≈ 378 : 外幣折算日
  x ≈ 419 : 消費地 (TW)
  x ≈ 458 : 幣別
  x ≈ 495 : 外幣金額

多行交易的 Y 座標排列：
  y=326  [x=160] 全家便利商店－工業三店M5030   ← 描述第 1 行（在日期行上方）
  y=331  [x= 75] 115/03/03 [x=117] 115/03/05 [x=342] 60  ← 日期 + 金額行
  y=336  [x=160] TAIPEI                          ← 描述第 2 行（在日期行下方）
"""

import re
import pdfplumber
import pandas as pd
from io import BytesIO
from datetime import datetime


def _decrypt_pdf(pdf_bytes: bytes, password: str = None) -> BytesIO:
    """用 pikepdf 解密 PDF"""
    import pikepdf
    try:
        src = pikepdf.open(BytesIO(pdf_bytes), password=password or "")
    except pikepdf.PasswordError:
        if password:
            raise ValueError("PDF 密碼錯誤，請確認密碼是否正確")
        raise ValueError("PDF 檔案有密碼保護，請輸入密碼（台新帳單通常為身分證字號）")

    buf = BytesIO()
    src.save(buf)
    src.close()
    buf.seek(0)
    return buf


def parse_taishin_pdf(pdf_bytes: bytes, billing_day: int = 27,
                      password: str = None) -> pd.DataFrame:
    """
    解析台新銀行信用卡 PDF 帳單

    主要策略：用 extract_words() 按 Y 座標分組，
    利用 X 座標判斷欄位，正確處理多行消費明細。
    """
    decrypted_buf = _decrypt_pdf(pdf_bytes, password)
    transactions = []

    try:
        with pdfplumber.open(decrypted_buf) as pdf:
            # 主方法：word 座標解析（最可靠）
            word_txns = _parse_from_words(pdf)

            # 備用方法：純文字解析
            all_text = ""
            for page in pdf.pages:
                text = page.extract_text()
                if text:
                    all_text += text + "\n"
            text_txns = _parse_roc_text(all_text) if all_text.strip() else []

            # 取結果較多的方法
            transactions = word_txns if len(word_txns) >= len(text_txns) else text_txns
    except Exception as e:
        raise ValueError(f"PDF 解析失敗: {e}") from e

    if not transactions:
        return pd.DataFrame(columns=[
            '消費日期', '入帳起息日', '消費明細', '清算消費金額'
        ])

    df = pd.DataFrame(transactions)
    df['清算消費金額'] = pd.to_numeric(df['清算消費金額'], errors='coerce')
    df = df.dropna(subset=['清算消費金額'])
    df = df[df['清算消費金額'] > 0]  # 排除退款/扣繳
    df = df.sort_values('消費日期', ascending=False).reset_index(drop=True)
    return df


# ─────────────────────────────────────────
# 民國年日期工具
# ─────────────────────────────────────────
ROC_DATE_RE = re.compile(r'^(\d{2,3})/(\d{2})/(\d{2})$')


def _roc_to_date(year_str: str, month_str: str, day_str: str):
    """民國年轉西元 date"""
    year = int(year_str) + 1911
    return datetime(year, int(month_str), int(day_str)).date()


def _parse_roc_date_str(s: str):
    """解析 '115/03/12' 格式字串"""
    if not s:
        return None
    m = ROC_DATE_RE.match(s.strip())
    if m:
        return _roc_to_date(m.group(1), m.group(2), m.group(3))
    return None


def _should_skip(description: str) -> bool:
    """判斷是否為非交易行（標題、卡號、繳款等）"""
    skip_keywords = [
        '帳單', '繳款', '循環', '最低', '小計', '合計',
        '本期', '上期', '信用額度', '可用餘額', '頁',
        '自動轉帳扣繳', '卡號末四碼', 'Rewards', '御璽卡',
        '商務御璽', '消費明細', '新臺幣金額', '外幣折算',
        '消費地', '幣別', '外幣金額', '消費日', '入帳起息日',
    ]
    return any(kw in description for kw in skip_keywords)


# ─────────────────────────────────────────
# 主方法: extract_words() 座標解析
# ─────────────────────────────────────────
# 欄位 X 座標邊界（根據實際 PDF 分析）
COL_DATE1_MAX = 110       # 消費日 x < 110
COL_DATE2_MIN = 110       # 入帳起息日 110 <= x < 155
COL_DATE2_MAX = 155
COL_DESC_MIN = 155        # 消費明細 155 <= x < 295
COL_DESC_MAX = 295
COL_AMOUNT_MIN = 295      # 金額 x >= 295 (到 ~355)
COL_AMOUNT_MAX = 375
COL_LOCATION_MIN = 375    # 消費地 x >= 375

# 民國年日期 pattern
_ROC_WORD_RE = re.compile(r'^\d{2,3}/\d{2}/\d{2}$')


def _parse_from_words(pdf) -> list:
    """
    用 extract_words() 按 Y 座標分組，再按 X 座標分欄位。
    
    核心邏輯：
    1. 把所有 word 按 Y 座標分成「行」
    2. 找出含有民國年日期的「日期行」
    3. 日期行上方/下方沒有日期的行 → 多行消費明細的延續
    4. 組合成完整交易
    """
    all_rows = []  # [(y, [words...])]

    for page in pdf.pages:
        words = page.extract_words(
            x_tolerance=3, y_tolerance=3, keep_blank_chars=True
        )
        if not words:
            continue

        # 按 Y 座標分組（容差 3px 視為同一行）
        rows_by_y = {}
        for w in words:
            y_key = round(w['top'], 0)
            # 合併接近的 Y 值（±3px）
            merged = False
            for existing_y in list(rows_by_y.keys()):
                if abs(y_key - existing_y) <= 3:
                    rows_by_y[existing_y].append(w)
                    merged = True
                    break
            if not merged:
                rows_by_y[y_key] = [w]

        for y_key in sorted(rows_by_y.keys()):
            row_words = sorted(rows_by_y[y_key], key=lambda w: w['x0'])
            all_rows.append((y_key, row_words))

    return _assemble_transactions(all_rows)


def _classify_row(row_words: list) -> dict:
    """
    分析一行 words，分類各欄位內容。
    返回: {
        'has_date1': bool, 'date1': str,
        'has_date2': bool, 'date2': str,
        'desc_parts': [str], 'amount': str,
        'location': str
    }
    """
    result = {
        'has_date1': False, 'date1': '',
        'has_date2': False, 'date2': '',
        'desc_parts': [], 'amount': '',
        'location': '',
    }

    for w in row_words:
        x = w['x0']
        text = w['text'].strip()
        if not text:
            continue

        if x < COL_DATE1_MAX:
            # 消費日欄位
            if _ROC_WORD_RE.match(text):
                result['has_date1'] = True
                result['date1'] = text
        elif COL_DATE2_MIN <= x < COL_DATE2_MAX:
            # 入帳起息日欄位
            if _ROC_WORD_RE.match(text):
                result['has_date2'] = True
                result['date2'] = text
        elif COL_DESC_MIN <= x < COL_DESC_MAX:
            # 消費明細欄位
            result['desc_parts'].append(text)
        elif COL_AMOUNT_MIN <= x < COL_AMOUNT_MAX:
            # 金額欄位
            result['amount'] = text
        elif x >= COL_LOCATION_MIN:
            # 消費地等欄位
            if text not in ('', 'TW', 'US', 'JP', 'HK', 'SG'):
                pass  # 忽略
            result['location'] = text

    return result


# Y 座標距離閾值：同一筆交易的行間距 ~5-6pt，不同交易間距 ~12-13pt
_Y_PROXIMITY = 8


def _assemble_transactions(all_rows: list) -> list:
    """
    從分類後的行組裝交易。
    
    用 Y 座標距離判斷：日期行上下 _Y_PROXIMITY (8pt) 以內的
    非日期行屬於同一筆交易的多行描述，超過則屬於不同交易。
    """
    transactions = []

    # 先分類每一行
    classified = []
    for y_key, row_words in all_rows:
        info = _classify_row(row_words)
        info['y'] = y_key
        classified.append(info)

    # 找所有日期行的索引
    date_indices = []
    for i, row in enumerate(classified):
        if row['has_date1'] and row['has_date2']:
            date_indices.append(i)

    for idx in date_indices:
        row = classified[idx]
        txn_date = _parse_roc_date_str(row['date1'])
        posting_date = _parse_roc_date_str(row['date2'])
        if not txn_date or not posting_date:
            continue

        date_y = row['y']
        desc_parts = list(row['desc_parts'])

        # ── 向上收集：Y 距離 ≤ _Y_PROXIMITY 的描述行 ──
        upper_parts = []
        j = idx - 1
        while j >= 0:
            prev = classified[j]
            if prev['has_date1'] and prev['has_date2']:
                break
            if date_y - prev['y'] > _Y_PROXIMITY:
                break  # 距離太遠，屬於不同交易
            if prev['desc_parts']:
                candidate = ' '.join(prev['desc_parts'])
                if _should_skip(candidate):
                    break
                upper_parts.append(candidate)
            j -= 1
        upper_parts.reverse()

        # ── 向下收集：Y 距離 ≤ _Y_PROXIMITY 的描述行 ──
        lower_parts = []
        k = idx + 1
        while k < len(classified):
            nxt = classified[k]
            if nxt['has_date1'] and nxt['has_date2']:
                break
            if nxt['y'] - date_y > _Y_PROXIMITY:
                break  # 距離太遠，屬於不同交易
            if nxt['desc_parts']:
                lower_parts.append(' '.join(nxt['desc_parts']))
            k += 1

        # 組合描述
        all_desc = upper_parts + desc_parts + lower_parts
        description = ''.join(all_desc).strip()
        description = _normalize_fullwidth(description)

        if not description or _should_skip(description):
            continue

        # 金額
        amount_str = row['amount']
        if not amount_str:
            for check_idx in range(max(0, idx - 2), min(len(classified), idx + 3)):
                if check_idx != idx and classified[check_idx]['amount']:
                    if not classified[check_idx]['has_date1']:
                        amount_str = classified[check_idx]['amount']
                        break

        if not amount_str:
            continue

        amount_val = _parse_amount(amount_str)
        if amount_val is None:
            continue

        # 排除負數金額（退款、自動扣繳等）
        if amount_val <= 0:
            continue

        transactions.append({
            '消費日期': txn_date,
            '入帳起息日': posting_date,
            '消費明細': description,
            '清算消費金額': amount_val,
        })

    return transactions


def _normalize_fullwidth(s: str) -> str:
    """將全形英數轉為半形"""
    result = []
    for ch in s:
        code = ord(ch)
        # 全形英數 (Ａ-Ｚ, ａ-ｚ, ０-９)
        if 0xFF01 <= code <= 0xFF5E:
            result.append(chr(code - 0xFEE0))
        # 全形空格
        elif code == 0x3000:
            result.append(' ')
        else:
            result.append(ch)
    return ''.join(result)


def _parse_amount(s: str):
    """解析金額字串，回傳 float 或 None"""
    cleaned = s.replace(',', '').replace(' ', '').strip()
    if not cleaned:
        return None
    try:
        return float(cleaned)
    except ValueError:
        return None


# ─────────────────────────────────────────
# 備用方法: 文字解析 fallback
# ─────────────────────────────────────────
def _parse_roc_text(text: str) -> list:
    """
    從 extract_text() 的純文字中解析民國年交易。

    台新 PDF 的 extract_text() 多行交易格式：
    - 描述行（無日期）
    - 日期行：115/03/03 115/03/05 ... 金額 TW
    - 描述行（TAIPEI 等）
    
    也可能單行完整：
    - 115/02/24 115/03/02 統一超商－勝福M5030 TAIPEI 165 TW
    """
    transactions = []
    lines = text.split('\n')

    roc_date_p = re.compile(r'(\d{2,3}/\d{2}/\d{2})')
    amount_re = re.compile(r'-?[\d,]+')

    # 找所有含 ≥2 個日期的行（=日期行）
    date_line_indices = []
    for i, line in enumerate(lines):
        dates = roc_date_p.findall(line)
        if len(dates) >= 2:
            d1 = _parse_roc_date_str(dates[0])
            d2 = _parse_roc_date_str(dates[1])
            if d1 and d2:
                date_line_indices.append(i)

    for di, line_idx in enumerate(date_line_indices):
        line = lines[line_idx].strip()
        dates = roc_date_p.findall(line)
        txn_date = _parse_roc_date_str(dates[0])
        posting_date = _parse_roc_date_str(dates[1])

        # 日期行中去除日期後的剩餘文字
        rest = line
        for d in dates[:2]:
            rest = rest.replace(d, '', 1)
        rest = rest.strip()

        # 向上收集描述行（上一個日期行之後到此行之前）
        prev_boundary = date_line_indices[di - 1] + 1 if di > 0 else 0
        upper_desc = []
        for j in range(line_idx - 1, prev_boundary - 1, -1):
            prev_line = lines[j].strip()
            if not prev_line:
                break
            if roc_date_p.search(prev_line):
                break
            if _should_skip(prev_line):
                break  # 卡號標題等不應收集
            upper_desc.insert(0, prev_line)

        # 向下收集描述行（此行之後到下一個日期行之前）
        next_boundary = date_line_indices[di + 1] if di + 1 < len(date_line_indices) else len(lines)
        lower_desc = []
        for j in range(line_idx + 1, next_boundary):
            next_line = lines[j].strip()
            if not next_line:
                break
            if roc_date_p.search(next_line):
                break
            lower_desc.append(next_line)

        # 組合所有文字
        combined = ' '.join(upper_desc) + ' ' + rest + ' ' + ' '.join(lower_desc)
        combined = combined.strip()
        combined = _normalize_fullwidth(combined)

        # 提取金額：TW 前面的數字
        # 先嘗試 "數字 TW" 模式
        tw_match = re.search(r'(-?[\d,]+)\s+TW', combined)
        if tw_match:
            amount_str = tw_match.group(1).replace(',', '')
            desc = combined[:tw_match.start()].strip()
        else:
            # fallback: 取最後一組數字
            nums = list(amount_re.finditer(combined))
            if not nums:
                continue
            best = nums[-1]
            amount_str = best.group().replace(',', '')
            desc = combined[:best.start()].strip()

        desc = re.sub(r'\s+', ' ', desc).strip()

        if not desc or _should_skip(desc):
            continue

        try:
            amount = float(amount_str)
        except ValueError:
            continue

        transactions.append({
            '消費日期': txn_date,
            '入帳起息日': posting_date,
            '消費明細': desc,
            '清算消費金額': amount,
        })

    return transactions


# ─────────────────────────────────────────
# 帳單期間提取
# ─────────────────────────────────────────
def extract_billing_period(pdf_bytes: bytes, password: str = None) -> dict | None:
    """從 PDF 提取帳單期間"""
    try:
        decrypted_buf = _decrypt_pdf(pdf_bytes, password)
        with pdfplumber.open(decrypted_buf) as pdf:
            if pdf.pages:
                text = pdf.pages[0].extract_text() or ""
            else:
                return None
    except Exception:
        return None

    period_patterns = [
        # 民國年: 115/02/28～115/03/27
        re.compile(r'(\d{2,3})/(\d{2})/(\d{2})\s*[~～至]\s*(\d{2,3})/(\d{2})/(\d{2})'),
        # 西元年
        re.compile(r'(\d{4})/(\d{2})/(\d{2})\s*[~～至]\s*(\d{4})/(\d{2})/(\d{2})'),
    ]

    for p in period_patterns:
        m = p.search(text)
        if m:
            groups = m.groups()
            end_year = int(groups[3])
            if end_year < 200:
                end_year += 1911
            return {'year': end_year, 'month': int(groups[4])}

    month_pattern = re.compile(r'(\d{4})\s*年?\s*(\d{1,2})\s*月')
    m = month_pattern.search(text)
    if m:
        return {'year': int(m.group(1)), 'month': int(m.group(2))}

    return None
