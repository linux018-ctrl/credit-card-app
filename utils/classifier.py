"""
交易分類引擎
根據消費明細自動判斷：消費類別（四大超商/一般）、Owner（Alan/Lydia）、結算月份
"""

import re
from datetime import datetime, date


# ============================================================
# 四大超商關鍵字（統一超商、全家、萊爾富、OK超商）
# ============================================================
CONVENIENCE_STORE_KEYWORDS = [
    '統一超商',
    '7-ELEVEN', '7-11', '７－ＥＬＥＶＥＮ',
    '全家便利商店', '全家',
    '萊爾富',
    'ＯＫ超商', 'OK超商', 'ＯＫ便利商店', 'OK便利商店', 'ＯＫ　ｍａｒｔ',
]

# ============================================================
# Alan 的機器代碼（出現在消費明細中）
# ============================================================
ALAN_DEVICE_CODE = 'M5030'

# ============================================================
# 特定帳單歸屬規則
# 有些固定費用不含 M5030，但根據帳號可以判斷歸屬
# key: 消費明細中的關鍵字, value: Owner
# ============================================================
FIXED_OWNER_RULES = {
    # Alan 的固定費用（電話號碼 09219 開頭）
    '09219': 'Alan',
    # 水費帳號 5Y287
    '5Y287': 'Alan',
    # 電費帳號 09-35-0056-27-8
    '09-35-0056-27-8': 'Alan',
    # IKEA 家具 (Alan 的消費)
    'ＩＫＥＡ': 'Alan',
    'IKEA': 'Alan',
}

# ============================================================
# 回饋費率設定
# ============================================================
REWARD_RATES = {
    '四大超商': 0.10,   # 10%
    '一般': 0.01,       # 1%
}

# 四大超商回饋上限（每月）
CONVENIENCE_REWARD_CAP = 200


def classify_category(description: str) -> str:
    """
    判斷消費類別

    Args:
        description: 消費明細文字

    Returns:
        '四大超商' 或 '一般'
    """
    desc_upper = description.upper()
    # 全形轉半形做比對
    desc_normalized = _fullwidth_to_halfwidth(desc_upper)

    for keyword in CONVENIENCE_STORE_KEYWORDS:
        kw_normalized = _fullwidth_to_halfwidth(keyword.upper())
        if kw_normalized in desc_normalized or keyword in description:
            return '四大超商'

    return '一般'


def classify_owner(description: str) -> str:
    """
    判斷消費 Owner

    規則：
    1. 消費明細含 M5030 → Alan
    2. 符合 FIXED_OWNER_RULES 的特定帳號 → 對應 Owner
    3. 其他 → Lydia

    Args:
        description: 消費明細文字

    Returns:
        'Alan' 或 'Lydia'
    """
    # Rule 1: M5030 代碼
    if ALAN_DEVICE_CODE in description:
        return 'Alan'

    # Rule 2: 特定帳號/關鍵字
    for keyword, owner in FIXED_OWNER_RULES.items():
        if keyword in description:
            return owner

    # Rule 3: 預設為 Lydia
    return 'Lydia'


def determine_billing_month(txn_date, billing_day: int = 27,
                            posting_date=None) -> tuple[int, int]:
    """
    根據入帳起息日和結帳日，判斷該筆消費歸屬的帳單月份

    規則：每月 27 號結帳，以「入帳起息日」為準
    - 入帳日 day <= 27 → 歸屬入帳日當月
    - 入帳日 day > 27  → 歸屬下個月

    Args:
        txn_date: 消費日期 (date 或 datetime)，作為備用
        billing_day: 結帳日（預設 27）
        posting_date: 入帳起息日（優先使用）

    Returns:
        (year, month) 結算年月
    """
    # 優先使用入帳日，否則用消費日
    ref_date = posting_date if posting_date is not None else txn_date

    if isinstance(ref_date, datetime):
        ref_date = ref_date.date()
    elif isinstance(ref_date, str):
        ref_date = datetime.strptime(str(ref_date)[:10], '%Y-%m-%d').date()

    day = ref_date.day
    month = ref_date.month
    year = ref_date.year

    if day <= billing_day:
        # 在結帳日之前（含）→ 歸屬當月
        return (year, month)
    else:
        # 超過結帳日 → 歸屬下個月
        if month == 12:
            return (year + 1, 1)
        else:
            return (year, month + 1)


def classify_transaction(description: str, txn_date, billing_day: int = 27,
                         posting_date=None) -> dict:
    """
    完整分類單筆交易

    Returns:
        dict with keys: 消費類別, Owner, 結算月份, 結算年份
    """
    category = classify_category(description)
    owner = classify_owner(description)
    bill_year, bill_month = determine_billing_month(
        txn_date, billing_day, posting_date=posting_date
    )

    return {
        '消費類別': category,
        'Owner': owner,
        '結算月份': bill_month,
        '結算年份': bill_year,
    }


def calculate_rewards(monthly_data: dict,
                      reward_rates: dict | None = None,
                      convenience_cap: float | None = None) -> dict:
    """
    計算月度回饋金額

    Args:
        monthly_data: dict with keys '四大超商' and '一般', values are total amounts
        reward_rates: 可選，自訂回饋費率 {'四大超商': 0.10, '一般': 0.01}
        convenience_cap: 可選，四大超商每月回饋上限

    Returns:
        dict with reward amounts per category and total
    """
    rates = reward_rates or REWARD_RATES
    cap = float(convenience_cap) if convenience_cap is not None else float(CONVENIENCE_REWARD_CAP)

    conv_amount = float(monthly_data.get('四大超商', 0))
    general_amount = float(monthly_data.get('一般', 0))

    conv_reward = min(
        conv_amount * float(rates.get('四大超商', 0.10)),
        cap
    )
    general_reward = general_amount * float(rates.get('一般', 0.01))

    return {
        '四大超商_消費': conv_amount,
        '四大超商_回饋': round(conv_reward),
        '一般_消費': general_amount,
        '一般_回饋': round(general_reward),
        '回饋總額': round(conv_reward + general_reward),
    }


def _fullwidth_to_halfwidth(text: str) -> str:
    """全形字元轉半形"""
    result = []
    for char in text:
        code = ord(char)
        if 0xFF01 <= code <= 0xFF5E:
            result.append(chr(code - 0xFEE0))
        elif code == 0x3000:
            result.append(' ')
        else:
            result.append(char)
    return ''.join(result)
