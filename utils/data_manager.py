"""
資料管理模組
負責：交易紀錄的讀取、儲存、匯入、合併去重
"""

import json
import os
import pandas as pd
from datetime import datetime, date
from pathlib import Path

DATA_DIR = Path(__file__).parent.parent / 'data'
RECORDS_FILE = DATA_DIR / 'records.json'


def _json_serial(obj):
    """JSON 序列化輔助 - 處理 date/datetime"""
    if isinstance(obj, (datetime, date)):
        return obj.isoformat()
    raise TypeError(f"Type {type(obj)} not serializable")


def load_records() -> pd.DataFrame:
    """載入所有已儲存的交易紀錄"""
    if not RECORDS_FILE.exists():
        return _empty_df()

    try:
        with open(RECORDS_FILE, 'r', encoding='utf-8') as f:
            data = json.load(f)
        if not data:
            return _empty_df()
        df = pd.DataFrame(data)
        # 確保日期欄位格式
        for col in ['消費日期', '入帳起息日']:
            if col in df.columns:
                df[col] = pd.to_datetime(df[col]).dt.date
        return df
    except (json.JSONDecodeError, KeyError):
        return _empty_df()


def save_records(df: pd.DataFrame):
    """儲存交易紀錄到 JSON"""
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    records = df.to_dict('records')
    with open(RECORDS_FILE, 'w', encoding='utf-8') as f:
        json.dump(records, f, ensure_ascii=False, indent=2, default=_json_serial)


def merge_records(existing: pd.DataFrame, new_records: pd.DataFrame) -> pd.DataFrame:
    """
    合併新舊紀錄並去重

    去重邏輯：消費日期 + 消費明細 + 清算消費金額 三者完全相同視為重複
    """
    if existing.empty:
        return new_records
    if new_records.empty:
        return existing

    combined = pd.concat([existing, new_records], ignore_index=True)

    # 去重：基於 消費日期 + 消費明細 + 金額
    dedup_cols = ['消費日期', '消費明細', '清算消費金額']
    available_cols = [c for c in dedup_cols if c in combined.columns]

    if available_cols:
        combined = combined.drop_duplicates(subset=available_cols, keep='last')

    combined = combined.sort_values('消費日期', ascending=False).reset_index(drop=True)
    return combined


def import_from_excel(excel_path: str, sheet_name: str = None) -> pd.DataFrame:
    """
    從現有的信用卡明細 Excel 匯入資料

    Args:
        excel_path: Excel 檔案路徑
        sheet_name: 工作表名稱（None 則自動偵測）

    Returns:
        DataFrame with 標準化欄位
    """
    if sheet_name is None:
        # 嘗試讀取所有 sheet 名稱，找包含 "消費紀錄" 的
        import openpyxl
        wb = openpyxl.load_workbook(excel_path, read_only=True)
        for name in wb.sheetnames:
            if '消費紀錄' in name or '信用卡' in name:
                sheet_name = name
                break
        wb.close()
        if sheet_name is None:
            sheet_name = 0  # 使用第一個 sheet

    df = pd.read_excel(excel_path, sheet_name=sheet_name, header=None)

    # 找到明細區域的起始行（找 "消費日期" 的位置）
    header_row = None
    for i in range(min(20, len(df))):
        for j in range(min(10, len(df.columns))):
            val = df.iloc[i, j]
            if val == '消費日期':
                header_row = i
                break
        if header_row is not None:
            break

    if header_row is None:
        raise ValueError("無法在 Excel 中找到消費明細表頭")

    # 提取明細資料
    detail_df = df.iloc[header_row + 1:].copy()
    # 取表頭行的值，NaN 的欄位用 extra_N 命名
    raw_headers = list(df.iloc[header_row])
    col_names = []
    extra_idx = 0
    for h in raw_headers:
        if pd.notna(h):
            col_names.append(str(h))
        else:
            col_names.append(f'extra_{extra_idx}')
            extra_idx += 1
    detail_df.columns = col_names

    # 只保留需要的欄位（支援不同表頭名稱）
    rename_map = {'消費明細(含消費地)': '消費明細'}
    detail_df = detail_df.rename(columns=rename_map)

    keep_cols = ['消費日期', '入帳起息日', '消費明細', '清算消費金額',
                 '消費類別', 'Owner', '結算月份', '結算年份']
    available = [c for c in keep_cols if c in detail_df.columns]
    detail_df = detail_df[available].copy()

    # 清理空行
    detail_df = detail_df.dropna(subset=['消費日期'])

    # 轉換日期
    for col in ['消費日期', '入帳起息日']:
        if col in detail_df.columns:
            detail_df[col] = pd.to_datetime(detail_df[col], errors='coerce').dt.date

    # 轉換數字
    if '清算消費金額' in detail_df.columns:
        detail_df['清算消費金額'] = pd.to_numeric(detail_df['清算消費金額'], errors='coerce')

    if '結算月份' in detail_df.columns:
        detail_df['結算月份'] = pd.to_numeric(detail_df['結算月份'], errors='coerce').astype('Int64')

    if '結算年份' in detail_df.columns:
        detail_df['結算年份'] = pd.to_numeric(detail_df['結算年份'], errors='coerce').astype('Int64')

    detail_df = detail_df.dropna(subset=['清算消費金額'])
    return detail_df.reset_index(drop=True)


def get_monthly_summary(df: pd.DataFrame, year: int = None) -> pd.DataFrame:
    """
    產生月度摘要（對應 Excel 上方的摘要區域）

    Returns:
        DataFrame with columns: 結算月份, Alan, Lydia, 總金額
    """
    if df.empty:
        return pd.DataFrame()

    if year:
        df = df[df['結算年份'] == year]

    summary = df.pivot_table(
        values='清算消費金額',
        index='結算月份',
        columns='Owner',
        aggfunc='sum',
        fill_value=0
    )

    summary['總金額'] = summary.sum(axis=1)
    summary = summary.reset_index()
    summary['結算月份'] = summary['結算月份'].astype(int)
    summary = summary.sort_values('結算月份')

    return summary


def get_category_summary(df: pd.DataFrame, year: int = None) -> pd.DataFrame:
    """
    產生類別摘要（四大超商 / 一般）

    Returns:
        DataFrame with columns: 結算月份, 四大超商, 一般
    """
    if df.empty:
        return pd.DataFrame()

    if year:
        df = df[df['結算年份'] == year]

    summary = df.pivot_table(
        values='清算消費金額',
        index='結算月份',
        columns='消費類別',
        aggfunc='sum',
        fill_value=0
    )
    summary = summary.reset_index()
    summary['結算月份'] = summary['結算月份'].astype(int)
    return summary.sort_values('結算月份')


def _empty_df() -> pd.DataFrame:
    """回傳空的標準 DataFrame"""
    return pd.DataFrame(columns=[
        '消費日期', '入帳起息日', '消費明細', '清算消費金額',
        '消費類別', 'Owner', '結算月份', '結算年份'
    ])
