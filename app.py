"""
台新信用卡帳單分析 App
=============================
功能：
1. 從 Google Drive 自動下載 PDF 帳單
2. 解析 PDF 提取消費明細
3. 自動分類（四大超商/一般）、判定 Owner（Alan/Lydia）
4. 產生與 Excel 相同的月度摘要、回饋計算
5. 支援手動上傳 PDF、匯入既有 Excel 資料
"""

import streamlit as st
import pandas as pd
import json
import os
from datetime import datetime
from pathlib import Path

from utils.classifier import (
    classify_transaction, calculate_rewards,
    REWARD_RATES, CONVENIENCE_REWARD_CAP
)
from utils.data_manager import (
    load_records, save_records, merge_records,
    import_from_excel, get_monthly_summary,
    get_category_summary, _empty_df
)
from utils.pdf_parser import parse_taishin_pdf, extract_billing_period
from utils.charts import (
    monthly_bar_chart, category_bar_chart, reward_chart,
    owner_pie_chart, category_pie_chart, trend_line_chart
)
from utils.email_sender import send_report_email

# ─────────────────────────────────────────────
# 頁面設定
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="台新信用卡帳單分析",
    page_icon="💳",
    layout="wide",
    initial_sidebar_state="expanded"
)

APP_DIR = Path(__file__).parent
DRIVE_CONFIG_PATH = APP_DIR / 'drive_config.json'
EMAIL_CONFIG_PATH = APP_DIR / 'email_config.json'
BILLING_DAY = 27

# ─────────────────────────────────────────────
# 雲端 / 本地 自動偵測
# ─────────────────────────────────────────────
def _is_cloud() -> bool:
    """判斷是否在 Streamlit Cloud 上執行（有 secrets 設定）"""
    try:
        return "gdrive_credentials" in st.secrets
    except Exception:
        return False


def _get_cloud_credentials() -> bytes | None:
    """從 st.secrets 取得 Google Drive 憑證 bytes"""
    try:
        cred = st.secrets["gdrive_credentials"]
        if isinstance(cred, str):
            return cred.encode("utf-8")
        else:
            return json.dumps(dict(cred)).encode("utf-8")
    except Exception:
        return None


def _get_cloud_folder_id() -> str:
    """從 st.secrets 取得 Drive folder ID"""
    try:
        return st.secrets["gdrive_folder_id"]
    except Exception:
        return ""


def _get_cloud_pdf_password() -> str:
    """從 st.secrets 取得 PDF 密碼"""
    try:
        return st.secrets["pdf_password"]
    except Exception:
        return ""


def _secret(key: str, default: str = "") -> str:
    """安全讀取 st.secrets 的值"""
    try:
        return st.secrets[key]
    except Exception:
        return default


def _get_cloud_email_config() -> dict:
    """從 st.secrets 取得 Email 設定"""
    return {
        'sender_email': _secret("email_sender"),
        'app_password': _secret("email_app_password"),
        'recipients': {
            'Alan': _secret("email_recipient_alan"),
            'Lydia': _secret("email_recipient_lydia"),
        }
    }


IS_CLOUD = _is_cloud()

# 確保 data 目錄存在（雲端部署首次啟動時可能不存在）
(APP_DIR / 'data').mkdir(parents=True, exist_ok=True)


def load_email_config():
    """讀取 Email 設定（雲端優先）"""
    if IS_CLOUD:
        cloud_cfg = _get_cloud_email_config()
        if cloud_cfg.get('sender_email'):
            return cloud_cfg
    if EMAIL_CONFIG_PATH.exists():
        with open(EMAIL_CONFIG_PATH, 'r', encoding='utf-8') as f:
            return json.load(f)
    return {'sender_email': '', 'recipients': {'Alan': '', 'Lydia': ''}}


def save_email_config(config):
    """儲存 Email 設定（雲端模式下僅 session 內有效）"""
    if IS_CLOUD:
        return  # 雲端不寫入本地檔案
    with open(EMAIL_CONFIG_PATH, 'w', encoding='utf-8') as f:
        json.dump(config, f, ensure_ascii=False, indent=2)


# ─────────────────────────────────────────────
# Session State 初始化
# ─────────────────────────────────────────────
if 'records' not in st.session_state:
    st.session_state.records = load_records()
if 'last_import' not in st.session_state:
    st.session_state.last_import = None


# ─────────────────────────────────────────────
# 輔助函式
# ─────────────────────────────────────────────
def load_drive_config():
    """讀取 Drive 設定（雲端優先）"""
    if IS_CLOUD:
        return {
            'folder_id': _get_cloud_folder_id(),
            'credentials_path': '__cloud__',
        }
    if DRIVE_CONFIG_PATH.exists():
        with open(DRIVE_CONFIG_PATH, 'r', encoding='utf-8') as f:
            return json.load(f)
    return {}


def save_drive_config(config):
    """儲存 Drive 設定（雲端模式下僅 session 內有效）"""
    if IS_CLOUD:
        return
    with open(DRIVE_CONFIG_PATH, 'w', encoding='utf-8') as f:
        json.dump(config, f, ensure_ascii=False, indent=2)


def load_credentials(cred_path: str) -> bytes | None:
    """載入 Service Account 憑證（雲端優先）"""
    if IS_CLOUD:
        return _get_cloud_credentials()
    # 支援相對路徑（相對於 APP_DIR）
    path = Path(cred_path)
    if not path.is_absolute():
        path = APP_DIR / path
    if path.exists():
        return path.read_bytes()
    return None


def process_pdf(pdf_bytes: bytes, filename: str = "",
                password: str = None) -> pd.DataFrame:
    """處理 PDF：解析 → 分類 → 回傳完整 DataFrame"""
    # 1. 解析 PDF
    raw_df = parse_taishin_pdf(pdf_bytes, BILLING_DAY, password=password)

    if raw_df.empty:
        st.warning(f"⚠️ 無法從 PDF 中提取交易 ({filename})")
        return _empty_df()

    # 2. 嘗試從 PDF 取得帳單期間
    period = extract_billing_period(pdf_bytes, password=password)

    # 3. 對每筆交易進行分類
    classifications = raw_df['消費明細'].apply(
        lambda desc: classify_transaction(desc, raw_df.loc[
            raw_df['消費明細'] == desc, '消費日期'
        ].iloc[0] if not raw_df[raw_df['消費明細'] == desc].empty else datetime.now(), BILLING_DAY)
    )

    # 更精確的分類 - 逐行處理
    categories = []
    owners = []
    bill_months = []
    bill_years = []

    for idx, row in raw_df.iterrows():
        posting = row.get('入帳起息日', None)
        result = classify_transaction(
            row['消費明細'],
            row['消費日期'],
            BILLING_DAY,
            posting_date=posting
        )
        categories.append(result['消費類別'])
        owners.append(result['Owner'])
        bill_months.append(result['結算月份'])
        bill_years.append(result['結算年份'])

    raw_df['消費類別'] = categories
    raw_df['Owner'] = owners
    raw_df['結算月份'] = bill_months
    raw_df['結算年份'] = bill_years

    return raw_df


# ─────────────────────────────────────────────
# 側邊欄
# ─────────────────────────────────────────────
with st.sidebar:
    st.title("💳 台新信用卡分析")
    st.divider()

    # ── 資料來源 ──
    st.subheader("📂 資料來源")
    data_source = st.radio(
        "選擇資料來源",
        ["🔄 Google Drive 同步", "📤 手動上傳 PDF", "📊 匯入 Excel"],
        label_visibility="collapsed"
    )

    if data_source == "🔄 Google Drive 同步":
        config = load_drive_config()

        if IS_CLOUD:
            st.info("☁️ 使用 Streamlit Cloud Secrets 設定")
            folder_id = config.get('folder_id', '')
            cred_path = '__cloud__'
            if folder_id:
                st.success(f"✅ Folder ID: ...{folder_id[-8:]}")
        else:
            folder_id = st.text_input(
                "Drive 資料夾 ID",
                value=config.get('folder_id', ''),
                help="放置信用卡帳單 PDF 的 Google Drive 資料夾 ID"
            )
            cred_path = st.text_input(
                "憑證路徑",
                value=config.get('credentials_path', '../budget_app/credentials.json'),
                help="Google Service Account JSON 檔案路徑"
            )

            if st.button("💾 儲存設定", use_container_width=True):
                save_drive_config({
                    'folder_id': folder_id,
                    'credentials_path': cred_path,
                    'billing_day': BILLING_DAY,
                })
                st.success("✅ 設定已儲存")

        # PDF 密碼：雲端從 secrets 讀取，本地讓使用者輸入
        cloud_pdf_pwd = _get_cloud_pdf_password() if IS_CLOUD else ""
        drive_pdf_password = cloud_pdf_pwd or st.text_input(
            "PDF 密碼",
            type="password",
            help="台新帳單通常以身分證字號作為密碼",
            placeholder="輸入身分證字號...",
            key="drive_pdf_password"
        )

        if st.button("🔄 從 Drive 同步帳單", type="primary", use_container_width=True):
            if not folder_id:
                st.error("請先設定 Drive 資料夾 ID")
            else:
                cred_bytes = load_credentials(cred_path)
                if cred_bytes is None:
                    st.error(f"找不到憑證檔案: {cred_path}")
                else:
                    with st.spinner("正在從 Google Drive 下載帳單..."):
                        try:
                            from utils.drive_sync import get_all_pdfs, download_pdf
                            pdf_files = get_all_pdfs(cred_bytes, folder_id)
                            if not pdf_files:
                                st.warning("Drive 資料夾中沒有 PDF 檔案")
                            else:
                                st.info(f"找到 {len(pdf_files)} 個 PDF 檔案")
                                all_new = _empty_df()
                                progress = st.progress(0)
                                for i, f in enumerate(pdf_files):
                                    pdf_bytes = download_pdf(
                                        cred_bytes, folder_id, f['id']
                                    )
                                    new_df = process_pdf(
                                        pdf_bytes, f['name'],
                                        password=drive_pdf_password or None
                                    )
                                    if not new_df.empty:
                                        all_new = pd.concat(
                                            [all_new, new_df], ignore_index=True
                                        )
                                    progress.progress((i + 1) / len(pdf_files))

                                if not all_new.empty:
                                    st.session_state.records = merge_records(
                                        st.session_state.records, all_new
                                    )
                                    save_records(st.session_state.records)
                                    st.success(
                                        f"✅ 匯入 {len(all_new)} 筆交易"
                                    )
                                else:
                                    st.warning("未能從 PDF 中提取交易資料")
                        except Exception as e:
                            st.error(f"同步失敗: {e}")

    elif data_source == "📤 手動上傳 PDF":
        uploaded_files = st.file_uploader(
            "上傳台新信用卡帳單 PDF",
            type=['pdf'],
            accept_multiple_files=True,
            help="支援一次上傳多個月份的帳單"
        )
        pdf_password = st.text_input(
            "PDF 密碼",
            type="password",
            help="台新帳單通常以身分證字號作為密碼",
            placeholder="輸入身分證字號..."
        )

        if uploaded_files and st.button("🔍 預覽 PDF 文字", use_container_width=True):
            for f in uploaded_files:
                pdf_bytes = f.read()
                try:
                    from utils.pdf_parser import _decrypt_pdf, _parse_from_words, _parse_roc_text
                    import pdfplumber as _pb
                    buf = _decrypt_pdf(pdf_bytes, pdf_password or None)
                    with _pb.open(buf) as pdf:
                        st.write(f"📄 **{f.name}** — {len(pdf.pages)} 頁")

                        word_txns = _parse_from_words(pdf)

                        # 每頁的原始文字 + words 分析
                        all_text = ""
                        raw_pages = []
                        for pi, page in enumerate(pdf.pages):
                            text = page.extract_text()
                            if text:
                                all_text += text + "\n"
                                raw_pages.append((pi + 1, text))

                        text_txns = _parse_roc_text(all_text) if all_text else []

                        st.success(
                            f"**Words 座標解析: {len(word_txns)} 筆** | "
                            f"**文字解析: {len(text_txns)} 筆**"
                        )

                        best = word_txns if len(word_txns) >= len(text_txns) else text_txns
                        import pandas as _pd2
                        if best:
                            preview_df = _pd2.DataFrame(best)
                            st.dataframe(preview_df, height=400)

                        # ── 原始文字 dump（debug 用）──
                        with st.expander("📝 原始文字 (Raw Text)", expanded=False):
                            for page_num, page_text in raw_pages:
                                st.markdown(f"**── 第 {page_num} 頁 ──**")
                                numbered = []
                                for li, ln in enumerate(page_text.split('\n'), 1):
                                    numbered.append(f"{li:3d} | {ln}")
                                st.code('\n'.join(numbered), language=None)

                        # ── extract_words 分析 ──
                        with st.expander("🔤 Words 座標分析", expanded=False):
                            buf2 = _decrypt_pdf(pdf_bytes, pdf_password or None)
                            with _pb.open(buf2) as pdf2:
                                for pi, page in enumerate(pdf2.pages):
                                    words = page.extract_words(
                                        x_tolerance=3, y_tolerance=3,
                                        keep_blank_chars=True
                                    )
                                    if not words:
                                        continue
                                    st.markdown(f"**── 第 {pi+1} 頁 ({len(words)} words) ──**")
                                    rows_by_y = {}
                                    for w in words:
                                        y_key = round(w['top'], 0)
                                        rows_by_y.setdefault(y_key, []).append(w)
                                    word_lines = []
                                    for y_key in sorted(rows_by_y.keys()):
                                        row_words = sorted(rows_by_y[y_key], key=lambda w: w['x0'])
                                        text_items = [f"[x={int(w['x0']):3d}]{w['text']}" for w in row_words]
                                        word_lines.append(f"y={int(y_key):4d}: {' '.join(text_items)}")
                                    st.code('\n'.join(word_lines[:80]), language=None)

                except ValueError as e:
                    st.error(f"❌ {e}")
                f.seek(0)

        if uploaded_files and st.button("📥 匯入 PDF", type="primary", use_container_width=True):
            all_new = _empty_df()
            for f in uploaded_files:
                pdf_bytes = f.read()
                try:
                    new_df = process_pdf(pdf_bytes, f.name, password=pdf_password or None)
                except ValueError as e:
                    st.error(f"❌ {e}")
                    new_df = _empty_df()
                if not new_df.empty:
                    all_new = pd.concat([all_new, new_df], ignore_index=True)

            if not all_new.empty:
                st.session_state.records = merge_records(
                    st.session_state.records, all_new
                )
                save_records(st.session_state.records)
                st.success(f"✅ 匯入 {len(all_new)} 筆交易")
                st.session_state.last_import = datetime.now().strftime('%Y-%m-%d %H:%M')

    elif data_source == "📊 匯入 Excel":
        excel_file = st.file_uploader(
            "上傳信用卡明細 Excel",
            type=['xlsx', 'xls'],
            help="選擇你現有的信用卡明細 Excel 檔案"
        )

        if excel_file and st.button("📥 匯入 Excel", type="primary", use_container_width=True):
            with st.spinner("正在匯入 Excel..."):
                try:
                    # 儲存暫存檔
                    tmp_path = APP_DIR / 'data' / 'tmp_upload.xlsx'
                    tmp_path.parent.mkdir(parents=True, exist_ok=True)
                    with open(tmp_path, 'wb') as f:
                        f.write(excel_file.read())

                    excel_df = import_from_excel(str(tmp_path))

                    if not excel_df.empty:
                        st.session_state.records = merge_records(
                            st.session_state.records, excel_df
                        )
                        save_records(st.session_state.records)
                        st.success(f"✅ 匯入 {len(excel_df)} 筆交易")
                    else:
                        st.warning("Excel 中沒有有效資料")

                    # 清理暫存
                    if tmp_path.exists():
                        tmp_path.unlink()
                except Exception as e:
                    st.error(f"匯入失敗: {e}")

    st.divider()

    # ── 篩選 ──
    st.subheader("🔍 篩選條件")
    records = st.session_state.records

    if not records.empty and '結算年份' in records.columns:
        years = sorted(records['結算年份'].dropna().unique(), reverse=True)
        selected_year = st.selectbox("年份", years if years else [2026], index=0)
    else:
        selected_year = 2026

    selected_month = st.selectbox(
        "月份",
        [0] + list(range(1, 13)),
        format_func=lambda x: '全年' if x == 0 else f'{x} 月',
    )

    selected_owner = st.selectbox(
        "Owner",
        ['全部', 'Alan', 'Lydia'],
    )

    st.divider()

    # ── 資料統計 ──
    st.subheader("📊 資料統計")
    st.metric("總交易筆數", len(records))
    if not records.empty:
        total_amount = records['清算消費金額'].sum()
        st.metric("總消費金額", f"${total_amount:,.0f}")

    # ── Email 設定 ──
    st.divider()
    st.subheader("📧 Email 設定")
    email_cfg = load_email_config()

    # 判斷是否有雲端 Email 設定（直接檢查 secrets）
    _cloud_email = _secret("email_sender") if IS_CLOUD else ""

    if _cloud_email:
        st.info("☁️ 使用 Streamlit Cloud Secrets 的 Email 設定")
        email_sender = _cloud_email
        email_app_pwd = _secret("email_app_password")
        email_alan = _secret("email_recipient_alan")
        email_lydia = _secret("email_recipient_lydia")
        # 把密碼存入 session_state 供寄信功能使用
        st.session_state['email_app_pwd'] = email_app_pwd
        with st.expander("目前設定", expanded=False):
            st.write(f"寄件者: {email_sender}")
            st.write(f"Alan: {email_alan}")
            st.write(f"Lydia: {email_lydia}")
    else:
        if IS_CLOUD:
            st.caption("⚠️ Secrets 中未找到 email_sender，請確認已設定")
        with st.expander("設定寄件資訊", expanded=False):
            email_sender = st.text_input(
                "寄件者 Gmail",
                value=email_cfg.get('sender_email', ''),
                placeholder="your_email@gmail.com"
            )
            email_app_pwd = st.text_input(
                "應用程式密碼",
                type="password",
                help="請到 Google 帳戶 → 安全性 → 兩步驟驗證 → 應用程式密碼，產生一組 16 碼密碼",
                placeholder="xxxx xxxx xxxx xxxx",
                key="email_app_pwd"
            )
            email_alan = st.text_input(
                "Alan 的 Email",
                value=email_cfg.get('recipients', {}).get('Alan', ''),
                placeholder="alan@example.com"
            )
            email_lydia = st.text_input(
                "Lydia 的 Email",
                value=email_cfg.get('recipients', {}).get('Lydia', ''),
                placeholder="lydia@example.com"
            )
            if st.button("💾 儲存 Email 設定", use_container_width=True):
                save_email_config({
                    'sender_email': email_sender,
                    'recipients': {'Alan': email_alan, 'Lydia': email_lydia}
                })
                st.success("✅ Email 設定已儲存")

    # ── 資料管理 ──
    st.divider()
    if st.button("🗑️ 清除所有資料", use_container_width=True):
        st.session_state.records = _empty_df()
        save_records(st.session_state.records)
        st.success("已清除所有資料")
        st.rerun()


# ─────────────────────────────────────────────
# 資料篩選
# ─────────────────────────────────────────────
df = st.session_state.records.copy()
if not df.empty:
    df = df[df['結算年份'] == selected_year]
    if selected_month > 0:
        df = df[df['結算月份'] == selected_month]
    if selected_owner != '全部':
        df = df[df['Owner'] == selected_owner]


# ─────────────────────────────────────────────
# 主畫面
# ─────────────────────────────────────────────
st.title("💳 台新信用卡帳單分析")

if df.empty and st.session_state.records.empty:
    st.info("""
    ### 🚀 開始使用
    請透過左側面板匯入資料：

    1. **Google Drive 同步** — 設定 Drive 資料夾 ID，自動下載並分析 PDF 帳單
    2. **手動上傳 PDF** — 直接上傳台新帳單 PDF
    3. **匯入 Excel** — 從你現有的「信用卡明細.xlsx」匯入歷史資料

    💡 **建議先匯入 Excel** 將既有資料導入系統，之後每月再用 Google Drive 同步新帳單。
    """)
    st.stop()

# ═══════════════════════════════════════════════
# Tab 1: 月報總覽
# ═══════════════════════════════════════════════
tab1, tab2, tab3, tab4 = st.tabs([
    "📊 月報總覽", "📋 消費明細", "🎁 回饋分析", "📈 趨勢分析"
])

with tab1:
    st.header(f"📊 {selected_year} 年度月報總覽")

    # ── 月度摘要表格（對應 Excel Row 3-5）──
    all_year_df = st.session_state.records[
        st.session_state.records['結算年份'] == selected_year
    ]
    summary = get_monthly_summary(all_year_df, selected_year)

    if not summary.empty:
        # KPI 卡片
        period_label = f"{selected_month}月" if selected_month > 0 else "全年"
        col1, col2, col3, col4 = st.columns(4)

        period_df = df  # 已篩選
        with col1:
            st.metric(
                f"📅 {period_label} 總消費",
                f"${period_df['清算消費金額'].sum():,.0f}"
            )
        with col2:
            alan_total = period_df[period_df['Owner'] == 'Alan']['清算消費金額'].sum()
            st.metric("👤 Alan", f"${alan_total:,.0f}")
        with col3:
            lydia_total = period_df[period_df['Owner'] == 'Lydia']['清算消費金額'].sum()
            st.metric("👤 Lydia", f"${lydia_total:,.0f}")
        with col4:
            st.metric("📝 交易筆數", f"{len(period_df)} 筆")

        st.divider()

        # ── 每月消費摘要表（對應 Excel Row 2-5）──
        st.subheader("每月消費摘要")

        # 建構完整12月的表格
        full_months = pd.DataFrame({'結算月份': range(1, 13)})
        display_summary = full_months.merge(summary, on='結算月份', how='left').fillna(0)

        # 加上年度總額
        num_cols = [c for c in display_summary.columns if c != '結算月份']
        total_dict = {c: display_summary[c].sum() for c in num_cols}
        total_dict['結算月份'] = '年度總額'
        display_table = pd.concat(
            [display_summary, pd.DataFrame([total_dict])], ignore_index=True
        )
        display_table['結算月份'] = display_table['結算月份'].apply(
            lambda x: f'{int(x)}月' if isinstance(x, (int, float)) else x
        )

        st.dataframe(
            display_table.style.format({
                c: '${:,.0f}' for c in display_table.columns if c != '結算月份'
            }),
            use_container_width=True,
            hide_index=True,
        )

        # ── 圖表 ──
        col_left, col_right = st.columns(2)
        with col_left:
            fig1 = monthly_bar_chart(summary, selected_year)
            st.plotly_chart(fig1, use_container_width=True)
        with col_right:
            cat_summary = get_category_summary(all_year_df, selected_year)
            fig2 = category_bar_chart(cat_summary, selected_year)
            st.plotly_chart(fig2, use_container_width=True)
        # ── 📧 寄送報告 ──
        st.divider()
        st.subheader("📧 寄送消費分析報告")

        email_cfg_send = load_email_config()
        sender = email_cfg_send.get('sender_email', '')
        recipients = email_cfg_send.get('recipients', {})
        # 密碼：優先從 cloud config 取，其次從 session_state
        _email_pwd = (
            email_cfg_send.get('app_password', '')
            or st.session_state.get('email_app_pwd', '')
        )

        if not sender:
            st.warning("請先在左側欄設定寄件者 Gmail 和應用程式密碼")
        else:
            _cat_summary = get_category_summary(all_year_df, selected_year)

            send_col1, send_col2, send_col3 = st.columns(3)

            with send_col1:
                alan_email = recipients.get('Alan', '')
                alan_df = all_year_df[all_year_df['Owner'] == 'Alan']
                if selected_month > 0:
                    alan_df = alan_df[alan_df['結算月份'] == selected_month]
                alan_count = len(alan_df)
                alan_total = alan_df['清算消費金額'].sum() if not alan_df.empty else 0
                st.markdown(f"**Alan** — {alan_count} 筆 / ${alan_total:,.0f}")
                if st.button(
                    f"📧 寄送給 Alan",
                    disabled=(not alan_email or alan_df.empty),
                    use_container_width=True,
                    key="send_alan"
                ):
                    with st.spinner("寄送中..."):
                        result = send_report_email(
                            sender, _email_pwd,
                            alan_email, 'Alan',
                            selected_year, selected_month,
                            alan_df, summary, _cat_summary
                        )
                        if result.startswith('✅'):
                            st.success(result)
                        else:
                            st.error(result)

            with send_col2:
                lydia_email = recipients.get('Lydia', '')
                lydia_df = all_year_df[all_year_df['Owner'] == 'Lydia']
                if selected_month > 0:
                    lydia_df = lydia_df[lydia_df['結算月份'] == selected_month]
                lydia_count = len(lydia_df)
                lydia_total_amt = lydia_df['清算消費金額'].sum() if not lydia_df.empty else 0
                st.markdown(f"**Lydia** — {lydia_count} 筆 / ${lydia_total_amt:,.0f}")
                if st.button(
                    f"📧 寄送給 Lydia",
                    disabled=(not lydia_email or lydia_df.empty),
                    use_container_width=True,
                    key="send_lydia"
                ):
                    with st.spinner("寄送中..."):
                        result = send_report_email(
                            sender, _email_pwd,
                            lydia_email, 'Lydia',
                            selected_year, selected_month,
                            lydia_df, summary, _cat_summary
                        )
                        if result.startswith('✅'):
                            st.success(result)
                        else:
                            st.error(result)

            with send_col3:
                both_ready = (alan_email and lydia_email and not all_year_df.empty)
                st.markdown("**全部寄送**")
                if st.button(
                    "📧 同時寄送兩人",
                    disabled=not both_ready,
                    type="primary",
                    use_container_width=True,
                    key="send_both"
                ):
                    with st.spinner("寄送中..."):
                        for name, email_addr in [('Alan', alan_email), ('Lydia', lydia_email)]:
                            owner_data = all_year_df[all_year_df['Owner'] == name]
                            if selected_month > 0:
                                owner_data = owner_data[owner_data['結算月份'] == selected_month]
                            if owner_data.empty:
                                st.warning(f"{name} 無消費資料，跳過")
                                continue
                            result = send_report_email(
                                sender, _email_pwd, email_addr, name,
                                selected_year, selected_month,
                                owner_data, summary, _cat_summary
                            )
                            if result.startswith('✅'):
                                st.success(result)
                            else:
                                st.error(result)

    else:
        st.info("目前沒有資料，請先匯入帳單")


# ═══════════════════════════════════════════════
# Tab 2: 消費明細
# ═══════════════════════════════════════════════
with tab2:
    st.header("📋 消費明細")

    if not df.empty:
        # 搜尋/篩選
        col1, col2 = st.columns([3, 1])
        with col1:
            search = st.text_input("🔍 搜尋消費明細", placeholder="輸入商家名稱...")
        with col2:
            cat_filter = st.selectbox("消費類別", ['全部', '四大超商', '一般'])

        filtered = df.copy()
        if search:
            filtered = filtered[
                filtered['消費明細'].str.contains(search, case=False, na=False)
            ]
        if cat_filter != '全部':
            filtered = filtered[filtered['消費類別'] == cat_filter]

        # 顯示篩選後統計
        st.info(
            f"共 **{len(filtered)}** 筆 | "
            f"總金額 **${filtered['清算消費金額'].sum():,.0f}** | "
            f"四大超商 **{len(filtered[filtered['消費類別'] == '四大超商'])}** 筆 | "
            f"一般 **{len(filtered[filtered['消費類別'] == '一般'])}** 筆"
        )

        # 顯示明細表格（Owner 欄可編輯）
        display_cols = ['消費日期', '入帳起息日', '消費明細', '清算消費金額',
                        '消費類別', 'Owner', '結算月份']
        available_cols = [c for c in display_cols if c in filtered.columns]

        # 建立欄位設定：Owner 用下拉選單，其他欄位唯讀
        column_config = {
            '清算消費金額': st.column_config.NumberColumn(
                '清算消費金額', format="$%d",
            ),
            'Owner': st.column_config.SelectboxColumn(
                'Owner',
                options=['Alan', 'Lydia'],
                required=True,
                width='small',
            ),
        }
        disabled_cols = [c for c in available_cols if c != 'Owner']

        edited_df = st.data_editor(
            filtered[available_cols],
            column_config=column_config,
            disabled=disabled_cols,
            use_container_width=True,
            hide_index=True,
            height=500,
            key="owner_editor",
        )

        # 偵測 Owner 是否被手動修改，同步回 session_state.records
        if edited_df is not None and 'Owner' in edited_df.columns:
            # 比對篩選後 df 的原始 index，回寫到 records
            original_owners = filtered['Owner'].values
            edited_owners = edited_df['Owner'].values
            if len(original_owners) == len(edited_owners):
                changed = False
                for i, (orig, edit) in enumerate(zip(original_owners, edited_owners)):
                    if orig != edit:
                        # 找到在 session_state.records 中對應的行
                        rec_idx = filtered.index[i]
                        st.session_state.records.at[rec_idx, 'Owner'] = edit
                        changed = True
                if changed:
                    save_records(st.session_state.records)
                    st.toast("✅ Owner 已更新並儲存", icon="💾")
                    st.rerun()

        # 下載按鈕
        csv_data = filtered[available_cols].to_csv(index=False, encoding='utf-8-sig')
        st.download_button(
            "📥 下載 CSV",
            data=csv_data,
            file_name=f"信用卡明細_{selected_year}_{selected_month or '全年'}.csv",
            mime='text/csv',
        )
    else:
        st.info("無消費明細資料")


# ═══════════════════════════════════════════════
# Tab 3: 回饋分析（對應 Excel Row 7-11）
# ═══════════════════════════════════════════════
with tab3:
    st.header("🎁 回饋金額分析")

    if not all_year_df.empty:
        # 回饋規則說明
        with st.expander("📖 回饋規則", expanded=False):
            st.markdown(f"""
            | 消費類別 | 回饋率 | 說明 |
            |---------|--------|------|
            | 四大超商 | **{REWARD_RATES['四大超商']*100:.0f}%** | 每月上限 ${CONVENIENCE_REWARD_CAP:,.0f} |
            | 一般消費 | **{REWARD_RATES['一般']*100:.0f}%** | 無上限 |
            """)

        # 計算每月回饋
        rewards_by_month = {}
        for month in sorted(all_year_df['結算月份'].dropna().unique()):
            month_data = all_year_df[all_year_df['結算月份'] == month]
            cat_totals = month_data.groupby('消費類別')['清算消費金額'].sum().to_dict()
            rewards_by_month[int(month)] = calculate_rewards(cat_totals)

        # 回饋摘要表格（對應 Excel Row 7-11）
        st.subheader("📊 月度回饋明細")
        reward_rows = []
        for month, r in sorted(rewards_by_month.items()):
            reward_rows.append({
                '月份': f'{month}月',
                '四大超商消費': r['四大超商_消費'],
                '四大超商回饋': r['四大超商_回饋'],
                '一般消費': r['一般_消費'],
                '一般回饋': r['一般_回饋'],
                '當月回饋總額': r['回饋總額'],
            })

        if reward_rows:
            reward_df = pd.DataFrame(reward_rows)

            # 加總行
            totals = reward_df.select_dtypes(include='number').sum()
            totals['月份'] = '年度合計'
            reward_df = pd.concat(
                [reward_df, pd.DataFrame([totals])], ignore_index=True
            )

            st.dataframe(
                reward_df.style.format({
                    '四大超商消費': '${:,.0f}',
                    '四大超商回饋': '${:,.1f}',
                    '一般消費': '${:,.0f}',
                    '一般回饋': '${:,.1f}',
                    '當月回饋總額': '${:,.1f}',
                }),
                use_container_width=True,
                hide_index=True,
            )

            # 回饋圖表
            fig = reward_chart(rewards_by_month, selected_year)
            st.plotly_chart(fig, use_container_width=True)

            # 年度回饋 KPI
            total_reward = sum(r['回饋總額'] for r in rewards_by_month.values())
            total_spend = all_year_df['清算消費金額'].sum()
            effective_rate = (total_reward / total_spend * 100) if total_spend > 0 else 0

            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("💰 年度回饋總額", f"${total_reward:,.1f}")
            with col2:
                st.metric("📊 實際回饋率", f"{effective_rate:.2f}%")
            with col3:
                st.metric("💳 年度總消費", f"${total_spend:,.0f}")
    else:
        st.info("無資料可分析")


# ═══════════════════════════════════════════════
# Tab 4: 趨勢分析
# ═══════════════════════════════════════════════
with tab4:
    st.header("📈 消費趨勢分析")

    if not all_year_df.empty:
        # 趨勢折線圖
        fig = trend_line_chart(all_year_df, selected_year)
        st.plotly_chart(fig, use_container_width=True)

        # 圓餅圖
        col1, col2 = st.columns(2)
        with col1:
            month_for_pie = selected_month if selected_month > 0 else None
            fig_pie1 = owner_pie_chart(
                st.session_state.records, selected_year, month_for_pie
            )
            st.plotly_chart(fig_pie1, use_container_width=True)
        with col2:
            fig_pie2 = category_pie_chart(
                st.session_state.records, selected_year, month_for_pie
            )
            st.plotly_chart(fig_pie2, use_container_width=True)

        # 商家排行
        st.subheader("🏪 消費商家排行 Top 15")
        top_merchants = (
            all_year_df.groupby('消費明細')['清算消費金額']
            .agg(['sum', 'count'])
            .rename(columns={'sum': '總金額', 'count': '次數'})
            .sort_values('總金額', ascending=False)
            .head(15)
            .reset_index()
        )
        top_merchants['平均每次'] = (
            top_merchants['總金額'] / top_merchants['次數']
        ).round(0)

        st.dataframe(
            top_merchants.style.format({
                '總金額': '${:,.0f}',
                '平均每次': '${:,.0f}',
            }),
            use_container_width=True,
            hide_index=True,
        )
    else:
        st.info("無資料可分析")
