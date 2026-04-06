"""
Email 寄送模組 — 透過 Gmail SMTP 發送信用卡消費分析報告
使用 Gmail App Password（應用程式密碼）驗證
"""

import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from datetime import date
import pandas as pd


SMTP_SERVER = 'smtp.gmail.com'
SMTP_PORT = 587


def send_report_email(
    sender_email: str,
    app_password: str,
    recipient_email: str,
    owner_name: str,
    year: int,
    month: int,
    owner_df: pd.DataFrame,
    summary_df: pd.DataFrame,
    category_df: pd.DataFrame,
) -> str:
    """
    寄送消費分析報告給指定 owner。

    Args:
        sender_email: 寄件者 Gmail
        app_password: Gmail 應用程式密碼
        recipient_email: 收件者 email
        owner_name: 持卡人名稱 (Alan / Lydia)
        year: 年份
        month: 月份 (0 = 全年)
        owner_df: 該 owner 的消費明細 DataFrame
        summary_df: 月度摘要 DataFrame
        category_df: 類別摘要 DataFrame

    Returns:
        成功/失敗訊息
    """
    period = f"{year}年{month}月" if month > 0 else f"{year}年度"
    subject = f"💳 {period} 信用卡消費分析 — {owner_name}"

    html_body = _build_html_report(
        owner_name, year, month, owner_df, summary_df, category_df
    )

    msg = MIMEMultipart('alternative')
    msg['Subject'] = subject
    msg['From'] = sender_email
    msg['To'] = recipient_email

    # 純文字備用
    text_part = MIMEText(
        f"{period} 信用卡消費分析報告\n\n"
        f"持卡人：{owner_name}\n"
        f"總消費：${owner_df['清算消費金額'].sum():,.0f}\n"
        f"交易筆數：{len(owner_df)} 筆\n\n"
        f"詳細內容請查看 HTML 版本。",
        'plain', 'utf-8'
    )
    html_part = MIMEText(html_body, 'html', 'utf-8')

    msg.attach(text_part)
    msg.attach(html_part)

    try:
        # 清理密碼（去除空格，Gmail App Password 常複製到帶空格）
        clean_password = app_password.replace(' ', '').strip()

        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.ehlo()
            server.starttls()
            server.ehlo()
            server.login(sender_email.strip(), clean_password)
            server.sendmail(sender_email.strip(), recipient_email.strip(), msg.as_string())
        return f"✅ 已成功寄送至 {recipient_email}"
    except smtplib.SMTPAuthenticationError as e:
        return (
            f"❌ Gmail 驗證失敗: {e}\n\n"
            "請確認：\n"
            "1. 已開啟 Google 兩步驟驗證\n"
            "2. 使用的是「應用程式密碼」(16碼) 而非一般密碼\n"
            "3. 到 https://myaccount.google.com/apppasswords 產生密碼"
        )
    except smtplib.SMTPException as e:
        return f"❌ 寄送失敗: {e}"
    except Exception as e:
        return f"❌ 未預期錯誤: {e}"


def _build_html_report(
    owner_name: str,
    year: int,
    month: int,
    owner_df: pd.DataFrame,
    summary_df: pd.DataFrame,
    category_df: pd.DataFrame,
) -> str:
    """產生 HTML 格式的消費分析報告"""

    period = f"{year}年{month}月" if month > 0 else f"{year}年度"
    total = owner_df['清算消費金額'].sum()
    count = len(owner_df)
    today = date.today().strftime('%Y-%m-%d')

    # 類別統計
    cat_stats = owner_df.groupby('消費類別')['清算消費金額'].agg(['sum', 'count'])
    cat_rows = ""
    for cat, row in cat_stats.iterrows():
        cat_rows += f"""
        <tr>
            <td style="padding:8px 12px;border-bottom:1px solid #eee;">{cat}</td>
            <td style="padding:8px 12px;border-bottom:1px solid #eee;text-align:right;">${row['sum']:,.0f}</td>
            <td style="padding:8px 12px;border-bottom:1px solid #eee;text-align:center;">{int(row['count'])} 筆</td>
        </tr>"""

    # 回饋計算
    convenience_total = 0
    general_total = 0
    if '消費類別' in owner_df.columns:
        convenience_total = owner_df[owner_df['消費類別'] == '四大超商']['清算消費金額'].sum()
        general_total = owner_df[owner_df['消費類別'] == '一般']['清算消費金額'].sum()

    convenience_reward = min(convenience_total * 0.10, 200)
    general_reward = general_total * 0.01
    total_reward = convenience_reward + general_reward

    # 交易明細表
    detail_rows = ""
    display_df = owner_df.sort_values('消費日期', ascending=False)
    for _, row in display_df.iterrows():
        txn_date = row['消費日期']
        if hasattr(txn_date, 'strftime'):
            txn_date = txn_date.strftime('%Y-%m-%d')
        detail_rows += f"""
        <tr>
            <td style="padding:6px 10px;border-bottom:1px solid #f0f0f0;font-size:13px;">{txn_date}</td>
            <td style="padding:6px 10px;border-bottom:1px solid #f0f0f0;font-size:13px;">{row['消費明細']}</td>
            <td style="padding:6px 10px;border-bottom:1px solid #f0f0f0;font-size:13px;text-align:right;">${row['清算消費金額']:,.0f}</td>
            <td style="padding:6px 10px;border-bottom:1px solid #f0f0f0;font-size:13px;text-align:center;">{row['消費類別']}</td>
        </tr>"""

    # 月度摘要表
    summary_rows = ""
    if not summary_df.empty:
        for _, row in summary_df.iterrows():
            m = int(row['結算月份'])
            alan_val = row.get('Alan', 0)
            lydia_val = row.get('Lydia', 0)
            total_val = row.get('總金額', 0)
            highlight = ' style="background:#fff3cd;"' if month > 0 and m == month else ''
            summary_rows += f"""
            <tr{highlight}>
                <td style="padding:6px 10px;border-bottom:1px solid #eee;text-align:center;">{m}月</td>
                <td style="padding:6px 10px;border-bottom:1px solid #eee;text-align:right;">${alan_val:,.0f}</td>
                <td style="padding:6px 10px;border-bottom:1px solid #eee;text-align:right;">${lydia_val:,.0f}</td>
                <td style="padding:6px 10px;border-bottom:1px solid #eee;text-align:right;font-weight:bold;">${total_val:,.0f}</td>
            </tr>"""

    # 商家排行 Top 10
    merchant_ranking = (
        owner_df.groupby('消費明細')['清算消費金額']
        .agg(['sum', 'count'])
        .rename(columns={'sum': '總金額', 'count': '次數'})
        .sort_values('總金額', ascending=False)
        .head(10)
        .reset_index()
    )
    merchant_ranking['平均每次'] = (
        merchant_ranking['總金額'] / merchant_ranking['次數']
    ).round(0)

    merchant_rows = ""
    for rank, (_, row) in enumerate(merchant_ranking.iterrows(), 1):
        medal = "🥇" if rank == 1 else "🥈" if rank == 2 else "🥉" if rank == 3 else f"{rank}"
        merchant_rows += f"""
        <tr>
            <td style="padding:6px 10px;border-bottom:1px solid #f0f0f0;text-align:center;font-size:13px;">{medal}</td>
            <td style="padding:6px 10px;border-bottom:1px solid #f0f0f0;font-size:13px;">{row['消費明細']}</td>
            <td style="padding:6px 10px;border-bottom:1px solid #f0f0f0;text-align:right;font-size:13px;font-weight:bold;">${row['總金額']:,.0f}</td>
            <td style="padding:6px 10px;border-bottom:1px solid #f0f0f0;text-align:center;font-size:13px;">{int(row['次數'])} 次</td>
            <td style="padding:6px 10px;border-bottom:1px solid #f0f0f0;text-align:right;font-size:13px;">${row['平均每次']:,.0f}</td>
        </tr>"""

    html = f"""
    <!DOCTYPE html>
    <html>
    <head><meta charset="utf-8"></head>
    <body style="font-family:'Microsoft JhengHei','Segoe UI',Arial,sans-serif;background:#f5f5f5;margin:0;padding:20px;">
    <div style="max-width:700px;margin:0 auto;background:white;border-radius:12px;overflow:hidden;box-shadow:0 2px 8px rgba(0,0,0,0.1);">

        <!-- Header -->
        <div style="background:linear-gradient(135deg,#667eea 0%,#764ba2 100%);padding:30px;text-align:center;">
            <h1 style="color:white;margin:0;font-size:24px;">💳 信用卡消費分析報告</h1>
            <p style="color:rgba(255,255,255,0.9);margin:10px 0 0;font-size:16px;">{period} — {owner_name}</p>
        </div>

        <!-- KPI Cards -->
        <div style="padding:20px;display:flex;gap:12px;">
            <div style="flex:1;background:#e8f5e9;border-radius:8px;padding:16px;text-align:center;">
                <div style="font-size:12px;color:#666;">總消費</div>
                <div style="font-size:22px;font-weight:bold;color:#2e7d32;">${total:,.0f}</div>
            </div>
            <div style="flex:1;background:#e3f2fd;border-radius:8px;padding:16px;text-align:center;">
                <div style="font-size:12px;color:#666;">交易筆數</div>
                <div style="font-size:22px;font-weight:bold;color:#1565c0;">{count} 筆</div>
            </div>
            <div style="flex:1;background:#fff3e0;border-radius:8px;padding:16px;text-align:center;">
                <div style="font-size:12px;color:#666;">預估回饋</div>
                <div style="font-size:22px;font-weight:bold;color:#e65100;">${total_reward:,.0f}</div>
            </div>
        </div>

        <!-- 類別統計 -->
        <div style="padding:0 20px 20px;">
            <h2 style="font-size:16px;color:#333;border-bottom:2px solid #667eea;padding-bottom:8px;">📊 消費類別統計</h2>
            <table style="width:100%;border-collapse:collapse;">
                <tr style="background:#f8f9fa;">
                    <th style="padding:8px 12px;text-align:left;">類別</th>
                    <th style="padding:8px 12px;text-align:right;">金額</th>
                    <th style="padding:8px 12px;text-align:center;">筆數</th>
                </tr>
                {cat_rows}
                <tr style="background:#f0f0f0;font-weight:bold;">
                    <td style="padding:8px 12px;">合計</td>
                    <td style="padding:8px 12px;text-align:right;">${total:,.0f}</td>
                    <td style="padding:8px 12px;text-align:center;">{count} 筆</td>
                </tr>
            </table>
        </div>

        <!-- 回饋明細 -->
        <div style="padding:0 20px 20px;">
            <h2 style="font-size:16px;color:#333;border-bottom:2px solid #667eea;padding-bottom:8px;">🎁 回饋金額</h2>
            <table style="width:100%;border-collapse:collapse;">
                <tr>
                    <td style="padding:8px 12px;border-bottom:1px solid #eee;">四大超商 (10%, 上限$200)</td>
                    <td style="padding:8px 12px;border-bottom:1px solid #eee;text-align:right;">消費 ${convenience_total:,.0f}</td>
                    <td style="padding:8px 12px;border-bottom:1px solid #eee;text-align:right;color:#e65100;font-weight:bold;">回饋 ${convenience_reward:,.0f}</td>
                </tr>
                <tr>
                    <td style="padding:8px 12px;border-bottom:1px solid #eee;">一般消費 (1%)</td>
                    <td style="padding:8px 12px;border-bottom:1px solid #eee;text-align:right;">消費 ${general_total:,.0f}</td>
                    <td style="padding:8px 12px;border-bottom:1px solid #eee;text-align:right;color:#e65100;font-weight:bold;">回饋 ${general_reward:,.0f}</td>
                </tr>
                <tr style="background:#fff3e0;font-weight:bold;">
                    <td style="padding:8px 12px;" colspan="2">回饋合計</td>
                    <td style="padding:8px 12px;text-align:right;color:#e65100;">${total_reward:,.0f}</td>
                </tr>
            </table>
        </div>

        <!-- 月度摘要 -->
        {"" if summary_df.empty else f'''
        <div style="padding:0 20px 20px;">
            <h2 style="font-size:16px;color:#333;border-bottom:2px solid #667eea;padding-bottom:8px;">📅 月度摘要</h2>
            <table style="width:100%;border-collapse:collapse;">
                <tr style="background:#f8f9fa;">
                    <th style="padding:6px 10px;text-align:center;">月份</th>
                    <th style="padding:6px 10px;text-align:right;">Alan</th>
                    <th style="padding:6px 10px;text-align:right;">Lydia</th>
                    <th style="padding:6px 10px;text-align:right;">總金額</th>
                </tr>
                {summary_rows}
            </table>
        </div>
        '''}

        <!-- 交易明細 -->
        <div style="padding:0 20px 20px;">
            <h2 style="font-size:16px;color:#333;border-bottom:2px solid #667eea;padding-bottom:8px;">📋 消費明細</h2>
            <table style="width:100%;border-collapse:collapse;">
                <tr style="background:#f8f9fa;">
                    <th style="padding:6px 10px;text-align:left;font-size:13px;">日期</th>
                    <th style="padding:6px 10px;text-align:left;font-size:13px;">明細</th>
                    <th style="padding:6px 10px;text-align:right;font-size:13px;">金額</th>
                    <th style="padding:6px 10px;text-align:center;font-size:13px;">類別</th>
                </tr>
                {detail_rows}
            </table>
        </div>

        <!-- 商家排行 -->
        <div style="padding:0 20px 20px;">
            <h2 style="font-size:16px;color:#333;border-bottom:2px solid #667eea;padding-bottom:8px;">🏪 消費商家排行 Top 10</h2>
            <table style="width:100%;border-collapse:collapse;">
                <tr style="background:#f8f9fa;">
                    <th style="padding:6px 10px;text-align:center;font-size:13px;">排名</th>
                    <th style="padding:6px 10px;text-align:left;font-size:13px;">商家</th>
                    <th style="padding:6px 10px;text-align:right;font-size:13px;">總金額</th>
                    <th style="padding:6px 10px;text-align:center;font-size:13px;">次數</th>
                    <th style="padding:6px 10px;text-align:right;font-size:13px;">平均/次</th>
                </tr>
                {merchant_rows}
            </table>
        </div>

        <!-- Footer -->
        <div style="background:#f8f9fa;padding:16px;text-align:center;font-size:12px;color:#999;">
            此報告由台新信用卡帳單分析系統自動產生 | {today}
        </div>
    </div>
    </body>
    </html>
    """
    return html
