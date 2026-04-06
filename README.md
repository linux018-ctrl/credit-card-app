# 💳 台新信用卡帳單分析 App

自動解析台新銀行信用卡 PDF 帳單，產生消費分類、回饋計算、趨勢分析報告。

## ✨ 功能

- 🔄 **Google Drive 自動同步** — 從 Drive 資料夾自動下載 PDF 帳單
- 📄 **PDF 智能解析** — 支援加密台新帳單，座標式精準提取交易明細
- 🏷️ **自動分類** — 四大超商(10%回饋) / 一般消費(1%回饋)
- 👤 **Owner 判定** — 自動辨識 Alan / Lydia 的消費
- 📊 **月報總覽** — KPI、每月摘要表、消費圖表
- 🎁 **回饋分析** — 計算每月現金回饋、年度回饋率
- 📈 **趨勢分析** — 消費趨勢折線圖、商家排行
- 📧 **Email 報告** — 自動寄送 HTML 格式消費報告
- ☁️ **Streamlit Cloud** — 支援雲端部署

## 🚀 本地開發

```bash
pip install -r requirements.txt
streamlit run app.py --server.port 8502
```

## ☁️ Streamlit Cloud 部署

1. Push 到 GitHub
2. 前往 [share.streamlit.io](https://share.streamlit.io) 連結 GitHub repo
3. 設定 **Secrets** (在 App Settings → Secrets):

```toml
# Google Drive 設定
gdrive_folder_id = "你的Drive資料夾ID"

[gdrive_credentials]
type = "service_account"
project_id = "your-project-id"
private_key_id = "..."
private_key = "-----BEGIN RSA PRIVATE KEY-----\n...\n-----END RSA PRIVATE KEY-----\n"
client_email = "your-sa@project.iam.gserviceaccount.com"
client_id = "..."
auth_uri = "https://accounts.google.com/o/oauth2/auth"
token_uri = "https://oauth2.googleapis.com/token"

# PDF 密碼
pdf_password = "身分證字號"

# Email 設定
email_sender = "your_email@gmail.com"
email_app_password = "xxxxxxxxxxxxxxxx"
email_recipient_alan = "alan@example.com"
email_recipient_lydia = "lydia@example.com"
```

## 📁 專案結構

```
credit_card_app/
├── app.py              # Streamlit 主程式
├── requirements.txt    # Python 依賴
├── .streamlit/
│   └── config.toml     # Streamlit 設定
├── utils/
│   ├── pdf_parser.py   # PDF 解析引擎
│   ├── classifier.py   # 分類 & 回饋計算
│   ├── data_manager.py # 資料 CRUD
│   ├── drive_sync.py   # Google Drive 同步
│   ├── charts.py       # Plotly 圖表
│   └── email_sender.py # Gmail 報告寄送
└── data/
    └── records.json    # 交易資料 (gitignored)
```
