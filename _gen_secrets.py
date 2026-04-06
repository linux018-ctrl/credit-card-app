"""一次性腳本：從 credentials.json 產生 Streamlit Cloud secrets.toml 格式"""
import json, os

cred_path = os.path.join(os.path.dirname(__file__), '..', 'budget_app', 'credentials.json')
d = json.load(open(cred_path))

print("# ═══ 請將以下內容複製貼到 Streamlit Cloud → App Settings → Secrets ═══")
print()
print('# --- Google Drive ---')
print('gdrive_folder_id = "請貼上你的Drive資料夾ID"')
print()
print('[gdrive_credentials]')
for k, v in d.items():
    if k == 'private_key':
        print(f'{k} = """')
        print(v.strip())
        print('"""')
    else:
        print(f'{k} = "{v}"')

print()
print('# --- PDF 密碼 ---')
print('pdf_password = "請填入身分證字號"')
print()
print('# --- Email ---')
print('email_sender = "請填入Gmail"')
print('email_app_password = "請填入16碼應用程式密碼"')
print('email_recipient_alan = "請填入Alan的Email"')
print('email_recipient_lydia = "請填入Lydia的Email"')
