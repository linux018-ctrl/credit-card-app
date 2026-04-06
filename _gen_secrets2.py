import json, os

cred_path = os.path.join(os.path.dirname(__file__), '..', 'budget_app', 'credentials.json')
d = json.load(open(cred_path))
out_path = os.path.join(os.path.dirname(__file__), 'secrets_template.toml')

lines = []
lines.append('gdrive_folder_id = "請貼上你的Drive資料夾ID"')
lines.append('')
lines.append('[gdrive_credentials]')
for k, v in d.items():
    if k == 'private_key':
        lines.append(f'{k} = """')
        lines.append(v.strip())
        lines.append('"""')
    else:
        lines.append(f'{k} = "{v}"')

lines.append('')
lines.append('pdf_password = "請填入身分證字號"')
lines.append('')
lines.append('email_sender = "請填入Gmail"')
lines.append('email_app_password = "請填入16碼應用程式密碼"')
lines.append('email_recipient_alan = "請填入Alan的Email"')
lines.append('email_recipient_lydia = "請填入Lydia的Email"')

with open(out_path, 'w', encoding='utf-8') as f:
    f.write('\n'.join(lines))

print(f"OK -> {out_path}")
