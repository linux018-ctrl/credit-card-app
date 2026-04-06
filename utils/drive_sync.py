"""
Google Drive 同步模組 - 複用 budget_app 的 Service Account 憑證
負責：監控 Drive 資料夾、下載 PDF 帳單
"""

import io
import json
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
from google.oauth2 import service_account

SCOPES = ['https://www.googleapis.com/auth/drive.readonly']


def get_drive_service(credentials_json_bytes: bytes):
    """從 Service Account JSON 建立 Drive v3 API client"""
    info = json.load(io.BytesIO(credentials_json_bytes))
    creds = service_account.Credentials.from_service_account_info(
        info, scopes=SCOPES
    )
    return build('drive', 'v3', credentials=creds)


def list_pdf_files(service, folder_id: str) -> list:
    """列出指定 Drive 資料夾中的所有 PDF 檔案，依建立時間排序（最新在前）"""
    results = service.files().list(
        q=f"'{folder_id}' in parents and trashed=false",
        fields="files(id, name, mimeType, createdTime, modifiedTime)",
        orderBy="createdTime desc"
    ).execute()
    files = results.get('files', [])
    pdf_files = [f for f in files if f['name'].lower().endswith('.pdf')]
    return pdf_files


def download_file(service, file_id: str) -> bytes:
    """下載 Drive 檔案並回傳 bytes"""
    request = service.files().get_media(fileId=file_id)
    fh = io.BytesIO()
    downloader = MediaIoBaseDownload(fh, request)
    done = False
    while not done:
        _, done = downloader.next_chunk()
    fh.seek(0)
    return fh.read()


def get_all_pdfs(credentials_json_bytes: bytes, folder_id: str) -> list:
    """取得資料夾中所有 PDF 的清單（含 id, name, createdTime）"""
    service = get_drive_service(credentials_json_bytes)
    return list_pdf_files(service, folder_id)


def download_pdf(credentials_json_bytes: bytes, folder_id: str, file_id: str) -> bytes:
    """下載指定的 PDF 檔案"""
    service = get_drive_service(credentials_json_bytes)
    return download_file(service, file_id)


def get_latest_pdf(credentials_json_bytes: bytes, folder_id: str):
    """取得最新的 PDF 及其內容"""
    service = get_drive_service(credentials_json_bytes)
    files = list_pdf_files(service, folder_id)
    if not files:
        raise FileNotFoundError('Google Drive 資料夾中找不到任何 PDF 帳單')
    latest = files[0]
    pdf_bytes = download_file(service, latest['id'])
    return pdf_bytes, latest['name']
