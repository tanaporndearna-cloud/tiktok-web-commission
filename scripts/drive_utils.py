# -*- coding: utf-8 -*-
"""
drive_utils.py — Google Drive helper functions สำหรับ TikTok Commission App
"""

import io
import os

import streamlit as st

STOCK_FILE_ID     = "1p4-DX-Sq-Mvo_FNAkYtGMkdEBvPDCuPM_sZ_b-_Tm-Y"
OUTPUT_FOLDER_ID  = "10YIkvJliq0OZvxTfT-ZfzQHn9wVELQeO"
OAUTH_SCOPES      = ["https://www.googleapis.com/auth/drive"]


def get_drive_service(credentials_info: dict):
    """สร้าง Drive service จาก service account credentials"""
    from google.oauth2 import service_account
    from googleapiclient.discovery import build

    creds = service_account.Credentials.from_service_account_info(
        credentials_info, scopes=OAUTH_SCOPES
    )
    return build('drive', 'v3', credentials=creds)


def get_auto_drive_service():
    """ใช้ stored refresh_token จาก secrets — ไม่ต้อง login"""
    from google.oauth2.credentials import Credentials
    from google.auth.transport.requests import Request
    from googleapiclient.discovery import build

    try:
        oauth = {}
        if "gdrive_oauth" in st.secrets:
            oauth = dict(st.secrets["gdrive_oauth"])
        refresh_token = oauth.get("refresh_token", "")
        if refresh_token:
            creds = Credentials(
                token=None,
                refresh_token=refresh_token,
                token_uri="https://oauth2.googleapis.com/token",
                client_id=oauth["client_id"],
                client_secret=oauth["client_secret"],
                scopes=OAUTH_SCOPES,
            )
            creds.refresh(Request())
            return build('drive', 'v3', credentials=creds)
    except Exception:
        pass
    return None


def download_stock(service, output_path: str, log_fn=None) -> str:
    """
    ดาวน์โหลด Google Sheets STOCK เป็น .xlsx แล้วบันทึกที่ output_path
    คืน output_path
    """
    from googleapiclient.http import MediaIoBaseDownload

    def log(msg):
        print(msg)
        if log_fn:
            log_fn(msg)

    log(f"  Downloading STOCK (fileId={STOCK_FILE_ID})...")
    request = service.files().export_media(
        fileId=STOCK_FILE_ID,
        mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )
    fh = io.BytesIO()
    downloader = MediaIoBaseDownload(fh, request)
    done = False
    while not done:
        status, done = downloader.next_chunk()
        if status:
            log(f"  Download {int(status.progress() * 100)}%...")
    fh.seek(0)
    with open(output_path, 'wb') as f:
        f.write(fh.read())
    size_kb = os.path.getsize(output_path) // 1024
    log(f"  ✅ STOCK บันทึกที่ {output_path} ({size_kb} KB)")
    return output_path


def upload_to_drive(service, file_path: str, filename: str, log_fn=None) -> dict:
    """อัปโหลดไฟล์ไปยัง Google Drive — ใช้ OAuth user ถ้ามี refresh_token"""
    from googleapiclient.http import MediaFileUpload

    def log(msg):
        print(msg)
        if log_fn:
            log_fn(msg)

    # ใช้ OAuth service (user credentials) แทน service account เพื่อหลีกเลี่ยง quota error
    upload_service = get_auto_drive_service()
    if upload_service is None:
        upload_service = service  # fallback ถึง service account

    log(f"  Uploading {filename} → Drive folder {OUTPUT_FOLDER_ID}...")

    # ลบไฟล์เดิมที่ชื่อเดียวกัน (ถ้ามี)
    try:
        existing = upload_service.files().list(
            q=f"name='{filename}' and '{OUTPUT_FOLDER_ID}' in parents and trashed=false",
            fields="files(id,name)"
        ).execute()
        for f in existing.get('files', []):
            upload_service.files().delete(fileId=f['id']).execute()
            log(f"  🗑️ ลบไฟล์เดิม: {f['name']} ({f['id']})")
    except Exception as e:
        log(f"  ⚠️ ไม่สามารถตรวจสอบไฟล์เดิม: {e}")

    # ตัดนามสกุล .xlsx ออกจากชื่อไฟล์ (เพราะจะแปลงเป็น Google Sheets)
    sheets_name = filename.replace('.xlsx', '').replace('.XLSX', '')
    file_metadata = {
        'name': sheets_name,
        'parents': [OUTPUT_FOLDER_ID],
        'mimeType': 'application/vnd.google-apps.spreadsheet'
    }
    media = MediaFileUpload(
        file_path,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        resumable=True
    )
    result = upload_service.files().create(
        body=file_metadata,
        media_body=media,
        fields='id,webViewLink,name'
    ).execute()

    log(f"  ✅ อัปโหลดสำเร็จ: {result.get('webViewLink', '')}")
    return result
