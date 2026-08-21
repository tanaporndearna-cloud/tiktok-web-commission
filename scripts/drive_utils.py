# -*- coding: utf-8 -*-
"""
drive_utils.py — Google Drive helper functions สำหรับ TikTok Commission App
"""

import io
import json
import os


STOCK_FILE_ID     = "1p4-DX-Sq-Mvo_FNAkYtGMkdEBvPDCuPM_sZ_b-_Tm-Y"
OUTPUT_FOLDER_ID  = "10YIkvJliq0OZvxTfT-ZfzQHn9wVELQeO"


def get_drive_service(credentials_info: dict):
    """
    สร้าง Google Drive API service จาก service account credentials
    credentials_info: dict ที่ได้จาก JSON ของ service account
    """
    from google.oauth2 import service_account
    from googleapiclient.discovery import build

    SCOPES = [
        'https://www.googleapis.com/auth/drive',
    ]
    creds = service_account.Credentials.from_service_account_info(
        credentials_info, scopes=SCOPES
    )
    return build('drive', 'v3', credentials=creds)


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
    """
    อัปโหลดไฟล์ไปยัง Google Drive folder OUTPUT_FOLDER_ID
    คืน dict ที่มี id และ webViewLink
    """
    from googleapiclient.http import MediaFileUpload

    def log(msg):
        print(msg)
        if log_fn:
            log_fn(msg)

    log(f"  Uploading {filename} → Drive folder {OUTPUT_FOLDER_ID}...")

    # ลบไฟล์เดิมที่ชื่อเดียวกัน (ถ้ามี) เพื่อป้องกันไฟล์ซ้ำ
    try:
        existing = service.files().list(
            q=f"name='{filename}' and '{OUTPUT_FOLDER_ID}' in parents and trashed=false",
            fields="files(id,name)"
        ).execute()
        for f in existing.get('files', []):
            service.files().delete(fileId=f['id']).execute()
            log(f"  🗑️ ลบไฟล์เดิม: {f['name']} ({f['id']})")
    except Exception as e:
        log(f"  ⚠️ ไม่สามารถตรวจสอบไฟล์เดิม: {e}")

    file_metadata = {
        'name': filename,
        'parents': [OUTPUT_FOLDER_ID]
    }
    media = MediaFileUpload(
        file_path,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        resumable=True
    )
    result = service.files().create(
        body=file_metadata,
        media_body=media,
        fields='id,webViewLink,name'
    ).execute()

    log(f"  ✅ อัปโหลดสำเร็จ: {result.get('webViewLink', '')}")
    return result
