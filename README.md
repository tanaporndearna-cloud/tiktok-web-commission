# TRC Motorsport — TikTok Commission Calculator

Streamlit web app สำหรับคำนวณค่าคอมมิชชั่น TikTok รายเดือน

## วิธี Deploy บน Streamlit Cloud

### 1. ตั้งค่า Google Service Account

1. ไปที่ [Google Cloud Console](https://console.cloud.google.com)
2. สร้าง Service Account ใหม่
3. เปิดใช้งาน **Google Drive API**
4. ดาวน์โหลด JSON key ของ Service Account
5. Share Google Sheets STOCK (`1p4-DX-Sq-Mvo_FNAkYtGMkdEBvPDCuPM_sZ_b-_Tm-Y`) ให้ service account email (Viewer)
6. Share Google Drive folder output ให้ service account email (Editor)

### 2. Push code ขึ้น GitHub

```bash
git init
git add .
git commit -m "TRC TikTok Commission Calculator"
git remote add origin https://github.com/YOUR_USERNAME/trc-tiktok-commission.git
git push -u origin main
```

### 3. Deploy บน Streamlit Cloud

1. ไปที่ [share.streamlit.io](https://share.streamlit.io)
2. เชื่อมต่อ GitHub repo
3. ตั้งค่า Main file path: `app.py`
4. ไปที่ **Advanced settings → Secrets** แล้วใส่:

```toml
GOOGLE_CREDENTIALS = """
{ ... paste service account JSON ทั้งหมดที่นี่ ... }
"""
```

### 4. ใช้งาน

1. เปิด URL ของ Streamlit app
2. อัปโหลดไฟล์ทั้ง 4:
   - Data ERP (.xlsx)
   - ABBHTT (.xlsx) — อัปโหลดได้หลายไฟล์
   - TikTok TRC (.xlsx)
   - TikTok MAFIA (.xlsx)
3. กด **คำนวณค่าคอมมิชชั่น**
4. ดาวน์โหลดไฟล์ผลลัพธ์ หรือเปิดจาก Google Drive

## โครงสร้างไฟล์ Output

ไฟล์ `Data ERP YYMM.xlsx` มี **6 ชีท**:

| ชีท | เนื้อหา |
|-----|--------|
| ใบปะหน้า TIKTOK | สรุปยอดขาย + ค่าคอมมิชชั่น |
| เงินเข้าบริษัท | วิเคราะห์กำไร/ขาดทุนรายคำสั่งซื้อ |
| STOCK | ข้อมูล STOCK จาก Google Sheets |
| Data ERP | ข้อมูล ERP + lookup TikTok order# |
| ABBHTT | ข้อมูล ABBHTT รวม |
| TikTok | ข้อมูล TikTok orders รวม 2 ร้าน + SHOP QC |

## ข้อมูล Google Drive

- **STOCK:** `1p4-DX-Sq-Mvo_FNAkYtGMkdEBvPDCuPM_sZ_b-_Tm-Y` (ชีท 2026)
- **Output folder:** `10YIkvJliq0OZvxTfT-ZfzQHn9wVELQeO`
