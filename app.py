# -*- coding: utf-8 -*-
"""
TRC Motorsport — TikTok Commission Calculator
Streamlit Web App
"""

import streamlit as st
import tempfile
import os
import sys
import json
from pathlib import Path

# ── Path setup ──────────────────────────────────────────────────────────────
APP_DIR = Path(__file__).parent
sys.path.insert(0, str(APP_DIR))

from scripts.build_tiktok_erp import run_build
from scripts.build_tiktok_analysis import main as run_analysis
from scripts.drive_utils import get_drive_service, download_stock, upload_to_drive

# ── Page config ─────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="TRC TikTok Commission",
    page_icon="🏎️",
    layout="wide",
    initial_sidebar_state="collapsed",
)

OUTPUT_FOLDER_ID = "10YIkvJliq0OZvxTfT-ZfzQHn9wVELQeO"
DRIVE_FOLDER_URL = f"https://drive.google.com/drive/folders/{OUTPUT_FOLDER_ID}"

# ── Custom CSS ───────────────────────────────────────────────────────────────
st.markdown("""
<style>
    .main-title {
        font-size: 2rem;
        font-weight: 700;
        color: #1F3864;
        margin-bottom: 0.2rem;
    }
    .sub-title {
        color: #666;
        font-size: 0.95rem;
        margin-bottom: 1.5rem;
    }
    .section-header {
        background: #1F3864;
        color: white;
        padding: 0.5rem 1rem;
        border-radius: 6px;
        font-weight: 600;
        margin-bottom: 1rem;
    }
    .log-box {
        background: #1a1a2e;
        color: #00ff88;
        font-family: monospace;
        font-size: 0.8rem;
        padding: 1rem;
        border-radius: 8px;
        max-height: 300px;
        overflow-y: auto;
    }
    .status-ok  { color: #28a745; font-weight: 600; }
    .status-err { color: #dc3545; font-weight: 600; }
    div[data-testid="stFileUploader"] label { font-weight: 600; }
</style>
""", unsafe_allow_html=True)

# ── Header ───────────────────────────────────────────────────────────────────
st.markdown('<div class="main-title">🏎️ TRC Motorsport</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-title">TikTok Commission Calculator — คำนวณค่าคอมมิชชั่น TikTok รายเดือน</div>',
            unsafe_allow_html=True)
st.divider()

# ── Google Drive Credentials ─────────────────────────────────────────────────
def get_credentials():
    """อ่าน credentials จาก st.secrets หรือ environment variable"""
    try:
        # Streamlit Cloud: เก็บใน st.secrets["GOOGLE_CREDENTIALS"]
        raw = st.secrets["GOOGLE_CREDENTIALS"]
        if isinstance(raw, str):
            return json.loads(raw)
        return dict(raw)
    except Exception:
        pass
    try:
        # Fallback: environment variable
        raw = os.environ.get("GOOGLE_CREDENTIALS", "")
        if raw:
            return json.loads(raw)
    except Exception:
        pass
    return None


# ── Layout ───────────────────────────────────────────────────────────────────
col_left, col_right = st.columns([3, 2])

with col_left:
    st.markdown('<div class="section-header">📂 อัปโหลดไฟล์ Input</div>', unsafe_allow_html=True)

    erp_file = st.file_uploader(
        "1️⃣  Data ERP (.xlsx) — ไฟล์ ERP รายเดือน",
        type=["xlsx"],
        key="erp",
        help="ไฟล์ Excel export จากระบบ ERP รายเดือน (ชีทแรก)"
    )

    abbhtt_files = st.file_uploader(
        "2️⃣  ABBHTT (.xlsx) — อัปโหลดได้หลายไฟล์",
        type=["xlsx"],
        key="abbhtt",
        accept_multiple_files=True,
        help="ไฟล์ ABBHTT รายเดือน เช่น ABBHTT04.xlsx, ABBHTT06.xlsx"
    )

    tt_trc_file = st.file_uploader(
        "3️⃣  TikTok TRC (.xlsx) — ร้าน TRC Motorsport",
        type=["xlsx"],
        key="tt_trc",
        help="ไฟล์ Excel จาก TikTok Seller Center ร้าน TRC Motorsport (ชีท รายละเอียดคำสั่งซื้อ)"
    )

    tt_mafia_file = st.file_uploader(
        "4️⃣  TikTok MAFIA (.xlsx) — ร้าน MAFIA",
        type=["xlsx"],
        key="tt_mafia",
        help="ไฟล์ Excel จาก TikTok Seller Center ร้าน MAFIA (ชีท รายละเอียดคำสั่งซื้อ)"
    )

with col_right:
    st.markdown('<div class="section-header">⚙️ ตั้งค่า</div>', unsafe_allow_html=True)

    upload_gdrive = st.checkbox(
        "📤 อัปโหลดผลลัพธ์ไปยัง Google Drive",
        value=True,
        help="ส่งไฟล์ Excel ผลลัพธ์ไปยังโฟลเดอร์ที่กำหนด"
    )

    st.markdown(f"""
    <div style='background:#e8f4fd;border-radius:8px;padding:0.8rem 1rem;margin-top:0.5rem;font-size:0.85rem;'>
        <b>📁 โฟลเดอร์ปลายทาง:</b><br>
        <a href="{DRIVE_FOLDER_URL}" target="_blank" style='word-break:break-all;'>
            {DRIVE_FOLDER_URL}
        </a>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("")
    st.markdown("**📋 ไฟล์ที่ STOCK ดึงอัตโนมัติ:**")
    st.info("🔗 STOCK Google Sheets\n`1p4-DX-Sq-Mvo_FNAkYtGMkdEBvPDCuPM_sZ_b-_Tm-Y`\n\nดึงชีท `2026` คอลัมน์ A–U")

    # แสดงสถานะ credentials
    creds = get_credentials()
    if creds:
        st.success("✅ Google Credentials พร้อมใช้งาน")
    else:
        st.warning("⚠️ ไม่พบ Google Credentials\nกรุณาตั้งค่า `GOOGLE_CREDENTIALS` ใน Streamlit Secrets")

# ── Status bar ───────────────────────────────────────────────────────────────
all_uploaded = erp_file and abbhtt_files and tt_trc_file and tt_mafia_file

if all_uploaded:
    st.success(f"✅ ไฟล์ครบแล้ว: ERP + {len(abbhtt_files)} ABBHTT + TT-TRC + TT-MAFIA")
else:
    missing = []
    if not erp_file:        missing.append("Data ERP")
    if not abbhtt_files:    missing.append("ABBHTT")
    if not tt_trc_file:     missing.append("TikTok TRC")
    if not tt_mafia_file:   missing.append("TikTok MAFIA")
    st.warning(f"⏳ รอไฟล์: {', '.join(missing)}")

st.divider()

# ── Run button ───────────────────────────────────────────────────────────────
btn_col, _ = st.columns([1, 3])
with btn_col:
    run_btn = st.button(
        "🚀 คำนวณค่าคอมมิชชั่น",
        disabled=not all_uploaded,
        type="primary",
        use_container_width=True,
    )

# ── Processing ───────────────────────────────────────────────────────────────
if run_btn:
    log_placeholder  = st.empty()
    prog_placeholder = st.empty()
    logs = []

    def log(msg: str):
        logs.append(msg)
        log_placeholder.code("\n".join(logs[-60:]), language=None)

    def set_progress(pct: int, label: str):
        prog_placeholder.progress(pct, text=label)

    try:
        with tempfile.TemporaryDirectory() as tmpdir:
            tmp = Path(tmpdir)

            # ── บันทึกไฟล์ที่อัปโหลด ─────────────────────────────────
            set_progress(5, "💾 บันทึกไฟล์ที่อัปโหลด...")
            log("=" * 55)
            log("TRC Motorsport — TikTok Commission Calculator")
            log("=" * 55)
            log("")
            log("💾 บันทึกไฟล์ที่อัปโหลด...")

            erp_path = str(tmp / "Data_ERP.xlsx")
            with open(erp_path, "wb") as f:
                f.write(erp_file.getbuffer())
            log(f"  ✅ ERP: {erp_file.name}")

            abbhtt_paths = []
            for i, afile in enumerate(abbhtt_files):
                p = str(tmp / f"ABBHTT_{i}.xlsx")
                with open(p, "wb") as f:
                    f.write(afile.getbuffer())
                abbhtt_paths.append(p)
                log(f"  ✅ ABBHTT: {afile.name}")

            tt_trc_path = str(tmp / "TT-TRC.xlsx")
            with open(tt_trc_path, "wb") as f:
                f.write(tt_trc_file.getbuffer())
            log(f"  ✅ TT-TRC: {tt_trc_file.name}")

            tt_mafia_path = str(tmp / "TT-MAFIA.xlsx")
            with open(tt_mafia_path, "wb") as f:
                f.write(tt_mafia_file.getbuffer())
            log(f"  ✅ TT-MAFIA: {tt_mafia_file.name}")

            # ── ดาวน์โหลด STOCK ──────────────────────────────────────
            set_progress(15, "📥 ดาวน์โหลด STOCK จาก Google Drive...")
            log("")
            log("📥 ดาวน์โหลด STOCK จาก Google Drive...")

            creds = get_credentials()
            if not creds:
                st.error("❌ ไม่พบ Google Credentials — ไม่สามารถดาวน์โหลด STOCK ได้")
                st.stop()

            service = get_drive_service(creds)
            stock_path = str(tmp / "stock_full.xlsx")
            download_stock(service, stock_path, log_fn=log)

            # ── รัน Step 1: Build ERP ─────────────────────────────────
            set_progress(35, "⚙️ สร้างไฟล์ Excel (ขั้นตอนที่ 1/2)...")
            log("")
            log("⚙️ ขั้นตอนที่ 1 — สร้างไฟล์หลัก (4 ชีท)...")

            tmp_output = str(tmp / "Data ERP output.xlsx")
            result = run_build(
                erp_path=erp_path,
                abbhtt_paths=abbhtt_paths,
                tt_trc_path=tt_trc_path,
                tt_mafia_path=tt_mafia_path,
                stock_path=stock_path,
                output_path=tmp_output,
                log_fn=log,
            )

            # ตั้งชื่อไฟล์ output จริง
            auto_name = result.get("output_filename", "Data ERP output.xlsx")
            final_output = str(tmp / auto_name)
            if tmp_output != final_output:
                import shutil
                shutil.copy(tmp_output, final_output)

            log(f"\n  📄 ชื่อไฟล์ output: {auto_name}")
            log(f"  📊 TikTok rows: {result['tt_rows']}")
            log(f"  🔴 Added from ABBHTT: {result['added_red']} แถว")
            log(f"  🔵 Fixed shop name: {result['fixed_blue']} แถว")

            # ── รัน Step 2: Analysis ──────────────────────────────────
            set_progress(65, "📊 สร้างชีทวิเคราะห์ (ขั้นตอนที่ 2/2)...")
            log("")
            log("📊 ขั้นตอนที่ 2 — สร้างชีท เงินเข้าบริษัท + ใบปะหน้า TIKTOK...")

            analysis_result = run_analysis(final_output, log_fn=log)

            set_progress(85, "✅ ประมวลผลเสร็จสิ้น...")

            # ── อ่านไฟล์ผลลัพธ์ ──────────────────────────────────────
            with open(final_output, "rb") as f:
                output_bytes = f.read()

            # ── อัปโหลดไปยัง Google Drive ─────────────────────────────
            drive_link = None
            if upload_gdrive:
                set_progress(90, "📤 อัปโหลดไปยัง Google Drive...")
                log("")
                log("📤 อัปโหลดไปยัง Google Drive...")
                try:
                    uploaded = upload_to_drive(service, final_output, auto_name, log_fn=log)
                    drive_link = uploaded.get("webViewLink", "")
                except Exception as e:
                    log(f"  ⚠️ อัปโหลด Drive ล้มเหลว: {e}")

            set_progress(100, "🎉 เสร็จสมบูรณ์!")
            log("")
            log("=" * 55)
            log("🎉 เสร็จสมบูรณ์!")
            log("=" * 55)

        # ── แสดงผลลัพธ์ ───────────────────────────────────────────────
        prog_placeholder.empty()
        st.divider()
        st.markdown("### 📊 สรุปผล")

        res_col1, res_col2, res_col3, res_col4 = st.columns(4)
        with res_col1:
            st.metric("TikTok Orders", f"{result['tt_rows']:,}")
        with res_col2:
            st.metric("เพิ่มจาก ABBHTT 🔴", f"{result['added_red']:,}")
        with res_col3:
            st.metric("แก้ชื่อร้าน 🔵", f"{result['fixed_blue']:,}")
        with res_col4:
            st.metric("ชีท ERP rows", f"{analysis_result.get('rows', 0):,}")

        st.markdown("### 📥 ดาวน์โหลดผลลัพธ์")
        dl_col, link_col = st.columns([1, 1])

        with dl_col:
            st.download_button(
                label=f"⬇️  ดาวน์โหลด {auto_name}",
                data=output_bytes,
                file_name=auto_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True,
            )

        with link_col:
            if drive_link:
                st.success(f"✅ อัปโหลด Google Drive สำเร็จ")
                st.markdown(f"[📂 เปิดไฟล์ใน Google Drive]({drive_link})")
            elif upload_gdrive:
                st.warning("⚠️ อัปโหลด Google Drive ไม่สำเร็จ — ดาวน์โหลดจากปุ่มด้านซ้ายแทน")

        st.markdown(
            f"**ไฟล์มี 6 ชีท:** ใบปะหน้า TIKTOK · เงินเข้าบริษัท · STOCK · Data ERP · ABBHTT · TikTok"
        )
        st.balloons()

    except Exception as e:
        import traceback
        log(f"\n❌ ERROR: {e}")
        log(traceback.format_exc())
        st.error(f"❌ เกิดข้อผิดพลาด: {e}")
        with st.expander("🔍 รายละเอียด Error"):
            st.code(traceback.format_exc())

# ── Footer ───────────────────────────────────────────────────────────────────
st.divider()
st.markdown(
    "<small style='color:#aaa;'>TRC Motorsport Internal Tool · "
    "STOCK ดึงจาก Google Sheets อัตโนมัติ · "
    "ผลลัพธ์ส่งไปยัง Google Drive</small>",
    unsafe_allow_html=True,
)
