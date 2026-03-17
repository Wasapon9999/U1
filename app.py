import streamlit as st
import pandas as pd
import os
import time
from io import BytesIO
from reportlab.lib.pagesizes import A4
from reportlab.platypus import (
    SimpleDocTemplate, Table, TableStyle,
    Image as RLImage, Paragraph, Spacer, PageBreak, KeepTogether
)
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from PIL import Image, ExifTags

# Google Drive API
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaIoBaseUpload
from google.oauth2 import service_account

# --- 1. การตั้งค่าหน้าเว็บและ Config ---
st.set_page_config(page_title="USO1-Report Manager", layout="wide")

# ✅ ID โฟลเดอร์ใน Shared Drive (จากลิงก์ที่คุณส่งมา)
GOOGLE_DRIVE_FOLDER_ID = '1-4OwgP-ODbelbtwSg5-m-rm4cyOTcW7O'

def get_drive_service():
    """เชื่อมต่อ Google Drive API"""
    try:
        if "gcp_service_account" in st.secrets:
            info = st.secrets["gcp_service_account"]
            creds = service_account.Credentials.from_service_account_info(
                info, scopes=['https://www.googleapis.com/auth/drive']
            )
        elif os.path.exists("service_account.json"):
            creds = service_account.Credentials.from_service_account_file(
                'service_account.json', scopes=['https://www.googleapis.com/auth/drive']
            )
        else: return None
        return build('drive', 'v3', credentials=creds)
    except Exception as e:
        st.error(f"⚠️ Drive Connection Error: {e}")
        return None

def init_fonts():
    try:
        pdfmetrics.registerFont(TTFont('THSarabun', 'THSarabunNew.ttf'))
        pdfmetrics.registerFont(TTFont('THSarabun-Bold', 'THSarabunNew Bold.ttf'))
        return 'THSarabun', 'THSarabun-Bold'
    except: return 'Helvetica', 'Helvetica-Bold'

F_REG, F_BOLD = init_fonts()

# --- 2. Google Drive Helpers (Pagination & Fuzzy Matching) ---

def normalize_filename(name):
    """ล้างชื่อไฟล์ให้เหลือแค่ตัวเลขและอักษรสำคัญเพื่อใช้เปรียบเทียบ"""
    if not name or pd.isna(name) or str(name).strip() in ["0", "nan", ""]: return None
    base = os.path.splitext(str(name).strip().lower())[0]
    return base.replace("__", "_").replace(" ", "").replace("-", "")

@st.cache_data(ttl=600) # แคชไว้ 10 นาทีเพราะไฟล์เยอะ
def get_all_files_in_drive():
    """กวาดรายชื่อไฟล์ทั้งหมดใน Shared Drive (รองรับหลักหมื่นไฟล์)"""
    service = get_drive_service()
    if not service: return []
    
    all_files = []
    page_token = None
    try:
        while True:
            results = service.files().list(
                q=f"'{GOOGLE_DRIVE_FOLDER_ID}' in parents and trashed = false",
                fields="nextPageToken, files(id, name)",
                pageSize=1000,
                supportsAllDrives=True,
                includeItemsFromAllDrives=True,
                pageToken=page_token
            ).execute()
            all_files.extend(results.get('files', []))
            page_token = results.get('nextPageToken')
            if not page_token: break
    except Exception as e:
        st.error(f"Error fetching file list: {e}")
    return all_files

def download_image_from_drive(file_name):
    """ค้นหาและดาวน์โหลดรูป (Smart Search)"""
    all_items = get_all_files_in_drive()
    search_target = normalize_filename(file_name)
    if not all_items or not search_target: return None
    
    target_id = None
    # 1. หาแบบเป๊ะๆ (หลัง Normalize)
    for item in all_items:
        if normalize_filename(item['name']) == search_target:
            target_id = item['id']
            break
            
    # 2. ถ้าไม่เจอ ลองหาแบบ Fuzzy (ใช้รหัสวันที่ส่วนท้าย เช่น _200326_01)
    if not target_id:
        # ดึงส่วนท้ายมาลองหา (เช่น 200326_01)
        suffix = search_target.split("_")[-1] if "_" in search_target else search_target
        for item in all_items:
            if suffix in normalize_filename(item['name']):
                target_id = item['id']
                break

    if target_id:
        try:
            service = get_drive_service()
            request = service.files().get_media(fileId=target_id)
            fh = BytesIO()
            downloader = MediaIoBaseDownload(fh, request)
            done = False
            while not done:
                _, done = downloader.next_chunk()
            fh.seek(0)
            return fh
        except: pass
    return None

def upload_image_to_drive(file_name, content_bytes):
    """อัปโหลดรูปใหม่เข้า Shared Drive"""
    service = get_drive_service()
    if not service: return
    try:
        clean_name = str(file_name).strip()
        file_metadata = {'name': clean_name, 'parents': [GOOGLE_DRIVE_FOLDER_ID]}
        media = MediaIoBaseUpload(BytesIO(content_bytes), mimetype='image/jpeg', resumable=True)
        service.files().create(body=file_metadata, media_body=media, supportsAllDrives=True).execute()
        st.cache_data.clear() # ล้างแคชเพื่อให้เห็นไฟล์ใหม่
    except Exception as e:
        st.error(f"❌ อัปโหลดล้มเหลว: {str(e)}")

# --- 3. UI & Business Logic ---

if 'main_df' not in st.session_state:
    try:
        st.session_state.main_df = pd.read_csv("03-2026.csv").fillna("")
    except:
        st.error("❌ ไม่พบไฟล์ 03-2026.csv")
        st.stop()

st.sidebar.title("เมนู")
centers = st.session_state.main_df['file_name'].unique()
sel_center = st.sidebar.selectbox("เลือกศูนย์", centers)

# ตรวจสอบไฟล์ใน Shared Drive (Debug)
if st.sidebar.checkbox("🔍 ตรวจสอบไฟล์ใน Shared Drive"):
    with st.sidebar.status("กำลังอ่านรายชื่อไฟล์หลักหมื่น..."):
        files = get_all_files_in_drive()
    st.sidebar.write(f"พบทั้งหมด {len(files)} ไฟล์")
    # แสดงแค่ 50 ไฟล์แรก
    st.sidebar.code("\n".join([f['name'] for f in files[:50]]))

if st.sidebar.button("💾 บันทึก CSV", width='stretch'):
    st.session_state.main_df.to_csv("03-2026.csv", index=False)
    st.sidebar.success("บันทึกสำเร็จ!")

df_idx = st.session_state.main_df[st.session_state.main_df['file_name'] == sel_center].index

for idx in df_idx:
    row = st.session_state.main_df.loc[idx]
    with st.expander(f"📅 {row['date']} - {row['name']}"):
        c = st.columns([2, 2, 1, 1])
        st.session_state.main_df.at[idx, 'name'] = c[0].text_input("ชื่อ", row['name'], key=f"n_{idx}")
        st.session_state.main_df.at[idx, 'status'] = c[1].text_input("ตำแหน่ง", row['status'], key=f"s_{idx}")
        st.session_state.main_df.at[idx, 'time_in'] = c[2].text_input("เข้า", row['time_in'], key=f"i_{idx}")
        st.session_state.main_df.at[idx, 'time_out'] = c[3].text_input("ออก", row['time_out'], key=f"o_{idx}")

        c_img = st.columns(2)
        for i, col in enumerate(["img_in1", "img_out1"]):
            img_val = str(row[col])
            img_d = download_image_from_drive(img_val)
            
            if img_d:
                c_img[i].image(img_d, caption=f"Found: {img_val}", use_container_width=True)
            else:
                c_img[i].warning(f"❌ ไม่พบรูป: {img_val}")
            
            new_f = c_img[i].file_uploader(f"เปลี่ยนรูป {col}", type=['jpg','png','jpeg'], key=f"u_{col}_{idx}")
            if new_f:
                with st.spinner("Uploading..."):
                    upload_image_to_drive(img_val, new_f.getbuffer())
                    st.toast("อัปโหลดสำเร็จ! กำลังรีเฟรช...")
                    time.sleep(1)
                    st.rerun()

st.divider()
if st.button("🖨️ ออกรายงาน PDF", width='stretch', type="primary"):
    # (โค้ดส่วน PDF Generator เดิม)
    st.info("ระบบกำลังสร้าง PDF...")
