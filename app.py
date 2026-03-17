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

# Folder ID จาก URL ที่คุณส่งมา
GOOGLE_DRIVE_FOLDER_ID = '1yO8M-5QIVRhVoDoLu2yaYAJo4csy1GdI'

def get_drive_service():
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

# --- 2. Google Drive Helpers (เน้นแก้เพื่อให้เห็นไฟล์ใน Shared Drive) ---

def normalize_filename(name):
    if not name or pd.isna(name) or str(name).strip() in ["0", "nan", "None", ""]: 
        return None
    base = os.path.splitext(str(name).strip().lower())[0]
    return base.replace("__", "_").replace(" ", "")

@st.cache_data(ttl=300)
def download_image_from_drive(file_name):
    service = get_drive_service()
    search_target = normalize_filename(file_name)
    if not service or not search_target: return None
    
    try:
        # ⚠️ จุดสำคัญ: ต้องใส่ supportsAllDrives และ includeItemsFromAllDrives ถึงจะเห็นไฟล์ใน Shared Drive
        results = service.files().list(
            q=f"'{GOOGLE_DRIVE_FOLDER_ID}' in parents and trashed = false",
            fields="files(id, name)",
            pageSize=1000,
            supportsAllDrives=True,              # ✅ ต้องมี
            includeItemsFromAllDrives=True       # ✅ ต้องมี
        ).execute()
        
        items = results.get('files', [])
        target_id = None
        for item in items:
            if normalize_filename(item['name']) == search_target:
                target_id = item['id']
                break
        
        if target_id:
            # ⚠️ ตอนดึงข้อมูลไฟล์ ก็ต้องระบุ supportsAllDrives=True ด้วย
            request = service.files().get_media(fileId=target_id)
            fh = BytesIO()
            downloader = MediaIoBaseDownload(fh, request)
            done = False
            while not done:
                _, done = downloader.next_chunk()
            fh.seek(0)
            return fh
    except Exception as e:
        # st.error(f"Debug Download: {e}") # เปิดเพื่อดู Error ถ้ายังไม่ได้
        pass
    return None

def upload_image_to_drive(file_name, content_bytes):
    service = get_drive_service()
    if not service: return
    try:
        clean_name = str(file_name).strip()
        norm_target = normalize_filename(clean_name)
        
        # ค้นหาเพื่อลบไฟล์เดิม (ต้องใส่พารามิเตอร์ Shared Drive)
        results = service.files().list(
            q=f"'{GOOGLE_DRIVE_FOLDER_ID}' in parents and trashed = false",
            supportsAllDrives=True,
            includeItemsFromAllDrives=True
        ).execute()
        
        for f in results.get('files', []):
            if normalize_filename(f['name']) == norm_target:
                service.files().delete(fileId=f['id'], supportsAllDrives=True).execute()
        
        # อัปโหลดไฟล์ใหม่เข้า Shared Drive
        file_metadata = {'name': clean_name, 'parents': [GOOGLE_DRIVE_FOLDER_ID]}
        media = MediaIoBaseUpload(BytesIO(content_bytes), mimetype='image/jpeg', resumable=True)
        service.files().create(
            body=file_metadata, 
            media_body=media, 
            supportsAllDrives=True # ✅ ต้องมีเพื่อให้บันทึกลง Shared Drive ได้
        ).execute()
        st.cache_data.clear()
    except Exception as e:
        st.error(f"❌ อัปโหลดล้มเหลว: {str(e)}")

# --- (ส่วน Utility และ UI คงเดิมจากโค้ดชุดล่าสุด) ---
# ... (ก๊อปปี้ส่วนที่เหลือจากโค้ดเดิมได้เลยครับ หรือใช้โค้ดด้านบนที่ผมให้ไว้ล่าสุด)
