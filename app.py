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

# ✅ รหัส Folder ใน Shared Drive
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
    """ลงทะเบียนฟอนต์ไทย"""
    try:
        pdfmetrics.registerFont(TTFont('THSarabun', 'THSarabunNew.ttf'))
        pdfmetrics.registerFont(TTFont('THSarabun-Bold', 'THSarabunNew Bold.ttf'))
        return 'THSarabun', 'THSarabun-Bold'
    except: return 'Helvetica', 'Helvetica-Bold'

F_REG, F_BOLD = init_fonts()

# --- 2. Google Drive Helpers (Shared Drive & Pagination) ---

def normalize_filename(name):
    """ล้างชื่อไฟล์ให้เหลือแค่ตัวเลขและอักษรสำคัญเพื่อใช้เปรียบเทียบ"""
    if not name or pd.isna(name) or str(name).strip() in ["0", "nan", ""]: return None
    base = os.path.splitext(str(name).strip().lower())[0]
    # ล้างอักขระพิเศษเพื่อทำ Fuzzy Matching
    return base.replace("__", "_").replace(" ", "").replace("-", "").replace(".", "")

@st.cache_data(ttl=600)
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
    # 1. หาแบบชื่อตรงกันเป๊ะ
    for item in all_items:
        if normalize_filename(item['name']) == search_target:
            target_id = item['id']
            break
            
    # 2. ถ้าไม่เจอ หาแบบ "รหัสวันที่ส่วนท้าย" (Fuzzy Match)
    if not target_id:
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
        # ตรวจสอบและลบไฟล์เก่าถ้าชื่อซ้ำ
        results = service.files().list(
            q=f"'{GOOGLE_DRIVE_FOLDER_ID}' in parents and name = '{clean_name}'",
            supportsAllDrives=True, includeItemsFromAllDrives=True
        ).execute()
        for f in results.get('files', []):
            service.files().delete(fileId=f['id'], supportsAllDrives=True).execute()

        file_metadata = {'name': clean_name, 'parents': [GOOGLE_DRIVE_FOLDER_ID]}
        media = MediaIoBaseUpload(BytesIO(content_bytes), mimetype='image/jpeg', resumable=True)
        service.files().create(body=file_metadata, media_body=media, supportsAllDrives=True).execute()
        st.cache_data.clear() # ล้างแคชเพื่อให้เห็นไฟล์ใหม่ทันที
    except Exception as e:
        st.error(f"❌ อัปโหลดล้มเหลว: {str(e)}")

# --- 3. PDF Generator ---

def generate_pdf_original_style(df, center_name):
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=20, leftMargin=20, topMargin=25, bottomMargin=15)
    styles = {
        "Normal": ParagraphStyle("N", fontName=F_REG, fontSize=14, leading=18, alignment=1),
        "Title": ParagraphStyle("T", fontName=F_BOLD, fontSize=18, leading=24, alignment=1),
        "Heading2": ParagraphStyle("H2", fontName=F_BOLD, fontSize=14, leading=20, alignment=1),
        "H_Table": ParagraphStyle("HT", fontName=F_BOLD, fontSize=10, leading=11, alignment=1),
        "C_Table": ParagraphStyle("CT", fontName=F_REG, fontSize=10, leading=11, alignment=1),
    }
    story = []
    story.append(Paragraph("รายงานเวลาปฏิบัติงาน USO1-Renew", styles["Title"]))
    story.append(Paragraph(f"ศูนย์ : {center_name}", styles["Title"]))
    
    # ตารางสรุป
    t_data = [[Paragraph(h, styles["H_Table"]) for h in ["ลำดับ", "วันที่", "ชื่อ - นามสกุล", "เวลาเข้า", "เวลาออก", "ตำแหน่ง", "หมายเหตุ"]]]
    for i, row in df.iterrows():
        t_data.append([
            Paragraph(str(i+1), styles["C_Table"]), Paragraph(str(row['date']), styles["C_Table"]),
            Paragraph(str(row['name']), styles["C_Table"]), Paragraph(str(row['time_in']), styles["C_Table"]),
            Paragraph(str(row['time_out']), styles["C_Table"]), Paragraph(str(row['status']), styles["C_Table"]),
            Paragraph("", styles["C_Table"])
        ])
    tbl = Table(t_data, colWidths=[35, 90, 130, 60, 60, 80, 75], repeatRows=1)
    tbl.setStyle(TableStyle([('GRID',(0,0),(-1,-1),0.5,colors.black),('ALIGN',(0,0),(-1,-1),'CENTER')]))
    story.append(tbl)

    # รายละเอียดรูปภาพรายวัน
    for _, r in df.iterrows():
        story.append(PageBreak())
        story.append(Paragraph(f"วันที่ : {r['date']}", styles["Heading2"]))
        story.append(Paragraph(f"ชื่อ : {r['name']} | ตำแหน่ง : {r['status']}", styles["Normal"]))
        
        for label, col_img, col_time in [("เข้า (เช้า)", "img_in1", "time_in"), ("ออก (เย็น)", "img_out1", "time_out")]:
            img_b = download_image_from_drive(r[col_img])
            if img_b:
                try:
                    with Image.open(img_b) as p_img:
                        p_img = p_img.convert('RGB')
                        t_io = BytesIO()
                        p_img.save(t_io, format="JPEG", quality=80)
                        t_io.seek(0)
                        im = RLImage(t_io)
                        im._restrictSize(350, 280)
                        story.append(im)
                except: pass
            story.append(Paragraph(f"เวลา{label} : {r[col_time]}", styles["Normal"]))
            story.append(Spacer(1, 15))

    doc.build(story)
    return buffer.getvalue()

# --- 4. Main UI ---

st.title("🚀 USO1 Report Manager (Shared Drive)")

if 'main_df' not in st.session_state:
    try:
        st.session_state.main_df = pd.read_csv("03-2026.csv").fillna("")
    except:
        st.error("❌ ไม่พบไฟล์ 03-2026.csv")
        st.stop()

st.sidebar.title("เมนู")
centers = st.session_state.main_df['file_name'].unique()
sel_center = st.sidebar.selectbox("เลือกศูนย์", centers)

# Debug ตรวจสอบไฟล์ใน Drive
if st.sidebar.checkbox("🔍 ตรวจสอบไฟล์ใน Shared Drive"):
    with st.sidebar.status("กำลังอ่านรายชื่อไฟล์..."):
        files = get_all_files_in_drive()
    st.sidebar.write(f"พบทั้งหมด {len(files)} ไฟล์")
    st.sidebar.code("\n".join([f['name'] for f in files[:50]]) + "\n...")

if st.sidebar.button("💾 บันทึก CSV", use_container_width=True):
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
                c_img[i].image(img_d, caption=f"Drive: {img_val}", use_container_width=True)
            else:
                c_img[i].warning(f"❌ ไม่พบรูป: {img_val}")
            
            # --- ส่วนแก้ไขป้องกัน Infinite Loop ---
            new_f = c_img[i].file_uploader(f"เลือกรูปใหม่ {col}", type=['jpg','png','jpeg'], key=f"u_{col}_{idx}")
            
            if new_f is not None:
                if c_img[i].button(f"ยืนยันอัปโหลด {col}", key=f"btn_{col}_{idx}"):
                    with st.spinner("กำลังประมวลผล..."):
                        # ถ้าชื่อเดิมว่าง ให้ตั้งชื่อตามศูนย์_วันที่
                        target_name = img_val if img_val not in ["0", "nan", ""] else f"{sel_center}_{idx}_{col}.jpg"
                        upload_image_to_drive(target_name, new_f.getbuffer())
                        st.session_state.main_df.at[idx, col] = target_name
                        st.success("สำเร็จ!")
                        time.sleep(1)
                        st.rerun()

st.divider()
if st.button("🖨️ ออกรายงาน PDF", use_container_width=True, type="primary"):
    with st.spinner("กำลังสร้างไฟล์..."):
        pdf = generate_pdf_original_style(st.session_state.main_df.loc[df_idx], sel_center)
        st.download_button("📥 ดาวน์โหลด PDF", pdf, f"{sel_center}.pdf", "application/pdf", use_container_width=True)
