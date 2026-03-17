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
from googleapiclient.http import MediaIoBaseDownload, MediaFileUpload
from google.oauth2 import service_account

# --- 1. การตั้งค่าหน้าเว็บและ Config ---
st.set_page_config(page_title="USO1-Report Manager", layout="wide")

# Folder ID ของคุณ
GOOGLE_DRIVE_FOLDER_ID = '0ANHEDviy1Mq0Uk9PVA'

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

# --- 2. Google Drive Helpers (Smart Matching Mode) ---

def normalize_filename(name):
    """ฟังก์ชันสำหรับล้างชื่อไฟล์ให้เปรียบเทียบกันได้ง่ายขึ้น"""
    if not name: return ""
    # ตัดนามสกุล, ทำเป็นตัวเล็ก, ตัดช่องว่าง, และยุบ __ ให้เหลือ _ อันเดียว
    base = os.path.splitext(str(name).strip().lower())[0]
    return base.replace("__", "_").replace(" ", "")

@st.cache_data(ttl=300)
def download_image_from_drive(file_name):
    service = get_drive_service()
    if not service or not file_name or pd.isna(file_name): return None
    try:
        # 1. เตรียมชื่อที่จะค้นหาจาก CSV (ล้างให้สะอาด)
        search_target = normalize_filename(file_name)
        
        # 2. ดึงรายชื่อไฟล์ทั้งหมดในโฟลเดอร์
        results = service.files().list(
            q=f"'{GOOGLE_DRIVE_FOLDER_ID}' in parents and trashed = false",
            fields="files(id, name)",
            pageSize=1000
        ).execute()
        items = results.get('files', [])
        
        target_id = None
        # 3. วนลูปเช็คชื่อไฟล์แบบยืดหยุ่น (ยุบขีดล่างเหมือนกันก่อนเทียบ)
        for item in items:
            drive_file_normalized = normalize_filename(item['name'])
            if drive_file_normalized == search_target:
                target_id = item['id']
                break
        
        # 4. ถ้ายังหาไม่เจอ ลองหาแบบ "มีชื่อนี้เป็นส่วนหนึ่ง" (Partial Match)
        if not target_id:
            for item in items:
                if search_target in normalize_filename(item['name']):
                    target_id = item['id']
                    break

        if target_id:
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
    service = get_drive_service()
    if not service: return
    clean_name = str(file_name).strip()
    norm_name = normalize_filename(clean_name)
    
    # ลบไฟล์เดิมที่มีชื่อหลักเดียวกัน (ไม่สนขีดล่างซ้ำ)
    results = service.files().list(q=f"'{GOOGLE_DRIVE_FOLDER_ID}' in parents and trashed = false").execute()
    for f in results.get('files', []):
        if normalize_filename(f['name']) == norm_name:
            service.files().delete(fileId=f['id']).execute()
    
    file_metadata = {'name': clean_name, 'parents': [GOOGLE_DRIVE_FOLDER_ID]}
    with open("temp.jpg", "wb") as f: f.write(content_bytes)
    media = MediaFileUpload("temp.jpg", mimetype='image/jpeg')
    service.files().create(body=file_metadata, media_body=media).execute()
    if os.path.exists("temp.jpg"): os.remove("temp.jpg")
    st.cache_data.clear()

# --- 3. Utility ---
def apply_exif_orientation(img):
    try:
        exif = img._getexif()
        if exif:
            for tag, value in exif.items():
                if ExifTags.TAGS.get(tag) == 'Orientation':
                    if value == 3: img = img.transpose(Image.ROTATE_180)
                    elif value == 6: img = img.transpose(Image.ROTATE_270)
                    elif value == 8: img = img.transpose(Image.ROTATE_90)
                    break
    except: pass
    return img

def fmt_time(t):
    if not t or pd.isna(t) or str(t).strip() == "": return ""
    t = str(t).strip().replace(".", ":")
    try:
        parts = t.split(":")
        return f"{int(parts[0]):02d}:{int(parts[1]):02d}"
    except: return t

def parse_thai_date_simple(s):
    m_thai = {1: "มกราคม", 2: "กุมภาพันธ์", 3: "มีนาคม", 4: "เมษายน", 5: "พฤษภาคม", 6: "มิถุนายน",
              7: "กรกฎาคม", 8: "สิงหาคม", 9: "กันยายน", 10: "ตุลาคม", 11: "พฤศจิกายน", 12: "ธันวาคม"}
    m_map = {"มกราคม": "01", "กุมภาพันธ์": "02", "มีนาคม": "03", "เมษายน": "04", "พฤษภาคม": "05", "มิถุนายน": "06",
             "กรกฎาคม": "07", "สิงหาคม": "08", "กันยายน": "09", "ตุลาคม": "10", "พฤศจิกายน": "11", "ธันวาคม": "12"}
    if not s or pd.isna(s): return pd.NaT, ""
    try:
        s_clean = str(s).strip()
        for k, v in m_map.items():
            if k in s_clean: s_clean = s_clean.replace(k, v)
        parts = s_clean.split()
        if len(parts) == 3:
            d, m, y = parts
            y_int = int(y)
            if y_int > 2500: y_int -= 543
            dt = pd.to_datetime(f"{y_int}-{m}-{d}")
            return dt, f"{d} {m_thai[int(m)]} {int(y)}"
    except: pass
    return pd.NaT, str(s)

# --- 4. PDF Generator ---
def generate_pdf_original_style(df, center_name):
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=20, leftMargin=20, topMargin=25, bottomMargin=15)
    styles = {
        "Normal": ParagraphStyle("N", fontName=F_REG, fontSize=14, leading=18, alignment=1),
        "Title": ParagraphStyle("T", fontName=F_BOLD, fontSize=18, leading=24, alignment=1),
        "Heading2": ParagraphStyle("H2", fontName=F_BOLD, fontSize=14, leading=20, alignment=1),
        "Signature": ParagraphStyle("S", fontName=F_REG, fontSize=14, leading=18, alignment=1),
        "H_Table": ParagraphStyle("HT", fontName=F_BOLD, fontSize=10, leading=11, alignment=1),
        "C_Table": ParagraphStyle("CT", fontName=F_REG, fontSize=10, leading=11, alignment=1),
    }
    story = []
    story.append(Paragraph("รายงานเวลาปฏิบัติงาน USO1-Renew", styles["Title"]))
    story.append(Paragraph(f"ศูนย์ : {center_name}", styles["Title"]))
    dt_f, d_str = parse_thai_date_simple(df.iloc[0]['date'])
    if pd.notna(dt_f): story.append(Paragraph(f"เดือน : {d_str.split(' ', 1)[1]}", styles["Heading2"]))
    emp = df["name"].loc[df["name"].str.strip() != ""].iloc[0] if not df["name"].empty else ""
    story.append(Paragraph(f"เจ้าหน้าที่ดูแลประจำศูนย์ : {emp}", styles["Heading2"]))
    story.append(Spacer(1, 10))
    t_data = [[Paragraph(h, styles["H_Table"]) for h in ["ลำดับ", "วันที่", "ชื่อ - นามสกุล", "เวลาเข้า", "เวลาออก", "ตำแหน่ง", "หมายเหตุ"]]]
    for i, row in df.iterrows():
        _, d_t = parse_thai_date_simple(row['date'])
        t_data.append([Paragraph(str(i+1), styles["C_Table"]), Paragraph(d_t, styles["C_Table"]), Paragraph(row['name'], styles["C_Table"]), Paragraph(fmt_time(row['time_in']), styles["C_Table"]), Paragraph(fmt_time(row['time_out']), styles["C_Table"]), Paragraph(row['status'], styles["C_Table"]), Paragraph("", styles["C_Table"])])
    tbl = Table(t_data, colWidths=[35, 100, 130, 60, 60, 80, 70], repeatRows=1)
    tbl.setStyle(TableStyle([('GRID',(0,0),(-1,-1),0.5,colors.black),('BACKGROUND',(0,0),(-1,0),colors.lightgrey),('ALIGN',(0,0),(-1,-1),'CENTER'),('VALIGN',(0,0),(-1,-1),'MIDDLE')]))
    story.append(tbl)
    story.append(Spacer(1, 30))
    sig_l = [Paragraph("....................................", styles["Signature"]), Spacer(1, 6), Paragraph(f"( {emp} )", styles["Signature"]), Paragraph("ผดล.ประจำศูนย์", styles["Signature"])]
    sig_r = [Paragraph("....................................", styles["Signature"]), Spacer(1, 6), Paragraph("( ...................................... )", styles["Signature"]), Paragraph("ตำแหน่ง_______________________", styles["Signature"])]
    story.append(KeepTogether(Table([[sig_l, sig_r]], colWidths=[260, 260])))
    for _, r in df.iterrows():
        story.append(PageBreak())
        _, d_t = parse_thai_date_simple(r['date'])
        story.append(Paragraph(f"วันที่ : {d_t}", styles["Heading2"]))
        story.append(Spacer(1, 12))
        story.append(Paragraph(f"ชื่อ : <b>{r['name']}</b> &nbsp; ตำแหน่ง : <b>{r['status']}</b>", styles["Normal"]))
        for label, col_img, col_time in [("เข้า (เช้า)", "img_in1", "time_in"), ("ออก (เย็น)", "img_out1", "time_out")]:
            img_b = download_image_from_drive(r[col_img])
            if img_b:
                try:
                    with Image.open(img_b) as p_img:
                        p_img = apply_exif_orientation(p_img)
                        t_io = BytesIO()
                        p_img.convert('RGB').save(t_io, format="JPEG", quality=85)
                        t_io.seek(0)
                        im = RLImage(t_io)
                        im._restrictSize(310, 260)
                        i_tbl = Table([[im]], colWidths=[450])
                        i_tbl.setStyle(TableStyle([('ALIGN',(0,0),(-1,-1),'CENTER')]))
                        story.append(i_tbl)
                except: pass
            story.append(Paragraph(f"เวลา{label} : <b>{fmt_time(r[col_time])}</b>", styles["Normal"]))
            story.append(Spacer(1, 18))
    doc.build(story)
    return buffer.getvalue()

# --- 5. Main UI ---

if 'main_df' not in st.session_state:
    st.session_state.main_df = pd.read_csv("03-2026.csv")

st.sidebar.title("เมนู")
centers = st.session_state.main_df['file_name'].unique()
sel_center = st.sidebar.selectbox("เลือกศูนย์", centers)

# ตรวจสอบไฟล์ใน Drive
if st.sidebar.checkbox("🔍 ตรวจสอบไฟล์ใน Drive"):
    service = get_drive_service()
    if service:
        res = service.files().list(q=f"'{GOOGLE_DRIVE_FOLDER_ID}' in parents and trashed = false", fields="files(name)").execute()
        files_in_drive = [f['name'] for f in res.get('files', [])]
        st.sidebar.write(f"ไฟล์ที่พบ ({len(files_in_drive)}):")
        st.sidebar.code("\n".join(files_in_drive))

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
            img_name = str(row[col])
            img_d = download_image_from_drive(img_name)
            if img_d:
                with Image.open(img_d) as im_disp:
                    im_disp = apply_exif_orientation(im_disp)
                    c_img[i].image(im_disp, caption=img_name, width='stretch')
            else:
                c_img[i].warning(f"❌ ไม่พบ {img_name}")

            new_f = c_img[i].file_uploader(f"เปลี่ยนรูป {col}", type=['jpg','png','jpeg'], key=f"u_{col}_{idx}")
            if new_f:
                with st.spinner("Uploading..."):
                    upload_image_to_drive(img_name, new_f.getbuffer())
                    st.toast("อัปโหลดสำเร็จ!")
                    time.sleep(1)
                    st.rerun()

st.divider()
if st.button("🖨️ ออกรายงาน PDF", width='stretch', type="primary"):
    with st.spinner("Generating..."):
        pdf = generate_pdf_original_style(st.session_state.main_df.loc[df_idx], sel_center)
        st.download_button("📥 ดาวน์โหลด PDF", pdf, f"{sel_center}.pdf", "application/pdf", width='stretch')
