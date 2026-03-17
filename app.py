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

# ✅ ID โฟลเดอร์ใน Shared Drive ของคุณ
GOOGLE_DRIVE_FOLDER_ID = '1yO8M-5QIVRhVoDoLu2yaYAJo4csy1GdI'

def get_drive_service():
    """เชื่อมต่อ Google Drive API ผ่าน Secrets"""
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
        else:
            return None
        return build('drive', 'v3', credentials=creds)
    except Exception as e:
        st.error(f"⚠️ Drive Connection Error: {e}")
        return None

def init_fonts():
    """ลงทะเบียนฟอนต์ภาษาไทย"""
    try:
        pdfmetrics.registerFont(TTFont('THSarabun', 'THSarabunNew.ttf'))
        pdfmetrics.registerFont(TTFont('THSarabun-Bold', 'THSarabunNew Bold.ttf'))
        return 'THSarabun', 'THSarabun-Bold'
    except:
        return 'Helvetica', 'Helvetica-Bold'

F_REG, F_BOLD = init_fonts()

# --- 2. Google Drive Helpers (Full Shared Drive & Smart Matching) ---

def normalize_filename(name):
    """ทำความสะอาดชื่อไฟล์เพื่อใช้เปรียบเทียบ (ตัดนามสกุล, ยุบขีดล่าง, ตัดช่องว่าง)"""
    if not name or pd.isna(name) or str(name).strip() in ["0", "nan", "None", ""]: 
        return None
    base = os.path.splitext(str(name).strip().lower())[0]
    return base.replace("__", "_").replace(" ", "")

@st.cache_data(ttl=300)
def download_image_from_drive(file_name):
    """ดาวน์โหลดรูปภาพจาก Shared Drive โดยใช้การค้นหาแบบยืดหยุ่น"""
    service = get_drive_service()
    search_target = normalize_filename(file_name)
    if not service or not search_target: return None
    
    try:
        # ⚠️ ต้องระบุพารามิเตอร์ Shared Drive เพื่อให้ API มองเห็นไฟล์
        results = service.files().list(
            q=f"'{GOOGLE_DRIVE_FOLDER_ID}' in parents and trashed = false",
            fields="files(id, name)",
            pageSize=1000,
            supportsAllDrives=True,
            includeItemsFromAllDrives=True
        ).execute()
        
        items = results.get('files', [])
        target_id = None
        
        # ค้นหาแบบเป๊ะๆ (หลัง Normalize)
        for item in items:
            if normalize_filename(item['name']) == search_target:
                target_id = item['id']
                break
        
        # ถ้าไม่เจอ ลองหาแบบ Partial Match (มีชื่อนี้อยู่ในชื่อไฟล์)
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
    except Exception:
        pass
    return None

def upload_image_to_drive(file_name, content_bytes):
    """อัปโหลดรูปภาพใหม่เข้า Shared Drive (แทนที่ไฟล์เดิมถ้ามีชื่อซ้ำ)"""
    service = get_drive_service()
    if not service: return
    try:
        clean_name = str(file_name).strip()
        if clean_name in ["0", "nan", ""]: clean_name = "uploaded_image.jpg"
        
        norm_target = normalize_filename(clean_name)
        
        # ค้นหาไฟล์เดิมใน Shared Drive เพื่อลบทิ้งก่อน
        results = service.files().list(
            q=f"'{GOOGLE_DRIVE_FOLDER_ID}' in parents and trashed = false",
            fields="files(id, name)",
            supportsAllDrives=True,
            includeItemsFromAllDrives=True
        ).execute()
        
        for f in results.get('files', []):
            if normalize_filename(f['name']) == norm_target:
                service.files().delete(fileId=f['id'], supportsAllDrives=True).execute()
        
        # อัปโหลดไฟล์ใหม่ผ่าน RAM (MediaIoBaseUpload)
        file_metadata = {'name': clean_name, 'parents': [GOOGLE_DRIVE_FOLDER_ID]}
        media = MediaIoBaseUpload(BytesIO(content_bytes), mimetype='image/jpeg', resumable=True)
        service.files().create(
            body=file_metadata, 
            media_body=media, 
            supportsAllDrives=True
        ).execute()
        
        st.cache_data.clear()
    except Exception as e:
        st.error(f"❌ อัปโหลดล้มเหลว: {str(e)}")

# --- 3. Utility Functions ---

def apply_exif_orientation(img):
    """ปรับหมุนรูปภาพตามค่า EXIF (แก้ปัญหารูปกลับหัว)"""
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
    """ฟอร์แมตเวลาให้สวยงาม (HH:mm)"""
    if not t or pd.isna(t) or str(t).strip() in ["", "0", "nan"]: return ""
    t = str(t).strip().replace(".", ":")
    try:
        parts = t.split(":")
        return f"{int(parts[0]):02d}:{int(parts[1]):02d}"
    except: return t

def parse_thai_date_simple(s):
    """แปลงวันที่ภาษาไทยจาก CSV เป็นรูปแบบที่ใช้งานได้"""
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
    if pd.notna(dt_f):
        story.append(Paragraph(f"เดือน : {d_str.split(' ', 1)[1] if len(d_str.split()) > 1 else d_str}", styles["Heading2"]))

    emp = df["name"].loc[df["name"].str.strip() != ""].iloc[0] if not df["name"].empty else ""
    story.append(Paragraph(f"เจ้าหน้าที่ดูแลประจำศูนย์ : {emp}", styles["Heading2"]))
    story.append(Spacer(1, 10))

    # ตารางสรุปรายเดือน
    t_data = [[Paragraph(h, styles["H_Table"]) for h in ["ลำดับ", "วันที่", "ชื่อ - นามสกุล", "เวลาเข้า", "เวลาออก", "ตำแหน่ง", "หมายเหตุ"]]]
    for i, row in df.iterrows():
        _, d_t = parse_thai_date_simple(row['date'])
        t_data.append([
            Paragraph(str(i+1), styles["C_Table"]), Paragraph(d_t, styles["C_Table"]),
            Paragraph(str(row['name']), styles["C_Table"]), Paragraph(fmt_time(row['time_in']), styles["C_Table"]),
            Paragraph(fmt_time(row['time_out']), styles["C_Table"]), Paragraph(str(row['status']), styles["C_Table"]),
            Paragraph("", styles["C_Table"])
        ])
    tbl = Table(t_data, colWidths=[35, 100, 130, 60, 60, 80, 70], repeatRows=1)
    tbl.setStyle(TableStyle([('GRID',(0,0),(-1,-1),0.5,colors.black),('BACKGROUND',(0,0),(-1,0),colors.lightgrey),('ALIGN',(0,0),(-1,-1),'CENTER'),('VALIGN',(0,0),(-1,-1),'MIDDLE')]))
    story.append(tbl)
    story.append(Spacer(1, 30))

    # ส่วนลายเซ็น
    sig_l = [Paragraph("....................................", styles["Signature"]), Spacer(1, 6), Paragraph(f"( {emp} )", styles["Signature"]), Paragraph("ผดล.ประจำศูนย์", styles["Signature"])]
    sig_r = [Paragraph("....................................", styles["Signature"]), Spacer(1, 6), Paragraph("( ...................................... )", styles["Signature"]), Paragraph("ตำแหน่ง_______________________", styles["Signature"])]
    story.append(KeepTogether(Table([[sig_l, sig_r]], colWidths=[260, 260])))

    # หน้าแนบรูปภาพรายวัน
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
                        p_img = p_img.convert('RGB')
                        t_io = BytesIO()
                        p_img.save(t_io, format="JPEG", quality=85)
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

st.title("🚀 ระบบจัดการรายงาน USO1 (Shared Drive Support)")

if 'main_df' not in st.session_state:
    try:
        st.session_state.main_df = pd.read_csv("03-2026.csv").fillna("")
    except:
        st.error("❌ ไม่พบไฟล์ 03-2026.csv")
        st.stop()

st.sidebar.title("เมนู")
centers = st.session_state.main_df['file_name'].unique()
sel_center = st.sidebar.selectbox("เลือกศูนย์", centers)

# ส่วน Debug ตรวจเช็คไฟล์ใน Shared Drive
if st.sidebar.checkbox("🔍 ตรวจสอบไฟล์ใน Shared Drive"):
    service = get_drive_service()
    if service:
        res = service.files().list(
            q=f"'{GOOGLE_DRIVE_FOLDER_ID}' in parents and trashed = false", 
            fields="files(name)",
            supportsAllDrives=True,
            includeItemsFromAllDrives=True
        ).execute()
        files_in_drive = [f['name'] for f in res.get('files', [])]
        st.sidebar.write(f"พบ {len(files_in_drive)} ไฟล์ในไดรฟ์")
        if files_in_drive:
            st.sidebar.code("\n".join(files_in_drive))
        else:
            st.sidebar.warning("ไม่พบไฟล์ใดๆ (โปรดเช็คสิทธิ์ Service Account)")

if st.sidebar.button("💾 บันทึกการแก้ไขลง CSV", width='stretch'):
    st.session_state.main_df.to_csv("03-2026.csv", index=False)
    st.sidebar.success("บันทึกเรียบร้อย!")

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
            if img_val and img_val not in ["0", "nan", ""]:
                img_d = download_image_from_drive(img_val)
                if img_d:
                    with Image.open(img_d) as im_disp:
                        c_img[i].image(im_disp, caption=img_val, width='stretch')
                else:
                    c_img[i].warning(f"❌ ไม่พบรูป: {img_val}")
            else:
                c_img[i].info("ยังไม่มีข้อมูลรูปภาพ")

            new_f = c_img[i].file_uploader(f"อัปโหลด {col}", type=['jpg','png','jpeg'], key=f"u_{col}_{idx}")
            if new_f:
                with st.spinner("กำลังอัปโหลด..."):
                    # ถ้าชื่อไฟล์เดิมว่าง ให้ตั้งชื่อใหม่ตามศูนย์และวันที่
                    target_name = img_val if img_val not in ["0", "nan", ""] else f"{sel_center}_{idx}_{col}.jpg"
                    upload_image_to_drive(target_name, new_f.getbuffer())
                    st.session_state.main_df.at[idx, col] = target_name
                    st.rerun()

st.divider()
if st.button("🖨️ ออกรายงาน PDF (รวมรูปภาพ)", width='stretch', type="primary"):
    with st.spinner("กำลังสร้าง PDF..."):
        pdf = generate_pdf_original_style(st.session_state.main_df.loc[df_idx], sel_center)
        st.download_button("📥 ดาวน์โหลดรายงาน PDF", pdf, f"{sel_center}.pdf", "application/pdf", width='stretch')
