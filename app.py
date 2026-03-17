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

# Google Drive Libraries
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaFileUpload
from google.oauth2 import service_account

# --- 1. การตั้งค่าหน้าเว็บและ Google Drive ---
st.set_page_config(page_title="USO1-Report Online", layout="wide")

# ID ของโฟลเดอร์ Google Drive ที่คุณให้มา
# GOOGLE_DRIVE_FOLDER_ID = '1YACVXpINnadQpX6DaOaBZsiMscGv9ERg'
GOOGLE_DRIVE_FOLDER_ID = '1-4OwgP-ODbelbtwSg5-m-rm4cyOTcW7O'

def get_drive_service():
    """เชื่อมต่อกับ Google Drive API โดยใช้ Secrets หรือไฟล์ Local"""
    try:
        if "gcp_service_account" in st.secrets:
            # สำหรับรันบน Streamlit Cloud
            info = st.secrets["gcp_service_account"]
            creds = service_account.Credentials.from_service_account_info(
                info, scopes=['https://www.googleapis.com/auth/drive']
            )
        else:
            # สำหรับรันในเครื่อง (ถ้ามีไฟล์ .json)
            creds = service_account.Credentials.from_service_account_file(
                'service_account.json', scopes=['https://www.googleapis.com/auth/drive']
            )
        return build('drive', 'v3', credentials=creds)
    except Exception as e:
        st.error(f"การเชื่อมต่อ Google Drive ล้มเหลว: {e}")
        return None


def init_fonts():
    try:
        pdfmetrics.registerFont(TTFont('THSarabun', 'THSarabunNew.ttf'))
        pdfmetrics.registerFont(
            TTFont('THSarabun-Bold', 'THSarabunNew Bold.ttf'))
        return 'THSarabun', 'THSarabun-Bold'
    except:
        return 'Helvetica', 'Helvetica-Bold'


F_REG, F_BOLD = init_fonts()

# --- 2. Google Drive Helper Functions ---


@st.cache_data(ttl=600)  # Cache รูปไว้ 10 นาที
def download_image_from_drive(file_name):
    service = get_drive_service()
    if not service:
        return None
    try:
        query = f"name = '{file_name}' and '{GOOGLE_DRIVE_FOLDER_ID}' in parents and trashed = false"
        results = service.files().list(q=query, fields="files(id)").execute()
        items = results.get('files', [])
        if not items:
            return None

        file_id = items[0]['id']
        request = service.files().get_media(fileId=file_id)
        fh = BytesIO()
        downloader = MediaIoBaseDownload(fh, request)
        done = False
        while not done:
            _, done = downloader.next_chunk()
        fh.seek(0)
        return fh
    except:
        return None


def upload_image_to_drive(file_name, content_bytes):
    service = get_drive_service()
    if not service:
        return
    # ลบไฟล์เดิมถ้าชื่อซ้ำเพื่อป้องกันไฟล์ขยะ
    query = f"name = '{file_name}' and '{GOOGLE_DRIVE_FOLDER_ID}' in parents"
    results = service.files().list(q=query, fields="files(id)").execute()
    for f in results.get('files', []):
        service.files().delete(fileId=f['id']).execute()

    file_metadata = {'name': file_name, 'parents': [GOOGLE_DRIVE_FOLDER_ID]}
    # สร้างไฟล์ชั่วคราวเพื่ออัปโหลด
    with open("temp_upload.jpg", "wb") as f:
        f.write(content_bytes)
    media = MediaFileUpload("temp_upload.jpg", mimetype='image/jpeg')
    service.files().create(body=file_metadata, media_body=media).execute()
    if os.path.exists("temp_upload.jpg"):
        os.remove("temp_upload.jpg")
    st.cache_data.clear()  # ล้าง Cache เพื่อให้รูปใหม่แสดงผลทันที

# --- 3. Image & Date Helpers ---


def apply_exif_orientation(img):
    try:
        exif = img._getexif()
        if exif:
            orientation_tag = next(
                (k for k, v in ExifTags.TAGS.items() if v == 'Orientation'), None)
            if orientation_tag in exif:
                orient = exif[orientation_tag]
                if orient == 3:
                    img = img.transpose(Image.ROTATE_180)
                elif orient == 6:
                    img = img.transpose(Image.ROTATE_270)
                elif orient == 8:
                    img = img.transpose(Image.ROTATE_90)
    except:
        pass
    return img


def fmt_time(t):
    if not t or pd.isna(t) or str(t).strip() == "":
        return ""
    t = str(t).strip().replace(".", ":")
    try:
        parts = t.split(":")
        return f"{int(parts[0]):02d}:{int(parts[1]):02d}"
    except:
        return t


def parse_thai_date_simple(s):
    month_thai_name = {1: "มกราคม", 2: "กุมภาพันธ์", 3: "มีนาคม", 4: "เมษายน", 5: "พฤษภาคม", 6: "มิถุนายน",
                       7: "กรกฎาคม", 8: "สิงหาคม", 9: "กันยายน", 10: "ตุลาคม", 11: "พฤศจิกายน", 12: "ธันวาคม"}
    thai_months_map = {"มกราคม": "01", "กุมภาพันธ์": "02", "มีนาคม": "03", "เมษายน": "04", "พฤษภาคม": "05", "มิถุนายน": "06",
                       "กรกฎาคม": "07", "สิงหาคม": "08", "กันยายน": "09", "ตุลาคม": "10", "พฤศจิกายน": "11", "ธันวาคม": "12"}
    if not s or pd.isna(s):
        return pd.NaT, ""
    try:
        s_clean = str(s).strip()
        for k, v in thai_months_map.items():
            if k in s_clean:
                s_clean = s_clean.replace(k, v)
        parts = s_clean.split()
        if len(parts) == 3:
            day, month, year = parts
            y_int = int(year)
            if y_int > 2500:
                y_int -= 543
            dt = pd.to_datetime(f"{y_int}-{month}-{day}")
            return dt, f"{day} {month_thai_name[int(month)]} {int(year)}"
    except:
        pass
    return pd.NaT, str(s)

# --- 4. PDF Generator ---


def generate_pdf_original_style(df, center_name):
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=20,
                            leftMargin=20, topMargin=25, bottomMargin=15)
    thai_styles = {
        "Normal": ParagraphStyle("ThaiNormal", fontName=F_REG, fontSize=14, leading=18, alignment=1),
        "Title": ParagraphStyle("ThaiTitle", fontName=F_BOLD, fontSize=18, leading=24, alignment=1),
        "Heading2": ParagraphStyle("ThaiHeading2", fontName=F_BOLD, fontSize=14, leading=20, alignment=1),
        "Signature": ParagraphStyle("ThaiSignature", fontName=F_REG, fontSize=14, leading=18, alignment=1),
        "HeaderStyle": ParagraphStyle("H", fontName=F_BOLD, fontSize=10, leading=11, alignment=1),
        "CellStyle": ParagraphStyle("C", fontName=F_REG, fontSize=10, leading=11, alignment=1),
    }

    story = []
    story.append(
        Paragraph("รายงานเวลาปฏิบัติงาน USO1-Renew", thai_styles["Title"]))
    story.append(Paragraph(f"ศูนย์ : {center_name}", thai_styles["Title"]))

    dt_first, date_str = parse_thai_date_simple(df.iloc[0]['date'])
    if pd.notna(dt_first):
        story.append(
            Paragraph(f"เดือน : {date_str.split(' ', 1)[1]}", thai_styles["Heading2"]))

    valid_names = df["name"].loc[df["name"].str.strip() != ""]
    emp_name = valid_names.iloc[0] if not valid_names.empty else ""
    story.append(
        Paragraph(f"เจ้าหน้าที่ดูแลประจำศูนย์ : {emp_name}", thai_styles["Heading2"]))

    # ส่วนตารางหน้าแรก
    table_data = [[Paragraph(h, thai_styles["HeaderStyle"]) for h in [
        "ลำดับ", "วันที่", "ชื่อ - นามสกุล", "เวลาเข้า", "เวลาออก", "ตำแหน่ง", "หมายเหตุ"]]]
    for i, row in df.iterrows():
        _, d_thai = parse_thai_date_simple(row['date'])
        table_data.append([
            Paragraph(str(i+1), thai_styles["CellStyle"]
                      ), Paragraph(d_thai, thai_styles["CellStyle"]),
            Paragraph(row['name'], thai_styles["CellStyle"]), Paragraph(
                fmt_time(row['time_in']), thai_styles["CellStyle"]),
            Paragraph(fmt_time(row['time_out']), thai_styles["CellStyle"]), Paragraph(
                row['status'], thai_styles["CellStyle"]),
            Paragraph("", thai_styles["CellStyle"])
        ])
    tbl = Table(table_data, colWidths=[
                35, 100, 130, 60, 60, 80, 70], repeatRows=1)
    tbl.setStyle(TableStyle([('GRID', (0, 0), (-1, -1), 0.5, colors.black), ('BACKGROUND', (0, 0), (-1, 0),
                 colors.lightgrey), ('ALIGN', (0, 0), (-1, -1), 'CENTER'), ('VALIGN', (0, 0), (-1, -1), 'MIDDLE')]))
    story.append(tbl)
    story.append(Spacer(1, 30))

    # ส่วนลายเซ็น
    sig_style = thai_styles["Signature"]
    sig_left = [Paragraph("....................................", sig_style), Spacer(
        1, 6), Paragraph(f"( {emp_name} )", sig_style), Paragraph("ผดล.ประจำศูนย์", sig_style)]
    sig_right = [Paragraph("....................................", sig_style), Spacer(1, 6), Paragraph(
        "( ...................................... )", sig_style), Paragraph("ตำแหน่ง_______________________", sig_style)]
    story.append(KeepTogether(
        Table([[sig_left, sig_right]], colWidths=[260, 260])))

    # ส่วนรายละเอียดรายวันพร้อมรูปภาพจาก Drive
    for _, r in df.iterrows():
        story.append(PageBreak())
        _, d_thai = parse_thai_date_simple(r['date'])
        story.append(Paragraph(f"วันที่ : {d_thai}", thai_styles["Heading2"]))
        story.append(Spacer(1, 12))
        story.append(Paragraph(
            f"ชื่อ : <b>{r['name']}</b> &nbsp; ตำแหน่ง : <b>{r['status']}</b>", thai_styles["Normal"]))

        for label, col_img, col_time in [("เข้า (เช้า)", "img_in1", "time_in"), ("ออก (เย็น)", "img_out1", "time_out")]:
            img_name = str(r[col_img])
            img_bytes = download_image_from_drive(img_name)

            if img_bytes:
                try:
                    with Image.open(img_bytes) as PIL_img:
                        PIL_img = apply_exif_orientation(PIL_img)
                        temp_io = BytesIO()
                        PIL_img.convert('RGB').save(
                            temp_io, format="JPEG", quality=85)
                        temp_io.seek(0)
                        im = RLImage(temp_io)
                        im._restrictSize(310, 260)
                        img_tbl = Table([[im]], colWidths=[450])
                        img_tbl.setStyle(TableStyle(
                            [('ALIGN', (0, 0), (-1, -1), 'CENTER')]))
                        story.append(img_tbl)
                except:
                    pass

            story.append(Paragraph(
                f"เวลา{label} : <b>{fmt_time(r[col_time])}</b>", thai_styles["Normal"]))
            story.append(Spacer(1, 18))

    doc.build(story)
    return buffer.getvalue()

# --- 5. Main UI ---


st.title("🚀 USO1 Report Manager (Google Drive Edition)")

if 'main_df' not in st.session_state:
    st.session_state.main_df = pd.read_csv("03-2026.csv")

if 'last_uploaded_id' not in st.session_state:
    st.session_state.last_uploaded_id = ""

centers = st.session_state.main_df['file_name'].unique()
selected_center = st.sidebar.selectbox("เลือกศูนย์", centers)

if st.sidebar.button("💾 บันทึก CSV (เครื่อง)", use_container_width=True):
    st.session_state.main_df.to_csv("03-2026.csv", index=False)
    st.sidebar.success("บันทึกสำเร็จ!")

df_idx = st.session_state.main_df[st.session_state.main_df['file_name']
                                  == selected_center].index

for idx in df_idx:
    row = st.session_state.main_df.loc[idx]
    with st.expander(f"📅 {row['date']} - {row['name']}"):
        c = st.columns([2, 2, 1, 1])
        st.session_state.main_df.at[idx, 'name'] = c[0].text_input(
            "ชื่อ", row['name'], key=f"n_{idx}")
        st.session_state.main_df.at[idx, 'status'] = c[1].text_input(
            "ตำแหน่ง", row['status'], key=f"s_{idx}")
        st.session_state.main_df.at[idx, 'time_in'] = c[2].text_input(
            "เข้า", row['time_in'], key=f"i_{idx}")
        st.session_state.main_df.at[idx, 'time_out'] = c[3].text_input(
            "ออก", row['time_out'], key=f"o_{idx}")

        c_img = st.columns(2)
        for i, col in enumerate(["img_in1", "img_out1"]):
            img_name = str(row[col])
            img_data = download_image_from_drive(img_name)

            if img_data:
                with Image.open(img_data) as img_display:
                    img_display = apply_exif_orientation(img_display)
                    c_img[i].image(img_display, caption=img_name,
                                   use_container_width=True)
            else:
                c_img[i].warning(f"ไม่พบไฟล์ {img_name} ใน Drive")

            u_key = f"u_{col}_{idx}"
            new_f = c_img[i].file_uploader(f"เปลี่ยนรูป {col}", type=[
                                           'jpg', 'png', 'jpeg'], key=u_key)
            if new_f is not None:
                current_file_id = f"{u_key}_{new_f.name}_{new_f.size}"
                if st.session_state.last_uploaded_id != current_file_id:
                    with st.spinner(f"กำลังอัปโหลด {img_name}..."):
                        upload_image_to_drive(img_name, new_f.getbuffer())
                        st.session_state.last_uploaded_id = current_file_id
                        st.toast(f"อัปโหลด {img_name} สำเร็จ!")
                        time.sleep(1)
                        st.rerun()

st.divider()
if st.button("🖨️ ออกรายงาน PDF (ดึงรูปจาก Drive)", use_container_width=True, type="primary"):
    with st.spinner("กำลังสร้าง PDF..."):
        pdf = generate_pdf_original_style(
            st.session_state.main_df.loc[df_idx], selected_center)
        st.download_button(
            "📥 ดาวน์โหลด PDF", pdf, f"{selected_center}.pdf", "application/pdf", use_container_width=True)
