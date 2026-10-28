from datetime import datetime
import io
from google import genai
from google.genai import types
import openpyxl
from openpyxl.styles import Alignment, Border, Font, Side
import pandas as pd
import streamlit as st

# ตั้งค่าหน้าเว็บ
st.set_page_config(
    page_title="ระบบสรุปวันลาประจำเดือน", page_icon="📅", layout="wide"
)

st.title("📌 ระบบสรุปวันลาประจำเดือน (อัตโนมัติ & AI OCR)")

# -------------------------------------------------------------------------
# 1. ส่วนตั้งค่า API Key สำหรับ Google Gemini
# -------------------------------------------------------------------------
with st.sidebar:
  st.header("⚙️ ตั้งค่าระบบ")
  api_key_input = st.text_input(
      "Google Gemini API Key",
      type="password",
      help="ใส่ API Key ของคุณที่นี่ หรือตั้งค่าใน Environment Variable",
  )
  api_key = api_key_input or os.environ.get("GEMINI_API_KEY", "")

  st.markdown("---")
  st.markdown(
      "**คำแนะนำการใช้งาน:**\n1. ตรวจสอบเดือนและปีปัจจุบัน\n2. อัปโหลดรูปภาพหรือสแกนใบลาที่มีตารางพนักงาน\n3. ตรวจสอบและแก้ไขข้อมูลในตาราง\n4. กดดาวน์โหลดไฟล์ Excel ตามแบบฟอร์มราชการ"
  )

# -------------------------------------------------------------------------
# 2. ส่วนเลือกเดือนและปี (Default เป็นปัจจุบัน)
# -------------------------------------------------------------------------
thai_months = [
    "มกราคม",
    "กุมภาพันธ์",
    "มีนาคม",
    "เมษายน",
    "พฤษภาคม",
    "มิถุนายน",
    "กรกฎาคม",
    "สิงหาคม",
    "กันยายน",
    "ตุลาคม",
    "พฤศจิกายน",
    "ธันวาคม",
]

now = datetime.now()
default_month_idx = now.month - 1
default_year = now.year + 543  # แปลงเป็นพ.ศ.

col1, col2, col3 = st.columns([2, 2, 3])
with col1:
  selected_month = st.selectbox(
      "ประจำเดือน", thai_months, index=default_month_idx
  )
with col2:
  selected_year = st.number_input("ปี พ.ศ.", value=default_year)
with col3:
  employee_type = st.selectbox(
      "ประเภทพนักงาน",
      ["พนักงานจ้างเหมาเอกชน", "ลูกจ้างประจำ", "พนักงานราชการ"],
  )

st.info(
    f"📂 กำลังจัดการข้อมูลของ: **{employee_type}** ประจำเดือน**{selected_month}"
    f" {selected_year}**"
)

# -------------------------------------------------------------------------
# 3. ข้อมูลตัวอย่างเริ่มต้น (Mock Data) และส่วนอัปโหลดไฟล์
# -------------------------------------------------------------------------
if "leave_data" not in st.session_state:
  st.session_state.leave_data = pd.DataFrame([
      {
          "ที่": 1,
          "ชื่อ - สกุล": "นางสาวณัฐพร  มิตรพิทักษ์",
          "ตำแหน่ง": "พนักงานจ้างปฏิบัติงานนโยบายและแผน",
          "ตั้งแต่วันที่": "5 Oct 26",
          "ถึงวันที่": "5 Oct 26",
          "ลาป่วย": 0.5,
          "ลากิจ": "-",
          "รวม (วัน)": 1.5,
          "หมายเหตุ": "ครึ่งวันเช้า",
      },
      {
          "ที่": 2,
          "ชื่อ - สกุล": "นางสาวณัฐธิชา  แถวทัด",
          "ตำแหน่ง": "พนักงานจ้างปฏิบัติงานนิเทศติดตามและประเมินผล",
          "ตั้งแต่วันที่": "-",
          "ถึงวันที่": "-",
          "ลาป่วย": "-",
          "ลากิจ": "-",
          "รวม (วัน)": "-",
          "หมายเหตุ": "-",
      },
  ])

st.subheader("1. อัปโหลดเอกสารใบลา (สแกน/รูปภาพ/PDF)")
uploaded_file = st.file_uploader(
    "เลือกไฟล์เอกสารบัญชีลงเวลา/ใบลา", type=["png", "jpg", "jpeg", "pdf"]
)

if uploaded_file is not None:
  st.image(
      uploaded_file,
      caption="ไฟล์เอกสารที่อัปโหลด",
      use_container_width=True,
  )
  if st.button("✨ ให้ AI (Gemini) ช่วยอ่านข้อมูลจากเอกสาร"):
    if not api_key:
      st.error("กรุณากรอก Google Gemini API Key ที่แถบด้านข้างก่อนครับ")
    else:
      try:
        with st.spinner("AI กำลังอ่านข้อมูลจากตาราง... รอสักครู่"):
          client = genai.Client(api_key=api_key)
          file_bytes = uploaded_file.getvalue()
          mime_type = uploaded_file.type

          prompt = (
              "คุณคือเจ้าหน้าที่บุคคล จงอ่านข้อมูลตารางใบลาจากรูปนี้ แล้วสกัดเป็น"
              " JSON รายชื่อพนักงาน ฟิลด์:ที่, ชื่อ-สกุล, ตำแหน่ง, ตั้งแต่วันที่,"
              " ถึงวันที่, ลาป่วย, ลากิจ, รวม(วัน), หมายเหตุ"
          )
          response = client.models.generate_content(
              model="gemini-2.5-flash",
              contents=[
                  types.Part.from_bytes(data=file_bytes, mime_type=mime_type),
                  prompt,
              ],
          )
          st.success("อ่านข้อมูลสำเร็จ!")
      except Exception as e:
        st.error(f"เกิดข้อผิดพลาด: {e}")

# -------------------------------------------------------------------------
# 4. ตารางตรวจสอบและแก้ไขข้อมูลก่อนออกรายงาน
# -------------------------------------------------------------------------
st.subheader("2. ตรวจสอบและแก้ไขข้อมูลก่อนสร้างรายงาน")
edited_df = st.data_editor(
    st.session_state.leave_data, num_rows="dynamic", use_container_width=True
)

# -------------------------------------------------------------------------
# 5. ฟังก์ชันสร้าง Excel ตาม Template ราชการ (เปิดตาราง 2 ชั้น + ลงชื่อ)
# -------------------------------------------------------------------------


def generate_excel_template(df, month, year, emp_type):
  wb = openpyxl.Workbook()
  ws = wb.active
  ws.title = "สรุปวันลา"

  font_header = Font(name="TH Sarabun PSK", size=16, bold=True)
  font_bold = Font(name="TH Sarabun PSK", size=14, bold=True)
  font_normal = Font(name="TH Sarabun PSK", size=14)

  thin_border = Border(
      left=Side(style="thin", color="000000"),
      right=Side(style="thin", color="000000"),
      top=Side(style="thin", color="000000"),
      bottom=Side(style="thin", color="000000"),
  )
  align_center = Alignment(
      horizontal="center", vertical="center", wrap_text=True
  )
  align_left = Alignment(horizontal="left", vertical="center")

  # หัวข้อรายงาน
  ws["A1"] = f"ตารางสรุปวันลา {emp_type} สำนักงานศึกษาธิการจังหวัดสมุทรสงคราม"
  ws["A1"].font = font_header
  ws.merge_cells("A1:I1")
  ws["A1"].alignment = align_center

  ws["A2"] = f"ประจำเดือน {month} {year}"
  ws["A2"].font = font_header
  ws.merge_cells("A2:I2")
  ws["A2"].alignment = align_center

  # หัวตาราง 2 ชั้น (Rows 4-5)
  ws.merge_cells("A4:A5")
  ws["A4"] = "ที่"
  ws.merge_cells("B4:B5")
  ws["B4"] = "ชื่อ - สกุล"
  ws.merge_cells("C4:C5")
  ws["C4"] = "ตำแหน่ง"
  ws.merge_cells("D4:D5")
  ws["D4"] = "ตั้งแต่วันที่"
  ws.merge_cells("E4:E5")
  ws["E4"] = "ถึงวันที่"

  ws.merge_cells("F4:G4")
  ws["F4"] = "ครั้งนี้"
  ws["F5"] = "ลาป่วย"
  ws["G5"] = "ลากิจ"

  ws.merge_cells("H4:H5")
  ws["H4"] = "รวม (วัน)"
  ws.merge_cells("I4:I5")
  ws["I4"] = "หมายเหตุ"

  for row in ws["A4:I5"]:
    for cell in row:
      cell.font = font_bold
      cell.alignment = align_center
      cell.border = thin_border

  # วนลูปข้อมูลพนักงาน
  start_row = 6
  last_row = start_row
  for i, row in df.iterrows():
    r = start_row + i
    ws[f"A{r}"] = row.get("ที่", i + 1)
    ws[f"B{r}"] = row.get("ชื่อ - สกุล", "")
    ws[f"C{r}"] = row.get("ตำแหน่ง", "")
    ws[f"D{r}"] = row.get("ตั้งแต่วันที่", "-")
    ws[f"E{r}"] = row.get("ถึงวันที่", "-")
    ws[f"F{r}"] = row.get("ลาป่วย", "-")
    ws[f"G{r}"] = row.get("ลากิจ", "-")
    ws[f"H{r}"] = row.get("รวม (วัน)", "-")
    ws[f"I{r}"] = row.get("หมายเหตุ", "-")

    for col_letter in ["A", "B", "C", "D", "E", "F", "G", "H", "I"]:
      cell = ws[f"{col_letter}{r}"]
      cell.font = font_normal
      cell.border = thin_border
      if col_letter in ["A", "D", "E", "F", "G", "H"]:
        cell.alignment = align_center
      else:
        cell.alignment = align_left
    last_row = r

  # ส่วนลงชื่อท้ายตาราง
  sig_row = last_row + 3
  ws.merge_cells(f"F{sig_row}:I{sig_row}")
  ws[f"F{sig_row}"] = "(ลงชื่อ)....................................................ผู้ตรวจสอบ"
  ws[f"F{sig_row}"].font = font_normal
  ws[f"F{sig_row}"].alignment = align_center

  ws.merge_cells(f"F{sig_row+1}:I{sig_row+1}")
  ws[f"F{sig_row+1}"] = "(นายวงศ์สกุล  เข่งคุ้ม)"
  ws[f"F{sig_row+1}"].font = font_normal
  ws[f"F{sig_row+1}"].alignment = align_center

  ws.merge_cells(f"F{sig_row+2}:I{sig_row+2}")
  ws[f"F{sig_row+2}"] = "ตำแหน่ง  นักทรัพยากรบุคคลปฏิบัติการ"
  ws[f"F{sig_row+2}"].font = font_normal
  ws[f"F{sig_row+2}"].alignment = align_center

  output = io.BytesIO()
  wb.save(output)
  output.seek(0)
  return output


# -------------------------------------------------------------------------
# 6. ปุ่มดาวน์โหลดรายงาน
# -------------------------------------------------------------------------
st.subheader("3. ดาวน์โหลดรายงานสรุปประจำเดือน")
if st.button("📥 สร้างและดาวน์โหลดไฟล์ Excel ตามแบบฟอร์ม"):
  excel_file = generate_excel_template(
      edited_df, selected_month, selected_year, employee_type
  )

  st.download_button(
      label="💾 คลิกลงเพื่อดาวน์โหลดไฟล์ .xlsx",
      data=excel_file,
      file_name=(
          f"สรุปวันลา_{employee_type}_{selected_month}_{selected_year}.xlsx"
      ),
      mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  )
  st.success("สร้างไฟล์ Excel เรียบร้อยแล้วครับ!")