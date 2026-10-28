from datetime import datetime
import os
from google import genai
from google.genai import types
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

  # ใช้ค่าจาก Streamlit Secrets หรือช่องกรอก
  api_key = api_key_input or os.environ.get("GEMINI_API_KEY")

  st.markdown("---")
  st.markdown(
      "**คำแนะนำการใช้งาน:**\n1. ตรวจสอบเดือนและปีปัจจุบัน\n2. อัปโหลดรูปภาพหรือสแกนใบลาที่มีตารางพนักงาน\n3. ให้ AI ช่วยดึงข้อมูลและตรวจสอบความถูกต้อง\n4. ดาวน์โหลดไฟล์ Excel ตามแบบฟอร์ม"
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
      "ประเภทพนักงาน", ["ลูกจ้างประจำ", "พนักงานราชการ"]
  )

st.info(
    f"📂 กำลังจัดการข้อมูลของ: **{employee_type}** ประจำเดือน**{selected_month} {selected_year}**"
)

# -------------------------------------------------------------------------
# 3. ส่วนอัปโหลดเอกสารและจำลองการทำ OCR ด้วย Gemini API
# -------------------------------------------------------------------------
st.subheader("1. อัปโหลดเอกสารใบลา (สแกน/รูปภาพ/PDF)")
uploaded_file = st.file_uploader(
    "เลือกไฟล์เอกสารบัญชีลงเวลา/ใบลา", type=["png", "jpg", "jpeg", "pdf"]
)

# ข้อมูลตัวอย่างเริ่มต้น (Mock Data) เผื่อยังไม่ได้ต่อ API
if "leave_data" not in st.session_state:
  st.session_state.leave_data = pd.DataFrame([
      {
          "ที่": 1,
          "ชื่อ - สกุล": "นางสาวณัฐพร มิตรพิทักษ์",
          "ตำแหน่ง": "พนักงานจ้างปฏิบัติงานนโยบายและแผน",
          "ตั้งแต่วันที่": "5 Oct 26",
          "ถึงวันที่": "5 Oct 26",
          "ลาป่วย": 0.5,
          "ลากิจ": "-",
          "รวม (วัน)": 0.5,
          "หมายเหตุ": "ครึ่งวันเช้า",
      },
      {
          "ที่": 2,
          "ชื่อ - สกุล": "นางสาวณัฐธิชา แถวทัด",
          "ตำแหน่ง": "พนักงานจ้างปฏิบัติงานนิเทศฯ",
          "ตั้งแต่วันที่": "-",
          "ถึงวันที่": "-",
          "ลาป่วย": "-",
          "ลากิจ": "-",
          "รวม (วัน)": "-",
          "หมายเหตุ": "-",
      },
  ])

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
          # เริ่มต้นใช้งาน Google GenAI SDK
          client = genai.Client(api_key=api_key)

          # อ่านไบต์ของไฟล์ที่อัปโหลด
          file_bytes = uploaded_file.getvalue()
          mime_type = uploaded_file.type

          prompt = (
              "คุณคือผู้ช่วยเจ้าหน้าที่บุคคล จงอ่านข้อมูลจากตารางใบลาในรูปภาพนี้"
              " แล้วสกัดออกมาเป็นโครงสร้าง JSON ที่มีฟิลด์ดังนี้: "
              "no (เลขลำดับ), name (ชื่อ-สกุล), position (ตำแหน่ง), "
              "start_date (ตั้งแต่วันที่), end_date (ถึงวันที่), "
              "sick_leave (ลาป่วย), personal_leave (ลากิจ), total_days"
              " (รวมวันลา), note (หมายเหตุ)."
              " ตอบกลับมาเป็นโครงสร้างข้อมูลที่นำไปแปลงเป็น DataFrame ได้ทันที"
          )

          response = client.models.generate_content(
              model="gemini-2.5-flash",
              contents=[
                  types.Part.from_bytes(data=file_bytes, mime_type=mime_type),
                  prompt,
              ],
          )

          st.success("อ่านข้อมูลสำเร็จ! (ตรวจสอบข้อมูลด้านล่างได้เลย)")
          # หมายเหตุ: ในการใช้งานจริงสามารถแปลง text response เป็น JSON แล้วยัดใส่ st.session_state.leave_data ได้ทันที

      except Exception as e:
        st.error(f"เกิดข้อผิดพลาดในการเชื่อมต่อ API: {e}")

# -------------------------------------------------------------------------
# 4. ตารางตรวจสอบและแก้ไขข้อมูลก่อนออกรายงาน
# -------------------------------------------------------------------------
st.subheader("2. ตรวจสอบและแก้ไขข้อมูลก่อนสร้างรายงาน")
st.markdown(
    "สามารถแก้ไขข้อมูลในตารางด้านล่างนี้ได้โดยตรง หาก AI อ่านข้อมูลผิดเพี้ยน"
)

edited_df = st.data_editor(
    st.session_state.leave_data, num_rows="dynamic", use_container_width=True
)

# -------------------------------------------------------------------------
# 5. ส่วนดาวน์โหลดไฟล์ Excel ตาม Template
# -------------------------------------------------------------------------
st.subheader("3. ดาวน์โหลดรายงานสรุปประจำเดือน")


def generate_excel(df, month, year, emp_type):
  # ฟังก์ชันสำหรับจัดรูปทรง Excel ตาม Template ราชการ
  import io

  output = io.BytesIO()
  with pd.ExcelWriter(output, engine="openpyxl") as writer:
    # เขียนข้อมูลลงใน Sheet
    df.to_excel(writer, sheet_name="สรุปวันลา", index=False, startrow=4)

    # สามารถใช้ openpyxl ปรับแต่งหัวตาราง ชื่อหน่วยงาน เพิ่มเติมตรงนี้ได้
    worksheet = writer.sheets["สรุปวันลา"]
    worksheet["A1"] = (
        f"ตารางสรุปวันลา {emp_type} สำนักงานศึกษาธิการจังหวัดสมุทรสงคราม"
    )
    worksheet["A2"] = f"ประจำเดือน {month} {year}"

  output.seek(0)
  return output


if st.button("📥 สร้างและดาวน์โหลดไฟล์ Excel"):
  excel_file = generate_excel(
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