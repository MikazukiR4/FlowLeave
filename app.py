from datetime import datetime
import io
import os
import openpyxl
from openpyxl.styles import Alignment, Border, Font, Side
import pandas as pd
import streamlit as st

# ตั้งค่าหน้าเว็บ
st.set_page_config(
    page_title="ระบบสรุปวันลาประจำเดือน", page_icon="📅", layout="wide"
)

st.title("📌 ระบบสรุปวันลาประจำเดือน (Template ราชการ)")

# -------------------------------------------------------------------------
# 1. ข้อมูลรายชื่อพนักงานตั้งต้น (Master List)
# -------------------------------------------------------------------------
DEFAULT_EMPLOYEES = {
    "พนักงานจ้างเหมาเอกชน": [
        {
            "name": "นางสาวณัฐพร  มิตรพิทักษ์",
            "position": "พนักงานจ้างปฏิบัติงานนโยบายและแผน",
        },
        {
            "name": "นางสาวณัฐธิชา  แถวทัด",
            "position": "พนักงานจ้างปฏิบัติงานนิเทศติดตามและประเมินผล",
        },
        {"name": "นางสาววิชุดา  ศรีคร้ามครัน", "position": "พนักงานจ้างปฏิบัติงานธุรการ"},
        {"name": "นางสาวยัสมี  ดาโอะ", "position": "พนักงานจ้างปฏิบัติงานธุรการ"},
        {"name": "นางสาวภาวินีย์  ตันสกุล", "position": "พนักงานจ้างทำความสะอาด"},
        {"name": "นายภานุวัฒน์  จันทร์ชีวา", "position": "พนักงานจ้างปฏิบัติงานขับรถยนต์"},
        {
            "name": "นางสาวหัทยาภักดิ์ แสงแก้ว",
            "position": "พนักงานจ้างปฏิบัติงานการเงินและบัญชี",
        },
        {
            "name": "นายหนึ่ง  สุนทรรัตน์",
            "position": "พนักงานจ้างปฏิบัติงานรักษาความปลอดภัย",
        },
    ],
    "พนักงานราชการ": [
        {
            "name": "นางสาวอมรารัตน์  พูลสวัสดิ์",
            "position": "พนักงานจัดการงานบุคคล",
        },
        {
            "name": "นายสมชาย  สุพรดวงชัย",
            "position": "พนักงานจัดการงานทั่วไป",
        },
    ],
}

# ฟังก์ชันแปลงวันที่เป็นรูปแบบไทย (เช่น 1 ม.ค. 69)
THAI_MONTH_ABBR = [
    "",
    "ม.ค.",
    "ก.พ.",
    "มี.ค.",
    "เม.ย.",
    "พ.ค.",
    "มิ.ย.",
    "ก.ค.",
    "ส.ค.",
    "ก.ย.",
    "ต.ค.",
    "พ.ย.",
    "ธ.ค.",
]


def format_thai_date(date_obj):
  if not date_obj:
    return "-"
  day = date_obj.day
  month = THAI_MONTH_ABBR[date_obj.month]
  year_short = str(date_obj.year + 43)[-2:]
  return f"{day} {month} {year_short}"


with st.sidebar:
  st.header("⚙️ คำแนะนำการใช้งาน")
  st.markdown(
      "1. เลือกประเภทพนักงาน\n2. เพิ่มหรือลบรายชื่อบุคลากรได้ที่เมนูด้านล่าง\n3. แก้ไขวันลาในตาราง (ช่องไหนไม่ลาให้ใส่ '-')\n4. กดสร้างและดาวน์โหลดไฟล์ Excel"
  )

# เลือกเดือน ปี และประเภทพนักงาน
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

col1, col2, col3 = st.columns([2, 2, 3])
with col1:
  selected_month = st.selectbox(
      "ประจำเดือน", thai_months, index=now.month - 1
  )
with col2:
  selected_year = st.number_input("ปี พ.ศ.", value=now.year + 543)
with col3:
  employee_type = st.selectbox(
      "ประเภทพนักงาน", list(DEFAULT_EMPLOYEES.keys())
  )

# -------------------------------------------------------------------------
# 2. จัดการ Session State สำหรับเก็บข้อมูลตารางและบุคลากร
# -------------------------------------------------------------------------
if (
    "current_emp_type" not in st.session_state
    or st.session_state.current_emp_type != employee_type
):
  st.session_state.current_emp_type = employee_type
  emp_list = DEFAULT_EMPLOYEES.get(employee_type, [])

  initial_rows = []
  for idx, emp in enumerate(emp_list, 1):
    initial_rows.append({
        "ที่": idx,
        "ชื่อ - สกุล": emp["name"],
        "ตำแหน่ง": emp["position"],
        "ตั้งแต่วันที่": "-",
        "ถึงวันที่": "-",
        "ลาป่วย": "-",
        "ลากิจ": "-",
        "รวม (วัน)": "-",
        "หมายเหตุ": "-",
    })
  st.session_state.leave_data = pd.DataFrame(initial_rows)

st.info(
    f"📂 กำลังจัดการข้อมูลของ: **{employee_type}** ประจำเดือน**{selected_month}"
    f" {selected_year}**"
)

# -------------------------------------------------------------------------
# 3. ตารางหลักสำหรับแก้ไขข้อมูล (Data Editor)
# -------------------------------------------------------------------------
st.subheader("1. ตารางสรุปข้อมูลบุคลากรและวันลา")
st.markdown(
    "💡 *เคล็ดลับ: สามารถคลิกแก้ข้อมูลในตารางได้โดยตรง หรือเพิ่ม/ลบรายชื่อด้านล่างได้*"
)

edited_df = st.data_editor(
    st.session_state.leave_data,
    num_rows="dynamic",
    use_container_width=True,
    key="employee_table_editor",
)
st.session_state.leave_data = edited_df

# -------------------------------------------------------------------------
# 4. ส่วนเพิ่ม / ลบ บุคลากร
# -------------------------------------------------------------------------
with st.expander("➕ จัดการรายชื่อบุคลากร (เพิ่ม / ลบ พนักงาน)"):
  col_add, col_del = st.columns(2)

  with col_add:
    st.markdown("#### เพิ่มบุคลากรใหม่")
    with st.form("add_emp_form", clear_on_submit=True):
      new_name = st.text_input("ชื่อ - สกุล")
      new_pos = st.text_input("ตำแหน่ง")
      submitted_add = st.form_submit_button("เพิ่มพนักงาน")
      if submitted_add and new_name:
        new_row = {
            "ที่": len(st.session_state.leave_data) + 1,
            "ชื่อ - สกุล": new_name,
            "ตำแหน่ง": new_pos or "-",
            "ตั้งแต่วันที่": "-",
            "ถึงวันที่": "-",
            "ลาป่วย": "-",
            "ลากิจ": "-",
            "รวม (วัน)": "-",
            "หมายเหตุ": "-",
        }
        st.session_state.leave_data = pd.concat(
            [st.session_state.leave_data, pd.DataFrame([new_row])],
            ignore_index=True,
        )
        st.session_state.leave_data["ที่"] = range(
            1, len(st.session_state.leave_data) + 1
        )
        st.success(f"เพิ่ม {new_name} เรียบร้อย!")
        st.rerun()

  with col_del:
    st.markdown("#### ลบพนักงานออก")
    current_names = st.session_state.leave_data["ชื่อ - สกุล"].tolist()
    if current_names:
      with st.form("del_emp_form"):
        target_to_delete = st.selectbox(
            "เลือกพนักงานที่ต้องการลบ", current_names
        )
        submitted_del = st.form_submit_button("🗑️ ลบพนักงานออก")
        if submitted_del and target_to_delete:
          st.session_state.leave_data = st.session_state.leave_data[
              st.session_state.leave_data["ชื่อ - สกุล"] != target_to_delete
          ].reset_index(drop=True)
          if not st.session_state.leave_data.empty:
            st.session_state.leave_data["ที่"] = range(
                1, len(st.session_state.leave_data) + 1
            )
          st.success(f"ลบ {target_to_delete} ออกแล้ว!")
          st.rerun()
    else:
      st.info("ไม่มีรายชื่อพนักงาน")

# -------------------------------------------------------------------------
# 5. ฟังก์ชันสร้าง Excel ใช้ฟอนต์ TH Sarabun PSK ตาม Template ราชการ
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

  col_widths = {
      "A": 6,
      "B": 25,
      "C": 35,
      "D": 15,
      "E": 15,
      "F": 10,
      "G": 10,
      "H": 10,
      "I": 18,
  }
  for col, width in col_widths.items():
    ws.column_dimensions[col].width = width

  # หัวข้อรายงาน
  ws["A1"] = (
      f"ตารางสรุปวันลา {emp_type} สำนักงานศึกษาธิการจังหวัดสมุทรสงคราม"
  )
  ws["A1"].font = font_header
  ws.merge_cells("A1:I1")
  ws["A1"].alignment = align_center

  ws["A2"] = f"ประจำเดือน {month} {year}"
  ws["A2"].font = font_header
  ws.merge_cells("A2:I2")
  ws["A2"].alignment = align_center

  # หัวตาราง 2 ชั้น
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
  ws["H4"] = "รวม\n(วัน)"
  ws.merge_cells("I4:I5")
  ws["I4"] = "หมายเหตุ"

  for r in [4, 5]:
    for col_idx in range(1, 10):
      cell = ws.cell(row=r, column=col_idx)
      cell.font = font_bold
      cell.alignment = align_center
      cell.border = thin_border

  # ข้อมูลพนักงาน
  start_row = 6
  last_row = start_row
  for i, row in df.iterrows():
    r = start_row + i
    ws[f"A{r}"] = i + 1
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
# 6. ปุ่มดาวน์โหลดไฟล์
# -------------------------------------------------------------------------
st.subheader("2. ดาวน์โหลดรายงาน")
if st.button("📥 สร้างไฟล์ Excel ตามแบบฟอร์มราชการ"):
  excel_file = generate_excel_template(
      edited_df, selected_month, selected_year, employee_type
  )

  data_as_bytes = excel_file.getvalue()
  st.download_button(
      label="💾 บันทึกไฟล์ .xlsx",
      data=data_as_bytes,
      file_name=(
          f"สรุปวันลา_{employee_type}_{selected_month}_{selected_year}.xlsx"
      ),
      mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  )
  st.success("สร้างไฟล์สำเร็จ! พร้อมนำไปใช้งานต่อได้ทันทีครับ")