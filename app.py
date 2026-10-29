import calendar
from datetime import datetime, timedelta
import io
import os
import re
import openpyxl
from openpyxl.styles import Alignment, Border, Font, Side
from openpyxl.styles.fonts import DEFAULT_FONT
import pandas as pd
import streamlit as st

# ป้องกันข้อผิดพลาดกรณีไม่มีไลบรารี holidays
try:
  import holidays

  HAS_HOLIDAYS = True
except ImportError:
  HAS_HOLIDAYS = False

# ตั้งค่าหน้าจอ
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
            "position": (
                "พนักงานจ้างปฏิบัติงานนิเทศติดตาม\nและประเมินผล"
            ),
        },
        {
            "name": "นางสาววิชุดา  ศรีคร้ามครัน",
            "position": "พนักงานจ้างปฏิบัติงานธุรการ (กลุ่มพัฒนาฯ)",
        },
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

THAI_MONTH_FULL = [
    "",
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


def format_thai_date(date_obj):
  if not date_obj:
    return "-"
  day = date_obj.day
  month = THAI_MONTH_FULL[date_obj.month]
  year_thai = date_obj.year + 543
  return f"{day} {month} {year_thai}"


# -------------------------------------------------------------------------
# ส่วนหัว: เลือกเดือน ปี และประเภทพนักงาน
# -------------------------------------------------------------------------
with st.sidebar:
  st.header("⚙️ คำแนะนำการใช้งาน")
  st.markdown(
      "1. **เลือกงวดเวลา:** กำหนดเดือนและปีที่ต้องการสรุปรายงานด้านบน\n2."
      " **บันทึกวันลา:** **คลิกเลือกแถวชื่อพนักงาน** ในตาราง ระบบจะเด้งหน้าต่าง"
      " Datepicker ขึ้นมาให้ทันที\n3. **ดาวน์โหลด:** ตรวจสอบข้อมูลแล้วกดสร้างไฟล์"
      " Excel ได้เลย"
  )

thai_months = THAI_MONTH_FULL[1:]
now = datetime.now()

st.markdown("### 🗓️ กำหนดงวดรายงานประจำเดือน")
header_col1, header_col2, header_col3 = st.columns([2, 2, 3])
with header_col1:
  selected_month = st.selectbox(
      "ประจำเดือน", thai_months, index=now.month - 1
  )
with header_col2:
  selected_year = st.number_input("ปี พ.ศ.", value=now.year + 543)
with header_col3:
  employee_type = st.selectbox(
      "ประเภทพนักงาน", list(DEFAULT_EMPLOYEES.keys())
  )

selected_month_num = thai_months.index(selected_month) + 1
selected_ce_year = selected_year - 543
last_day_of_month = calendar.monthrange(selected_ce_year, selected_month_num)[
    1
]

month_start_date = datetime(selected_ce_year, selected_month_num, 1).date()
month_end_date = datetime(
    selected_ce_year, selected_month_num, last_day_of_month
).date()

# -------------------------------------------------------------------------
# Session State Management
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
st.markdown(
    "💡 *เคล็ดลับ: **คลิกเลือกแถวพนักงานในตารางด้านล่าง** เพื่อเปิดหน้าต่างบันทึกวันลาแบบป๊อปอัปได้ทันที!*"
)


# -------------------------------------------------------------------------
# Modal Dialog สำหรับเลือกวันลา (เด้งขึ้นมาเมื่อคลิกแถวในตาราง)
# -------------------------------------------------------------------------
@st.dialog("📅 บันทึก / อัปเดตวันลาบุคลากร", width="large")
def open_leave_dialog(emp_name):
  st.markdown(f"### 👤 บุคลากร: **{emp_name}**")

  leave_category = st.selectbox("ประเภทการลา", ["ลาป่วย", "ลากิจ"], key="d_cat")

  col_d1, col_d2 = st.columns(2)
  with col_d1:
    d_start = st.date_input(
        "ตั้งแต่วันที่",
        value=month_start_date,
        min_value=month_start_date,
        max_value=month_end_date,
        key="d_start",
    )
  with col_d2:
    d_end = st.date_input(
        "ถึงวันที่",
        value=month_start_date,
        min_value=month_start_date,
        max_value=month_end_date,
        key="d_end",
    )

  note_choice = st.selectbox(
      "หมายเหตุ / ลักษณะการลา",
      [
          "- (ปกติ / เต็มวัน)",
          "ลาครึ่งวันเช้า",
          "ลาครึ่งวันบ่าย",
          "✏️ พิมพ์ระบุหมายเหตุเอง...",
      ],
      key="d_note_choice",
  )
  final_note = "-"
  if note_choice == "✏️ พิมพ์ระบุเอง...":
    final_note = st.text_input(
        "ระบุข้อความหมายเหตุเพิ่มเติม", value="", key="d_custom"
    )
  elif note_choice != "- (ปกติ / เต็มวัน)":
    final_note = note_choice

  action_mode = st.radio(
      "โหมดการบันทึกข้อมูล",
      ["เพิ่มสะสมจากวันลาเดิม", "เขียนทับข้อมูลเดิมทั้งหมด"],
      horizontal=True,
      key="d_mode",
  )

  is_date_invalid = d_end < d_start
  if is_date_invalid:
    st.error("⚠️ ข้อผิดพลาด: วันที่สิ้นสุดต้องไม่ก่อนวันที่เริ่มต้น")
  else:
    th_holidays = holidays.TH(years=selected_ce_year) if HAS_HOLIDAYS else set()
    working_days = 0
    curr = d_start
    while curr <= d_end:
      if curr.weekday() < 5 and curr not in th_holidays:
        working_days += 1
      curr += timedelta(days=1)

    calc_days = max(1, working_days) if working_days > 0 else 1
    if d_start == d_end and "ครึ่งวัน" in final_note:
      calc_days = 0.5
    elif "ครึ่งวัน" in final_note and calc_days >= 0.5:
      calc_days -= 0.5

    st.info(f"📌 **จำนวนวันที่คำนวณได้:** `{calc_days} วัน`")

  col_btn1, col_btn2 = st.columns(2)
  with col_btn1:
    if st.button(
        "💾 บันทึกข้อมูล",
        type="primary",
        disabled=is_date_invalid,
        use_container_width=True,
    ):
      formatted_start = format_thai_date(d_start)
      formatted_end = format_thai_date(d_end)

      df = st.session_state.leave_data
      idx = df[df["ชื่อ - สกุล"] == emp_name].index
      if not idx.empty:
        i = idx[0]
        old_start = str(df.loc[i, "ตั้งแต่วันที่"])
        old_end = str(df.loc[i, "ถึงวันที่"])
        old_sick = str(df.loc[i, "ลาป่วย"])
        old_kit = str(df.loc[i, "ลากิจ"])
        old_total = str(df.loc[i, "รวม (วัน)"])
        old_note = str(df.loc[i, "หมายเหตุ"])

        if "เพิ่มสะสม" in action_mode and old_start != "-":
          df.loc[i, "ตั้งแต่วันที่"] = f"{old_start}, {formatted_start}"
          df.loc[i, "ถึงวันที่"] = f"{old_end}, {formatted_end}"
          curr_s = float(old_sick) if old_sick != "-" else 0.0
          curr_k = float(old_kit) if old_kit != "-" else 0.0
          curr_t = float(old_total) if old_total != "-" else 0.0
          if leave_category == "ลาป่วย":
            df.loc[i, "ลาป่วย"] = str(curr_s + calc_days)
          else:
            df.loc[i, "ลากิจ"] = str(curr_k + calc_days)
          df.loc[i, "รวม (วัน)"] = str(curr_t + calc_days)
          df.loc[i, "หมายเหตุ"] = (
              f"{old_note}, {final_note}"
              if old_note != "-" and final_note != "-"
              else (final_note if final_note != "-" else old_note)
          )
        else:
          df.loc[i, "ตั้งแต่วันที่"] = formatted_start
          df.loc[i, "ถึงวันที่"] = formatted_end
          df.loc[i, "ลาป่วย"] = (
              str(calc_days) if leave_category == "ลาป่วย" else "-"
          )
          df.loc[i, "ลากิจ"] = (
              str(calc_days) if leave_category == "ลากิจ" else "-"
          )
          df.loc[i, "รวม (วัน)"] = str(calc_days)
          df.loc[i, "หมายเหตุ"] = final_note

      st.success(f"บันทึกข้อมูลของ {emp_name} สำเร็จ!")
      st.rerun()
  with col_btn2:
    if st.button("❌ ปิดหน้าต่าง", use_container_width=True):
      st.rerun()


# -------------------------------------------------------------------------
# 4. ตารางแสดงข้อมูลหลัก (คลิกแถวเพื่อเด้ง Dialog)
# -------------------------------------------------------------------------
st.subheader("📊 ตารางสรุปข้อมูลบุคลากรและวันลา")
st.markdown(
    "👉 **คลิกเลือกแถว (Row) ของพนักงานที่ต้องการบันทึกวันลา ระบบจะเปิดหน้าต่างป๊อปอัปให้ทันที**"
)

event = st.dataframe(
    st.session_state.leave_data,
    use_container_width=True,
    hide_index=True,
    selection_mode="single-row",
    on_select="rerun",
    key="employee_table_selection",
)

# ตรวจสอบการคลิกเลือกแถวในตาราง
selected_rows = event.selection.rows if hasattr(event, "selection") else []
if selected_rows:
  row_idx = selected_rows[0]
  selected_emp_name = st.session_state.leave_data.iloc[row_idx]["ชื่อ - สกุล"]
  # เรียกใช้งาน Modal Dialog
  open_leave_dialog(selected_emp_name)

# จัดการเพิ่ม/ลบรายชื่อบุคลากร
with st.expander("➕ เพิ่มรายชื่อพนักงานใหม่ / 🗑️ ลบรายชื่อพนักงาน"):
  add_col, del_col = st.columns(2)
  with add_col:
    with st.form("add_form", clear_on_submit=True):
      new_n = st.text_input("ชื่อ - สกุลใหม่")
      new_p = st.text_input("ตำแหน่งใหม่")
      if st.form_submit_button("➕ เพิ่มพนักงาน") and new_n:
        new_row = {
            "ที่": len(st.session_state.leave_data) + 1,
            "ชื่อ - สกุล": new_n,
            "ตำแหน่ง": new_p or "-",
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
        st.success("เพิ่มรายชื่อสำเร็จ!")
        st.rerun()
  with del_col:
    current_names = st.session_state.leave_data["ชื่อ - สกุล"].tolist()
    if current_names:
      with st.form("del_form"):
        del_n = st.selectbox("เลือกรายชื่อที่ต้องการลบออก", current_names)
        if st.form_submit_button("🗑️ ลบพนักงานออก") and del_n:
          st.session_state.leave_data = st.session_state.leave_data[
              st.session_state.leave_data["ชื่อ - สกุล"] != del_n
          ].reset_index(drop=True)
          if not st.session_state.leave_data.empty:
            st.session_state.leave_data["ที่"] = range(
                1, len(st.session_state.leave_data) + 1
            )
          st.success("ลบรายชื่อสำเร็จ!")
          st.rerun()

# -------------------------------------------------------------------------
# 5. ฟังก์ชันสร้างไฟล์ Excel ตามแบบฟอร์มราชการ
# -------------------------------------------------------------------------
DEFAULT_FONT.name = "TH SarabunPSK"
DEFAULT_FONT.size = 14


def generate_excel_template(df, month, year, emp_type):
  wb = openpyxl.Workbook()
  ws = wb.active
  ws.title = "สรุปวันลา"

  font_header = Font(name="TH SarabunPSK", size=16, bold=True)
  font_bold = Font(name="TH SarabunPSK", size=14, bold=True)
  font_normal = Font(name="TH SarabunPSK", size=14)

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
      "B": 28,
      "C": 42,
      "D": 16,
      "E": 16,
      "F": 12,
      "G": 12,
      "H": 12,
      "I": 22,
  }
  for col, width in col_widths.items():
    ws.column_dimensions[col].width = width

  ws["A1"] = (
      f"ตารางสรุปวันลา {emp_type} สำนักงานศึกษาธิการจังหวัดสมุทรสงคราม"
  )
  ws["A1"].font = font_header
  ws.merge_cells("A1:I1")
  ws["A1"].alignment = align_center

  ws["A2"] = f"ประจำเดือน{month} {year}"
  ws["A2"].font = font_header
  ws.merge_cells("A2:I2")
  ws["A2"].alignment = align_center

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

  sig_row = last_row + 3
  ws.merge_cells(f"F{sig_row}:I{sig_row}")
  ws[f"F{sig_row}"] = "(ลงชื่อ)....................................................ผู้ตรวจสอบ"
  ws[f"F{sig_row}"].font = font_normal
  ws[f"F{sig_row}"].alignment = align_center

  output = io.BytesIO()
  wb.save(output)
  output.seek(0)
  return output


# -------------------------------------------------------------------------
# 6. ส่วนดาวน์โหลดรายงาน
# -------------------------------------------------------------------------
st.subheader("📥 ดาวน์โหลดรายงานประจำเดือน")
if st.button("📥 สร้างและดาวน์โหลดไฟล์ Excel ตามแบบฟอร์มราชการ", type="primary"):
  excel_file = generate_excel_template(
      df=st.session_state.leave_data,
      month=selected_month,
      year=selected_year,
      emp_type=employee_type,
  )
  data_as_bytes = excel_file.getvalue()
  st.download_button(
      label="💾 บันทึกไฟล์ .xlsx ลงเครื่อง",
      data=data_as_bytes,
      file_name=(
          f"สรุปวันลา_{employee_type}_{selected_month}_{selected_year}.xlsx"
      ),
      mime=(
          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
      ),
  )
  st.success(
      "สร้างไฟล์ Excel สำเร็จ! รูปแบบถูกต้องตามมาตรฐานงานสารบรรณราชการเรียบร้อย"
      "ครับ"
  )