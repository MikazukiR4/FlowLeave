import calendar
import json
from pathlib import Path
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
# Auto-save (B4): data persists even after page/browser restart
# -------------------------------------------------------------------------
SAVE_FILE = Path(__file__).resolve().parent / "flowleave_data.json"


def load_saved_data():
  """Read saved employee data from file .json (if file exists)"""
  try:
    with open(SAVE_FILE, encoding="utf-8") as f:
      raw = json.load(f)
    master = raw.get("master", {})
    records = {
      name: [
        {
          "start": datetime.strptime(rec["start"], "%Y-%m-%d").date(),
          "end": datetime.strptime(rec["end"], "%Y-%m-%d").date(),
          "category": rec["category"],
          "days": rec["days"],
          "note": rec.get("note", "-"),
        }
        for rec in recs
      ]
      for name, recs in raw.get("records", {}).items()
    }
    return master, records
  except (FileNotFoundError, json.JSONDecodeError, KeyError, ValueError):
    return {}, {}


def save_all_data():
  """Auto-save employee data + leave records to file .json"""
  payload = {
    "master": DEFAULT_EMPLOYEES,
    "records": {
      name: [
        {**rec, "start": str(rec["start"]), "end": str(rec["end"])}
        for rec in recs
      ]
      for name, recs in st.session_state.leave_records.items()
    },
  }
  try:
    with open(SAVE_FILE, "w", encoding="utf-8") as f:
      json.dump(payload, f, ensure_ascii=False, indent=2)
  except OSError:
    st.warning("Cannot save data file — check your folder permissions")


# -------------------------------------------------------------------------
# ส่วนหัว: เลือกเดือน ปี และประเภทพนักงาน
# -------------------------------------------------------------------------
with st.sidebar:
  st.header("⚙️ คำแนะนำการใช้งาน")
  st.markdown(
      "1. **เลือกงวดเวลา:** กำหนดเดือนและปีที่ต้องการสรุปรายงานด้านบน\n2."
      " **บันทึกวันลา:** คลิกเลือกแถวชื่อพนักงานในตาราง เพื่อเปิดหน้าต่างจัดการประวัติการลา\n3."
      " **หมายเหตุอัจฉริยะ:** พิมพ์คำว่า 'ครึ่งวัน' ระบบจะหัก 0.5 วันให้อัตโนมัติทันทีแบบเรียลไทม์\n4."
      " **ดาวน์โหลด:** ตรวจสอบข้อมูลแล้วกดสร้างไฟล์ Excel ได้เลย"
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
# Load data from save file so employee list edits persist across page opens
_saved_master, _saved_records = load_saved_data()
if _saved_master:
  DEFAULT_EMPLOYEES.clear()
  DEFAULT_EMPLOYEES.update(_saved_master)

if "leave_records" not in st.session_state:
  st.session_state.leave_records = {}

for name, recs in _saved_records.items():
  st.session_state.leave_records.setdefault(name, recs)

# Ensure employees of current type have a record list
# (do not clear data when changing employee type)
for emp in DEFAULT_EMPLOYEES.get(employee_type, []):
  st.session_state.leave_records.setdefault(emp["name"], [])

if "selected_employee" not in st.session_state:
  st.session_state.selected_employee = None


def build_summary_dataframe():
  emp_list = DEFAULT_EMPLOYEES.get(employee_type, [])
  rows = []
  for idx, emp in enumerate(emp_list, 1):
    name = emp["name"]
    position = emp["position"]
    records = st.session_state.leave_records.get(name, [])

    if not records:
      rows.append({
          "ที่": idx,
          "ชื่อ - สกุล": name,
          "ตำแหน่ง": position,
          "ตั้งแต่วันที่": "-",
          "ถึงวันที่": "-",
          "ลาป่วย": "-",
          "ลากิจ": "-",
          "รวม (วัน)": "-",
          "หมายเหตุ": "-",
      })
    else:
      starts = "\n".join([format_thai_date(r["start"]) for r in records])
      ends = "\n".join([format_thai_date(r["end"]) for r in records])
      # One line per range (keep "-" placeholder) so from/to/note align row-per-row
      notes = "\n".join(
          [
              f"{(r['note'] if r['note'] and r['note'] != '-' else '-')} "
              f"({r['days']} วัน)"
              for r in records
          ]
      )

      total_sick = sum([r["days"] for r in records if r["category"] == "ลาป่วย"])
      total_kit = sum([r["days"] for r in records if r["category"] == "ลากิจ"])
      total_all = total_sick + total_kit

      rows.append({
          "ที่": idx,
          "ชื่อ - สกุล": name,
          "ตำแหน่ง": position,
          "ตั้งแต่วันที่": starts,
          "ถึงวันที่": ends,
          "ลาป่วย": str(total_sick) if total_sick > 0 else "-",
          "ลากิจ": str(total_kit) if total_kit > 0 else "-",
          "รวม (วัน)": str(total_all) if total_all > 0 else "-",
          "หมายเหตุ": notes,
      })
  return pd.DataFrame(rows)


st.info(
    f"📂 กำลังจัดการข้อมูลของ: **{employee_type}** ประจำเดือน**{selected_month}"
    f" {selected_year}**"
)


# -------------------------------------------------------------------------
# Modal Dialog สำหรับจัดการประวัติการลา (ตรวจจับคำอัตโนมัติแบบเรียลไทม์)
# -------------------------------------------------------------------------
@st.dialog("📅 จัดการประวัติการลาบุคลากร", width="large")
def open_leave_dialog(emp_name):
  st.markdown(f"### 👤 บุคลากร: **{emp_name}**")

  current_records = st.session_state.leave_records.get(emp_name, [])

  st.markdown("#### 📋 รายการลาที่บันทึกไว้แล้วในเดือนนี้:")
  if not current_records:
    st.info("ยังไม่มีประวัติการลาในเดือนนี้")
  else:
    for idx, rec in enumerate(current_records):
      cols = st.columns([3, 2, 1, 3, 1])
      with cols[0]:
        st.markdown(f"**{rec['category']}** ({rec['days']} วัน)")
      with cols[1]:
        st.text(f"{format_thai_date(rec['start'])}")
      with cols[2]:
        st.text("ถึง")
      with cols[3]:
        st.text(f"{format_thai_date(rec['end'])} [{rec['note']}]")
      with cols[4]:
        if st.button("🗑️ ลบ", key=f"del_rec_{emp_name}_{idx}"):
          st.session_state.leave_records[emp_name].pop(idx)
          save_all_data()
          st.success("ลบรายการสำเร็จ!")
          st.rerun()
  st.markdown("---")
  st.markdown("#### ➕ เพิ่มรายการลาใหม่:")

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

  # ใช้ on_change หรือดึงค่าจาก session state ทันทีเพื่อให้ระบบรีเฟรชผลคำนวณแบบสดๆ
  custom_note = st.text_input(
      "หมายเหตุ / ลักษณะการลา",
      value="-",
      placeholder="เช่น ลาครึ่งวันเช้า, วันแรกเต็ม วันที่สองครึ่งวันบ่าย",
      key="d_custom_note",
  )

  is_date_invalid = d_end < d_start
  if is_date_invalid:
    st.error("⚠️ วันที่สิ้นสุดต้องไม่ก่อนวันที่เริ่มต้น")
  else:
    th_holidays = holidays.TH(years=selected_ce_year) if HAS_HOLIDAYS else set()
    working_days = 0
    curr = d_start
    while curr <= d_end:
      if curr.weekday() < 5 and curr not in th_holidays:
        working_days += 1
      curr += timedelta(days=1)

    calc_days = max(1, working_days) if working_days > 0 else 1

    # ตรวจจับคำว่าครึ่งวันหรือ 0.5 ทันทีที่พิมพ์โดยไม่ต้องกด Enter
    note_lower = custom_note.lower()
    if "ครึ่งวัน" in note_lower or "0.5" in note_lower:
      if calc_days >= 0.5:
        calc_days -= 0.5

    st.info(f"📌 **จำนวนวันที่คำนวณได้สำหรับรายการนี้:** `{calc_days} วัน`")

  if st.button(
      "➕ บันทึกเพิ่มรายการนี้",
      type="primary",
      disabled=is_date_invalid,
      use_container_width=True,
  ):
    new_entry = {
        "start": d_start,
        "end": d_end,
        "category": leave_category,
        "days": calc_days,
        "note": custom_note if custom_note else "-",
    }
    st.session_state.leave_records[emp_name].append(new_entry)
    save_all_data()
    st.success("เพิ่รายการสำเร็จ!")
    st.rerun()

# -------------------------------------------------------------------------
# 4. ตารางแสดงข้อมูลหลักและเรียก Dialog
# -------------------------------------------------------------------------
st.subheader("📊 ตารางสรุปข้อมูลบุคลากรและวันลาประจำเดือน")
st.markdown(
    "👉 **คลิกเลือกแถว (Row) ของพนักงานในตาราง เพื่อเปิดหน้าต่างจัดการประวัติการลา**"
)

df_display = build_summary_dataframe()
if df_display.empty:
  st.info(
      "📭 No personnel data found in this group. "
      "➕ Add new personnel from the menu below."
  )
event = st.dataframe(
    df_display,
    use_container_width=True,
    hide_index=True,
    selection_mode="single-row",
    on_select="rerun",
    key="employee_table_selection",
)

selected_rows = event.selection.rows if hasattr(event, "selection") else []
if selected_rows:
  row_idx = selected_rows[0]
  st.session_state.selected_employee = df_display.iloc[row_idx]["ชื่อ - สกุล"]

if st.session_state.selected_employee:
  open_leave_dialog(st.session_state.selected_employee)

# จัดการเพิ่ม/ลบรายชื่อบุคลากรในสังกัด
with st.expander("➕ เพิ่มรายชื่อพนักงานใหม่ / 🗑️ ลบรายชื่อพนักงาน"):
  add_col, del_col = st.columns(2)
  with add_col:
    with st.form("add_form", clear_on_submit=True):
      new_n = st.text_input("ชื่อ - สกุลใหม่")
      new_p = st.text_input("ตำแหน่งใหม่")
      if st.form_submit_button("➕ เพิ่มพนักงาน") and new_n:
        if new_n not in st.session_state.leave_records:
          save_all_data()
          st.session_state.leave_records[new_n] = []
          DEFAULT_EMPLOYEES[employee_type].append(
              {"name": new_n, "position": new_p or "-"}
          )
          st.success("เพิ่มรายชื่อสำเร็จ!")
          st.rerun()
        else:
          st.warning("มีชื่อนี้อยู่ในระบบแล้ว")
  with del_col:
    current_names = [emp["name"] for emp in DEFAULT_EMPLOYEES.get(employee_type, [])]
    if current_names:
      with st.form("del_form"):
        del_n = st.selectbox("เลือกรายชื่อที่ต้องการลบออก", current_names, key="del_select")
        if st.form_submit_button("🗑️ ลบพนักงานออก") and del_n:
          DEFAULT_EMPLOYEES[employee_type] = [
              emp
              for emp in DEFAULT_EMPLOYEES[employee_type]
              if emp["name"] != del_n
          ]
          save_all_data()
          if del_n in st.session_state.leave_records:
            del st.session_state.leave_records[del_n]
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
  align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)

  col_widths = {
      "A": 6,
      "B": 28,
      "C": 42,
      "D": 20,
      "E": 20,
      "F": 12,
      "G": 12,
      "H": 12,
      "I": 28,
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
      df=df_display,
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
      "Excel file is ready - click '💾 Save file .xlsx to computer' "
        "to save to your computer."
    )
