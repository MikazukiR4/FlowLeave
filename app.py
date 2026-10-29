from datetime import datetime, timedelta
import re
import io
import os
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

st.set_page_config(
    page_title="ระบบสรุปวันลาประจำเดือน", page_icon="📅", layout="wide"
)
st.title("📌 ระบบสรุปวันลาประจำเดือน (Template ราชการ)")

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


with st.sidebar:
  st.header("⚙️ คำแนะนำการใช้งาน")
  st.markdown(
      "1. เลือกเดือน ปี และประเภทพนักงาน\n2. เลือกช่วงวันลา (ระบบป้องกันการเลือกวันย้อนหลัง)\n3. เลือกหมายเหตุ/ลักษณะการลา\n4. กดสร้างและดาวน์โหลดไฟล์ Excel"
  )

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

selected_month_num = thai_months.index(selected_month) + 1
selected_ce_year = selected_year - 543
default_picker_date = datetime(selected_ce_year, selected_month_num, 1)

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

st.subheader("📅 บันทึก/อัปเดตวันลาประจำเดือน")
with st.container(border=True):
  current_names = st.session_state.leave_data["ชื่อ - สกุล"].tolist()
  if current_names:
    tab_calendar, tab_quick = st.tabs([
        "🗓️ เลือกช่วงวันที่จากปฏิทิน (Datepicker)",
        "⚡ พิมพ์ระบุวันที่แบบย่อ (เช่น 1 หรือ 1-3)",
    ])

    with tab_calendar:
      col_p1, col_p2, col_p3, col_p4 = st.columns(4)
      with col_p1:
        target_emp = st.selectbox("เลือกบุคลากร", current_names, key="c_emp")
      with col_p2:
        leave_category = st.selectbox(
            "ประเภทการลา", ["ลาป่วย", "ลากิจ"], key="c_cat"
        )
      with col_p3:
        d_start = st.date_input(
            "ตั้งแต่วันที่", value=default_picker_date, key="c_start"
        )
      with col_p4:
        d_end = st.date_input(
            "ถึงวันที่", value=default_picker_date, key="c_end"
        )

      is_date_invalid = d_end < d_start
      if is_date_invalid:
        st.error(
            "⚠️ ข้อผิดพลาด: วันที่สิ้นสุดต้องไม่ก่อนวันที่เริ่มต้น"
            " กรุณาเลือกใหม่"
        )

      col_n1, col_n2 = st.columns([2, 1])
      with col_n1:
        note_choice = st.selectbox(
            "หมายเหตุ / ลักษณะการลา",
            [
                "- (ปกติ / เต็มวัน)",
                "ลาครึ่งวันเช้า",
                "ลาครึ่งวันบ่าย",
                "✏️ พิมพ์ระบุหมายเหตุเอง...",
            ],
            key="c_note_choice",
        )
        final_note = "-"
        if note_choice == "✏️ พิมพ์ระบุเอง...":
          final_note = st.text_input(
              "ระบุข้อความหมายเหตุ", value="", key="c_custom_note"
          )
        elif note_choice != "- (ปกติ / เต็มวัน)":
          final_note = note_choice

      with col_n2:
        action_mode = st.radio(
            "รูปแบบการบันทึก",
            ["เพิ่มสะสมจากเดิม", "เขียนทับใหม่"],
            horizontal=True,
            key="c_mode",
        )

      calc_days = 0
      if not is_date_invalid:
        th_holidays = holidays.TH(years=d_start.year) if HAS_HOLIDAYS else set()
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

        st.success(
            f"📌 **ระบบคำนวณจำนวนวันรอบนี้ให้อัตโนมัติ:** `{calc_days}` วัน"
        )

      if st.button(
          "💾 บันทึกวันลาลงตาราง", key="btn_c", disabled=is_date_invalid
      ):
        formatted_start = format_thai_date(d_start)
        formatted_end = format_thai_date(d_end)

        df = st.session_state.leave_data
        idx = df[df["ชื่อ - สกุล"] == target_emp].index
        if not idx.empty:
          i = idx[0]
          old_start = str(df.loc[i, "ตั้งแต่วันที่"])
          old_end = str(df.loc[i, "ถึงวันที่"])
          old_sick = str(df.loc[i, "ลาป่วย"])
          old_kit = str(df.loc[i, "ลากิจ"])
          old_total = str(df.loc[i, "รวม (วัน)"])
          old_note = str(df.loc[i, "หมายเหตุ"])

          if action_mode == "เพิ่มสะสมจากเดิม" and old_start != "-":
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

          st.success(f"บันทึกวันลาของ {target_emp} ({calc_days} วัน) สำเร็จ!")
          st.rerun()

    with tab_quick:
      st.markdown("💡 *พิมพ์ระบุวันที่แบบย่อ เช่น `1` หรือ `1-3`*")
      q_emp = st.selectbox("เลือกบุคลากร", current_names, key="q_emp")
      q_cat = st.selectbox("ประเภทการลา", ["ลาป่วย", "ลากิจ"], key="q_cat")

      col_q1, col_q2 = st.columns([1, 1])
      with col_q1:
        day_input_str = st.text_input("ระบุวันที่ (เช่น 1 หรือ 1-3)", value="1")
      with col_q2:
        q_note_choice = st.selectbox(
            "หมายเหตุ / ลักษณะการลา",
            [
                "- (ปกติ / เต็มวัน)",
                "ลาครึ่งวันเช้า",
                "ลาครึ่งวันบ่าย",
                "✏️ พิมพ์ระบุเอง...",
            ],
            key="q_note_choice",
        )
        q_final_note = "-"
        if q_note_choice == "✏️ พิมพ์ระบุเอง...":
          q_final_note = st.text_input(
              "ระบุข้อความหมายเหตุ", value="", key="q_custom_note"
          )
        elif q_note_choice != "- (ปกติ / เต็มวัน)":
          q_final_note = q_note_choice

      if st.button("🚀 ประมวลผลและบันทึกด่วน", key="btn_q"):
        try:
          numbers = [int(s) for s in re.findall(r"\d+", day_input_str)]
          if not numbers:
            st.error("กรุณาระบุตัวเลขวันที่ให้ถูกต้อง เช่น 1 หรือ 1-3")
          else:
            start_day = numbers[0]
            end_day = numbers[-1]

            d_start_q = datetime(selected_ce_year, selected_month_num, start_day)
            d_end_q = datetime(selected_ce_year, selected_month_num, end_day)

            if d_end_q < d_start_q:
              st.error("⚠️ วันที่สิ้นสุดต้องไม่ก่อนวันที่เริ่มต้น")
            else:
              formatted_start = format_thai_date(d_start_q)
              formatted_end = format_thai_date(d_end_q)

              th_holidays = (
                  holidays.TH(years=selected_ce_year)
                  if HAS_HOLIDAYS
                  else set()
              )
              w_days = 0
              curr = d_start_q
              while curr <= d_end_q:
                if curr.weekday() < 5 and curr not in th_holidays:
                  w_days += 1
                curr += timedelta(days=1)

              calc_d = max(1, w_days) if w_days > 0 else 1
              if start_day == end_day and "ครึ่งวัน" in q_final_note:
                calc_d = 0.5
              elif "ครึ่งวัน" in q_final_note and calc_d >= 0.5:
                calc_d -= 0.5

              df = st.session_state.leave_data
              idx = df[df["ชื่อ - สกุล"] == q_emp].index
              if not idx.empty:
                i = idx[0]
                old_start = str(df.loc[i, "ตั้งแต่วันที่"])
                old_end = str(df.loc[i, "ถึงวันที่"])
                old_sick = str(df.loc[i, "ลาป่วย"])
                old_kit = str(df.loc[i, "ลากิจ"])
                old_total = str(df.loc[i, "รวม (วัน)"])
                old_note = str(df.loc[i, "หมายเหตุ"])

                if old_start != "-":
                  df.loc[i, "ตั้งแต่วันที่"] = (
                      f"{old_start}, {formatted_start}"
                  )
                  df.loc[i, "ถึงวันที่"] = f"{old_end}, {formatted_end}"

                  curr_s = float(old_sick) if old_sick != "-" else 0.0
                  curr_k = float(old_kit) if old_kit != "-" else 0.0
                  curr_t = float(old_total) if old_total != "-" else 0.0

                  if q_cat == "ลาป่วย":
                    df.loc[i, "ลาป่วย"] = str(curr_s + calc_d)
                  else:
                    df.loc[i, "ลากิจ"] = str(curr_k + calc_d)
                  df.loc[i, "รวม (วัน)"] = str(curr_t + calc_d)
                  df.loc[i, "หมายเหตุ"] = (
                      f"{old_note}, {q_final_note}"
                      if old_note != "-" and q_final_note != "-"
                      else (q_final_note if q_final_note != "-" else old_note)
                  )
                else:
                  df.loc[i, "ตั้งแต่วันที่"] = formatted_start
                  df.loc[i, "ถึงวันที่"] = formatted_end
                  df.loc[i, "ลาป่วย"] = (
                      str(calc_d) if q_cat == "ลาป่วย" else "-"
                  )
                  df.loc[i, "ลากิจ"] = str(calc_d) if q_cat == "ลากิจ" else "-"
                  df.loc[i, "รวม (วัน)"] = str(calc_d)
                  df.loc[i, "หมายเหตุ"] = q_final_note

                st.success(
                    f"บันทึกวันลาวันที่ {day_input_str} ({calc_d} วัน) สำเร็จ!"
                )
                st.rerun()
        except Exception as e:
          st.error(f"เกิดข้อผิดพลาดในการประมวลผล: {e}")
  else:
    st.warning("ไม่มีรายชื่อพนักงานในระบบ")

st.subheader("1. ตารางสรุปข้อมูลบุคลากรและวันลาปัจจุบัน")
edited_df = st.data_editor(
    st.session_state.leave_data,
    num_rows="dynamic",
    use_container_width=True,
    key="employee_table_editor",
)
st.session_state.leave_data = edited_df

with st.expander("➕ เพิ่ม / 🗑️ ลบ รายชื่อพนักงาน"):
  col_add, col_del = st.columns(2)
  with col_add:
    with st.form("add_form", clear_on_submit=True):
      new_n = st.text_input("ชื่อ - สกุลใหม่")
      new_p = st.text_input("ตำแหน่งใหม่")
      if st.form_submit_button("เพิ่มพนักงาน") and new_n:
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
        st.success("เพิ่มสำเร็จ!")
        st.rerun()
  with col_del:
    if current_names:
      with st.form("del_form"):
        del_n = st.selectbox("เลือกคนที่จะลบ", current_names)
        if st.form_submit_button("ลบพนักงานออก") and del_n:
          st.session_state.leave_data = st.session_state.leave_data[
              st.session_state.leave_data["ชื่อ - สกุล"] != del_n
          ].reset_index(drop=True)
          if not st.session_state.leave_data.empty:
            st.session_state.leave_data["ที่"] = range(
                1, len(st.session_state.leave_data) + 1
            )
          st.success("ลบสำเร็จ!")
          st.rerun()

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
  ws[f"F{sig_row}"] = (
      "(ลงชื่อ)....................................................ผู้ตรวจสอบ"
  )
  ws[f"F{sig_row}"].font = font_normal
  ws[f"F{sig_row}"].alignment = align_center

  output = io.BytesIO()
  wb.save(output)
  output.seek(0)
  return output


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
      mime=(
          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
      ),
  )
  st.success("สร้างไฟล์สำเร็จ! พร้อมนำไปใช้งานต่อได้ทันทีครับ")