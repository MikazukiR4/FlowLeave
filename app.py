import calendar
import copy
import io
import json
import os
from datetime import date, datetime, timedelta, timezone
from pathlib import Path

import openpyxl
import pandas as pd
import requests
import streamlit as st
from openpyxl.styles import Alignment, Border, Font, Side
from openpyxl.styles.fonts import DEFAULT_FONT

try:
  import holidays
except ImportError:
  holidays = None

# ตั้งค่าฟอนต์เริ่มต้นสำหรับ openpyxl
DEFAULT_FONT.name = "TH SarabunPSK"
DEFAULT_FONT.size = 14

st.set_page_config(
    page_title="ระบบสรุปวันลาประจำเดือน", page_icon="📅", layout="wide"
)
st.title("📌 ระบบสรุปวันลาประจำเดือน (Template ราชการ)")

if msg := st.session_state.pop("flash", None):
  st.toast(msg)

# -------------------------------------------------------------------------
# 1. ค่าคงที่
# -------------------------------------------------------------------------
SEED_EMPLOYEES = {
    "พนักงานจ้างเหมาเอกชน": [
        {"name": "นางสาวณัฐพร  มิตรพิทักษ์", "position": "พนักงานจ้างปฏิบัติงานนโยบายและแผน"},
        {"name": "นางสาวณัฐธิชา  แถวทัด", "position": "พนักงานจ้างปฏิบัติงานนิเทศติดตาม\nและประเมินผล"},
        {"name": "นางสาววิชุดา  ศรีคร้ามครัน", "position": "พนักงานจ้างปฏิบัติงานธุรการ (กลุ่มพัฒนาฯ)"},
        {"name": "นางสาวยัสมี  ดาโอะ", "position": "พนักงานจ้างปฏิบัติงานธุรการ"},
        {"name": "นางสาวภาวินีย์  ตันสกุล", "position": "พนักงานจ้างทำความสะอาด"},
        {"name": "นายภานุวัฒน์  จันทร์ชีวา", "position": "พนักงานจ้างปฏิบัติงานขับรถยนต์"},
        {"name": "นางสาวหัทยาภักดิ์ แสงแก้ว", "position": "พนักงานจ้างปฏิบัติงานการเงินและบัญชี"},
        {"name": "นายหนึ่ง  สุนทรรัตน์", "position": "พนักงานจ้างปฏิบัติงานรักษาความปลอดภัย"},
    ],
    "พนักงานราชการ": [
        {"name": "นางสาวอมรารัตน์  พูลสวัสดิ์", "position": "พนักงานจัดการงานบุคคล"},
        {"name": "นายสมชาย  สุพรดวงชัย", "position": "พนักงานจัดการงานทั่วไป"},
    ],
}

THAI_MONTHS = [
    "มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน",
    "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม",
]
COLUMNS = [
    "ที่", "ชื่อ - สกุล", "ตำแหน่ง", "ตั้งแต่วันที่", "ถึงวันที่",
    "ลาป่วย", "ลากิจ", "รวม (วัน)", "หมายเหตุ",
]
FULL, AM, PM = "เต็มวัน", "ครึ่งวันเช้า", "ครึ่งวันบ่าย"

# -------------------------------------------------------------------------
# 2. ที่เก็บข้อมูล
# -------------------------------------------------------------------------
SUPABASE_URL = os.getenv("SUPABASE_URL", "").rstrip("/")
SUPABASE_KEY = os.getenv("SUPABASE_KEY", "")
USE_SUPABASE = bool(SUPABASE_URL and SUPABASE_KEY)
DATA_DIR = os.getenv("DATA_DIR")
SAVE_FILE = Path(DATA_DIR or Path(__file__).resolve().parent) / "flowleave_data.json"

if USE_SUPABASE:
  STORAGE_LABEL = "🟢 Supabase (ข้อมูลถาวร)"
elif DATA_DIR:
  STORAGE_LABEL = f"🟢 Persistent Disk ({SAVE_FILE})"
else:
  STORAGE_LABEL = "🟡 ไฟล์ในเครื่องเซิร์ฟเวอร์"

EPHEMERAL = bool(os.getenv("RENDER")) and not (USE_SUPABASE or DATA_DIR)


def _sb_headers():
  h = {"apikey": SUPABASE_KEY, "Content-Type": "application/json"}
  if not SUPABASE_KEY.startswith("sb_"):
    h["Authorization"] = f"Bearer {SUPABASE_KEY}"
  return h


@st.cache_data(ttl=15, show_spinner=False)
def read_raw():
  try:
    if USE_SUPABASE:
      resp = requests.get(
          f"{SUPABASE_URL}/rest/v1/app_state",
          params={"id": "eq.1", "select": "data"},
          headers=_sb_headers(),
          timeout=5,
      )
      if resp.status_code == 200:
        rows = resp.json()
        return rows[0]["data"] if rows else {}
    elif SAVE_FILE.exists():
      return json.loads(SAVE_FILE.read_text(encoding="utf-8"))
  except Exception:
    pass
  return {}


def write_raw(payload):
  if USE_SUPABASE:
    resp = requests.post(
        f"{SUPABASE_URL}/rest/v1/app_state",
        json={
            "id": 1,
            "data": payload,
            "updated_at": datetime.now(timezone.utc).isoformat(),
        },
        headers={**_sb_headers(), "Prefer": "resolution=merge-duplicates"},
        timeout=10,
    )
    resp.raise_for_status()
  else:
    SAVE_FILE.parent.mkdir(parents=True, exist_ok=True)
    tmp = SAVE_FILE.with_suffix(".tmp")
    tmp.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    tmp.replace(SAVE_FILE)
  read_raw.clear()


def parse_payload(raw):
  if not raw or not isinstance(raw, dict):
    return copy.deepcopy(SEED_EMPLOYEES), {}
  employees = raw.get("master") or copy.deepcopy(SEED_EMPLOYEES)
  records = {}
  for name, recs in raw.get("records", {}).items():
    records[name] = []
    if isinstance(recs, list):
      for r in recs:
        try:
          note = str(r.get("note") or "").strip()
          records[name].append({
              "start": date.fromisoformat(r["start"]),
              "end": date.fromisoformat(r["end"]),
              "category": r["category"],
              "days": float(r["days"]),
              "first": r.get("first", FULL),
              "last": r.get("last", FULL),
              "note": "" if note == "-" else note,
          })
        except Exception:
          continue
  return employees, records


def build_payload():
  return {
      "master": employees,
      "records": {
          name: [
              {**r, "start": r["start"].isoformat(), "end": r["end"].isoformat()}
              for r in recs
          ]
          for name, recs in records.items()
          if recs
      },
  }


def save_data():
  try:
    write_raw(build_payload())
  except Exception as e:
    st.session_state.flash = f"⚠️ บันทึกไม่สำเร็จ: {e}"


def rerun_with(msg, open_for=None):
  st.session_state.setdefault("flash", msg)
  if open_for:
    st.session_state.open_for = open_for
  st.rerun()


try:
  employees, records = parse_payload(read_raw())
except Exception:
  employees, records = copy.deepcopy(SEED_EMPLOYEES), {}


# -------------------------------------------------------------------------
# 3. ฟังก์ชันวันที่และการคำนวณ
# -------------------------------------------------------------------------
@st.cache_resource
def th_holidays(year):
  return holidays.TH(years=year) if holidays else {}


def is_workday(d):
  return d.weekday() < 5 and d not in th_holidays(d.year)


def calc_leave_days(start, end, first=FULL, last=FULL):
  span = (start + timedelta(days=i) for i in range((end - start).days + 1))
  workdays = sum(1 for d in span if is_workday(d))
  days = float(workdays)
  if start == end:
    if first == AM or first == PM:
      days -= 0.5
  else:
    if first != FULL and is_workday(start):
      days -= 0.5
    if last != FULL and is_workday(end):
      days -= 0.5
  return days, workdays


def format_thai_date(d):
  return f"{d.day} {THAI_MONTHS[d.month - 1]} {d.year + 543}"


def format_range(r):
  if r["start"] == r["end"]:
    return format_thai_date(r["start"])
  return f"{format_thai_date(r['start'])} ถึง {format_thai_date(r['end'])}"


def fmt_days(v):
  return f"{v:g}" if v else "-"


def describe(r):
  parts = []
  if r["start"] == r["end"]:
    if r["first"] != FULL:
      parts.append(r["first"])
  else:
    if r["first"] != FULL:
      parts.append(r["first"])
    if r["last"] != FULL:
      parts.append(r["last"])
  if r["note"]:
    parts.append(r["note"])
  return ", ".join(parts)


def month_records(name):
  return sorted(
      (r for r in records.get(name, []) if month_start <= r["start"] <= month_end),
      key=lambda r: r["start"],
  )


# -------------------------------------------------------------------------
# 4. Sidebar: คำแนะนำ + สำรอง/กู้คืนข้อมูล
# -------------------------------------------------------------------------
st.session_state.setdefault("restore_ver", 0)
with st.sidebar:
  st.header("⚙️ คำแนะนำการใช้งาน")
  st.markdown(
      "1. **เลือกงวดเวลา:** กำหนดเดือนและปีที่ต้องการสรุปรายงาน\n"
      "2. **บันทึกวันลา:** แตะแถวชื่อพนักงานในตารางเพื่อเปิดหน้าต่างจัดการ\n"
      "3. **ครึ่งวัน:** ถ้าระยะเวลามากกว่า 1 วัน ระบบให้เลือกช่วงเวลาวันแรก/วันสุดท้ายได้\n"
      "4. **ดาวน์โหลด:** ตรวจสอบตารางแล้วกดดาวน์โหลดไฟล์ Excel"
  )

  st.divider()
  st.subheader("💾 สำรอง / กู้คืนข้อมูล")
  st.caption(f"ที่เก็บข้อมูล: {STORAGE_LABEL}")
  if EPHEMERAL:
    st.error(
        "ข้อมูลตอนนี้เก็บแบบชั่วคราว **จะหายเมื่อเซิร์ฟเวอร์รีสตาร์ท** "
        "กรุณาตั้งค่า Supabase หรือกดสำรองข้อมูลเป็นประจำ"
    )

  st.download_button(
      "⬇️ ดาวน์โหลดไฟล์สำรอง (.json)",
      data=json.dumps(build_payload(), ensure_ascii=False, indent=2),
      file_name=f"flowleave_backup_{date.today().isoformat()}.json",
      mime="application/json",
      use_container_width=True,
  )

  up = st.file_uploader(
      "กู้คืนจากไฟล์สำรอง", type="json",
      key=f"restore_{st.session_state.restore_ver}",
  )
  if up is not None:
    st.warning("การกู้คืนจะ **แทนที่ข้อมูลปัจจุบันทั้งหมด**")
    if st.button("♻️ ยืนยันกู้คืนข้อมูล", use_container_width=True):
      try:
        payload = json.loads(up.getvalue().decode("utf-8"))
        if not isinstance(payload, dict) or "master" not in payload:
          raise ValueError("ไม่ใช่ไฟล์สำรองของระบบนี้")
        parse_payload(payload)
        write_raw(payload)
      except Exception as e:
        st.error(f"กู้คืนไม่สำเร็จ: {e}")
      else:
        st.session_state.restore_ver += 1
        rerun_with("✅ กู้คืนข้อมูลสำเร็จ")

# -------------------------------------------------------------------------
# 5. ส่วนหัว: งวดรายงาน
# -------------------------------------------------------------------------
today = date.today()
st.markdown("### 🗓️ กำหนดงวดรายงานประจำเดือน")
c1, c2, c3 = st.columns([2, 2, 3])
selected_month = c1.selectbox("ประจำเดือน", THAI_MONTHS, index=today.month - 1)
selected_year = int(c2.number_input(
    "ปี พ.ศ.", value=today.year + 543, min_value=2560, max_value=2700,
    step=1, format="%d",
))
employee_type = c3.selectbox("ประเภทพนักงาน", list(employees.keys()))

month_num = THAI_MONTHS.index(selected_month) + 1
ce_year = selected_year - 543
month_start = date(ce_year, month_num, 1)
month_end = date(ce_year, month_num, calendar.monthrange(ce_year, month_num)[1])

st.info(
    f"📂 กำลังจัดการข้อมูลของ: **{employee_type}** "
    f"ประจำเดือน**{selected_month} {selected_year}**"
)


# -------------------------------------------------------------------------
# 6. ตารางสรุปบน UI
# -------------------------------------------------------------------------
def build_summary_dataframe():
  rows = []
  for idx, emp in enumerate(employees.get(employee_type, []), 1):
    recs = month_records(emp["name"])
    sick = sum(r["days"] for r in recs if r["category"] == "ลาป่วย")
    kit = sum(r["days"] for r in recs if r["category"] == "ลากิจ")
    lines = []
    for r in recs:
      desc = describe(r)
      lines.append(
          f"{r['category']} {r['days']:g} วัน" + (f" ({desc})" if desc else "")
      )
    rows.append({
        "ที่": idx,
        "ชื่อ - สกุล": emp["name"],
        "ตำแหน่ง": emp["position"],
        "ตั้งแต่วันที่": "\n".join(format_thai_date(r["start"]) for r in recs) or "-",
        "ถึงวันที่": "\n".join(format_thai_date(r["end"]) for r in recs) or "-",
        "ลาป่วย": fmt_days(sick),
        "ลากิจ": fmt_days(kit),
        "รวม (วัน)": fmt_days(sick + kit),
        "หมายเหตุ": "\n".join(lines) or "-",
    })
  return pd.DataFrame(rows, columns=COLUMNS)


# -------------------------------------------------------------------------
# 7. ป๊อปอัปจัดการวันลา
# -------------------------------------------------------------------------
@st.dialog("📅 จัดการประวัติการลาบุคลากร", width="large")
def leave_dialog(emp_name):
  st.markdown(f"### 👤 {emp_name}")
  st.markdown(f"#### 📋 รายการลาเดือน{selected_month} {selected_year}")

  recs = month_records(emp_name)
  if not recs:
    st.info("ยังไม่มีประวัติการลาในเดือนนี้")
  for i, rec in enumerate(recs):
    a, b = st.columns([9, 1], vertical_alignment="center")
    a.markdown(f"**{rec['category']} {rec['days']:g} วัน** · {format_range(rec)}")
    if desc := describe(rec):
      a.caption(desc)
    if b.button("🗑️", key=f"del_{emp_name}_{i}", help="ลบรายการนี้"):
      records[emp_name].remove(rec)
      save_data()
      rerun_with("🗑️ ลบรายการสำเร็จ", open_for=emp_name)

  st.divider()
  st.markdown("#### ➕ เพิ่มรายการลาใหม่")
  category = st.radio("ประเภทการลา", ["ลาป่วย", "ลากิจ"], horizontal=True, key="d_cat")

  a, b = st.columns(2)
  d_start = a.date_input(
      "ตั้งแต่วันที่", value=month_start, min_value=month_start,
      max_value=month_end, format="DD/MM/YYYY", key="d_start",
  )
  d_end = b.date_input(
      "ถึงวันที่", value=month_start, min_value=month_start,
      max_value=month_end, format="DD/MM/YYYY", key="d_end",
  )

  if d_end < d_start:
    st.error("⚠️ วันที่สิ้นสุดต้องไม่ก่อนวันที่เริ่มต้น")
    return

  if d_start == d_end:
    first = st.radio("ลักษณะการลา", [FULL, AM, PM], horizontal=True, key="d_single")
    last = FULL
  else:
    st.markdown("##### ⏱️ รายละเอียดช่วงเวลาที่ลา")
    c1, c2 = st.columns(2)
    first = c1.selectbox(
        f"วันแรก ({format_thai_date(d_start)})",
        [FULL, PM],
        format_func=lambda x: "เต็มวัน (เริ่มเช้า)" if x == FULL else "ครึ่งวันบ่าย (เริ่มบ่าย)",
        key="d_first",
    )
    last = c2.selectbox(
        f"วันสุดท้าย ({format_thai_date(d_end)})",
        [FULL, AM],
        format_func=lambda x: "เต็มวัน (ลาถึงเย็น)" if x == FULL else "ครึ่งวันเช้า (กลับมาทำงานบ่าย)",
        key="d_last",
    )

  note = st.text_input(
      "หมายเหตุเพิ่มเติม (ไม่บังคับ)",
      placeholder="เช่น ไปพบแพทย์, ธุระส่วนตัว",
      key="d_note",
  ).strip()

  days, workdays = calc_leave_days(d_start, d_end, first, last)
  span = (d_end - d_start).days + 1

  if workdays == 0:
    st.warning("⚠️ ช่วงวันที่เลือกเป็นวันเสาร์-อาทิตย์/วันหยุดราชการทั้งหมด จึงไม่มีวันลาที่ต้องบันทึก")
    return

  detail = [f"ช่วงที่เลือก {span} วัน"]
  if span - workdays:
    detail.append(f"หักวันหยุด {span - workdays} วัน")
  if workdays - days:
    detail.append(f"หักครึ่งวัน {workdays - days:g} วัน")
  st.info(f"📌 **รวม {days:g} วัน**  \n{' · '.join(detail)}")

  if st.button("➕ บันทึกเพิ่มรายการนี้", type="primary", use_container_width=True):
    records.setdefault(emp_name, []).append({
        "start": d_start, "end": d_end, "category": category,
        "days": days, "first": first, "last": last, "note": note,
    })
    save_data()
    for k in ("d_single", "d_first", "d_last", "d_note"):
      st.session_state.pop(k, None)
    rerun_with("✅ เพิ่มรายการลาสำเร็จ", open_for=emp_name)


# -------------------------------------------------------------------------
# 8. แสดงตาราง + เปิดป๊อปอัป
# -------------------------------------------------------------------------
st.subheader("📊 ตารางสรุปข้อมูลบุคลากรและวันลาประจำเดือน")
st.markdown("👉 **แตะแถวพนักงานในตาราง เพื่อเปิดหน้าต่างจัดการประวัติการลา**")

df_display = build_summary_dataframe()
if df_display.empty:
  st.info("📭 ยังไม่มีบุคลากรในกลุ่มนี้ — เพิ่มได้จากเมนูด้านล่าง")

st.session_state.setdefault("table_ver", 0)
event = st.dataframe(
    df_display,
    use_container_width=True,
    hide_index=True,
    selection_mode="single-row",
    on_select="rerun",
    key=f"employee_table_{st.session_state.table_ver}",
)
if event.selection.rows:
  row = event.selection.rows[0]
  if row < len(df_display):
    st.session_state.open_for = df_display.iloc[row]["ชื่อ - สกุล"]
  st.session_state.table_ver += 1
  st.rerun()

if name := st.session_state.pop("open_for", None):
  leave_dialog(name)

# -------------------------------------------------------------------------
# 9. เพิ่ม / ลบรายชื่อพนักงาน
# -------------------------------------------------------------------------
with st.expander("➕ เพิ่มรายชื่อพนักงานใหม่ / 🗑️ ลบรายชื่อพนักงาน"):
  add_col, del_col = st.columns(2)

  with add_col:
    with st.form("add_form", clear_on_submit=True):
      new_n = st.text_input("ชื่อ - สกุลใหม่").strip()
      new_p = st.text_input("ตำแหน่งใหม่").strip()
      if st.form_submit_button("➕ เพิ่มพนักงาน") and new_n:
        all_names = {e["name"] for group in employees.values() for e in group}
        if new_n in all_names:
          st.warning("มีชื่อนี้อยู่ในระบบแล้ว")
        else:
          employees.setdefault(employee_type, []).append(
              {"name": new_n, "position": new_p or "-"}
          )
          save_data()
          rerun_with("✅ เพิ่มรายชื่อสำเร็จ")

  with del_col:
    names = [e["name"] for e in employees.get(employee_type, [])]
    if names:
      del_n = st.selectbox("เลือกรายชื่อที่ต้องการลบ", names, key="del_select")
      n_recs = len(records.get(del_n, []))
      st.warning(
          f"จะลบ **{del_n}** ออกจากระบบ"
          + (f" พร้อมประวัติการลา **{n_recs} รายการ** (ทุกเดือน)" if n_recs else "")
      )
      confirm = st.checkbox("ยืนยันการลบ (ย้อนกลับไม่ได้)", key=f"confirm_{del_n}")
      if st.button("🗑️ ลบพนักงานออก", disabled=not confirm):
        employees[employee_type] = [
            e for e in employees[employee_type] if e["name"] != del_n
        ]
        records.pop(del_n, None)
        save_data()
        rerun_with("🗑️ ลบรายชื่อสำเร็จ")

# -------------------------------------------------------------------------
# 10. สร้างไฟล์ Excel ตามแบบฟอร์มราชการ (หัวตารางและชื่อผู้ตรวจสอบตรงต้นฉบับเป๊ะ)
# -------------------------------------------------------------------------
def format_excel_date(d):
  short_months = [
      "", "ม.ค.", "ก.พ.", "มี.ค.", "เม.ย.", "พ.ค.", "มิ.ย.",
      "ก.ค.", "ส.ค.", "ก.ย.", "ต.ค.", "พ.ย.", "ธ.ค."
  ]
  return f"{d.day} {short_months[d.month]} {str(d.year + 543)[-2:]}"


def generate_excel(employees_list, records_dict, month_str, year_num, emp_type):
  wb = openpyxl.Workbook()
  ws = wb.active
  ws.title = "สรุปวันลา"

  ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT
  ws.page_setup.paperSize = ws.PAPERSIZE_A4
  ws.sheet_properties.pageSetUpPr.fitToPage = True
  ws.page_setup.fitToWidth = 1
  ws.page_setup.fitToHeight = 0

  side = Side(style="thin", color="000000")
  border = Border(left=side, right=side, top=side, bottom=side)
  center = Alignment(horizontal="center", vertical="center", wrap_text=True)
  left = Alignment(horizontal="left", vertical="center", wrap_text=True)
  
  f_title = Font(name="TH SarabunPSK", size=16, bold=True)
  f_bold = Font(name="TH SarabunPSK", size=14, bold=True)
  f_normal = Font(name="TH SarabunPSK", size=14)

  # กำหนดความกว้างคอลัมน์ A ถึง I เหมือนต้นฉบับเป๊ะ
  col_widths = {"A": 6, "B": 28, "C": 42, "D": 16, "E": 16, "F": 10, "G": 10, "H": 10, "I": 28}
  for col, width in col_widths.items():
    ws.column_dimensions[col].width = width

  # หัวเรื่อง 2 บรรทัด
  titles = [
      f"ตารางสรุปวันลา {emp_type} สำนักงานศึกษาธิการจังหวัดสมุทรสงคราม",
      f"ประจำเดือน{month_str} {year_num}",
  ]
  for row, text in enumerate(titles, start=1):
    ws.merge_cells(f"A{row}:I{row}")
    cell = ws[f"A{row}"]
    cell.value, cell.font, cell.alignment = text, f_title, center

  # หัวตาราง 2 ชั้น (จัดตำแหน่ง ลาป่วย / ลากิจ ใต้ครั้งนี้ ให้ตรงต้นฉบับเป๊ะ)
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
  ws["F4"] = "ครั้งนี้\n(วัน)"
  ws["F5"] = "ลาป่วย"
  ws["G5"] = "ลากิจ"

  ws.merge_cells("H4:H5")
  ws["H4"] = "รวม\n(วัน)"
  ws.merge_cells("I4:I5")
  ws["I4"] = "หมายเหตุ"

  for row in ws.iter_rows(min_row=4, max_row=5, max_col=9):
    for cell in row:
      cell.font, cell.alignment, cell.border = f_bold, center, border

  start_row = 6
  current_row = start_row

  for idx, emp in enumerate(employees_list, 1):
    name = emp["name"]
    position = emp["position"]
    recs = [
        r for r in records_dict.get(name, [])
        if month_start <= r["start"] <= month_end
    ]

    if not recs:
      ws.cell(row=current_row, column=1, value=idx)
      ws.cell(row=current_row, column=2, value=name)
      ws.cell(row=current_row, column=3, value=position)
      for c in range(4, 10):
        ws.cell(row=current_row, column=c, value="-")
      
      for c in range(1, 10):
        cell = ws.cell(row=current_row, column=c)
        cell.font = f_normal
        cell.border = border
        cell.alignment = left if c in (2, 3, 9) else center
      current_row += 1
    else:
      emp_start_row = current_row
      total_sick = 0
      total_kit = 0

      for r in recs:
        ws.cell(row=current_row, column=4, value=format_excel_date(r["start"]))
        ws.cell(row=current_row, column=5, value=format_excel_date(r["end"]))
        
        s_val = r["days"] if r["category"] == "ลาป่วย" else 0
        k_val = r["days"] if r["category"] == "ลากิจ" else 0
        total_sick += s_val
        total_kit += k_val

        ws.cell(row=current_row, column=6, value=fmt_days(s_val))
        ws.cell(row=current_row, column=7, value=fmt_days(k_val))
        ws.cell(row=current_row, column=8, value="")
        ws.cell(row=current_row, column=9, value=describe(r) or "-")

        for c in range(4, 10):
          cell = ws.cell(row=current_row, column=c)
          cell.font = f_normal
          cell.alignment = left if c == 9 else center
          t_line = side if current_row == emp_start_row else None
          b_line = side if current_row == emp_start_row + len(recs) - 1 else None
          cell.border = Border(left=side, right=side, top=t_line, bottom=b_line)

        current_row += 1

      emp_end_row = current_row - 1

      ws.cell(row=emp_start_row, column=1, value=idx)
      ws.cell(row=emp_start_row, column=2, value=name)
      ws.cell(row=emp_start_row, column=3, value=position)
      grand_total = total_sick + total_kit
      ws.cell(row=emp_start_row, column=8, value=fmt_days(grand_total))

      for col_idx in [1, 2, 3, 8]:
        if emp_start_row != emp_end_row:
          ws.merge_cells(
              start_row=emp_start_row, start_column=col_idx,
              end_row=emp_end_row, end_column=col_idx
          )

      for r_idx in range(emp_start_row, emp_end_row + 1):
        for c_idx in [1, 2, 3, 8]:
          cell = ws.cell(row=r_idx, column=c_idx)
          cell.font = f_normal
          cell.alignment = left if c_idx in (2, 3) else center
          t_line = side if r_idx == emp_start_row else None
          b_line = side if r_idx == emp_end_row else None
          cell.border = Border(left=side, right=side, top=t_line, bottom=b_line)

  # ส่วนลงชื่อผู้ตรวจสอบ (ดึงชื่อ นายวงค์สกุล เข่งคุ้ม และตำแหน่งมาแสดงตามต้นฉบับเป๊ะ)
  sig_row = current_row + 2
  ws.merge_cells(f"F{sig_row}:I{sig_row}")
  cell = ws[f"F{sig_row}"]
  cell.value = "(ลงชื่อ)....................................................ผู้ตรวจสอบ"
  cell.font, cell.alignment = f_normal, center

  name_row = sig_row + 2
  ws.merge_cells(f"F{name_row}:I{name_row}")
  cell_name = ws[f"F{name_row}"]
  cell_name.value = "(นายวงค์สกุล เข่งคุ้ม)"
  cell_name.font, cell_name.alignment = f_normal, center

  pos_row = name_row + 1
  ws.merge_cells(f"F{pos_row}:I{pos_row}")
  cell_pos = ws[f"F{pos_row}"]
  cell_pos.value = "ตำแหน่ง นักทรัพยากรบุคคลปฏิบัติการ"
  cell_pos.font, cell_pos.alignment = f_normal, center

  buf = io.BytesIO()
  wb.save(buf)
  return buf.getvalue()


st.subheader("📥 ดาวน์โหลดรายงานประจำเดือน")
st.download_button(
    "📥 ดาวน์โหลดไฟล์ Excel ตามแบบฟอร์มราชการ",
    data=generate_excel(
        employees.get(employee_type, []), records,
        selected_month, selected_year, employee_type
    ),
    file_name=f"สรุปวันลา_{employee_type}_{selected_month}_{selected_year}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    type="primary",
)
