# -------------------------------------------------------------------------
# 10. สร้างไฟล์ Excel ตามแบบฟอร์มราชการ (หน้าตาและความกว้างเหมือนต้นฉบับเป๊ะ)
# -------------------------------------------------------------------------
def generate_excel(employees_list, records_dict, month_str, year_num, emp_type):
  wb = openpyxl.Workbook()
  ws = wb.active
  ws.title = "สรุปวันลา"

  # ตั้งค่าหน้ากระดาษ A4 แนวตั้ง และสั่ง Fit to 1 page width
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

  # กำหนดความกว้างคอลัมน์ A ถึง I ให้สอดคล้องกับแบบฟอร์มราชการ
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

  # หัวตาราง 2 ชั้นตามแบบฟอร์ม
  header_cols = zip(
      "ABCDEHI",
      ["ที่", "ชื่อ - สกุล", "ตำแหน่ง", "ตั้งแต่วันที่", "ถึงวันที่", "รวม\n(วัน)", "หมายเหตุ"],
  )
  for col, text in header_cols:
    ws.merge_cells(f"{col}4:{col}5")
    ws[f"{col}4"] = text
  ws.merge_cells("F4:G4")
  ws["F4"], ws["F5"], ws["G5"] = "ครั้งนี้", "ลาป่วย", "ลากิจ"

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

        # ตีเส้นขอบซ้าย/ขวา และละเว้นเส้นขอบนอนระหว่างบรรทัดย่อยของคนเดียวกัน
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

      # Merge คอลัมน์ ที่, ชื่อ, ตำแหน่ง และรวมวัน ยาวลงมาตรงกลาง
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

  # ช่องลงชื่อด้านล่าง
  sig = current_row + 2
  ws.merge_cells(f"F{sig}:I{sig}")
  cell = ws[f"F{sig}"]
  cell.value = "(ลงชื่อ)....................................................ผู้ตรวจสอบ"
  cell.font, cell.alignment = f_normal, center

  buf = io.BytesIO()
  wb.save(buf)
  return buf.getvalue()
