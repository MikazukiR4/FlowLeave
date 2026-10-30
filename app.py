def generate_excel(employees_list, records_dict, month_str, year_num, emp_type):
  template_path = Path(__file__).resolve().parent / "template.xlsx"
  
  if template_path.exists():
    wb = openpyxl.load_workbook(template_path)
    ws = wb.active
  else:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "สรุปวันลา"

  # อัปเดตหัวข้อรายงานตามงวดเวลาที่เลือก
  ws["A1"] = f"ตารางสรุปวันลา {emp_type} สำนักงานศึกษาธิการจังหวัดสมุทรสงคราม"
  ws["A2"] = f"ประจำเดือน{month_str} {year_num}"

  side = Side(style="thin", color="000000")
  border = Border(left=side, right=side, top=side, bottom=side)
  center = Alignment(horizontal="center", vertical="center", wrap_text=True)
  left = Alignment(horizontal="left", vertical="center", wrap_text=True)
  f_normal = Font(name="TH SarabunPSK", size=14)

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

  buf = io.BytesIO()
  wb.save(buf)
  return buf.getvalue()
