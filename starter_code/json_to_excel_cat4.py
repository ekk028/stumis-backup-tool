"""Convert STU-MIS category 4 (home visit) JSON dump to Excel."""
import json, os, re
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

SRC = "/sessions/compassionate-hopeful-cori/mnt/D:--Downloads/stumis_cat4_homevisit.json"
DST_DIR = "/sessions/compassionate-hopeful-cori/mnt/15. backup stumis 2568/04_เยี่ยมบ้าน"
os.makedirs(DST_DIR, exist_ok=True)
DST = os.path.join(DST_DIR, "04_เยี่ยมบ้าน_ทุกห้อง_ปี2568.xlsx")

with open(SRC, encoding='utf-8') as f:
    data = json.load(f)

wb = Workbook()
wb.remove(wb.active)

thin = Side(border_style="thin", color="999999")
border = Border(left=thin, right=thin, top=thin, bottom=thin)
title_font = Font(name="TH Sarabun New", size=16, bold=True, color="FFFFFF")
title_fill = PatternFill("solid", fgColor="1F4E79")
header_font = Font(name="TH Sarabun New", size=12, bold=True)
header_fill = PatternFill("solid", fgColor="D9E2F3")
cell_font = Font(name="TH Sarabun New", size=12)
center = Alignment(horizontal="center", vertical="center", wrap_text=True)

# Summary sheet
ws = wb.create_sheet("สรุปภาพรวม")
ws.sheet_view.showGridLines = False
ws["A1"] = "สรุปการเยี่ยมบ้านนักเรียน — ครั้งที่ 1 ภาคเรียนที่ 1/2568"
ws["A1"].font = title_font
ws["A1"].fill = title_fill
ws["A1"].alignment = center
ncol = max(len(r) for r in data["summary"])
ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ncol)
ws.row_dimensions[1].height = 26

row_cursor = 3
for i, row in enumerate(data["summary"]):
    if not row or not any(c for c in row):
        continue
    is_header = (i == 0 or all(re.match(r'^[ก-๙a-zA-Z\s/]+$', c) for c in row if c) and not any(c.replace(".", "").replace("%","").isdigit() for c in row if c))
    for ci, val in enumerate(row, 1):
        cv = val
        # try numeric conversion
        if val and val.replace(".","").replace("%","").isdigit():
            try:
                cv = float(val.replace("%","")) if "." in val or "%" in val else int(val)
            except: pass
        c = ws.cell(row=row_cursor, column=ci, value=cv)
        c.font = header_font if is_header else cell_font
        c.fill = header_fill if is_header else PatternFill()
        c.alignment = center
        c.border = border
    row_cursor += 1

for col, w in enumerate([6, 14, 18, 22, 22, 22, 22], 1):
    ws.column_dimensions[get_column_letter(col)].width = w

def safe_sheet_name(name):
    return re.sub(r'[\\/?*\[\]:]', '-', name)[:31]

# Per-room detail sheets
for room, table in data.get("rooms", {}).items():
    if not table or len(table) < 2:
        continue
    sheet_name = safe_sheet_name(room)
    ws = wb.create_sheet(sheet_name)
    ws.sheet_view.showGridLines = False
    ws["A1"] = f"ข้อมูลการเยี่ยมบ้านนักเรียน — ชั้น {room}"
    ws["A1"].font = title_font
    ws["A1"].fill = title_fill
    ws["A1"].alignment = center
    ncol = max(len(r) for r in table) if table else 1
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ncol)
    ws.row_dimensions[1].height = 26

    for ri, row in enumerate(table):
        if not row or not any(c for c in row):
            continue
        is_header = ri == 0
        for ci, val in enumerate(row, 1):
            cv = val
            if val and val.replace(".", "").isdigit():
                try: cv = int(val) if val.isdigit() else float(val)
                except: pass
            c = ws.cell(row=ri + 3, column=ci, value=cv)
            c.font = header_font if is_header else cell_font
            c.fill = header_fill if is_header else PatternFill()
            c.alignment = center if (ci != 3 or is_header) else Alignment(horizontal="left", vertical="center", wrap_text=True)
            c.border = border

    # Widths
    ws.column_dimensions['A'].width = 6
    ws.column_dimensions['B'].width = 14
    ws.column_dimensions['C'].width = 30
    for ci in range(4, 12):
        ws.column_dimensions[get_column_letter(ci)].width = 14

wb.save(DST)
print(f"Saved: {DST}")
print(f"Size: {os.path.getsize(DST)/1024:.1f} KB")
print(f"Sheets: {len(wb.sheetnames)} — {wb.sheetnames[:6]}...")
