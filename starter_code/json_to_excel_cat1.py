"""Convert STU-MIS category 1 (morning attendance) JSON dump to Excel."""
import json, os, re
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

SRC = "/sessions/compassionate-hopeful-cori/mnt/D:--Downloads/stumis_cat1_morning_check.json"
DST_DIR = "/sessions/compassionate-hopeful-cori/mnt/15. backup stumis 2568/01_เช็กชื่อเช้า"
os.makedirs(DST_DIR, exist_ok=True)
DST = os.path.join(DST_DIR, "01_เช็กชื่อเช้า_ทุกห้อง_ปี2568.xlsx")

MONTH_NAMES = {
    "01": "มกราคม 2569", "02": "กุมภาพันธ์ 2569", "03": "มีนาคม 2569",
    "04": "เมษายน 2569", "05": "พฤษภาคม 2568", "06": "มิถุนายน 2568",
    "07": "กรกฎาคม 2568", "08": "สิงหาคม 2568", "09": "กันยายน 2568",
    "10": "ตุลาคม 2568", "11": "พฤศจิกายน 2568", "12": "ธันวาคม 2568"
}
MONTH_ORDER = ["05","06","07","08","09","10","11","12","01","02","03","04"]

with open(SRC, encoding='utf-8') as f:
    data = json.load(f)

wb = Workbook()
wb.remove(wb.active)

# Styles
thin = Side(border_style="thin", color="999999")
border = Border(left=thin, right=thin, top=thin, bottom=thin)
title_font = Font(name="TH Sarabun New", size=16, bold=True, color="FFFFFF")
title_fill = PatternFill("solid", fgColor="1F4E79")
month_font = Font(name="TH Sarabun New", size=14, bold=True, color="FFFFFF")
month_fill = PatternFill("solid", fgColor="2E75B6")
header_font = Font(name="TH Sarabun New", size=12, bold=True)
header_fill = PatternFill("solid", fgColor="D9E2F3")
cell_font = Font(name="TH Sarabun New", size=12)
center = Alignment(horizontal="center", vertical="center", wrap_text=True)

# Index sheet
idx = wb.create_sheet("สารบัญ")
idx["A1"] = "สารบัญ — รายงานเช็กชื่อเช้า (กิจกรรมหน้าเสาธง) ปีการศึกษา 2568"
idx["A1"].font = title_font
idx["A1"].fill = title_fill
idx["A1"].alignment = center
idx.merge_cells("A1:D1")
idx.row_dimensions[1].height = 24
idx.append([])
idx.append(["ห้องเรียน", "ครูประจำชั้น", "เดือนที่มีข้อมูล", "ลิงค์ไปที่ Sheet"])
for col, w in enumerate([14, 36, 32, 24], 1):
    idx.column_dimensions[get_column_letter(col)].width = w
for c in idx[3]:
    c.font = header_font
    c.fill = header_fill
    c.alignment = center
    c.border = border

def safe_sheet_name(name):
    # Excel sheet names can't contain: / \ ? * [ ] : (max 31 chars)
    s = re.sub(r'[\\/?*\[\]:]', '-', name)[:31]
    return s

for room in sorted(data.keys(), key=lambda r: (int(r.split('.')[1].split('/')[0]) if re.match(r'^ม\.\d', r) else 99,
                                                int(r.split('/')[-1]) if '/' in r and r.split('/')[-1].isdigit() else 99)):
    sheet_name = safe_sheet_name(room)
    ws = wb.create_sheet(sheet_name)
    ws.sheet_view.showGridLines = False

    months_with_data = []
    teacher_note = ""

    # Room title
    ws["A1"] = f"รายงานการเช็กชื่อ กิจกรรมหน้าเสาธง — ชั้น {room}"
    ws["A1"].font = title_font
    ws["A1"].fill = title_fill
    ws["A1"].alignment = center
    ws.merge_cells("A1:O1")
    ws.row_dimensions[1].height = 26

    row_cursor = 3
    for mm in MONTH_ORDER:
        bucket = data.get(room, {}).get(mm, {})
        rows = bucket.get("data", [])
        # Skip if no real student data (row count <= 2 means just headers)
        if len(rows) <= 2:
            continue
        months_with_data.append(MONTH_NAMES[mm])
        if not teacher_note:
            note = bucket.get("headerNote", "")
            # extract "ครูประจำชั้น ..."
            m = re.search(r'ครูประจำชั้น.*?(?=ที่|รหัส|$)', note, re.DOTALL)
            teacher_note = m.group(0).strip() if m else note

        # Month band
        ws.cell(row=row_cursor, column=1, value=f"เดือน {MONTH_NAMES[mm]}")
        ws.cell(row=row_cursor, column=1).font = month_font
        ws.cell(row=row_cursor, column=1).fill = month_fill
        ws.cell(row=row_cursor, column=1).alignment = center
        # The data has variable columns based on dates - determine
        clean_rows = rows[1:]  # skip the messy first row (merged)
        header_row = clean_rows[0] if clean_rows else []
        ncols = max(len(header_row), 4)
        ws.merge_cells(start_row=row_cursor, start_column=1, end_row=row_cursor, end_column=ncols)
        ws.row_dimensions[row_cursor].height = 22
        row_cursor += 1

        # Header row
        for ci, val in enumerate(header_row, 1):
            c = ws.cell(row=row_cursor, column=ci, value=val)
            c.font = header_font
            c.fill = header_fill
            c.alignment = center
            c.border = border
        ws.row_dimensions[row_cursor].height = 32
        row_cursor += 1

        # Data rows
        for r in clean_rows[1:]:
            for ci, val in enumerate(r, 1):
                # Try to convert to int for first column (running number)
                cv = val
                if ci == 1 and isinstance(val, str) and val.strip().isdigit():
                    cv = int(val.strip())
                c = ws.cell(row=row_cursor, column=ci, value=cv)
                c.font = cell_font
                c.alignment = center if ci != 3 else Alignment(horizontal="left", vertical="center", wrap_text=True)
                c.border = border
            row_cursor += 1
        # Spacer
        row_cursor += 1

    # Set column widths
    ws.column_dimensions['A'].width = 5
    ws.column_dimensions['B'].width = 12
    ws.column_dimensions['C'].width = 30
    for ci in range(4, 32):
        ws.column_dimensions[get_column_letter(ci)].width = 11

    # Add to index
    idx_row = idx.max_row + 1
    idx.cell(row=idx_row, column=1, value=room).font = cell_font
    idx.cell(row=idx_row, column=2, value=teacher_note.replace("ครูประจำชั้น", "").strip()).font = cell_font
    idx.cell(row=idx_row, column=3, value=", ".join(m.replace(" 2568","").replace(" 2569","") for m in months_with_data)).font = cell_font
    idx.cell(row=idx_row, column=4, value=f'=HYPERLINK("#{sheet_name}!A1","ไปที่ Sheet")').font = Font(name="TH Sarabun New", size=12, color="0563C1", underline="single")
    for col in range(1, 5):
        idx.cell(row=idx_row, column=col).alignment = Alignment(horizontal="left" if col in (2,3) else "center", vertical="center", wrap_text=True)
        idx.cell(row=idx_row, column=col).border = border

idx.freeze_panes = "A4"
wb.save(DST)
print(f"Saved: {DST}")
print(f"Size: {os.path.getsize(DST) / 1024:.1f} KB")
print(f"Sheets: {len(wb.sheetnames)} — {wb.sheetnames[:5]}...")
