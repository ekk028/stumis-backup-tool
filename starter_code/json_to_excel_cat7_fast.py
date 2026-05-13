"""Fast Cat 7 → Excel converter using xlsxwriter (much faster than openpyxl)."""
import json, os, re, time
from collections import defaultdict

# Try to use xlsxwriter (faster); fall back to openpyxl write_only
try:
    import xlsxwriter
    USE_XLSXWRITER = True
except ImportError:
    USE_XLSXWRITER = False

SRC = "/sessions/compassionate-hopeful-cori/mnt/D:--Downloads/stumis_cat7_class_check.json"
DST_DIR = "/sessions/compassionate-hopeful-cori/mnt/15. backup stumis 2568/07_เช็กชื่อรายคาบ"
os.makedirs(DST_DIR, exist_ok=True)
DST = os.path.join(DST_DIR, "07_เช็กชื่อรายคาบ_ทุกห้อง_ปี2568.xlsx")

THAI_MONTHS = {1:'ม.ค.',2:'ก.พ.',3:'มี.ค.',4:'เม.ย.',5:'พ.ค.',6:'มิ.ย.',7:'ก.ค.',8:'ส.ค.',9:'ก.ย.',10:'ต.ค.',11:'พ.ย.',12:'ธ.ค.'}

def fmt_thai_date(iso):
    y, m, d = iso.split('-')
    return f"{int(d)} {THAI_MONTHS[int(m)]} {int(y)+543}"

def safe_sheet_name(name):
    return re.sub(r'[\\/?*\[\]:]', '-', name)[:31]

t0 = time.time()
print(f"Loading JSON...")
with open(SRC, encoding='utf-8') as f:
    data = json.load(f)
print(f"  loaded in {time.time()-t0:.1f}s")

by_room = defaultdict(dict)
for date, rooms in data.items():
    for room, table in rooms.items():
        by_room[room][date] = table

room_order = sorted(by_room.keys(), key=lambda r: (
    int(re.match(r'^ม\.(\d+)', r).group(1)) if re.match(r'^ม\.\d', r) else 99,
    int(r.split('/')[-1]) if '/' in r and r.split('/')[-1].isdigit() else 99
))

print(f"Writing Excel using xlsxwriter...")
wb = xlsxwriter.Workbook(DST, {'constant_memory': True})  # stream mode

f_title = wb.add_format({'font_name':'TH Sarabun New','font_size':16,'bold':True,'font_color':'white','bg_color':'#1F4E79','align':'center','valign':'vcenter','border':1})
f_date = wb.add_format({'font_name':'TH Sarabun New','font_size':13,'bold':True,'font_color':'white','bg_color':'#2E75B6','align':'center','valign':'vcenter','border':1})
f_header = wb.add_format({'font_name':'TH Sarabun New','font_size':11,'bold':True,'bg_color':'#D9E2F3','align':'center','valign':'vcenter','border':1,'text_wrap':True})
f_cell = wb.add_format({'font_name':'TH Sarabun New','font_size':11,'align':'center','valign':'vcenter','border':1})
f_name = wb.add_format({'font_name':'TH Sarabun New','font_size':11,'align':'left','valign':'vcenter','border':1})
f_link = wb.add_format({'font_name':'TH Sarabun New','font_size':12,'color':'#0563C1','underline':1,'align':'center','valign':'vcenter','border':1})

# Index sheet
idx = wb.add_worksheet("สารบัญ")
idx.set_column('A:A', 12)
idx.set_column('B:B', 18)
idx.set_column('C:C', 36)
idx.set_column('D:D', 18)
idx.set_row(0, 24)
idx.merge_range(0, 0, 0, 3, "สารบัญ — รายงานเช็กชื่อรายคาบเรียน ปีการศึกษา 2568", f_title)
idx.write_row(2, 0, ["ห้องเรียน", "จำนวนวันที่มีข้อมูล", "ช่วงวันที่", "ลิงค์ Sheet"], f_header)
idx.freeze_panes(3, 0)

idx_row = 3
for room in room_order:
    sheet_name = safe_sheet_name(room)
    ws = wb.add_worksheet(sheet_name)
    ws.set_column('A:A', 5)
    ws.set_column('B:B', 12)
    ws.set_column('C:C', 28)
    ws.set_column('D:O', 11)

    dates_sorted = sorted(by_room[room].keys())

    ws.set_row(0, 24)
    ws.merge_range(0, 0, 0, 14, f"รายงานเช็กชื่อรายคาบเรียน — ชั้น {room}", f_title)

    row = 2
    for date in dates_sorted:
        table = by_room[room][date]
        if not table or len(table) < 3:
            continue
        clean_rows = table[1:]
        if not clean_rows:
            continue
        header = clean_rows[0]
        ncols = max(len(header), 4)

        ws.set_row(row, 20)
        ws.merge_range(row, 0, row, ncols-1, f"วันที่ {fmt_thai_date(date)}", f_date)
        row += 1

        ws.set_row(row, 28)
        for ci, val in enumerate(header):
            ws.write(row, ci, val, f_header)
        row += 1

        for r in clean_rows[1:]:
            for ci, val in enumerate(r):
                fmt = f_name if ci == 2 else f_cell
                cv = int(val.strip()) if (ci == 0 and isinstance(val, str) and val.strip().isdigit()) else val
                ws.write(row, ci, cv, fmt)
            row += 1
        row += 1  # spacer

    range_str = f"{fmt_thai_date(dates_sorted[0])}  ถึง  {fmt_thai_date(dates_sorted[-1])}" if dates_sorted else "-"
    idx.write(idx_row, 0, room, f_cell)
    idx.write(idx_row, 1, len(dates_sorted), f_cell)
    idx.write(idx_row, 2, range_str, f_name)
    idx.write_url(idx_row, 3, f"internal:'{sheet_name}'!A1", f_link, "ไปที่ Sheet")
    idx_row += 1

t1 = time.time()
print(f"  built in {t1-t0:.1f}s")
wb.close()
print(f"  closed in {time.time()-t1:.1f}s")
print(f"Saved: {DST}")
print(f"Size: {os.path.getsize(DST)/1024/1024:.2f} MB")
