"""Convert structured data → Excel files using xlsxwriter (fast) or openpyxl for Cat4."""
import os
import re
from pathlib import Path


# ── helpers ──────────────────────────────────────────────────────────────────

THAI_MONTHS_SHORT = {
    1: "ม.ค.", 2: "ก.พ.", 3: "มี.ค.", 4: "เม.ย.",
    5: "พ.ค.", 6: "มิ.ย.", 7: "ก.ค.", 8: "ส.ค.",
    9: "ก.ย.", 10: "ต.ค.", 11: "พ.ย.", 12: "ธ.ค.",
}

MONTH_FULL = {
    "01": "มกราคม", "02": "กุมภาพันธ์", "03": "มีนาคม",
    "04": "เมษายน",  "05": "พฤษภาคม",   "06": "มิถุนายน",
    "07": "กรกฎาคม", "08": "สิงหาคม",   "09": "กันยายน",
    "10": "ตุลาคม",  "11": "พฤศจิกายน", "12": "ธันวาคม",
}

MONTH_ORDER = ["05", "06", "07", "08", "09", "10", "11", "12", "01", "02", "03", "04"]


def safe_sheet_name(name: str) -> str:
    return re.sub(r'[\\/?*\[\]:]', "-", name)[:31]


def _room_sort_key(room: str):
    m = re.match(r"^ม\.(\d+)/(\d+)", room)
    return (int(m.group(1)), int(m.group(2))) if m else (99, 99)


def save_raw(content: bytes, filepath: str) -> None:
    Path(filepath).parent.mkdir(parents=True, exist_ok=True)
    with open(filepath, "wb") as f:
        f.write(content)


def _fmt_thai_date(iso: str) -> str:
    y, mo, d = iso.split("-")
    return f"{int(d)} {THAI_MONTHS_SHORT[int(mo)]} {int(y)+543}"


def _to_number(val: str):
    if not isinstance(val, str):
        return val
    v = val.strip()
    if v.isdigit():
        return int(v)
    try:
        return float(v)
    except ValueError:
        return val


# ── xlsxwriter format helpers ─────────────────────────────────────────────────

def _base_formats(wb):
    """Return dict of named xlsxwriter formats."""
    return {
        "title": wb.add_format({
            "font_name": "TH Sarabun New", "font_size": 16, "bold": True,
            "font_color": "white", "bg_color": "#1F4E79",
            "align": "center", "valign": "vcenter", "border": 1,
        }),
        "band": wb.add_format({
            "font_name": "TH Sarabun New", "font_size": 14, "bold": True,
            "font_color": "white", "bg_color": "#2E75B6",
            "align": "center", "valign": "vcenter", "border": 1,
        }),
        "header": wb.add_format({
            "font_name": "TH Sarabun New", "font_size": 12, "bold": True,
            "bg_color": "#D9E2F3",
            "align": "center", "valign": "vcenter", "border": 1, "text_wrap": True,
        }),
        "cell": wb.add_format({
            "font_name": "TH Sarabun New", "font_size": 12,
            "align": "center", "valign": "vcenter", "border": 1,
        }),
        "cell_left": wb.add_format({
            "font_name": "TH Sarabun New", "font_size": 12,
            "align": "left", "valign": "vcenter", "border": 1, "text_wrap": True,
        }),
        "link": wb.add_format({
            "font_name": "TH Sarabun New", "font_size": 12,
            "font_color": "#0563C1", "underline": 1,
            "align": "center", "valign": "vcenter", "border": 1,
        }),
    }


# ── Cat 1: เช็กชื่อเช้า ────────────────────────────────────────────────────────

def convert_cat1(data: dict, filepath: str, year: int = 2568) -> None:
    """
    data format: {room: {mm: {"data": [[row,...]], "headerNote": str}}}
    Produces one sheet per room + index sheet.
    """
    import xlsxwriter
    Path(filepath).parent.mkdir(parents=True, exist_ok=True)

    def month_label(mm: str) -> str:
        suffix = year + 1 if mm in ("01", "02", "03", "04") else year
        return f"{MONTH_FULL[mm]} {suffix}"

    wb = xlsxwriter.Workbook(filepath)
    fmt = _base_formats(wb)

    # Index sheet
    idx = wb.add_worksheet("สารบัญ")
    idx.set_column("A:A", 14)
    idx.set_column("B:B", 36)
    idx.set_column("C:C", 32)
    idx.set_column("D:D", 24)
    idx.set_row(0, 26)
    idx.merge_range(0, 0, 0, 3, f"สารบัญ — รายงานเช็กชื่อเช้า ปีการศึกษา {year}", fmt["title"])
    idx.write_row(2, 0, ["ห้องเรียน", "ครูประจำชั้น", "เดือนที่มีข้อมูล", "ลิงค์"], fmt["header"])
    idx.freeze_panes(3, 0)
    idx_row = 3

    for room in sorted(data.keys(), key=_room_sort_key):
        sheet_name = safe_sheet_name(room)
        ws = wb.add_worksheet(sheet_name)
        ws.set_column("A:A", 5)
        ws.set_column("B:B", 12)
        ws.set_column("C:C", 30)
        ws.set_column(3, 31, 11)
        ws.set_row(0, 26)
        ws.merge_range(0, 0, 0, 14, f"รายงานการเช็กชื่อ กิจกรรมหน้าเสาธง — ชั้น {room}", fmt["title"])

        row = 2
        months_done: list[str] = []
        teacher = ""

        for mm in MONTH_ORDER:
            bucket = data.get(room, {}).get(mm, {})
            rows = bucket.get("data", [])
            note = bucket.get("headerNote", "")
            if len(rows) <= 2:
                continue

            months_done.append(month_label(mm))
            if not teacher and note:
                m2 = re.search(r"ครูประจำชั้น[^\n\r]{0,100}", note)
                teacher = m2.group(0).replace("ครูประจำชั้น", "").strip() if m2 else ""

            clean = rows[1:]
            header_row = clean[0] if clean else []
            ncols = max(len(header_row), 4)

            ws.set_row(row, 22)
            ws.merge_range(row, 0, row, ncols - 1, f"เดือน {month_label(mm)}", fmt["band"])
            row += 1

            ws.set_row(row, 32)
            for ci, val in enumerate(header_row):
                ws.write(row, ci, val, fmt["header"])
            row += 1

            for data_row in clean[1:]:
                for ci, val in enumerate(data_row):
                    f = fmt["cell_left"] if ci == 2 else fmt["cell"]
                    ws.write(row, ci, _to_number(val) if ci != 2 else val, f)
                row += 1
            row += 1  # spacer

        # Write index entry
        idx.write(idx_row, 0, room, fmt["cell"])
        idx.write(idx_row, 1, teacher, fmt["cell_left"])
        idx.write(idx_row, 2, ", ".join(months_done), fmt["cell_left"])
        idx.write_url(idx_row, 3, f"internal:'{sheet_name}'!A1", fmt["link"], "ไปที่ Sheet")
        idx_row += 1

    wb.close()


# ── Cat 4: เยี่ยมบ้าน ─────────────────────────────────────────────────────────

def convert_cat4(data: dict, filepath: str, year: int = 2568) -> None:
    """
    data format: {"summary": [[rows]], "rooms": {room: [[rows]]}}
    """
    import xlsxwriter
    Path(filepath).parent.mkdir(parents=True, exist_ok=True)

    wb = xlsxwriter.Workbook(filepath)
    fmt = _base_formats(wb)

    summary_rows = data.get("summary", [])

    # Summary sheet
    ws_sum = wb.add_worksheet("สรุปภาพรวม")
    ws_sum.set_column("A:A", 6)
    ws_sum.set_column("B:B", 14)
    ws_sum.set_column("C:C", 22)
    ws_sum.set_column(3, 8, 18)
    ncols_sum = max((len(r) for r in summary_rows), default=1)
    ws_sum.set_row(0, 26)
    ws_sum.merge_range(0, 0, 0, max(ncols_sum - 1, 0),
                       f"สรุปการเยี่ยมบ้านนักเรียน ปีการศึกษา {year}", fmt["title"])

    for ri, row in enumerate(summary_rows):
        if not row or not any(row):
            continue
        is_header = ri == 0
        f = fmt["header"] if is_header else fmt["cell"]
        for ci, val in enumerate(row):
            ws_sum.write(ri + 2, ci, _to_number(val) if not is_header else val, f)

    # Per-room sheets
    rooms = data.get("rooms", {})
    for room in sorted(rooms.keys(), key=_room_sort_key):
        table = rooms[room]
        if not table or len(table) < 2:
            continue
        sheet_name = safe_sheet_name(room)
        ws = wb.add_worksheet(sheet_name)
        ws.set_column("A:A", 6)
        ws.set_column("B:B", 14)
        ws.set_column("C:C", 30)
        ws.set_column(3, 11, 14)
        ncols = max((len(r) for r in table), default=1)
        ws.set_row(0, 26)
        ws.merge_range(0, 0, 0, max(ncols - 1, 0),
                       f"ข้อมูลการเยี่ยมบ้านนักเรียน — ชั้น {room}", fmt["title"])

        for ri, row in enumerate(table):
            if not row or not any(row):
                continue
            is_header = ri == 0
            f_row = fmt["header"] if is_header else fmt["cell"]
            for ci, val in enumerate(row):
                if not is_header and ci == 2:
                    ws.write(ri + 2, ci, val, fmt["cell_left"])
                else:
                    ws.write(ri + 2, ci, _to_number(val) if not is_header else val, f_row)

    wb.close()


# ── Cat 8: ข้อมูลครูและบุคลากร ───────────────────────────────────────────────

def _write_table(ws, rows: list, fmt: dict, title: str, note: str = "") -> None:
    """Write a titled table onto a worksheet starting at row 0."""
    ncols = max((len(r) for r in rows), default=4)
    ws.set_row(0, 26)
    ws.merge_range(0, 0, 0, max(ncols - 1, 3), title, fmt["title"])
    row_offset = 1
    if note:
        note_fmt = fmt["title"].__class__  # fallback; overwritten below
        ws.write(1, 0, note)
        row_offset = 2

    for ri, row in enumerate(rows):
        is_header = ri == 0
        ws.set_row(ri + row_offset + (1 if note else 0), 22)
        for ci, val in enumerate(row):
            if is_header:
                ws.write(ri + row_offset + (1 if note else 0), ci, val, fmt["header"])
            elif ci == 2:
                ws.write(ri + row_offset + (1 if note else 0), ci, val, fmt["cell_left"])
            else:
                ws.write(ri + row_offset + (1 if note else 0), ci,
                         _to_number(val), fmt["cell"])


def convert_cat8(rows_main: list, rows_stat: list, filepath: str, year: int = 2568) -> None:
    """
    rows_main: from Teacher-main.php   (ที่, Username, ชื่อ-สกุล, ครูที่ปรึกษา)
    rows_stat: from Teacher-main-excel.php (ที่, ชื่อ, จำนวนชั่วโมง, Login สะสม, การเข้าสอน ...)
    """
    import xlsxwriter
    Path(filepath).parent.mkdir(parents=True, exist_ok=True)

    wb = xlsxwriter.Workbook(filepath)
    fmt = _base_formats(wb)
    note_fmt = wb.add_format({
        "font_name": "TH Sarabun New", "font_size": 11,
        "italic": True, "font_color": "#888888",
    })

    # ── Sheet 1: รายชื่อ + Username ──────────────────────────────────────────
    ws1 = wb.add_worksheet("รายชื่อครูและบุคลากร")
    ws1.set_column("A:A", 6)
    ws1.set_column("B:B", 20)
    ws1.set_column("C:C", 32)
    ws1.set_column("D:D", 20)
    ws1.set_column("E:Z", 16)
    ncols1 = max((len(r) for r in rows_main), default=4)
    ws1.set_row(0, 26)
    ws1.merge_range(0, 0, 0, max(ncols1 - 1, 3),
                    f"รายชื่อครูและบุคลากร ปีการศึกษา {year}", fmt["title"])
    ws1.write(1, 0, "หมายเหตุ: รหัสผ่านไม่สามารถดึงได้ (ระบบเก็บแบบเข้ารหัส)", note_fmt)
    for ri, row in enumerate(rows_main):
        is_hdr = ri == 0
        ws1.set_row(ri + 3, 22)
        for ci, val in enumerate(row):
            f = fmt["header"] if is_hdr else (fmt["cell_left"] if ci == 2 else fmt["cell"])
            ws1.write(ri + 3, ci, val if is_hdr else _to_number(val) if ci != 2 else val, f)

    # ── Sheet 2: สถิติ (จำนวนชั่วโมง, Login, การเข้าสอน) ────────────────────
    ws2 = wb.add_worksheet("สถิติการใช้งาน")
    ws2.set_column("A:A", 6)
    ws2.set_column("B:B", 32)
    ws2.set_column("C:Z", 18)
    ncols2 = max((len(r) for r in rows_stat), default=5)
    ws2.set_row(0, 26)
    ws2.merge_range(0, 0, 0, max(ncols2 - 1, 4),
                    f"สถิติการใช้งานระบบ ปีการศึกษา {year}", fmt["title"])
    for ri, row in enumerate(rows_stat):
        is_hdr = ri == 0
        ws2.set_row(ri + 2, 22)
        for ci, val in enumerate(row):
            f = fmt["header"] if is_hdr else (fmt["cell_left"] if ci == 1 else fmt["cell"])
            ws2.write(ri + 2, ci, val if is_hdr else _to_number(val) if ci != 1 else val, f)

    wb.close()


# ── Cat 7: เช็กชื่อรายคาบ ────────────────────────────────────────────────────

def convert_cat7(data: dict, filepath: str, year: int = 2568) -> None:
    """
    data format: {room: {date_iso: [[rows]]}}
    Uses xlsxwriter constant_memory for large files.
    """
    import xlsxwriter
    Path(filepath).parent.mkdir(parents=True, exist_ok=True)

    wb = xlsxwriter.Workbook(filepath, {"constant_memory": True})
    fmt = _base_formats(wb)

    # Collect rooms in sorted order
    room_order = sorted(data.keys(), key=_room_sort_key)

    # Index sheet must be first
    idx = wb.add_worksheet("สารบัญ")
    idx.set_column("A:A", 12)
    idx.set_column("B:B", 20)
    idx.set_column("C:C", 38)
    idx.set_column("D:D", 18)
    idx.set_row(0, 26)
    idx.merge_range(0, 0, 0, 3, f"สารบัญ — รายงานเช็กชื่อรายคาบ ปีการศึกษา {year}", fmt["title"])
    idx.write_row(2, 0, ["ห้องเรียน", "จำนวนวันที่มีข้อมูล", "ช่วงวันที่", "ลิงค์"], fmt["header"])
    idx.freeze_panes(3, 0)
    idx_row = 3

    for room in room_order:
        dates_dict = data[room]
        dates_sorted = sorted(d for d, rows in dates_dict.items() if rows and len(rows) >= 3)
        sheet_name = safe_sheet_name(room)
        ws = wb.add_worksheet(sheet_name)
        ws.set_column("A:A", 5)
        ws.set_column("B:B", 12)
        ws.set_column("C:C", 28)
        ws.set_column(3, 14, 11)
        ws.set_row(0, 26)
        ws.merge_range(0, 0, 0, 14, f"รายงานเช็กชื่อรายคาบเรียน — ชั้น {room}", fmt["title"])

        row = 2
        for date_iso in sorted(dates_dict.keys()):
            table = dates_dict[date_iso]
            if not table or len(table) < 3:
                continue
            clean = table[1:]
            header_row = clean[0]
            ncols = max(len(header_row), 4)

            ws.set_row(row, 20)
            ws.merge_range(row, 0, row, ncols - 1, f"วันที่ {_fmt_thai_date(date_iso)}", fmt["band"])
            row += 1

            ws.set_row(row, 28)
            for ci, val in enumerate(header_row):
                ws.write(row, ci, val, fmt["header"])
            row += 1

            for data_row in clean[1:]:
                for ci, val in enumerate(data_row):
                    f = fmt["cell_left"] if ci == 2 else fmt["cell"]
                    ws.write(row, ci, _to_number(val) if ci != 2 else val, f)
                row += 1
            row += 1  # spacer

        date_range = (
            f"{_fmt_thai_date(dates_sorted[0])}  ถึง  {_fmt_thai_date(dates_sorted[-1])}"
            if dates_sorted else "ไม่มีข้อมูล"
        )
        idx.write(idx_row, 0, room, fmt["cell"])
        idx.write(idx_row, 1, len(dates_sorted), fmt["cell"])
        idx.write(idx_row, 2, date_range, fmt["cell_left"])
        idx.write_url(idx_row, 3, f"internal:'{sheet_name}'!A1", fmt["link"], "ไปที่ Sheet")
        idx_row += 1

    wb.close()
