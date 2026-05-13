"""Background worker thread: orchestrates the full backup process."""
import json
import logging
import queue
import time
from datetime import date, timedelta
from pathlib import Path

from client import StumisClient, LoginError, NetworkError
from parsers import (
    parse_morning_attendance,
    parse_homevisit_summary,
    parse_homevisit_room,
    parse_period_attendance,
    parse_teacher_list,
)
from converters import save_raw, convert_cat1, convert_cat4, convert_cat7, convert_cat8

PROGRESS_FILE = "_progress_cat7.json"
STATUS_FILE = "_README_STATUS.md"


# ── date helpers ─────────────────────────────────────────────────────────────

def _academic_year_dates(thai_year: int):
    """Return (start_date, end_date) for the given Thai academic year."""
    ce = thai_year - 543
    return date(ce, 5, 15), date(ce + 1, 4, 30)


def _weekdays(start: date, end: date):
    d = start
    while d <= end:
        if d.weekday() < 5:
            yield d
        d += timedelta(days=1)


# ── retry helper ─────────────────────────────────────────────────────────────

def _retry(fn, tries: int = 2, delay: float = 5.0):
    last_exc = None
    for attempt in range(tries):
        try:
            return fn()
        except (NetworkError, Exception) as e:
            last_exc = e
            if attempt < tries - 1:
                time.sleep(delay)
    raise last_exc


# ── main worker ──────────────────────────────────────────────────────────────

def run_backup(
    client: StumisClient,
    save_dir: str,
    categories: set,
    year: int,
    msg_queue: queue.Queue,
    cancel_event,
    log_path: str | None = None,
) -> None:
    """
    Run in a background thread. Sends messages to msg_queue:
      ('progress', current, total, status_str)
      ('log', message_str)
      ('done', success_bool, message_str)
    """
    # Ensure save directory exists before anything else
    Path(save_dir).mkdir(parents=True, exist_ok=True)

    # Set up file logger
    logger = logging.getLogger("worker")
    logger.handlers.clear()
    if log_path:
        fh = logging.FileHandler(log_path, encoding="utf-8")
        fh.setFormatter(logging.Formatter("%(asctime)s %(message)s"))
        logger.addHandler(fh)
    logger.setLevel(logging.INFO)

    def log(msg: str):
        msg_queue.put(("log", msg))
        logger.info(msg)

    def check_cancel():
        if cancel_event.is_set():
            raise InterruptedError("ยกเลิกโดยผู้ใช้")

    save_path = Path(save_dir)
    errors: list[str] = []
    step = 0
    total = 0

    try:
        # ── discover rooms ────────────────────────────────────────────────
        log("กำลังโหลดรายชื่อห้องเรียน...")
        rooms = client.discover_rooms()
        log(f"พบ {len(rooms)} ห้อง: {', '.join(rooms[:5])}{'...' if len(rooms) > 5 else ''}")

        # ── calculate total steps ─────────────────────────────────────────
        bulk_cats = {0, 2, 3, 5, 6, 8} & categories
        total += len(bulk_cats)
        if 1 in categories:
            total += len(rooms) * 12
        if 4 in categories:
            total += 1 + len(rooms)
        if 7 in categories:
            start_d, end_d = _academic_year_dates(year)
            all_weekdays = list(_weekdays(start_d, end_d))
            total += len(rooms) * len(all_weekdays)
        if total == 0:
            total = 1

        def progress(status: str):
            nonlocal step
            step += 1
            msg_queue.put(("progress", step, total, status))

        # ── Cat 2: พฤติกรรมลบ ─────────────────────────────────────────────
        if 2 in categories:
            check_cancel()
            log("หมวด 2: กำลังดาวน์โหลดพฤติกรรมลบ...")
            try:
                content = _retry(lambda: client.get_bulk_excel("backup/report_bad_stu_excel.php"))
                d = save_path / "02_พฤติกรรมลบ"
                save_raw(content, str(d / f"02_พฤติกรรมลบ_ทุกห้อง_ปีการศึกษา{year}.xls"))
                log("หมวด 2: เสร็จแล้ว")
                progress("หมวด 2: พฤติกรรมลบ — เสร็จแล้ว")
            except Exception as e:
                msg = f"หมวด 2: ข้อผิดพลาด — {e}"
                log(msg); errors.append(msg)
                progress("หมวด 2: ข้อผิดพลาด (ข้ามไป)")

        # ── Cat 3: ความดี ─────────────────────────────────────────────────
        if 3 in categories:
            check_cancel()
            log("หมวด 3: กำลังดาวน์โหลดความดี...")
            try:
                content = _retry(lambda: client.get_bulk_excel("backup/report_good_stu_excel.php"))
                d = save_path / "03_ความดี"
                save_raw(content, str(d / f"03_ความดี_ทุกห้อง_ปีการศึกษา{year}.xls"))
                log("หมวด 3: เสร็จแล้ว")
                progress("หมวด 3: ความดี — เสร็จแล้ว")
            except Exception as e:
                msg = f"หมวด 3: ข้อผิดพลาด — {e}"
                log(msg); errors.append(msg)
                progress("หมวด 3: ข้อผิดพลาด (ข้ามไป)")

        # ── Cat 5: EQ ─────────────────────────────────────────────────────
        if 5 in categories:
            check_cancel()
            log("หมวด 5: กำลังดาวน์โหลด EQ...")
            try:
                content = _retry(lambda: client.get_bulk_excel("backup/report_eq-excel.php"))
                d = save_path / "05_EQ"
                save_raw(content, str(d / f"05_EQ_ทุกห้อง_ปีการศึกษา{year}.xls"))
                log("หมวด 5: เสร็จแล้ว")
                progress("หมวด 5: EQ — เสร็จแล้ว")
            except Exception as e:
                msg = f"หมวด 5: ข้อผิดพลาด — {e}"
                log(msg); errors.append(msg)
                progress("หมวด 5: ข้อผิดพลาด (ข้ามไป)")

        # ── Cat 6: SDQ ────────────────────────────────────────────────────
        if 6 in categories:
            check_cancel()
            log("หมวด 6: กำลังดาวน์โหลด SDQ...")
            try:
                content = _retry(lambda: client.get_bulk_excel("backup/report_sdq_excel.php"))
                d = save_path / "06_SDQ"
                save_raw(content, str(d / f"06_SDQ_ทุกห้อง_ปีการศึกษา{year}.xls"))
                log("หมวด 6: เสร็จแล้ว")
                progress("หมวด 6: SDQ — เสร็จแล้ว")
            except Exception as e:
                msg = f"หมวด 6: ข้อผิดพลาด — {e}"
                log(msg); errors.append(msg)
                progress("หมวด 6: ข้อผิดพลาด (ข้ามไป)")

        # ── Cat 1: เช็กชื่อเช้า ───────────────────────────────────────────
        if 1 in categories:
            log(f"หมวด 1: เช็กชื่อเช้า ({len(rooms)} ห้อง × 12 เดือน = {len(rooms)*12} คำขอ)")
            cat1_data: dict = {}
            for room in rooms:
                cat1_data[room] = {}
                for mm in ["01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12"]:
                    check_cancel()
                    try:
                        html = _retry(lambda r=room, m=mm: client.fetch_morning_attendance(r, m))
                        rows, note = parse_morning_attendance(html)
                        cat1_data[room][mm] = {"data": rows, "headerNote": note}
                    except Exception as e:
                        msg = f"หมวด 1: {room}/เดือน{mm} — {e}"
                        log(msg); errors.append(msg)
                        cat1_data[room][mm] = {"data": [], "headerNote": ""}
                    progress(f"หมวด 1 — {room} เดือน {mm}")
                    time.sleep(0.05)

            # Save raw JSON for debugging
            raw_dir = save_path / "_raw_json"
            raw_dir.mkdir(exist_ok=True)
            with open(raw_dir / "cat1.json", "w", encoding="utf-8") as f:
                json.dump(cat1_data, f, ensure_ascii=False)

            log("หมวด 1: กำลังสร้างไฟล์ Excel...")
            d = save_path / "01_เช็กชื่อเช้า"
            convert_cat1(cat1_data, str(d / f"01_เช็กชื่อเช้า_ทุกห้อง_ปีการศึกษา{year}.xlsx"), year)
            log("หมวด 1: เสร็จแล้ว")

        # ── Cat 4: เยี่ยมบ้าน ─────────────────────────────────────────────
        if 4 in categories:
            log("หมวด 4: เยี่ยมบ้าน")
            try:
                check_cancel()
                sum_html = _retry(client.fetch_homevisit_summary)
                summary_rows = parse_homevisit_summary(sum_html)
                hv_rooms = client.discover_homevisit_rooms()
                log(f"หมวด 4: พบ {len(hv_rooms)} ห้อง")

                cat4_data: dict = {"summary": summary_rows, "rooms": {}}
                progress("หมวด 4: สรุปภาพรวม")

                for params in hv_rooms:
                    check_cancel()
                    room_raw = params.get("room", "")
                    import urllib.parse
                    room_display = urllib.parse.unquote(room_raw)
                    try:
                        html = _retry(lambda p=params: client.fetch_homevisit_room(
                            p.get("homeid", ""), p.get("room", ""), p.get("yy", "")
                        ))
                        cat4_data["rooms"][room_display] = parse_homevisit_room(html)
                    except Exception as e:
                        msg = f"หมวด 4: {room_display} — {e}"
                        log(msg); errors.append(msg)
                        cat4_data["rooms"][room_display] = []
                    progress(f"หมวด 4 — ห้อง {room_display}")
                    time.sleep(0.05)

                raw_dir = save_path / "_raw_json"
                raw_dir.mkdir(exist_ok=True)
                with open(raw_dir / "cat4.json", "w", encoding="utf-8") as f:
                    json.dump(cat4_data, f, ensure_ascii=False)

                log("หมวด 4: กำลังสร้างไฟล์ Excel...")
                d = save_path / "04_เยี่ยมบ้าน"
                convert_cat4(cat4_data, str(d / f"04_เยี่ยมบ้าน_ทุกห้อง_ปีการศึกษา{year}.xlsx"), year)
                log("หมวด 4: เสร็จแล้ว")
            except InterruptedError:
                raise
            except Exception as e:
                msg = f"หมวด 4: ข้อผิดพลาด — {e}"
                log(msg); errors.append(msg)

        # ── Cat 7: เช็กชื่อรายคาบ (big one) ─────────────────────────────
        if 7 in categories:
            start_d, end_d = _academic_year_dates(year)
            all_weekdays = list(_weekdays(start_d, end_d))
            total_fetches = len(rooms) * len(all_weekdays)
            log(f"หมวด 7: {len(rooms)} ห้อง × {len(all_weekdays)} วัน = {total_fetches:,} คำขอ (อาจใช้เวลาหลายชั่วโมง)")

            progress_path = save_path / PROGRESS_FILE
            cat7_data: dict = {}
            done_pairs: set = set()

            # Resume from progress file if exists
            if progress_path.exists():
                try:
                    with open(progress_path, encoding="utf-8") as f:
                        cat7_data = json.load(f)
                    for room, dates_dict in cat7_data.items():
                        for iso, v in dates_dict.items():
                            if v is not None:
                                done_pairs.add((room, iso))
                    log(f"หมวด 7: ต่อจาก progress — {len(done_pairs):,} รายการที่เสร็จแล้ว")
                except Exception:
                    cat7_data = {}

            fetch_count = 0
            for room in rooms:
                if room not in cat7_data:
                    cat7_data[room] = {}
                for wd in all_weekdays:
                    check_cancel()
                    iso = wd.strftime("%Y-%m-%d")
                    if (room, iso) in done_pairs:
                        progress(f"หมวด 7 — {room} — {iso} (ข้าม)")
                        continue
                    try:
                        html = _retry(lambda r=room, d=iso: client.fetch_period_attendance(r, d), tries=2, delay=5)
                        rows = parse_period_attendance(html)
                        cat7_data[room][iso] = rows
                    except Exception as e:
                        msg = f"หมวด 7: {room}/{iso} — {e}"
                        log(msg); errors.append(msg)
                        cat7_data[room][iso] = []

                    fetch_count += 1
                    if fetch_count % 200 == 0:
                        with open(progress_path, "w", encoding="utf-8") as f:
                            json.dump(cat7_data, f, ensure_ascii=False)
                        log(f"หมวด 7: บันทึก progress แล้ว ({fetch_count:,} คำขอ)")

                    progress(f"หมวด 7 — {room} — {iso}")
                    time.sleep(0.05)

            raw_dir = save_path / "_raw_json"
            raw_dir.mkdir(exist_ok=True)
            with open(raw_dir / "cat7.json", "w", encoding="utf-8") as f:
                json.dump(cat7_data, f, ensure_ascii=False)

            log("หมวด 7: กำลังสร้างไฟล์ Excel...")
            d = save_path / "07_เช็กชื่อรายคาบ"
            convert_cat7(cat7_data, str(d / f"07_เช็กชื่อรายคาบ_ทุกห้อง_ปีการศึกษา{year}.xlsx"), year)
            if progress_path.exists():
                progress_path.unlink()
            log("หมวด 7: เสร็จแล้ว")

        # ── Cat 8: ข้อมูลครูและบุคลากร ───────────────────────────────────
        if 8 in categories:
            check_cancel()
            log("หมวด 8: กำลังดาวน์โหลดข้อมูลครูและบุคลากร...")
            try:
                html_main = _retry(client.fetch_teacher_list)
                rows_main = parse_teacher_list(html_main)
                log(f"หมวด 8: รายชื่อครู {max(len(rows_main)-1,0)} คน")

                html_stat = _retry(client.fetch_teacher_excel)
                rows_stat = parse_teacher_list(html_stat)
                log(f"หมวด 8: สถิติการใช้งาน {max(len(rows_stat)-1,0)} รายการ")

                d = save_path / "08_ข้อมูลครูและบุคลากร"
                convert_cat8(rows_main, rows_stat,
                             str(d / f"08_ข้อมูลครูและบุคลากร_ปีการศึกษา{year}.xlsx"), year)
                log("หมวด 8: เสร็จแล้ว")
                progress("หมวด 8: ข้อมูลครูและบุคลากร — เสร็จแล้ว")
            except Exception as e:
                msg = f"หมวด 8: ข้อผิดพลาด — {e}"
                log(msg); errors.append(msg)
                progress("หมวด 8: ข้อผิดพลาด (ข้ามไป)")

        # ── Cat 0: สรุปคะแนนพฤติกรรม (only 18:00-06:00) ──────────────────
        if 0 in categories:
            from datetime import datetime
            now_hour = datetime.now().hour
            if 6 <= now_hour < 18:
                msg = "หมวด 0: ข้ามเนื่องจากอยู่นอกช่วงเวลา 18:00–06:00"
                log(msg); errors.append(msg)
                progress("หมวด 0: ข้าม (นอกช่วงเวลา)")
            else:
                check_cancel()
                log("หมวด 0: กำลังดาวน์โหลดสรุปคะแนนพฤติกรรม...")
                try:
                    content = _retry(lambda: client.get_bulk_excel("indextecher-all-excel.php"))
                    d = save_path / "00_สรุปคะแนนพฤติกรรม"
                    save_raw(content, str(d / f"00_สรุปคะแนนพฤติกรรม_ทุกห้อง_ปีการศึกษา{year}.xls"))
                    log("หมวด 0: เสร็จแล้ว")
                    progress("หมวด 0: สรุปคะแนนพฤติกรรม — เสร็จแล้ว")
                except Exception as e:
                    msg = f"หมวด 0: ข้อผิดพลาด — {e}"
                    log(msg); errors.append(msg)
                    progress("หมวด 0: ข้อผิดพลาด")

        # ── Write status file ─────────────────────────────────────────────
        from datetime import datetime
        status_lines = [
            f"# STU-MIS Backup Status",
            f"",
            f"- ปีการศึกษา: {year}",
            f"- วันที่สำรอง: {datetime.now().strftime('%d/%m/%Y %H:%M')}",
            f"- URL: {client.base_url}",
            f"- จำนวนห้อง: {len(rooms)}",
            f"- หมวดที่เลือก: {sorted(categories)}",
            f"",
        ]
        if errors:
            status_lines += ["## ข้อผิดพลาด", ""] + [f"- {e}" for e in errors]
        else:
            status_lines += ["## ผลลัพธ์: สำเร็จทั้งหมด ✓"]
        with open(save_path / STATUS_FILE, "w", encoding="utf-8") as f:
            f.write("\n".join(status_lines))

        msg_queue.put(("done", True, f"ดาวน์โหลดเสร็จสมบูรณ์ ({len(errors)} ข้อผิดพลาด)"))

    except InterruptedError as e:
        msg_queue.put(("done", False, str(e)))
    except LoginError as e:
        msg_queue.put(("done", False, f"เข้าสู่ระบบล้มเหลว: {e}"))
    except Exception as e:
        import traceback
        msg_queue.put(("done", False, f"ข้อผิดพลาด: {e}\n{traceback.format_exc()}"))
