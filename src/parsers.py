"""Parse STU-MIS HTML pages into structured data (lists of rows)."""
import re
from bs4 import BeautifulSoup


def _soup(html: str) -> BeautifulSoup:
    try:
        return BeautifulSoup(html, "lxml")
    except Exception:
        return BeautifulSoup(html, "html.parser")


def _table_to_rows(table) -> list[list[str]]:
    rows = []
    for tr in table.find_all("tr"):
        row = [td.get_text(" ", strip=True) for td in tr.find_all(["td", "th"])]
        rows.append(row)
    return rows


def _find_largest_table(soup, min_rows: int = 2):
    tables = soup.find_all("table")
    if not tables:
        return None
    candidates = [t for t in tables if len(t.find_all("tr")) >= min_rows]
    if not candidates:
        return None
    return max(candidates, key=lambda t: len(t.find_all("tr")))


def parse_morning_attendance(html: str) -> tuple[list[list[str]], str]:
    """
    Returns (rows, header_note).
    rows[0] is skipped by converters (merged header), rows[1] is column headers, rows[2+] are data.
    """
    soup = _soup(html)
    # Extract teacher/class note from page text
    header_note = ""
    body_text = soup.get_text(" ", strip=True)
    m = re.search(r"ครูประจำชั้น[^\n\r]{0,120}", body_text)
    if m:
        header_note = m.group(0).strip()

    table = _find_largest_table(soup)
    if table is None:
        return [], header_note
    return _table_to_rows(table), header_note


def parse_homevisit_summary(html: str) -> list[list[str]]:
    soup = _soup(html)
    table = _find_largest_table(soup)
    return _table_to_rows(table) if table else []


def parse_homevisit_room(html: str) -> list[list[str]]:
    soup = _soup(html)
    table = _find_largest_table(soup)
    return _table_to_rows(table) if table else []


def parse_teacher_list(html: str) -> list[list[str]]:
    """Parse Teacher-main.php → rows with ที่, Username, ชื่อ-สกุล, ครูที่ปรึกษา"""
    soup = _soup(html)
    table = _find_largest_table(soup)
    return _table_to_rows(table) if table else []


def parse_period_attendance(html: str) -> list[list[str]]:
    """
    Returns all rows from the main table, or [] if the day has no data.
    Caller should check len(rows) >= 3 before using.
    """
    soup = _soup(html)
    table = _find_largest_table(soup, min_rows=3)
    if table is None:
        return []
    rows = _table_to_rows(table)
    return rows if len(rows) >= 3 else []
