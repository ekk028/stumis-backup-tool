"""HTTP client for STU-MIS with login, session management, and all fetch methods."""
import re
import urllib.parse
import requests
from bs4 import BeautifulSoup


class LoginError(Exception):
    pass


class NetworkError(Exception):
    pass


def normalize_url(url: str) -> str:
    """Return bare base URL: https://xxx.stumis06.com (no trailing slash, no /master/...)"""
    url = url.strip()
    if not url.startswith("http"):
        url = "https://" + url
    url = re.sub(r"/master(/.*)?$", "", url)
    return url.rstrip("/")


class StumisClient:
    def __init__(self, base_url: str):
        self.base_url = normalize_url(base_url)
        self.master_url = self.base_url + "/master"
        self.session = requests.Session()
        self.session.headers.update({"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"})
        self._username = ""
        self._password = ""

    def _url(self, path: str) -> str:
        return f"{self.master_url}/{path}"

    # Phrases specific to STU-MIS login/auth failure pages
    _LOGIN_FAIL_MARKERS = (
        ('type="password"',),             # any login form has a password input
        ("false.php",),
        ("กลับไปกรอกข้อมูลใหม่",),
        ("ไม่สามารถเข้าสู่ระบบ",),
    )

    def _is_login_page(self, html: str) -> bool:
        """Return True only if ALL tokens in any marker group appear in html."""
        if not html:
            return False
        for group in self._LOGIN_FAIL_MARKERS:
            if all(token in html for token in group):
                return True
        return False

    def login(self, username: str, password: str) -> None:
        self._username = username
        self._password = password
        login_url = self._url("index.php")
        try:
            # GET first: initialize PHPSESSID and read real form structure
            get_resp = self.session.get(login_url, timeout=30)
            get_resp.encoding = "utf-8"

            soup = BeautifulSoup(get_resp.text, "lxml")
            form = soup.find("form")

            # Build post_url from form action (fall back to login_url)
            post_url = login_url
            if form and form.get("action"):
                post_url = urllib.parse.urljoin(login_url, form["action"])

            # Collect ALL form fields (hidden + visible defaults)
            form_data: dict = {}
            if form:
                for inp in form.find_all("input"):
                    itype = (inp.get("type") or "text").lower()
                    name  = inp.get("name", "")
                    val   = inp.get("value", "")
                    if not name:
                        continue
                    if itype == "checkbox":
                        continue          # skip unchecked checkboxes
                    form_data[name] = val

            # Find actual field names for username (type=text) and password (type=password)
            user_field = "user"
            pass_field = "pass"
            if form:
                for inp in form.find_all("input"):
                    itype = (inp.get("type") or "text").lower()
                    name  = inp.get("name", "")
                    if not name:
                        continue
                    if itype == "text":
                        user_field = name
                    elif itype == "password":
                        pass_field = name

            form_data[user_field] = username
            form_data[pass_field] = password

            resp = self.session.post(
                post_url,
                data=form_data,
                headers={
                    "Referer": login_url,
                    "Origin":  self.base_url,
                    "Accept":  "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
                },
                allow_redirects=True,
                timeout=30,
            )
            resp.encoding = "utf-8"
        except requests.RequestException as e:
            raise NetworkError(f"เชื่อมต่อไม่ได้: {e}") from e

        html = resp.text
        # Explicit failure markers
        if self._is_login_page(html):
            raise LoginError("เข้าสู่ระบบไม่สำเร็จ — ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง")
        # If final URL is still index.php, login didn't redirect → failed
        if resp.url.rstrip("/").endswith("index.php"):
            snippet = html[:400].replace("\n", " ").strip()
            raise LoginError(
                f"เข้าสู่ระบบไม่สำเร็จ — เซิร์ฟเวอร์ไม่ยืนยันการเข้าสู่ระบบ\n"
                f"(URL ยังอยู่ที่: {resp.url})\n"
                f"ข้อความ: {snippet}"
            )

    def _get_html(self, path: str, method: str = "GET", data: dict = None, timeout: int = 30,
                  _relogin: bool = True) -> str:
        try:
            if method == "POST":
                resp = self.session.post(self._url(path), data=data, timeout=timeout)
            else:
                resp = self.session.get(self._url(path), timeout=timeout)
            resp.encoding = "utf-8"
        except requests.RequestException as e:
            raise NetworkError(f"เครือข่ายขัดข้อง: {e}") from e

        if _relogin and self._is_login_page(resp.text) and self._username:
            self.login(self._username, self._password)
            return self._get_html(path, method, data, timeout, _relogin=False)

        return resp.text

    # Magic bytes for valid Excel files
    _EXCEL_MAGIC = (
        b"\xd0\xcf\x11\xe0",  # .xls  (OLE2)
        b"PK\x03\x04",        # .xlsx (ZIP)
    )

    def get_bulk_excel(self, path: str, timeout: int = 120) -> bytes:
        try:
            resp = self.session.get(self._url(path), timeout=timeout)
        except requests.RequestException as e:
            raise NetworkError(f"เครือข่ายขัดข้อง: {e}") from e

        content = resp.content
        ct = resp.headers.get("Content-Type", "").lower()

        # Accept if server explicitly says it's an Excel/binary file
        is_excel_ct = any(t in ct for t in (
            "vnd.ms-excel", "spreadsheetml", "octet-stream", "application/excel",
        ))
        has_magic = any(content.startswith(m) for m in self._EXCEL_MAGIC)

        if is_excel_ct or has_magic:
            return content

        # Otherwise it's a plain HTML page → session error or wrong URL
        resp.encoding = "utf-8"
        html = resp.text
        if self._is_login_page(html):
            raise LoginError("session หมดอายุ — กรุณาตรวจสอบ Username/Password")
        snippet = html[:300].replace("\n", " ").strip()
        raise LoginError(
            f"เซิร์ฟเวอร์คืน HTML แทนไฟล์ Excel\n"
            f"URL: {resp.url}\n"
            f"Content-Type: {ct}\n"
            f"ข้อความ: {snippet}"
        )

    def discover_rooms(self) -> list[str]:
        html = self._get_html("report-check-MM-t.php")
        soup = BeautifulSoup(html, "lxml")
        sel = soup.find("select", {"name": "room"})
        if not sel:
            return []
        return [o.get("value", o.text.strip()) for o in sel.find_all("option") if o.get("value", "").strip()]

    def discover_homevisit_rooms(self) -> list[dict]:
        """Parse stu_home/index.php for room→query param mappings."""
        html = self._get_html("stu_home/index.php")
        soup = BeautifulSoup(html, "lxml")
        rooms = []
        seen = set()
        for a in soup.find_all("a", href=True):
            href = a["href"]
            m = re.search(r"process_home_room\.php\?(.+)", href)
            if not m:
                continue
            qs = m.group(1)
            params = dict(urllib.parse.parse_qsl(qs))
            key = params.get("room", "") + params.get("homeid", "")
            if key and key not in seen:
                seen.add(key)
                rooms.append(params)
        return rooms

    def fetch_morning_attendance(self, room: str, month: str) -> str:
        return self._get_html(
            "report-check-MM-t.php",
            method="POST",
            data={"room": room, "MM": month, "button": "ค้นหา"},
        )

    def fetch_homevisit_summary(self) -> str:
        return self._get_html("stu_home/index.php")

    def fetch_homevisit_room(self, homeid: str, room: str, yy: str) -> str:
        qs = urllib.parse.urlencode({"homeid": homeid, "room": room, "yy": yy})
        return self._get_html(f"stu_home/process_home_room.php?{qs}")

    def fetch_teacher_list(self) -> str:
        return self._get_html("Teacher-main.php")

    def fetch_teacher_excel(self) -> str:
        return self._get_html("Teacher-main-excel.php")

    def fetch_period_attendance(self, room: str, date_iso: str) -> str:
        return self._get_html(
            "vicha_report-room3.php",
            method="POST",
            data={"room": room, "date": date_iso, "button": "ค้นหา"},
        )
