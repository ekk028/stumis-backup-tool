# CLAUDE.md — STU-MIS Backup Tool

> **For Claude Code** — read this first, then SPEC.md and ENDPOINTS.md before writing code.

## Project goal
Build a Windows desktop tool that backs up data from any school running on the **STU-MIS** student management platform (stumis06.com). Schools must back up before the annual academic-year transition or data is lost.

The tool must work for **any school** on STU-MIS, not just one — the URL is user-provided.

## What's already done (by Claude in Cowork mode)
- Manually backed up one school (wangluangpitt.stumis06.com) for academic year 2568
- Discovered all the HTTP endpoints (see `ENDPOINTS.md`)
- Wrote JSON→Excel converters in Python (see `starter_code/`)
- Captured sample JSON dumps (see `sample_data/`) so you don't need to re-fetch to develop

## Tech stack
- **Python 3.10+**
- **GUI:** `customtkinter` (modern look) — fall back to `tkinter` if needed
- **HTTP:** `requests` (use `requests.Session()` for cookie/session handling)
- **HTML parsing:** `beautifulsoup4` + `lxml`
- **Excel output:** `xlsxwriter` (much faster than openpyxl for large files)
- **Packaging:** `pyinstaller` → single `.exe` for Windows

## Build order (start small, iterate)

1. **MVP — bulk Excel endpoints only** (Cat 2/3/5/6):
   - Login → GET each bulk URL → save bytes to disk → done
   - Build a barebones UI: URL/user/pass/save-folder/checkboxes/button
   - Build `.exe` with PyInstaller, test on a real school
2. **Add Cat 1** (เช็กชื่อเช้า): 22 rooms × 12 months ≈ 264 fetches; parse HTML tables; convert to one Excel with 23 sheets
3. **Add Cat 4** (เยี่ยมบ้าน): summary table + per-room detail (~23 fetches)
4. **Add Cat 7** (เช็กชื่อรายคาบ): the biggest one — ~5500 fetches over school weekdays; needs progress bar and resume-on-crash
5. **Add Cat 0** (สรุปคะแนนพฤติกรรม): single endpoint `indextecher-all-excel.php` — works 24/7 (the time-restricted version is the unrelated `backup/index.php` aggregate page; we don't use it)

## UI Layout

```
┌──────────────────────────────────────────────────────────────────┐
│ 🏫 STU-MIS Backup Tool                                           │
├──────────────────────────────────────────────────────────────────┤
│ Left pane:                       │  Right pane:                  │
│ URL: [https://___.stumis06.com/] │  ☑ 1. เช็กชื่อเช้า            │
│ User: [____________]             │  ☑ 2. พฤติกรรมลบ              │
│ Pass: [____________]             │  ☑ 3. ความดี                  │
│ ☑ จดจำรหัสผ่าน                   │  ☑ 4. เยี่ยมบ้าน               │
│                                   │  ☑ 5. EQ                     │
│ Save: [D:\backup\___]  [📁]      │  ☑ 6. SDQ                    │
│                                   │  ☑ 7. เช็กชื่อรายคาบ (ช้า)    │
│ [⬇ ดาวน์โหลด]                    │  ☑ 0. สรุปคะแนนพฤติกรรม      │
│                                   │                              │
│                                   │                              │
│ Progress: ███████░░░░ 47%                                       │
│ Status: หมวด 7 — ม.3/2 — 15 ก.ย. 68 (2,584 / 5,522)             │
└──────────────────────────────────────────────────────────────────┘
```

## Important things to handle

- **URL normalization:** user types `https://wangluangpitt.stumis06.com/` — auto-append `master/index.php` for login
- **Login form fields:** `user`, `pass`, `keep_login`, `button2` (POST to `/index.php`)
- **Session cookies:** use `requests.Session()`; if a request returns the login wall, auto-relogin once
- **Auto-discover rooms:** parse the room dropdown from any page so the tool works for any school
- **School year:** academic year 2568 spans May 2025 to April 2026 (Thai system); the user should be able to choose year (default current academic year)
- **Concurrency:** keep it sequential — server is slow and may rate-limit
- **Resume after crash:** save progress JSON every 200 fetches; resume detects the file and continues
- **"Remember password" checkbox:** save to a local config file (e.g. `%APPDATA%\stumis-backup\config.json`); for safety use Windows DPAPI (`pywin32`) or `cryptography` library to encrypt

## Repo layout (suggested)

```
stumis-backup-tool/
├── CLAUDE.md              ← this file
├── SPEC.md                ← detailed requirements
├── ENDPOINTS.md           ← API documentation
├── README.md              ← end-user docs (Thai)
├── requirements.txt
├── pyproject.toml
├── src/
│   ├── main.py            ← entry point (Tkinter app)
│   ├── client.py          ← StumisClient (login, fetch endpoints)
│   ├── parsers.py         ← HTML → structured data
│   ├── converters.py      ← structured data → xlsx
│   ├── config.py          ← save/load user config
│   └── gui.py             ← customtkinter UI
├── sample_data/           ← real JSON dumps from one school
└── starter_code/          ← reference converters from Cowork session
```

## Definition of "done"

- Builds to a single `.exe` (no Python install required by end user)
- Tested successfully on at least 2 different schools (different subdomains)
- All 8 categories produce valid Excel files when checked
- Progress bar accurate and pause/resume works
- Save location remembered between runs
- "Remember password" works and password is encrypted at rest
- Error messages are clear in Thai

## Distribution
- Build .exe with PyInstaller (one-file mode)
- Upload to GitHub Releases (or Google Drive if private)
- Write a short Thai installation/usage guide for non-technical teachers
