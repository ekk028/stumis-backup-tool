# SPEC — STU-MIS Backup Tool

## User stories

**As a teacher/school admin, I want to:**
1. Type my school's STU-MIS URL, username, password once and back up everything in one click
2. Choose which categories to back up (so I can skip slow ones if time-pressed)
3. See a clear progress indicator (especially for Cat 7 which takes hours)
4. Save my login next time (so I don't retype every year)
5. Get a folder with Excel files that I can hand to my principal or archive
6. Have the tool work even if my school's STU-MIS subdomain is different from a colleague's

## Functional requirements

### Input
- **Base URL** (text input) — auto-normalize:
  - Strip trailing `/`
  - Accept with/without `https://`
  - Accept full path (e.g. `.../master/index.php`) — strip to root
- **Username** (text input)
- **Password** (masked input)
- **Remember password** (checkbox) — saves to encrypted local config
- **Save location** (folder picker) — default: `Documents\STU-MIS Backup\<school>-<year>\`
- **Categories** (8 checkboxes) — all checked by default

### Validation before "ดาวน์โหลด"
- URL must match `*.stumis06.com` (warn but allow override for variations)
- Username/password not empty
- Save folder writable
- (No time restriction — all individual endpoints work 24/7)

### Execution
- Disable inputs during run
- Run categories sequentially in chosen order (0 → 7); 0 last so user can review others first
- Show progress bar (overall %) and status line (current category, current item, X / Y)
- Allow **cancel** at any time — partial results should still be saved
- On crash/network error: retry once with 5s delay; if still fails, log error and continue with next item

### Output structure

```
<save_folder>/
├── 00_สรุปคะแนนพฤติกรรม/
│   └── 00_สรุปคะแนนพฤติกรรม_ทุกห้อง_ปีการศึกษา<year>.xls
├── 01_เช็กชื่อเช้า/
│   └── 01_เช็กชื่อเช้า_ทุกห้อง_ปีการศึกษา<year>.xlsx
├── 02_พฤติกรรมลบ/
│   └── 02_พฤติกรรมลบ_ทุกห้อง_ปีการศึกษา<year>.xls
├── 03_ความดี/...
├── 04_เยี่ยมบ้าน/...
├── 05_EQ/...
├── 06_SDQ/...
├── 07_เช็กชื่อรายคาบ/...
├── _raw_json/        ← intermediate JSON dumps (for debugging or re-conversion)
└── _README_STATUS.md ← summary of what was backed up, when, errors
```

### Non-functional
- Single `.exe` (PyInstaller `--onefile`), no external Python install needed
- UI in Thai
- Total dependencies after install: < 50 MB
- Handle screens at 1366×768 minimum
- Logging to `<save_folder>/_log.txt` (rotate by date)

## Out of scope (for v1)
- Multi-school batch backup
- Cloud upload (Google Drive, etc.)
- Restoring data back to STU-MIS
- Year-other-than-current (admin UI doesn't expose; would need direct DB access)
- Mac/Linux builds

## Edge cases to test
1. Wrong password → show clear error
2. URL with typo (e.g. missing subdomain) → DNS error → clear message
3. Network drops mid-fetch → retry, then skip with logged error
4. Cat 7 interrupted at 60% → resume from progress JSON
5. Disk full → stop with clear error, leave partial saved
6. School with 5 rooms vs school with 80 rooms → both work
7. Categories not all available (some schools may not have SDQ or EQ data) → handle empty table gracefully
8. Special chars in room names (`/`, spaces) → sheet name sanitization
