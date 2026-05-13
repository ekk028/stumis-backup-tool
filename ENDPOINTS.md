# STU-MIS API Endpoints

All paths are relative to `<BASE_URL>/master/` unless noted. Verified against `wangluangpitt.stumis06.com` on 13 พ.ค. 2569.

## Authentication

| Endpoint | Method | Body fields | Returns |
|---|---|---|---|
| `/index.php` | POST | `user`, `pass`, `keep_login`, `button2` | redirects to `/master/main-master.php` on success; same page (with error) on failure |

Detect logged-in state: if a GET to any admin URL returns HTML containing `name="user"` and `name="pass"` and is < 15 KB, the session is expired — re-login.

## Master dashboard
- `main-master.php` — landing page after login; lists all menus

## Cat 0 — สรุปคะแนนพฤติกรรม (Behavior summary, all classes)
| Endpoint | Method | Notes |
|---|---|---|
| `indextecher-all-excel.php` | GET | Returns Excel directly. Works 24/7. |

**Note:** Do NOT confuse with `backup/index.php` (the green "Backup" button on the dashboard) — that page IS time-restricted (17:00–06:00 only) because it bundles everything server-side. We bypass it entirely by hitting the individual download endpoints below.

## Cat 1 — เช็กชื่อเช้า / กิจกรรมหน้าเสาธง (Morning attendance)
| Endpoint | Method | Body | Returns |
|---|---|---|---|
| `report-check-MM-t.php` | POST | `room` (e.g. `ม.1/1`), `MM` (`01`-`12`), `button=ค้นหา` | HTML page; the largest table with header `รหัสนักเรียน|ชื่อ สกุล|<dates>` is the data |

Rooms come from the `<select name="room">` on the page — auto-discover.
Months: `01` through `12` (calendar months, but in academic year scope the system maps them to the current year).

Year selector: **not exposed** in admin UI — only current academic year is accessible via this endpoint.

## Cat 2 — พฤติกรรมไม่เหมาะสม (Bad behavior, bulk)
| Endpoint | Method |
|---|---|
| `backup/report_bad_stu_excel.php` | GET |

Returns Excel binary directly (saves as `Report_bad.xls`).

## Cat 3 — ความดี (Good behavior, bulk)
| Endpoint | Method |
|---|---|
| `backup/report_good_stu_excel.php` | GET |

Returns Excel binary directly (`Report_good.xls`).

## Cat 4 — เยี่ยมบ้าน (Home visits)
| Endpoint | Method | Notes |
|---|---|---|
| `stu_home/index.php` | GET | Lists all rooms with per-row "ข้อมูลเยี่ยมบ้าน" link → `process_home_room.php?homeid=...&room=...&yy=...` |
| `stu_home/process_home_room.php` | GET with query string | Per-room detail; query params: `homeid`, `room`, `yy` |

Auto-discover the room→query mappings by parsing `stu_home/index.php`. The largest table on each detail page is the data.

## Cat 5 — EQ
| Endpoint | Method |
|---|---|
| `backup/report_eq-excel.php` | GET |

Returns Excel binary (`Report_eq.xls`).

## Cat 6 — SDQ
| Endpoint | Method |
|---|---|
| `backup/report_sdq_excel.php` | GET |

Returns Excel binary (`Report_sdq_all.xls`).

## Cat 7 — เช็กชื่อรายคาบ (Class period attendance)
| Endpoint | Method | Body |
|---|---|---|
| `vicha_report-room3.php` | POST | `room`, `date` (`YYYY-MM-DD`), `button=ค้นหา` |

Returns HTML with student×period attendance for that date. Empty days (school closed) return a page with no data table.

**Coverage strategy:** loop weekdays from `(year-543)-05-15` through `(year-542)-04-30` for each room. Detect empty days and skip storage. Expect ~5,000 fetches per school year × school. Use ~50ms delay between fetches.

## Misc endpoints (admin tabs, for reference, not part of backup)
- `Teacher-main.php` — manage teachers
- `stu-edit.php` — student data
- `setgoodbad/good_main.php` — score rules
- `sdq/sdq_home.php` — EQ/SDQ admin pages
- `backup/index.php` — server-side aggregate backup (only 17:00–06:00)

## Common pitfalls
- All Thai text is UTF-8 — set `response.encoding = 'utf-8'` if `requests` auto-detection fails
- Some pages redirect through a "false.php" wrapper when session is invalid — detect this and re-login
- `vicha_home.php` returned an error when we accessed it as admin during testing — avoid it; use `vicha_report-room3.php` instead
- Form POSTs require the `button` field with the literal Thai label `ค้นหา` — without it, the server returns an empty page
