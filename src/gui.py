"""STU-MIS Backup Tool — customtkinter UI."""
from __future__ import annotations

import os
import queue
import subprocess
import sys
import threading
from datetime import date, datetime
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox

import customtkinter as ctk

from client import StumisClient, LoginError, NetworkError
from config import decrypt_password, encrypt_password, load_config, save_config
import worker as backup_worker

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

TITLE = "STU-MIS Backup Tool"
VERSION = "v1.0"
WIN_W, WIN_H = 960, 760

CAT_LABELS = {
    0: "0. สรุปคะแนนพฤติกรรม  (เฉพาะ 18:00–06:00)",
    1: "1. เช็กชื่อเช้า / หน้าเสาธง",
    2: "2. พฤติกรรมไม่เหมาะสม",
    3: "3. ความดี",
    4: "4. เยี่ยมบ้าน",
    5: "5. EQ",
    6: "6. SDQ",
    7: "7. เช็กชื่อรายคาบ  (ช้ามาก)",
    8: "8. ข้อมูลครูและบุคลากร",
}
CAT_ORDER = [0, 1, 2, 3, 4, 5, 6, 7, 8]
CAT_RED   = {0, 7}   # show label in red


def _current_thai_year() -> int:
    today = date.today()
    # Academic year starts May 16; before that still belongs to previous year
    if today.month > 5 or (today.month == 5 and today.day >= 16):
        return today.year + 543
    else:
        return today.year + 542


def _default_save_dir(url: str, year: int) -> str:
    school = url.replace("https://", "").split(".")[0] if url else "school"
    return str(Path.home() / "Documents" / "STU-MIS Backup" / f"{school}-{year}")


# ─────────────────────────────────────────────────────────────────────────────

class App(ctk.CTk):
    def __init__(self) -> None:
        super().__init__()
        self.title(TITLE)
        self.geometry(f"{WIN_W}x{WIN_H}")
        self.minsize(860, 680)

        self._msg_queue: queue.Queue = queue.Queue()
        self._cancel_event = threading.Event()
        self._running = False

        self._build_ui()
        self._load_config()
        self.after(100, self._poll_queue)

    # ── layout ────────────────────────────────────────────────────────────────
    #
    #  row 0  │ title bar                              │ fixed 46 px
    #  row 1  │ left (credentials) + right (cats)      │ expands
    #  row 2  │ action bar: [⬇ ดาวน์โหลด] [✕ ยกเลิก] │ fixed 60 px
    #  row 3  │ progress + status + log + [📂] [ⓘ]    │ fixed ~230 px

    def _build_ui(self) -> None:
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)   # middle stretches

        self._build_title()
        middle = self._build_middle()
        self._build_left(middle)
        self._build_right(middle)
        self._build_action_bar()
        self._build_bottom()

    # ── title ─────────────────────────────────────────────────────────────────

    def _build_title(self) -> None:
        bar = ctk.CTkFrame(self, height=46, corner_radius=0)
        bar.grid(row=0, column=0, sticky="ew")
        ctk.CTkLabel(
            bar, text=f"  {TITLE}",
            font=ctk.CTkFont("TH Sarabun New", size=22, weight="bold"),
        ).pack(side="left", padx=8)
        ctk.CTkLabel(
            bar, text=VERSION,
            font=ctk.CTkFont("TH Sarabun New", size=13, weight="bold"),
            text_color=("gray50", "gray60"),
        ).pack(side="left", padx=(0, 8))

    # ── middle ────────────────────────────────────────────────────────────────

    def _build_middle(self) -> ctk.CTkFrame:
        mid = ctk.CTkFrame(self, fg_color="transparent")
        mid.grid(row=1, column=0, sticky="nsew", padx=12, pady=(8, 4))
        mid.grid_columnconfigure(0, weight=3)
        mid.grid_columnconfigure(1, weight=2)
        mid.grid_rowconfigure(0, weight=1)
        return mid

    # ── left: credentials ─────────────────────────────────────────────────────

    def _toggle_password(self) -> None:
        self.pass_entry.configure(show="" if self._show_pass_var.get() else "●")

    def _patch_entry(self, entry: ctk.CTkEntry) -> ctk.CTkEntry:
        """Enable Ctrl+C/V/X/A and right-click menu on any CTkEntry."""
        inner = entry._entry  # underlying tk.Entry

        # Keyboard shortcuts
        def _paste(e):
            inner.event_generate("<<Paste>>")
            return "break"
        def _copy(e):
            inner.event_generate("<<Copy>>")
            return "break"
        def _cut(e):
            inner.event_generate("<<Cut>>")
            return "break"
        def _select_all(e):
            inner.select_range(0, "end")
            inner.icursor("end")
            return "break"

        for widget in (entry, inner):
            widget.bind("<Control-v>", _paste)
            widget.bind("<Control-V>", _paste)
            widget.bind("<Control-c>", _copy)
            widget.bind("<Control-C>", _copy)
            widget.bind("<Control-x>", _cut)
            widget.bind("<Control-X>", _cut)
            widget.bind("<Control-a>", _select_all)
            widget.bind("<Control-A>", _select_all)

        # Right-click context menu
        menu = tk.Menu(self, tearoff=0, font=("TH Sarabun New", 13))
        menu.add_command(label="ตัด",          command=lambda: inner.event_generate("<<Cut>>"))
        menu.add_command(label="คัดลอก",       command=lambda: inner.event_generate("<<Copy>>"))
        menu.add_command(label="วาง",          command=lambda: inner.event_generate("<<Paste>>"))
        menu.add_separator()
        menu.add_command(label="เลือกทั้งหมด", command=lambda: inner.select_range(0, "end"))

        def _show_menu(e):
            try:
                menu.tk_popup(e.x_root, e.y_root)
            finally:
                menu.grab_release()

        for widget in (entry, inner):
            widget.bind("<Button-3>", _show_menu)

        return entry

    def _lbl(self, parent, text: str) -> ctk.CTkLabel:
        return ctk.CTkLabel(
            parent, text=text,
            font=ctk.CTkFont("TH Sarabun New", size=17, weight="bold"), anchor="w",
        )

    def _build_left(self, parent: ctk.CTkFrame) -> None:
        left = ctk.CTkFrame(parent)
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 6))
        left.grid_columnconfigure(0, weight=1)

        r = 0

        # URL
        self._lbl(left, "URL โรงเรียน:").grid(row=r, column=0, sticky="w", padx=12, pady=(12, 2))
        r += 1
        self.url_entry = self._patch_entry(ctk.CTkEntry(
            left, placeholder_text="https://xxx.stumis06.com/",
            font=ctk.CTkFont("TH Sarabun New", size=16, weight="bold"), height=32,
        ))
        self.url_entry.grid(row=r, column=0, sticky="ew", padx=12, pady=(0, 8))
        r += 1

        # Username
        self._lbl(left, "ชื่อผู้ใช้:").grid(row=r, column=0, sticky="w", padx=12, pady=(0, 2))
        r += 1
        self.user_entry = self._patch_entry(ctk.CTkEntry(
            left, font=ctk.CTkFont("TH Sarabun New", size=16, weight="bold"), height=32,
        ))
        self.user_entry.grid(row=r, column=0, sticky="ew", padx=12, pady=(0, 8))
        r += 1

        # Password
        self._lbl(left, "รหัสผ่าน:").grid(row=r, column=0, sticky="w", padx=12, pady=(0, 2))
        r += 1
        pass_row = ctk.CTkFrame(left, fg_color="transparent")
        pass_row.grid(row=r, column=0, sticky="ew", padx=12, pady=(0, 4))
        pass_row.grid_columnconfigure(0, weight=1)
        self.pass_entry = self._patch_entry(ctk.CTkEntry(
            pass_row, show="●", font=ctk.CTkFont("TH Sarabun New", size=16, weight="bold"), height=32,
        ))
        self.pass_entry.grid(row=0, column=0, sticky="ew", padx=(0, 6))
        self._show_pass_var = ctk.BooleanVar(value=False)
        ctk.CTkCheckBox(
            pass_row, text="แสดง", variable=self._show_pass_var,
            font=ctk.CTkFont("TH Sarabun New", size=14, weight="bold"),
            width=70,
            command=self._toggle_password,
        ).grid(row=0, column=1)
        r += 1

        # Remember password
        self.remember_var = ctk.BooleanVar(value=False)
        ctk.CTkCheckBox(
            left, text="จดจำรหัสผ่าน", variable=self.remember_var,
            font=ctk.CTkFont("TH Sarabun New", size=16, weight="bold"),
        ).grid(row=r, column=0, sticky="w", padx=12, pady=(0, 10))
        r += 1

        # Divider
        ctk.CTkFrame(left, height=1, fg_color="gray55").grid(
            row=r, column=0, sticky="ew", padx=12, pady=4,
        )
        r += 1

        # Save folder
        self._lbl(left, "บันทึกที่:").grid(row=r, column=0, sticky="w", padx=12, pady=(8, 2))
        r += 1
        folder_row = ctk.CTkFrame(left, fg_color="transparent")
        folder_row.grid(row=r, column=0, sticky="ew", padx=12, pady=(0, 8))
        folder_row.grid_columnconfigure(0, weight=1)
        self.save_entry = self._patch_entry(ctk.CTkEntry(
            folder_row, font=ctk.CTkFont("TH Sarabun New", size=15, weight="bold"), height=32,
        ))
        self.save_entry.grid(row=0, column=0, sticky="ew", padx=(0, 6))
        ctk.CTkButton(
            folder_row, text="📁", width=40, height=32,
            command=self._pick_folder,
        ).grid(row=0, column=1)

    # ── right: category checkboxes ────────────────────────────────────────────

    def _build_right(self, parent: ctk.CTkFrame) -> None:
        right = ctk.CTkFrame(parent)
        right.grid(row=0, column=1, sticky="nsew")

        ctk.CTkLabel(
            right, text="เลือกหมวดข้อมูล",
            font=ctk.CTkFont("TH Sarabun New", size=18, weight="bold"),
        ).pack(anchor="w", padx=14, pady=(14, 6))

        self.cat_vars: dict[int, ctk.BooleanVar] = {}
        for cat in CAT_ORDER:
            var = ctk.BooleanVar(value=True)
            self.cat_vars[cat] = var
            kw = {}
            if cat in CAT_RED:
                kw["text_color"] = ("#BB0000", "#FF6666")
            ctk.CTkCheckBox(
                right, text=CAT_LABELS[cat], variable=var,
                font=ctk.CTkFont("TH Sarabun New", size=17, weight="bold"),
                **kw,
            ).pack(anchor="w", padx=16, pady=4)

        sel_row = ctk.CTkFrame(right, fg_color="transparent")
        sel_row.pack(fill="x", padx=14, pady=(12, 0))
        ctk.CTkButton(
            sel_row, text="เลือกทั้งหมด", width=120,
            font=ctk.CTkFont("TH Sarabun New", size=15, weight="bold"),
            command=lambda: [v.set(True) for v in self.cat_vars.values()],
        ).pack(side="left", padx=(0, 6))
        ctk.CTkButton(
            sel_row, text="ยกเลิกทั้งหมด", width=120,
            font=ctk.CTkFont("TH Sarabun New", size=15, weight="bold"),
            fg_color="gray50", hover_color="gray35",
            command=lambda: [v.set(False) for v in self.cat_vars.values()],
        ).pack(side="left")

    # ── action bar: always-visible download / cancel ──────────────────────────

    def _build_action_bar(self) -> None:
        bar = ctk.CTkFrame(self, height=58, corner_radius=0, fg_color=("gray90", "gray17"))
        bar.grid(row=2, column=0, sticky="ew", padx=0, pady=0)
        bar.grid_columnconfigure(0, weight=1)
        bar.grid_columnconfigure(1, weight=0)
        bar.grid_propagate(False)

        self.dl_btn = ctk.CTkButton(
            bar,
            text="ดาวน์โหลด",
            font=ctk.CTkFont("TH Sarabun New", size=20, weight="bold"),
            height=40,
            command=self._start_download,
        )
        self.dl_btn.grid(row=0, column=0, sticky="ew", padx=(14, 8), pady=9)

        self.cancel_btn = ctk.CTkButton(
            bar,
            text="ยกเลิก",
            font=ctk.CTkFont("TH Sarabun New", size=20, weight="bold"),
            fg_color="#FF5252", hover_color="#D32F2F",
            text_color="white",
            text_color_disabled="white",
            width=130, height=40,
            command=self._cancel,
            state="disabled",
        )
        self.cancel_btn.grid(row=0, column=1, sticky="e", padx=(0, 14), pady=9)

    # ── bottom: progress + log ────────────────────────────────────────────────

    def _build_bottom(self) -> None:
        bot = ctk.CTkFrame(self, corner_radius=0)
        bot.grid(row=3, column=0, sticky="ew", padx=12, pady=(4, 8))
        bot.grid_columnconfigure(0, weight=1)
        bot.grid_columnconfigure(1, weight=0)

        # Progress bar
        self.progress_bar = ctk.CTkProgressBar(bot, height=16)
        self.progress_bar.grid(row=0, column=0, columnspan=2, sticky="ew", padx=10, pady=(10, 2))
        self.progress_bar.set(0)

        # Status label
        self.status_lbl = ctk.CTkLabel(
            bot, text="พร้อมใช้งาน",
            font=ctk.CTkFont("TH Sarabun New", size=16, weight="bold"), anchor="w",
        )
        self.status_lbl.grid(row=1, column=0, columnspan=2, sticky="ew", padx=10, pady=(0, 4))

        # Log textbox
        self.log_box = ctk.CTkTextbox(
            bot, height=110, font=ctk.CTkFont("TH Sarabun New", size=15, weight="bold"), wrap="word",
        )
        self.log_box.grid(row=2, column=0, columnspan=2, sticky="ew", padx=10, pady=(0, 8))
        self.log_box.configure(state="disabled")

        # Open-folder button (left)
        self.open_folder_btn = ctk.CTkButton(
            bot,
            text="📂  เปิดโฟลเดอร์ผลลัพธ์",
            font=ctk.CTkFont("TH Sarabun New", size=15, weight="bold"),
            command=self._open_save_folder,
            state="disabled",
        )
        self.open_folder_btn.grid(row=3, column=0, sticky="w", padx=10, pady=(0, 10))

        # ⓘ info button (right)
        ctk.CTkButton(
            bot,
            text="ⓘ", width=30, height=30,
            font=ctk.CTkFont(size=16, weight="bold"),
            fg_color="transparent",
            hover_color=("gray80", "gray30"),
            border_width=1,
            border_color=("gray65", "gray45"),
            text_color=("gray40", "gray65"),
            corner_radius=15,
            command=self._show_about,
        ).grid(row=3, column=1, sticky="e", padx=(4, 10), pady=(0, 10))

    # ── config ────────────────────────────────────────────────────────────────

    def _load_config(self) -> None:
        cfg = load_config()
        if cfg.get("url"):
            self.url_entry.insert(0, cfg["url"])
        if cfg.get("username"):
            self.user_entry.insert(0, cfg["username"])
        if cfg.get("enc_password"):
            self.pass_entry.insert(0, decrypt_password(cfg["enc_password"]))
            self.remember_var.set(True)
        if cfg.get("save_dir"):
            self.save_entry.delete(0, "end")
            self.save_entry.insert(0, cfg["save_dir"])
        if cfg.get("categories") is not None:
            saved = set(cfg["categories"])
            for cat, var in self.cat_vars.items():
                var.set(cat in saved)
        if not self.save_entry.get():
            self.save_entry.insert(
                0, _default_save_dir(self.url_entry.get(), _current_thai_year())
            )

    def _save_config_now(self) -> None:
        cfg: dict = {
            "url": self.url_entry.get().strip(),
            "username": self.user_entry.get().strip(),
            "save_dir": self.save_entry.get().strip(),
            "categories": [c for c, v in self.cat_vars.items() if v.get()],
        }
        if self.remember_var.get():
            pw = self.pass_entry.get()
            if pw:
                cfg["enc_password"] = encrypt_password(pw)
        save_config(cfg)

    # ── actions ───────────────────────────────────────────────────────────────

    def _pick_folder(self) -> None:
        folder = filedialog.askdirectory(title="เลือกโฟลเดอร์บันทึก")
        if folder:
            self.save_entry.delete(0, "end")
            self.save_entry.insert(0, folder)

    def _open_save_folder(self) -> None:
        path = self.save_entry.get().strip()
        if path and os.path.isdir(path):
            os.startfile(path) if sys.platform == "win32" else subprocess.Popen(["xdg-open", path])

    def _validate_inputs(self) -> bool:
        url  = self.url_entry.get().strip()
        user = self.user_entry.get().strip()
        pw   = self.pass_entry.get()
        save = self.save_entry.get().strip()
        cats = [c for c, v in self.cat_vars.items() if v.get()]

        if not url:
            messagebox.showerror("ข้อผิดพลาด", "กรุณากรอก URL โรงเรียน"); return False
        if not user:
            messagebox.showerror("ข้อผิดพลาด", "กรุณากรอกชื่อผู้ใช้"); return False
        if not pw:
            messagebox.showerror("ข้อผิดพลาด", "กรุณากรอกรหัสผ่าน"); return False
        if not save:
            messagebox.showerror("ข้อผิดพลาด", "กรุณาเลือกโฟลเดอร์บันทึก"); return False
        if not cats:
            messagebox.showerror("ข้อผิดพลาด", "กรุณาเลือกอย่างน้อย 1 หมวดข้อมูล"); return False

        return True

    def _start_download(self) -> None:
        if self._running:
            return
        self.status_lbl.configure(text="กำลังตรวจสอบข้อมูล...")
        self.update_idletasks()
        if not self._validate_inputs():
            self.status_lbl.configure(text="พร้อมใช้งาน")
            return
        self._save_config_now()
        self._cancel_event.clear()
        self._set_running(True)
        self._log_clear()
        self.progress_bar.set(0)
        self.status_lbl.configure(text="กำลังเข้าสู่ระบบ...")
        self.open_folder_btn.configure(state="disabled")

        url      = self.url_entry.get().strip()
        user     = self.user_entry.get().strip()
        pw       = self.pass_entry.get()
        save_dir = self.save_entry.get().strip()
        year     = _current_thai_year()
        cats     = {c for c, v in self.cat_vars.items() if v.get()}
        log_path = str(Path(save_dir) / "_log.txt")

        def _thread():
            try:
                client = StumisClient(url)
                client.login(user, pw)
                self._msg_queue.put(("log", "เข้าสู่ระบบสำเร็จ"))
                backup_worker.run_backup(
                    client=client, save_dir=save_dir, categories=cats,
                    year=year, msg_queue=self._msg_queue,
                    cancel_event=self._cancel_event, log_path=log_path,
                )
            except (LoginError, NetworkError) as e:
                self._msg_queue.put(("done", False, str(e)))
            except Exception as e:
                import traceback
                self._msg_queue.put(("done", False, f"ข้อผิดพลาด: {e}\n{traceback.format_exc()}"))

        threading.Thread(target=_thread, daemon=True).start()

    def _cancel(self) -> None:
        if self._running:
            self._cancel_event.set()
            self.status_lbl.configure(text="กำลังยกเลิก...")

    # ── queue polling ─────────────────────────────────────────────────────────

    def _poll_queue(self) -> None:
        try:
            while True:
                self._handle_msg(self._msg_queue.get_nowait())
        except queue.Empty:
            pass
        self.after(100, self._poll_queue)

    def _handle_msg(self, msg: tuple) -> None:
        kind = msg[0]
        if kind == "progress":
            _, cur, tot, status = msg
            pct = cur / tot if tot else 0
            self.progress_bar.set(pct)
            self.status_lbl.configure(text=f"{status}  ({cur} / {tot}  {pct*100:.0f}%)")
        elif kind == "log":
            self._log_append(msg[1])
        elif kind == "done":
            _, ok, message = msg
            self._set_running(False)
            if ok:
                self.progress_bar.set(1.0)
                self.status_lbl.configure(text=f"✓ {message}")
                self.open_folder_btn.configure(state="normal")
                messagebox.showinfo("เสร็จสิ้น", message)
            else:
                self.status_lbl.configure(text=f"✗ {message}")
                messagebox.showerror("ข้อผิดพลาด", message)

    # ── about popup ───────────────────────────────────────────────────────────

    def _show_about(self) -> None:
        win = ctk.CTkToplevel(self)
        win.title("เกี่ยวกับโปรแกรม")
        win.geometry("440x270")
        win.resizable(False, False)
        win.grab_set()
        win.focus_force()

        self.update_idletasks()
        px = self.winfo_x() + (self.winfo_width()  - 440) // 2
        py = self.winfo_y() + (self.winfo_height() - 270) // 2
        win.geometry(f"440x270+{px}+{py}")

        ctk.CTkLabel(
            win, text="ℹ  เกี่ยวกับโปรแกรม",
            font=ctk.CTkFont("TH Sarabun New", size=17, weight="bold"),
        ).pack(pady=(18, 4))

        ctk.CTkFrame(win, height=1, fg_color="gray55").pack(fill="x", padx=20, pady=(0, 6))

        for label, value in [
            ("พัฒนาโดย", "นายเอกชัย ทองพุ่ม"),
            ("ตำแหน่ง",  "ครูศิลปะโรงเรียนวางหลวงพิทยาสรรพ์"),
            ("โครงการ",  "ระบบสำรองข้อมูล STU-MIS  (Shub Chiwit Project 2026)"),
            ("ติดต่อ",   "ekkachai.th11222@gmail.com"),
        ]:
            row = ctk.CTkFrame(win, fg_color="transparent")
            row.pack(fill="x", padx=24, pady=2)
            ctk.CTkLabel(
                row, text=f"{label}:", width=76,
                font=ctk.CTkFont("TH Sarabun New", size=14, weight="bold"), anchor="w",
            ).pack(side="left")
            ctk.CTkLabel(
                row, text=value,
                font=ctk.CTkFont("TH Sarabun New", size=14, weight="bold"), anchor="w", justify="left",
            ).pack(side="left", padx=(4, 0))

        ctk.CTkButton(
            win, text="ปิด", width=90,
            font=ctk.CTkFont("TH Sarabun New", size=14, weight="bold"),
            command=win.destroy,
        ).pack(pady=(12, 16))

    # ── helpers ───────────────────────────────────────────────────────────────

    def _set_running(self, running: bool) -> None:
        self._running = running
        sm = "disabled" if running else "normal"
        sc = "normal"   if running else "disabled"
        self.dl_btn.configure(state=sm)
        self.cancel_btn.configure(state=sc)
        for w in [self.url_entry, self.user_entry, self.pass_entry, self.save_entry]:
            w.configure(state=sm)

    def _log_append(self, text: str) -> None:
        self.log_box.configure(state="normal")
        self.log_box.insert("end", f"[{datetime.now().strftime('%H:%M:%S')}] {text}\n")
        self.log_box.see("end")
        self.log_box.configure(state="disabled")

    def _log_clear(self) -> None:
        self.log_box.configure(state="normal")
        self.log_box.delete("1.0", "end")
        self.log_box.configure(state="disabled")
