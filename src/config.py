"""Load and save user config. Password is encrypted with Windows DPAPI (no extra deps)."""
import base64
import ctypes
import ctypes.wintypes
import json
import os
from pathlib import Path

CONFIG_DIR = Path(os.environ.get("APPDATA", Path.home())) / "stumis-backup"
CONFIG_FILE = CONFIG_DIR / "config.json"

# ── Windows DPAPI via ctypes ──────────────────────────────────────────────────

class _DATA_BLOB(ctypes.Structure):
    _fields_ = [("cbData", ctypes.wintypes.DWORD),
                ("pbData", ctypes.POINTER(ctypes.c_char))]


def encrypt_password(password: str) -> str:
    """Encrypt with DPAPI (Windows user account key). Falls back to base64."""
    try:
        data = password.encode("utf-8")
        buf_in = _DATA_BLOB(len(data), ctypes.cast(ctypes.c_char_p(data), ctypes.POINTER(ctypes.c_char)))
        buf_out = _DATA_BLOB()
        ok = ctypes.windll.Crypt32.CryptProtectData(
            ctypes.byref(buf_in), None, None, None, None, 0, ctypes.byref(buf_out)
        )
        if ok:
            encrypted = bytes(buf_out.pbData[:buf_out.cbData])
            ctypes.windll.Kernel32.LocalFree(buf_out.pbData)
            return "dpapi:" + base64.b64encode(encrypted).decode()
    except Exception:
        pass
    # Fallback: simple base64 obfuscation
    return "b64:" + base64.b64encode(password.encode("utf-8")).decode()


def decrypt_password(token: str) -> str:
    """Decrypt DPAPI or base64 token. Returns empty string on failure."""
    try:
        if token.startswith("dpapi:"):
            data = base64.b64decode(token[6:])
            buf_in = _DATA_BLOB(len(data), ctypes.cast(ctypes.c_char_p(data), ctypes.POINTER(ctypes.c_char)))
            buf_out = _DATA_BLOB()
            ok = ctypes.windll.Crypt32.CryptUnprotectData(
                ctypes.byref(buf_in), None, None, None, None, 0, ctypes.byref(buf_out)
            )
            if ok:
                plaintext = bytes(buf_out.pbData[:buf_out.cbData])
                ctypes.windll.Kernel32.LocalFree(buf_out.pbData)
                return plaintext.decode("utf-8")
        elif token.startswith("b64:"):
            return base64.b64decode(token[4:]).decode("utf-8")
        # Legacy: old Fernet token — ignore
    except Exception:
        pass
    return ""


def load_config() -> dict:
    if CONFIG_FILE.exists():
        try:
            with open(CONFIG_FILE, encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {}


def save_config(cfg: dict) -> None:
    CONFIG_DIR.mkdir(parents=True, exist_ok=True)
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)
