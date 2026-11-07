import json, re, subprocess, threading, queue, codecs, unicodedata, traceback
import sys, os, ctypes, webbrowser, tempfile, time
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import tkinter.font as tkfont
from io import BytesIO
from PIL import Image, ImageDraw, ImageFont, ImageColor
import urllib.request

try:  # Windows-only modules (guarded so the file still imports elsewhere)
    import winsound
    import winreg
except ImportError:  # pragma: no cover
    winsound = None
    winreg = None

IS_WIN = os.name == "nt"

APP_VERSION_ONLY = "v2.4"
APP_TITLE = "Windows App Updater"
APP_NAME_VERSION = f"{APP_TITLE} {APP_VERSION_ONLY}"
DATE_APP = "2026/11/07"
GITHUB_REPO_PAGE = "https://github.com/ilukezippo/Windows-App-Updater"
GITHUB_RELEASES_PAGE = GITHUB_REPO_PAGE + "/releases"
GITHUB_API_LATEST = "https://api.github.com/repos/ilukezippo/Windows-App-Updater/releases/latest"
DONATE_PAGE = "https://buymeacoffee.com/ilukezippo"
CONTACT_EMAIL = "ilukezippo@gmail.com"

CREATE_NO_WINDOW = 0x08000000

# If winget produces no output AND the progress % doesn't move for this many
# seconds, treat the package as stalled and give up on it (so the batch can
# continue instead of hanging at 0%). Activity resets the timer, so legitimately
# slow downloads/installs are unaffected — only a truly hung process trips it.
STALL_TIMEOUT_SECONDS = 180

LOG_MAX_LINES = 5000

SETTINGS_DIR = os.path.join(os.environ.get("APPDATA") or os.path.expanduser("~"), "WindowsAppUpdater")
SETTINGS_FILE = os.path.join(SETTINGS_DIR, "settings.json")


# =====================================================================
#  Small helpers
# =====================================================================
def _hidden_startupinfo():
    if not IS_WIN:
        return None
    si = subprocess.STARTUPINFO()
    si.dwFlags |= subprocess.STARTF_USESHOWWINDOW
    si.wShowWindow = 0
    return si


def _no_window_flags():
    return CREATE_NO_WINDOW if IS_WIN else 0


def resource_path(p):
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, p)


def is_admin():
    try:
        return bool(ctypes.windll.shell32.IsUserAnAdmin())
    except Exception:
        return False


def relaunch_as_admin():
    """Relaunch elevated. Returns True if the elevated copy started."""
    if is_admin() or not IS_WIN:
        return False
    app = sys.executable
    if getattr(sys, "frozen", False):
        params = " ".join(f'"{a}"' for a in sys.argv[1:])
    else:
        # Prefer pythonw.exe so the elevated copy doesn't open a console window.
        pyw = os.path.join(os.path.dirname(app), "pythonw.exe")
        if os.path.exists(pyw):
            app = pyw
        script = os.path.abspath(sys.argv[0])
        params = " ".join([f'"{script}"'] + [f'"{a}"' for a in sys.argv[1:]])
    # ShellExecuteW returns an HINSTANCE > 32 on success; <= 32 means it failed
    # (e.g. the user declined the UAC prompt).
    rc = ctypes.windll.shell32.ShellExecuteW(None, "runas", app, params, None, 1)
    return bool(rc) and int(rc) > 32


def enable_dpi_awareness():
    """Crisp text on high-DPI screens (must run before Tk() is created)."""
    if not IS_WIN:
        return
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(2)  # per-monitor
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass


def system_prefers_dark():
    if not winreg:
        return True
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER,
                            r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize") as k:
            return winreg.QueryValueEx(k, "AppsUseLightTheme")[0] == 0
    except OSError:
        return False


def set_titlebar_theme(win, dark, caption_hex=None):
    """Dark/light native title bar (Win10 1809+) and caption colour (Win11)."""
    if not IS_WIN:
        return
    try:
        win.update_idletasks()
        hwnd = ctypes.windll.user32.GetParent(win.winfo_id())
        dwm = ctypes.windll.dwmapi.DwmSetWindowAttribute
        val = ctypes.c_int(1 if dark else 0)
        for attr in (20, 19):  # DWMWA_USE_IMMERSIVE_DARK_MODE (new, old)
            if dwm(hwnd, attr, ctypes.byref(val), ctypes.sizeof(val)) == 0:
                break
        if caption_hex:
            r, g, b = ImageColor.getrgb(caption_hex)
            color = ctypes.c_int(r | (g << 8) | (b << 16))
            dwm(hwnd, 35, ctypes.byref(color), ctypes.sizeof(color))  # DWMWA_CAPTION_COLOR
    except Exception:
        pass


def set_app_icon(win):
    ico = resource_path("windows-updater.ico")
    if os.path.exists(ico):
        try:
            win.iconbitmap(ico)
            return ico
        except Exception:
            pass
    return None


def load_flag_image():
    png = resource_path("kuwait.png")
    if os.path.exists(png):
        try:
            return tk.PhotoImage(file=png)
        except Exception:
            pass
    return None


def play_success_sound():
    if not winsound:
        return
    wav = resource_path("success.wav")
    try:
        if os.path.exists(wav):
            winsound.PlaySound(wav, winsound.SND_FILENAME | winsound.SND_ASYNC)
        else:
            winsound.MessageBeep(winsound.MB_ICONASTERISK)
    except Exception:
        pass


def load_settings():
    try:
        with open(SETTINGS_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
            return data if isinstance(data, dict) else {}
    except Exception:
        return {}


def save_settings(data):
    try:
        os.makedirs(SETTINGS_DIR, exist_ok=True)
        tmp = SETTINGS_FILE + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2)
        os.replace(tmp, SETTINGS_FILE)
    except Exception:
        pass


def run(cmd, timeout=None):
    p = subprocess.run(cmd, capture_output=True, text=True, shell=False, encoding="utf-8", errors="replace",
                       stdin=subprocess.DEVNULL, startupinfo=_hidden_startupinfo(),
                       creationflags=_no_window_flags(), timeout=timeout)
    return p.returncode, (p.stdout or "").strip(), (p.stderr or "").strip()


def kill_process_tree(proc):
    """Stop winget AND the installer it spawned (terminate() alone orphans the installer,
    which then keeps running and blocks the next MSI install with error 1618)."""
    if not proc or proc.poll() is not None:
        return
    try:
        if IS_WIN:
            run(["taskkill", "/PID", str(proc.pid), "/T", "/F"], timeout=20)
        else:
            proc.terminate()
    except Exception:
        try:
            proc.terminate()
        except Exception:
            pass


# =====================================================================
#  winget output parsing
# =====================================================================
def _char_width(ch):
    if unicodedata.combining(ch):
        return 0
    return 2 if unicodedata.east_asian_width(ch) in ("W", "F") else 1


def _slice_by_columns(line, starts):
    """Cut `line` at *display-cell* positions (winget pads by cells, not code points,
    so names with wide characters would otherwise shift every following column)."""
    pieces = [""] * len(starts)
    col = 0
    idx = 0
    for ch in line:
        while idx + 1 < len(starts) and col >= starts[idx + 1]:
            idx += 1
        if col >= starts[0]:
            pieces[idx] += ch
        col += _char_width(ch)
    return pieces


_HEADER_WORDS = ("Name", "Id", "Version", "Available")
_INFO_LINE_RE = re.compile(
    r"cannot be determined|explicit targeting|have pins|--include-unknown|--include-pinned|winget pin",
    re.I)
_SUMMARY_RE = re.compile(r"^\d+\s+(upgrades?|packages?|package\(s\))\b", re.I)


def _header_columns(line):
    pos = {}
    for word in _HEADER_WORDS + ("Source",):
        m = re.search(rf"(?<!\S){word}(?!\S)", line)
        if m:
            pos[word] = sum(_char_width(c) for c in line[:m.start()])
    if all(w in pos for w in _HEADER_WORDS) and pos["Name"] < pos["Id"] < pos["Version"] < pos["Available"]:
        return pos
    return None


def parse_table_upgrade_output(text: str):
    """
    Parser for `winget upgrade` table output.
    - Uses header column positions (in display cells) to slice fields.
    - Understands multiple tables (e.g. the "require explicit targeting" section
      winget prints after the main list) instead of turning their headers,
      separators and notes into bogus rows.
    """
    items = []
    cols = None
    seen = set()
    for raw in (text or "").splitlines():
        ln = raw.rstrip()
        s = ln.strip()
        if not s:
            continue
        hdr = _header_columns(ln)
        if hdr:
            cols = hdr
            continue
        if cols is None:
            continue
        if set(s) <= set("-─—"):
            continue
        if _SUMMARY_RE.match(s) or _INFO_LINE_RE.search(s) or s.endswith(":"):
            # Notes/summaries end the current table; a new header may follow.
            cols = None
            continue

        starts = [cols["Name"], cols["Id"], cols["Version"], cols["Available"]]
        if "Source" in cols:
            starts.append(cols["Source"])
        parts = [p.strip() for p in _slice_by_columns(ln, starts)]
        name, pid, curr, avail = parts[0], parts[1], parts[2], parts[3]

        if not pid or not re.fullmatch(r"\S+", pid):
            continue
        # Drop cosmetic notes like "(August 2025)"
        curr = re.sub(r"\s*\(.*?\)", "", curr).strip()
        avail = re.sub(r"\s*\(.*?\)", "", avail).strip()
        if pid in seen:
            continue
        seen.add(pid)
        items.append({"name": name or pid, "id": pid, "current": curr, "available": avail})
    return items


_COSMETIC_SUFFIX_RE = re.compile(
    r"[-_ .]?(release|stable|final|ga|rtm|x64|x86|amd64|win64|win32|64-?bit|32-?bit)$")


def _normalize_version(v):
    """
    Normalize a winget version string for *equivalence* comparison.

    winget compares versions literally, so it treats '2.4.0-release' as different
    from (and lower than) '2.4.0' and offers a pointless upgrade. We strip only
    cosmetic suffixes. Real pre-release tags (beta, rc, preview…) are kept, so a
    genuine '2.4.0-beta' -> '2.4.0' upgrade is still offered.
    """
    if not v:
        return ""
    s = str(v).strip().lower()
    if s.startswith("v") and len(s) > 1 and s[1].isdigit():
        s = s[1:]
    s = s.split("+", 1)[0]  # build metadata never matters
    prev = None
    while prev != s:
        prev = s
        s = _COSMETIC_SUFFIX_RE.sub("", s).strip()
    parts = s.split(".")
    while len(parts) > 1 and parts[-1] in ("", "0"):
        parts.pop()
    return ".".join(parts)


def _is_noop_upgrade(current, available):
    c = _normalize_version(current)
    a = _normalize_version(available)
    return bool(c) and bool(a) and c == a


def _filter_noop_upgrades(items):
    return [it for it in items if not _is_noop_upgrade(it.get("current", ""), it.get("available", ""))]


def get_winget_upgrades(include_unknown):
    try:
        c, _, _ = run(["winget", "--version"], timeout=30)
    except (FileNotFoundError, OSError):
        c = 1
    if c != 0:
        raise RuntimeError("winget was not found.\nInstall “App Installer” from the Microsoft Store.")

    base = ["winget", "upgrade", "--accept-source-agreements", "--disable-interactivity"]
    if include_unknown:
        base.append("--include-unknown")
    last_err = "winget upgrade failed"
    # First the winget community source, then all sources (adds msstore apps).
    for cmd in (base + ["--source", "winget"], base):
        try:
            code, out, err = run(cmd, timeout=300)
        except subprocess.TimeoutExpired:
            last_err = "winget took too long to respond."
            continue
        parsed = parse_table_upgrade_output(out)
        if parsed:
            return _filter_noop_upgrades(parsed)
        if out and _header_columns_in(out) is not None:
            return []  # a valid (empty) table
        if re.search(r"no (installed package|available upgrade|applicable upgrade)", out, re.I):
            return []
        last_err = (err or out or last_err).strip()
    raise RuntimeError(last_err)


def _header_columns_in(text):
    for ln in (text or "").splitlines():
        h = _header_columns(ln)
        if h:
            return h
    return None


# ===== winget exit codes (https://aka.ms/winget-returncodes) =====
WINGET_OK = 0
WINGET_NO_APPLICABLE_UPDATE = 0x8A15002B
WINGET_NO_APPLICATIONS_FOUND = 0x8A150014
WINGET_NO_APPLICABLE_INSTALLER = 0x8A150010
WINGET_UPDATE_NOT_APPLICABLE_SET = {
    WINGET_NO_APPLICABLE_UPDATE, WINGET_NO_APPLICATIONS_FOUND, WINGET_NO_APPLICABLE_INSTALLER,
}
WINGET_INSTALL_FILE_IN_USE = 0x8A150101
WINGET_INSTALL_PACKAGE_IN_USE = 0x8A150111
WINGET_IN_USE_SET = {WINGET_INSTALL_FILE_IN_USE, WINGET_INSTALL_PACKAGE_IN_USE}
WINGET_CTRL_SIGNAL = 0x8A150005
WINGET_REBOOT_REQUIRED = {0x8A150109, 0x8A15010A}  # install OK, restart needed


def _normalize_winget_code(rc):
    if rc is None:
        return None
    try:
        rc = int(rc)
    except (TypeError, ValueError):
        return None
    if rc < 0:
        rc &= 0xFFFFFFFF
    return rc


STATUS = {
    "success": ("✓  Updated", "ok"),
    "reboot": ("✓  Restart needed", "ok"),
    "noupdate": ("–  No update", "skip"),
    "skipped": ("↷  Skipped", "skip"),
    "inuse": ("⚠  App in use", "fail"),
    "failed": ("✕  Failed", "fail"),
    "stalled": ("⚠  Stalled", "fail"),
    "canceled": ("■  Canceled", "cancel"),
    "updating": ("⟳  Updating…", "active"),
    "queued": ("•  Queued", ""),
}


def classify_winget_result(returncode, skip_now=False, canceled=False):
    """Outcome from winget's exit code. Returns a STATUS key."""
    if canceled:
        return "canceled"
    if skip_now:
        return "skipped"
    code = _normalize_winget_code(returncode)
    if code == WINGET_OK:
        return "success"
    if code in WINGET_REBOOT_REQUIRED:
        return "reboot"
    if code in WINGET_UPDATE_NOT_APPLICABLE_SET:
        return "noupdate"
    if code in WINGET_IN_USE_SET:
        return "inuse"
    if code == WINGET_CTRL_SIGNAL:
        return "canceled"
    return "failed"


RESULT_BUCKET = {  # STATUS key -> summary bucket
    "success": "success", "reboot": "success", "noupdate": "skipped", "skipped": "skipped",
    "inuse": "failed", "failed": "failed", "stalled": "stalled", "canceled": "canceled",
}


# =====================================================================
#  Finding the processes that block an update
# =====================================================================
def _norm(p):
    return os.path.normcase(os.path.normpath(p)) if p else ""


def _is_safe_app_folder(folder):
    """Never treat system-wide folders as an 'app folder' (we kill everything under it)."""
    if not folder or not os.path.isdir(folder):
        return False
    f = _norm(folder)
    _drive, rest = os.path.splitdrive(f)
    if not [part for part in rest.split(os.sep) if part]:
        return False  # a drive root is never an app folder
    env = os.environ
    blocked = set()
    for k in ("SystemRoot", "WINDIR", "ProgramFiles", "ProgramFiles(x86)", "ProgramW6432", "ProgramData",
              "CommonProgramFiles", "CommonProgramFiles(x86)", "CommonProgramW6432", "USERPROFILE",
              "APPDATA", "LOCALAPPDATA", "PUBLIC", "TEMP", "TMP", "ALLUSERSPROFILE"):
        if env.get(k):
            blocked.add(_norm(env[k]))
    if env.get("LOCALAPPDATA"):
        for sub in ("Programs", "Microsoft", r"Microsoft\WindowsApps", "Temp"):
            blocked.add(_norm(os.path.join(env["LOCALAPPDATA"], sub)))
    if env.get("SystemDrive"):
        blocked.add(_norm(env["SystemDrive"] + os.sep + "Users"))
    if f in blocked:
        return False
    windir = _norm(env.get("SystemRoot") or env.get("WINDIR") or r"C:\Windows")
    if f == windir or f.startswith(windir + os.sep):
        return False
    return True


def _exe_from_command(cmd):
    """First executable path from a registry command like '"C:\\x\\app.exe",0' or 'C:\\x\\u.exe /S'."""
    if not cmd:
        return ""
    cmd = cmd.strip()
    if cmd.startswith('"'):
        end = cmd.find('"', 1)
        path = cmd[1:end] if end > 0 else cmd[1:]
    else:
        m = re.match(r"(.+?\.exe)\b", cmd, re.I)
        path = m.group(1) if m else cmd.split(",")[0]
    path = path.split(",")[0].strip()
    return path


_GENERIC_EXES = {"msiexec.exe", "rundll32.exe", "cmd.exe", "powershell.exe", "explorer.exe", "regsvr32.exe"}


def _arp_entries():
    """Installed-programs entries (Add/Remove Programs) from the registry."""
    if not winreg:
        return []
    out = []
    path = r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
    views = [(winreg.HKEY_LOCAL_MACHINE, winreg.KEY_WOW64_64KEY),
             (winreg.HKEY_LOCAL_MACHINE, winreg.KEY_WOW64_32KEY),
             (winreg.HKEY_CURRENT_USER, 0)]
    for hive, view in views:
        try:
            k = winreg.OpenKey(hive, path, 0, winreg.KEY_READ | view)
        except OSError:
            continue
        with k:
            i = 0
            while True:
                try:
                    sub = winreg.EnumKey(k, i)
                except OSError:
                    break
                i += 1
                try:
                    with winreg.OpenKey(k, sub) as sk:
                        def q(n):
                            try:
                                return str(winreg.QueryValueEx(sk, n)[0] or "")
                            except OSError:
                                return ""

                        out.append({"key": sub, "name": q("DisplayName"), "loc": q("InstallLocation"),
                                    "icon": q("DisplayIcon"), "uninst": q("UninstallString")})
                except OSError:
                    continue
    return out


def find_install_folder(package_id, display_name):
    """
    Locate the folder a package is installed in.
    (`winget show` describes the *manifest*, not the installed copy, so the old
    approach never found anything and auto-close silently did nothing.)
    """
    # 1) Newer winget: `winget list --details` prints "Installed Location".
    try:
        code, out, _ = run(["winget", "list", "--id", package_id, "--exact", "--details",
                            "--accept-source-agreements", "--disable-interactivity"], timeout=60)
        if code == 0:
            for line in out.splitlines():
                m = re.match(r"\s*install(?:ed)?\s+location\s*:\s*(.+)$", line, re.I)
                if m:
                    loc = m.group(1).strip().strip('"')
                    if _is_safe_app_folder(loc):
                        return os.path.normpath(loc)
    except Exception:
        pass

    # 2) Registry (Add/Remove Programs), matched by key or display name.
    want = (display_name or "").strip()
    prefix = want[:-1].strip().lower() if want.endswith("…") else None
    pid_l = (package_id or "").lower()
    for e in _arp_entries():
        nm = e["name"].strip()
        hit = (e["key"].lower() == pid_l) or (nm and (nm.lower() == want.lower() or
                                                     (prefix and nm.lower().startswith(prefix))))
        if not hit:
            continue
        candidates = [e["loc"].strip().strip('"')]
        for cmd in (e["icon"], e["uninst"]):
            exe = _exe_from_command(cmd)
            if exe and os.path.basename(exe).lower() not in _GENERIC_EXES and os.path.isfile(exe):
                candidates.append(os.path.dirname(exe))
        for c in candidates:
            if c and _is_safe_app_folder(c):
                return os.path.normpath(c)
    return None


def running_processes_under(folder):
    """List (pid, exe_path) of running processes whose image lives inside `folder`."""
    if not _is_safe_app_folder(folder):
        return []
    folder_norm = _norm(folder) + os.sep  # trailing sep: "C:\App" must not match "C:\AppOther"
    own = {os.getpid(), os.getppid()}
    ps_cmd = ("Get-CimInstance Win32_Process | Where-Object { $_.ExecutablePath } | "
              "ForEach-Object { \"$($_.ProcessId)`t$($_.ExecutablePath)\" }")
    procs = []
    try:
        code, out, _ = run(["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", ps_cmd],
                           timeout=60)
    except Exception:
        return procs
    if code != 0:
        return procs
    for line in out.splitlines():
        if "\t" not in line:
            continue
        pid_s, path = line.split("\t", 1)
        try:
            pid = int(pid_s.strip())
        except ValueError:
            continue
        if pid in own:
            continue
        if _norm(path.strip()).startswith(folder_norm):
            procs.append((pid, path.strip()))
    return procs


def kill_pids(pids):
    killed = 0
    for pid in pids:
        try:
            code, _, _ = run(["taskkill", "/PID", str(pid), "/T", "/F"], timeout=20)
            if code == 0:
                killed += 1
        except Exception:
            pass
    return killed


# =====================================================================
#  Self-update download
# =====================================================================
def _download_file(url: str, dest_path: str, progress_cb=None):
    """Stream to a .part file, verify size + MZ header, then atomically move into place."""
    tmp = dest_path + ".part"
    try:
        if os.path.exists(tmp):
            os.remove(tmp)
    except Exception:
        pass

    req = urllib.request.Request(url, headers={"User-Agent": "Windows-App-Updater"})
    done = total = 0
    with urllib.request.urlopen(req, timeout=60) as r:
        total = int(r.headers.get("Content-Length") or 0)
        with open(tmp, "wb") as f:
            while True:
                chunk = r.read(1024 * 128)
                if not chunk:
                    break
                f.write(chunk)
                done += len(chunk)
                if progress_cb and total > 0:
                    try:
                        progress_cb(int(done * 100 / total))
                    except Exception:
                        pass

    try:
        if total > 0 and done != total:
            raise RuntimeError(f"Download incomplete ({done}/{total} bytes). Please retry.")
        with open(tmp, "rb") as f:
            if f.read(2) != b"MZ":
                raise RuntimeError("Downloaded file is not a valid Windows EXE.")
    except Exception:
        try:
            os.remove(tmp)
        except Exception:
            pass
        raise
    os.replace(tmp, dest_path)
    return dest_path


def _ps_quote(s):
    return "'" + str(s).replace("'", "''") + "'"


# =====================================================================
#  Theme
# =====================================================================
PALETTES = {
    "dark": {
        "bg": "#0f1116", "surface": "#171a21", "surface2": "#20242d", "surface3": "#2a2f3a",
        "border": "#2b303b", "text": "#e7e9ee", "muted": "#8f98a8", "faint": "#5b6475",
        "accent": "#4c8dff", "accent_hover": "#689fff", "accent_press": "#3b78e0", "on_accent": "#ffffff",
        "success": "#3fb950", "warn": "#d6a23a", "danger": "#f25f5c",
        "row_alt": "#191c24", "row_hover": "#1f2430", "row_checked": "#1c2a44",
        "ok_bg": "#15281c", "fail_bg": "#321b1d", "skip_bg": "#2c2617", "cancel_bg": "#2d2218",
        "active_bg": "#1b2d4d", "log_bg": "#0c0e12", "select": "#284777",
    },
    "light": {
        "bg": "#f3f5f9", "surface": "#ffffff", "surface2": "#eef1f6", "surface3": "#e2e6ee",
        "border": "#dde2ea", "text": "#1b1f27", "muted": "#5f6878", "faint": "#9aa2b1",
        "accent": "#2f6fed", "accent_hover": "#4a82f0", "accent_press": "#255fd0", "on_accent": "#ffffff",
        "success": "#1f8f3a", "warn": "#b07a10", "danger": "#d33a36",
        "row_alt": "#f8f9fc", "row_hover": "#f0f4fb", "row_checked": "#e6efff",
        "ok_bg": "#e7f6ea", "fail_bg": "#fdecec", "skip_bg": "#fdf6e3", "cancel_bg": "#fdf0e4",
        "active_bg": "#e3edff", "log_bg": "#fbfcfe", "select": "#cfe0ff",
    },
}

# Segoe Fluent Icons / Segoe MDL2 Assets code points
ICONS = {
    "refresh": "\uE72C", "download": "\uE896", "skip": "\uE893", "cancel": "\uE711",
    "folder": "\uE838", "delete": "\uE74D", "info": "\uE946", "shield": "\uEA18",
    "sun": "\uE706", "moon": "\uE708", "search": "\uE721", "save": "\uE74E",
    "chevron_down": "\uE70D", "chevron_up": "\uE70E", "heart": "\uEB51", "update": "\uE777",
}


def _rgba(hex_color, alpha=255):
    r, g, b = ImageColor.getrgb(hex_color)[:3]
    return (r, g, b, alpha)


def _to_photo(im):
    bio = BytesIO()
    im.save(bio, format="PNG")
    return tk.PhotoImage(data=bio.getvalue())


def _rounded_image(w, h, radius, fill, outline=None, width=1, ss=4):
    im = Image.new("RGBA", (w * ss, h * ss), (0, 0, 0, 0))
    d = ImageDraw.Draw(im)
    d.rounded_rectangle([0, 0, w * ss - 1, h * ss - 1], radius=radius * ss, fill=fill,
                        outline=outline, width=(max(1, int(width * ss)) if outline else 0))
    return im.resize((w, h), Image.LANCZOS)


class Theme:
    """A flat, image-based ttk theme (rounded buttons, switches, checkboxes).
    Images are tiny 9-slice bitmaps, so drawing stays as cheap as plain ttk."""

    def __init__(self, root, scale):
        self.root = root
        self.s = scale
        self.style = ttk.Style(root)
        self.style.theme_use("clam")
        self.images = {}  # keep references alive
        self.icon_font_path = self._find_icon_font()
        self._icon_cache = {}
        fams = set(tkfont.families(root))
        ui = "Segoe UI" if "Segoe UI" in fams else tkfont.nametofont("TkDefaultFont").actual("family")
        semi = "Segoe UI Semibold" if "Segoe UI Semibold" in fams else ui
        mono = next((f for f in ("Cascadia Mono", "Consolas", "DejaVu Sans Mono") if f in fams), "Courier")
        semi_w = "normal" if semi != ui else "bold"
        self.fonts = {
            "title": tkfont.Font(root, family=semi, size=17, weight=semi_w),
            "h2": tkfont.Font(root, family=semi, size=11, weight=semi_w),
            "body": tkfont.Font(root, family=ui, size=10),
            "semibold": tkfont.Font(root, family=semi, size=10, weight=semi_w),
            "small": tkfont.Font(root, family=ui, size=9),
            "small_semi": tkfont.Font(root, family=semi, size=9, weight=semi_w),
            "mono": tkfont.Font(root, family=mono, size=9),
            "link": tkfont.Font(root, family=ui, size=10, underline=True),
        }
        for name in ("TkDefaultFont", "TkTextFont", "TkMenuFont"):
            try:
                tkfont.nametofont(name).configure(family=ui, size=10)
            except Exception:
                pass
        self.name = None
        self.p = None

    def px(self, n):
        return max(1, int(round(n * self.s)))

    # ---------- icons ----------
    @staticmethod
    def _find_icon_font():
        fonts_dir = os.path.join(os.environ.get("WINDIR", r"C:\Windows"), "Fonts")
        for fn in ("SegoeIcons.ttf", "segmdl2.ttf"):
            p = os.path.join(fonts_dir, fn)
            if os.path.exists(p):
                return p
        return None

    def icon(self, name, color, size=16):
        if not self.icon_font_path or name not in ICONS:
            return None
        key = (name, color, size, self.s)
        if key in self._icon_cache:
            return self._icon_cache[key]
        px = self.px(size)
        try:
            ss = 4
            font = ImageFont.truetype(self.icon_font_path, int(px * 0.78 * ss))
            im = Image.new("RGBA", (px * ss, px * ss), (0, 0, 0, 0))
            ImageDraw.Draw(im).text((px * ss / 2, px * ss / 2), ICONS[name], font=font,
                                    fill=_rgba(color), anchor="mm")
            photo = _to_photo(im.resize((px, px), Image.LANCZOS))
        except Exception:
            photo = None
        self._icon_cache[key] = photo
        return photo

    # ---------- theme ----------
    def apply(self, name):
        self.name = name
        p = self.p = PALETTES[name]
        st = self.style
        px = self.px
        f = self.fonts

        st.configure(".", background=p["bg"], foreground=p["text"], font=f["body"],
                     bordercolor=p["border"], lightcolor=p["bg"], darkcolor=p["bg"],
                     troughcolor=p["surface2"], focuscolor=p["accent"], selectbackground=p["select"],
                     selectforeground=p["text"], fieldbackground=p["surface2"], insertcolor=p["text"])
        st.map(".", foreground=[("disabled", p["faint"])])

        # Frames & labels
        st.configure("TFrame", background=p["bg"])
        st.configure("Card.TFrame", background=p["surface"], relief="solid", borderwidth=1,
                     bordercolor=p["border"], lightcolor=p["border"], darkcolor=p["border"])
        st.configure("Inner.TFrame", background=p["surface"])
        st.configure("TLabel", background=p["bg"], foreground=p["text"])
        st.configure("Title.TLabel", font=f["title"])
        st.configure("Sub.TLabel", foreground=p["muted"])
        st.configure("Card.TLabel", background=p["surface"])
        st.configure("CardH.TLabel", background=p["surface"], font=f["h2"])
        st.configure("CardMuted.TLabel", background=p["surface"], foreground=p["muted"])
        st.configure("CardValue.TLabel", background=p["surface"], foreground=p["text"], font=f["small_semi"])
        st.configure("Empty.TLabel", background=p["surface"], foreground=p["muted"], font=f["body"],
                     justify="center")
        st.configure("Pill.TLabel", background=p["surface3"], foreground=p["muted"], font=f["small_semi"],
                     padding=(px(8), px(2)))
        st.configure("AdminPill.TLabel", background=p["ok_bg"], foreground=p["success"],
                     font=f["small_semi"], padding=(px(10), px(5)))
        st.configure("Link.TLabel", background=p["surface"], foreground=p["accent"], font=f["link"])

        # Buttons (rounded image elements, one set per palette)
        r = px(7)
        # ttk *tiles* the middle of a 9-slice image. A tiny image meant hundreds of
        # alpha-blended tiles per button on every redraw (= visible lag). A large
        # image covers any button in a single tile.
        bw, bh = px(420), px(64)
        variants = {
            "Accent": (p["accent"], p["accent_hover"], p["accent_press"], None, p["on_accent"]),
            "Secondary": (p["surface2"], p["surface3"], p["border"], p["border"], p["text"]),
            "Ghost": (None, p["surface2"], p["surface3"], None, p["text"]),
            "Danger": (p["danger"], _mix(p["danger"], "#ffffff", 0.15), _mix(p["danger"], "#000000", 0.15),
                       None, "#ffffff"),
        }
        for v, (normal, hover, press, outline, fg) in variants.items():
            elem = f"{name}.{v}.btn"
            if elem in self.images:  # element already exists: its images must stay alive
                imgs = self.images[elem]
            else:
                imgs = {}
                for state, col in (("normal", normal), ("active", hover), ("pressed", press)):
                    fill = _rgba(col) if col else (0, 0, 0, 0)
                    imgs[state] = _to_photo(_rounded_image(bw, bh, r, fill,
                                                           outline=_rgba(outline) if outline else None, ss=2))
                dis_fill = _rgba(p["surface2"]) if normal else (0, 0, 0, 0)
                imgs["disabled"] = _to_photo(_rounded_image(bw, bh, r, dis_fill,
                                                            outline=_rgba(outline) if outline else None, ss=2))
                self.images[elem] = imgs
            if elem not in st.element_names():
                st.element_create(elem, "image", imgs["normal"],
                                  ("disabled", imgs["disabled"]), ("pressed", imgs["pressed"]),
                                  ("active", imgs["active"]), border=r + 1, sticky="nsew",
                                  width=2 * r + px(4), height=2 * r + px(4))  # natural size stays small
            style_name = f"{v}.TButton"
            st.layout(style_name, [(elem, {"sticky": "nsew", "children": [
                ("Button.padding", {"sticky": "nsew", "children": [
                    ("Button.label", {"sticky": "nsew"})]})]})])
            dis_fg = p["faint"]
            # Same look on the window background and on card surfaces. The background
            # map overrides clam's own hover/press colours, which would otherwise show
            # through the transparent corners (and the whole Ghost button).
            for sn, bgc in ((style_name, p["bg"]), (f"Card.{style_name}", p["surface"])):
                if sn != style_name:
                    st.layout(sn, st.layout(style_name))
                st.configure(sn, foreground=fg, font=f["semibold"], padding=(px(14), px(7)),
                             anchor="center", background=bgc, width=0)
                st.map(sn, foreground=[("disabled", dis_fg), ("pressed", fg), ("active", fg)],
                       background=[("disabled", bgc), ("pressed", bgc), ("active", bgc)])
        st.layout("Small.Card.Ghost.TButton", st.layout("Ghost.TButton"))
        st.configure("Small.Card.Ghost.TButton", padding=(px(10), px(4)), font=f["small_semi"],
                     foreground=p["muted"], background=p["surface"], width=0)
        st.map("Small.Card.Ghost.TButton", foreground=[("disabled", p["faint"]), ("active", p["text"])],
               background=[("disabled", p["surface"]), ("pressed", p["surface"]), ("active", p["surface"])])

        # Toggle switch checkbutton
        sw_w, sw_h = px(36), px(20)
        elem = f"{name}.switch"
        sw = self.images.get(elem) or self._switch_images(sw_w, sw_h, p)
        self.images[elem] = sw
        if elem not in st.element_names():
            st.element_create(elem, "image", sw["off"],
                              ("selected", "disabled", sw["on_dis"]), ("disabled", sw["off_dis"]),
                              ("selected", "active", sw["on_hover"]), ("selected", sw["on"]),
                              ("active", sw["off_hover"]), sticky="w")
        st.layout("Switch.TCheckbutton", [("Checkbutton.padding", {"sticky": "nsew", "children": [
            (elem, {"side": "left", "sticky": ""}),
            ("Checkbutton.label", {"side": "left", "sticky": "nsew"})]})])
        st.configure("Switch.TCheckbutton", background=p["surface"], foreground=p["text"], font=f["body"],
                     padding=(px(2), px(4)), indicatormargin=0)
        st.map("Switch.TCheckbutton", background=[("active", p["surface"])],
               foreground=[("disabled", p["faint"])])
        st.configure("Plain.TCheckbutton", background=p["surface"], foreground=p["text"],
                     indicatorbackground=p["surface2"], indicatorforeground=p["accent"])
        st.map("Plain.TCheckbutton", background=[("active", p["surface"])],
               indicatorbackground=[("selected", p["surface2"])])

        # Entry (search box)
        st.configure("Search.TEntry", fieldbackground=p["surface2"], foreground=p["text"],
                     bordercolor=p["border"], lightcolor=p["surface2"], darkcolor=p["surface2"],
                     insertcolor=p["text"], padding=(px(8), px(5)))
        st.map("Search.TEntry", bordercolor=[("focus", p["accent"])], lightcolor=[("focus", p["accent"])])

        # Treeview
        st.layout("App.Treeview", [("Treeview.treearea", {"sticky": "nswe"})])
        st.configure("App.Treeview", background=p["surface"], fieldbackground=p["surface"],
                     foreground=p["text"], rowheight=px(34), borderwidth=0, relief="flat", font=f["body"])
        st.map("App.Treeview", background=[("selected", p["surface"])], foreground=[("selected", p["text"])])
        st.configure("App.Treeview.Heading", background=p["surface"], foreground=p["muted"],
                     font=f["small_semi"], relief="flat", borderwidth=0, padding=(px(8), px(8)),
                     bordercolor=p["surface"], lightcolor=p["surface"], darkcolor=p["surface"])
        st.map("App.Treeview.Heading", background=[("active", p["surface2"])],
               foreground=[("active", p["text"])])

        # Slim scrollbars (no arrows)
        for orient in ("Vertical", "Horizontal"):
            st.layout(f"Slim.{orient}.TScrollbar", [(f"{orient}.Scrollbar.trough", {"sticky": "nswe", "children": [
                (f"{orient}.Scrollbar.thumb", {"expand": "1", "sticky": "nswe"})]})])
            st.configure(f"Slim.{orient}.TScrollbar", troughcolor=p["surface"], background=p["surface3"],
                         bordercolor=p["surface"], lightcolor=p["surface3"], darkcolor=p["surface3"],
                         gripcount=0, arrowsize=px(8), width=px(8))
            st.map(f"Slim.{orient}.TScrollbar", background=[("active", p["faint"]), ("pressed", p["muted"])],
                   lightcolor=[("active", p["faint"])], darkcolor=[("active", p["faint"])])

        # Progress bars
        for nm, col in (("Accent", p["accent"]), ("Success", p["success"])):
            st.configure(f"{nm}.Horizontal.TProgressbar", troughcolor=p["surface2"], background=col,
                         bordercolor=p["surface2"], lightcolor=col, darkcolor=col, thickness=px(6))

        # Tree checkbox images
        if f"{name}.cb" not in self.images:
            self.images[f"{name}.cb"] = self._checkbox_images(px(18), p)
        self.images["cb"] = self.images[f"{name}.cb"]

    def _switch_images(self, w, h, p):
        out = {}
        ss = 4
        specs = {
            "off": (p["surface3"], p["muted"], False, 255),
            "off_hover": (p["border"], p["text"], False, 255),
            "on": (p["accent"], "#ffffff", True, 255),
            "on_hover": (p["accent_hover"], "#ffffff", True, 255),
            "off_dis": (p["surface3"], p["faint"], False, 120),
            "on_dis": (p["accent"], "#ffffff", True, 110),
        }
        pad = self.px(2)
        for key, (track, knob, on, alpha) in specs.items():
            im = Image.new("RGBA", ((w + pad * 2) * ss, h * ss), (0, 0, 0, 0))
            d = ImageDraw.Draw(im)
            d.rounded_rectangle([0, 0, w * ss - 1, h * ss - 1], radius=h * ss // 2, fill=_rgba(track, alpha))
            kr = int(h * ss * 0.32)
            cy = h * ss // 2
            cx = (w * ss - h * ss // 2) if on else h * ss // 2
            d.ellipse([cx - kr, cy - kr, cx + kr, cy + kr], fill=_rgba(knob, alpha))
            out[key] = _to_photo(im.resize((w + pad * 2, h), Image.LANCZOS))
        return out

    def _checkbox_images(self, size, p):
        ss = 4
        S = size * ss
        r = int(S * 0.24)
        un = Image.new("RGBA", (S, S), (0, 0, 0, 0))
        ImageDraw.Draw(un).rounded_rectangle([ss, ss, S - ss - 1, S - ss - 1], radius=r,
                                             fill=_rgba(p["surface"]), outline=_rgba(p["faint"]),
                                             width=int(1.6 * ss))
        ch = Image.new("RGBA", (S, S), (0, 0, 0, 0))
        d = ImageDraw.Draw(ch)
        d.rounded_rectangle([ss, ss, S - ss - 1, S - ss - 1], radius=r, fill=_rgba(p["accent"]))
        d.line([(S * 0.28, S * 0.52), (S * 0.44, S * 0.68), (S * 0.73, S * 0.36)],
               fill=(255, 255, 255, 255), width=int(S * 0.11), joint="curve")
        return (_to_photo(un.resize((size, size), Image.LANCZOS)),
                _to_photo(ch.resize((size, size), Image.LANCZOS)))


def _mix(c1, c2, t):
    a = ImageColor.getrgb(c1)
    b = ImageColor.getrgb(c2)
    return "#%02x%02x%02x" % tuple(int(a[i] + (b[i] - a[i]) * t) for i in range(3))


def make_donate_image(scale=1.0):
    w, h = int(170 * scale), int(40 * scale)
    ss = 4
    im = Image.new("RGBA", (w * ss, h * ss), (0, 0, 0, 0))
    grad = Image.new("RGBA", (w * ss, h * ss))
    gd = ImageDraw.Draw(grad)
    top, bot = (255, 196, 84), (242, 151, 24)
    for y in range(h * ss):
        t = y / (h * ss - 1)
        gd.line([(0, y), (w * ss, y)], fill=tuple(int(top[i] + (bot[i] - top[i]) * t) for i in range(3)) + (255,))
    mask = Image.new("L", (w * ss, h * ss), 0)
    ImageDraw.Draw(mask).rounded_rectangle([0, 0, w * ss - 1, h * ss - 1], radius=h * ss // 2, fill=255)
    im.paste(grad, (0, 0), mask)
    return _to_photo(im.resize((w, h), Image.LANCZOS))


class ToolTip:
    def __init__(self, app, w, text):
        self.app, self.w, self.text, self.tip, self._job = app, w, text, None, None
        w.bind("<Enter>", self._schedule, add="+")
        w.bind("<Leave>", self.hide, add="+")
        w.bind("<ButtonPress>", self.hide, add="+")

    def _schedule(self, _=None):
        self._job = self.w.after(450, self.show)

    def show(self):
        self._job = None
        if self.tip or not self.text:
            return
        p = self.app.theme.p
        x = self.w.winfo_rootx() + 8
        y = self.w.winfo_rooty() + self.w.winfo_height() + 6
        self.tip = tk.Toplevel(self.w)
        self.tip.wm_overrideredirect(True)
        self.tip.wm_geometry(f"+{x}+{y}")
        outer = tk.Frame(self.tip, bg=p["border"], padx=1, pady=1)
        outer.pack()
        tk.Label(outer, text=self.text, justify="left", bg=p["surface2"], fg=p["text"],
                 font=self.app.theme.fonts["small"], padx=10, pady=6).pack()

    def hide(self, _=None):
        if self._job:
            try:
                self.w.after_cancel(self._job)
            except Exception:
                pass
            self._job = None
        if self.tip:
            self.tip.destroy()
            self.tip = None


# =====================================================================
#  Main window
# =====================================================================
class WingetUpdaterUI:
    COLS = ("Name", "Id", "Current", "Available", "Result")

    def __init__(self, root):
        self.root = root
        root.withdraw()  # build hidden -> no white flash / layout jumps on startup
        self.settings = load_settings()
        try:
            scale = root.winfo_fpixels("1i") / 96.0
        except Exception:
            scale = 1.0
        self.theme = Theme(root, max(1.0, scale))
        self.px = self.theme.px
        theme_name = self.settings.get("theme") or ("dark" if system_prefers_dark() else "light")
        if theme_name not in PALETTES:
            theme_name = "dark"
        self.theme.apply(theme_name)

        root.title(APP_NAME_VERSION)
        self.window_icon_path = set_app_icon(root)
        sw, sh = root.winfo_screenwidth(), root.winfo_screenheight()
        self._win_w = min(self.px(1040), sw - self.px(40))
        self._win_h = min(self.px(760), sh - self.px(90))
        root.geometry(f"{self._win_w}x{self._win_h}")
        root.minsize(min(self.px(860), self._win_w), min(self.px(560), self._win_h))

        # State
        self.updating = False
        self.cancel_requested = False
        self.skip_requested = False
        self._can_skip = False
        self.current_proc = None
        self.loading_win = None
        self._remembered_close_choice = None
        self.checked_items = set()
        self.all_items = []  # master order (filter detaches rows from the tree)
        self.id_to_item = {}
        self.pkg_downloads = {}
        self._sort_state = (None, False)
        self._hover_item = None
        self._dialogs = []

        # Cross-thread plumbing: workers never touch Tk directly.
        self._uiq = queue.SimpleQueue()
        self._log_lock = threading.Lock()
        self._log_buf = []
        self._pct_pending = None

        self._themed_tk = []  # (widget, {option: palette_key}) for classic tk widgets
        self._build()
        self._retheme_tk()
        self._pump()

        root.protocol("WM_DELETE_WINDOW", self.on_close)
        root.update_idletasks()
        self.center_on_screen()
        set_titlebar_theme(root, self.theme.name == "dark", self.theme.p["bg"])
        root.deiconify()
        set_titlebar_theme(root, self.theme.name == "dark", self.theme.p["bg"])
        root.after(1500, self.check_latest_app_version_async)
        root.after(800, self._prewarm_other_theme)

    # ------------------------------------------------------------------
    #  Building the UI
    # ------------------------------------------------------------------
    def _tk(self, widget, **opts):
        self._themed_tk.append((widget, opts))
        return widget

    def _retheme_tk(self):
        p = self.theme.p
        for w, opts in list(self._themed_tk):
            try:
                w.configure(**{k: p[v] for k, v in opts.items()})
            except tk.TclError:
                self._themed_tk.remove((w, opts))
        self.root.configure(bg=p["bg"])

    def _card(self, parent, **grid):
        card = ttk.Frame(parent, style="Card.TFrame", padding=(self.px(16), self.px(12)))
        card.grid(**grid)
        return card

    def _btn(self, parent, text, cmd, style="Card.Secondary.TButton", icon=None, **kw):
        b = ttk.Button(parent, text=text, command=cmd, style=style, **kw)
        b._icon_name = icon
        self._icon_buttons.append(b)
        return b

    def _apply_button_icons(self):
        p = self.theme.p
        for b in self._icon_buttons:
            name = getattr(b, "_icon_name", None)
            if not name:
                continue
            style = str(b.cget("style"))
            color = p["on_accent"] if ("Accent" in style or "Danger" in style) else p["text"]
            img = self.theme.icon(name, color, 16)
            if getattr(b, "_icon_img", "unset") is img:
                continue  # unchanged -> no reconfigure, no redraw
            b._icon_img = img
            if img:
                b.configure(image=img, compound="left")
            else:
                b.configure(image="", compound="none")

    def _build(self):
        px = self.px
        self._icon_buttons = []
        outer = ttk.Frame(self.root, padding=(px(20), px(14), px(20), px(16)))
        outer.pack(fill="both", expand=True)
        outer.columnconfigure(0, weight=1)
        self.outer = outer

        # ===== Header =====
        header = ttk.Frame(outer)
        header.grid(row=0, column=0, sticky="ew", pady=(0, px(12)))
        header.columnconfigure(0, weight=1)
        title_box = ttk.Frame(header)
        title_box.grid(row=0, column=0, sticky="w")
        title_row = ttk.Frame(title_box)
        title_row.pack(anchor="w")
        ttk.Label(title_row, text=APP_TITLE, style="Title.TLabel").pack(side="left")
        ttk.Label(title_row, text=APP_VERSION_ONLY, style="Pill.TLabel").pack(side="left", padx=(px(10), 0),
                                                                              pady=(px(4), 0))
        ttk.Label(title_box, text="Keep every installed app up to date with winget — in one click.",
                  style="Sub.TLabel").pack(anchor="w", pady=(px(2), 0))

        hdr_right = ttk.Frame(header)
        hdr_right.grid(row=0, column=1, sticky="e")
        if is_admin():
            self.btn_admin = ttk.Label(hdr_right, text="●  Administrator", style="AdminPill.TLabel")
            self.btn_admin.pack(side="left", padx=(0, px(8)))
            ToolTip(self, self.btn_admin, "Running elevated — installers can run silently.")
        else:
            self.btn_admin = self._btn(hdr_right, "Run as admin", self.run_as_admin, "Secondary.TButton", "shield")
            self.btn_admin.pack(side="left", padx=(0, px(8)))
            ToolTip(self, self.btn_admin, "Restart elevated so installers can run silently\n"
                                          "(no UAC prompt for every app).")
        self.btn_theme = self._btn(hdr_right, "", self.toggle_theme, "Ghost.TButton", None)
        self.btn_theme.pack(side="left", padx=(0, px(4)))
        ToolTip(self, self.btn_theme, "Switch light / dark theme")
        self.btn_about = self._btn(hdr_right, "About", self.show_about, "Ghost.TButton", "info")
        self.btn_about.pack(side="left")

        # ===== Toolbar card =====
        bar = self._card(outer, row=1, column=0, sticky="ew", pady=(0, px(12)))
        bar.columnconfigure(3, weight=1)
        self.btn_check = self._btn(bar, "Check for updates", self.check_for_updates_async,
                                   "Card.Accent.TButton", "refresh")
        self.btn_check.grid(row=0, column=0, sticky="w")
        self.btn_update = self._btn(bar, "Update selected", self.update_selected_async,
                                    "Card.Secondary.TButton", "download", width=20)
        self.btn_update.grid(row=0, column=1, sticky="w", padx=(px(8), 0))
        self.btn_skip = self._btn(bar, "Skip current", self.skip_current, "Card.Secondary.TButton", "skip",
                                  state="disabled")
        self.btn_skip.grid(row=0, column=2, sticky="w", padx=(px(8), 0))
        ToolTip(self, self.btn_skip, "Stop the app that is updating right now and continue with the next one.")

        self.btn_temp = self._btn(bar, "Temp files", self._show_temp_menu, "Card.Ghost.TButton", "folder")
        self.btn_temp.grid(row=0, column=4, sticky="e")
        ToolTip(self, self.btn_temp, "Open or clean your Windows temp folder\n(leftover installers, caches…).")

        toggles = ttk.Frame(bar, style="Inner.TFrame")
        toggles.grid(row=1, column=0, columnspan=5, sticky="w", pady=(px(12), 0))
        self.include_unknown_var = tk.BooleanVar(value=bool(self.settings.get("include_unknown", False)))
        self.chk_unknown = ttk.Checkbutton(toggles, text="  Include apps with unknown version",
                                           variable=self.include_unknown_var, style="Switch.TCheckbutton",
                                           command=self._save_prefs)
        self.chk_unknown.pack(side="left")
        ToolTip(self, self.chk_unknown, "Also list apps whose installed version winget can't detect.")
        self.close_running_var = tk.BooleanVar(value=bool(self.settings.get("auto_close", False)))
        self.chk_close_running = ttk.Checkbutton(toggles, text="  Auto-close apps that block their update",
                                                 variable=self.close_running_var, style="Switch.TCheckbutton",
                                                 command=self._save_prefs)
        self.chk_close_running.pack(side="left", padx=(px(28), 0))
        ToolTip(self, self.chk_close_running,
                "When an app blocks its own update:\n"
                "  • Off (default): you're asked per app whether to close it.\n"
                "  • On: blocking apps are force-closed automatically.\n"
                "⚠ Closing an app loses unsaved work in it.")

        self.temp_menu = self._tk(tk.Menu(self.root, tearoff=0, bd=0, relief="flat",
                                          font=self.theme.fonts["body"]),
                                  bg="surface2", fg="text", activebackground="select",
                                  activeforeground="text")
        self.temp_menu.add_command(label="  Open temp folder", command=self.open_temp)
        self.temp_menu.add_command(label="  Clean temp files…", command=self.clear_temp_async)

        # ===== List card =====
        outer.rowconfigure(2, weight=5)
        lst = self._card(outer, row=2, column=0, sticky="nsew", pady=(0, px(12)))
        lst.columnconfigure(0, weight=1)
        lst.rowconfigure(1, weight=1)
        lhead = ttk.Frame(lst, style="Inner.TFrame")
        lhead.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, px(10)))
        lhead.columnconfigure(2, weight=1)
        ttk.Label(lhead, text="Available updates", style="CardH.TLabel").grid(row=0, column=0, sticky="w")
        self.counter_var = tk.StringVar(value="")
        ttk.Label(lhead, textvariable=self.counter_var, style="CardMuted.TLabel").grid(
            row=0, column=1, sticky="w", padx=(px(10), 0), pady=(px(2), 0))

        self.filter_var = tk.StringVar()
        self.search = ttk.Entry(lhead, textvariable=self.filter_var, style="Search.TEntry", width=26,
                                font=self.theme.fonts["body"])
        self.search.grid(row=0, column=3, sticky="e", padx=(0, px(8)))
        self._search_placeholder = True
        self.search.insert(0, "Filter apps…")
        self.search.bind("<FocusIn>", self._search_focus_in)
        self.search.bind("<FocusOut>", self._search_focus_out)
        self._filter_job = None
        self.filter_var.trace_add("write", lambda *_: self._schedule_filter())
        self.btn_sel_all = self._btn(lhead, "Select all", self.select_all, "Small.Card.Ghost.TButton",
                                     state="disabled")
        self.btn_sel_all.grid(row=0, column=4, sticky="e")
        self.btn_sel_none = self._btn(lhead, "Select none", self.select_none, "Small.Card.Ghost.TButton",
                                      state="disabled")
        self.btn_sel_none.grid(row=0, column=5, sticky="e", padx=(px(4), 0))

        self.tree = ttk.Treeview(lst, columns=self.COLS, show="tree headings", selectmode="none",
                                 style="App.Treeview", height=6)
        headings = {"Name": "NAME", "Id": "PACKAGE ID", "Current": "INSTALLED", "Available": "AVAILABLE",
                    "Result": "STATUS"}
        for c in self.COLS:
            self.tree.heading(c, text=headings[c], anchor="w", command=lambda c=c: self.sort_by(c))
        self._headings = headings
        self.tree.heading("#0", text="", anchor="center")
        self.tree.column("#0", width=px(46), minwidth=px(46), stretch=False, anchor="center")
        self.tree.column("Name", width=px(250), minwidth=px(140), stretch=True, anchor="w")
        self.tree.column("Id", width=px(250), minwidth=px(140), stretch=True, anchor="w")
        self.tree.column("Current", width=px(115), minwidth=px(80), stretch=False, anchor="w")
        self.tree.column("Available", width=px(115), minwidth=px(80), stretch=False, anchor="w")
        self.tree.column("Result", width=px(150), minwidth=px(110), stretch=False, anchor="w")
        ysb = ttk.Scrollbar(lst, orient="vertical", command=self.tree.yview, style="Slim.Vertical.TScrollbar")
        self.tree.configure(yscrollcommand=ysb.set)
        self.tree.grid(row=1, column=0, sticky="nsew")
        ysb.grid(row=1, column=1, sticky="ns", padx=(px(4), 0))
        self.tree.bind("<Button-1>", self._on_tree_click)
        self.tree.bind("<Double-Button-1>", self._on_double_click_header, add="+")
        self.tree.bind("<Motion>", self._on_tree_motion)
        self.tree.bind("<Leave>", lambda e: self._set_hover(None))
        self.tree.bind("<space>", self._on_space)
        self.tree.bind("<Button-3>", self._show_row_menu)

        self.empty_lbl = ttk.Label(lst, text="No scan yet\nClick “Check for updates” to find apps with newer versions.",
                                   style="Empty.TLabel", anchor="center", justify="center")
        self.empty_lbl.place(in_=self.tree, relx=0.5, rely=0.55, anchor="center")

        self.row_menu = self._tk(tk.Menu(self.root, tearoff=0, bd=0, relief="flat", font=self.theme.fonts["body"]),
                                 bg="surface2", fg="text", activebackground="select", activeforeground="text")
        self.row_menu.add_command(label="  Show downloaded installer(s)", command=self._menu_open_downloads)
        self.row_menu.add_command(label="  Delete downloaded installer(s)…", command=self._menu_delete_downloads)
        self.row_menu.add_separator()
        self.row_menu.add_command(label="  Copy package ID", command=self._menu_copy_id)
        self._menu_item = None

        # ===== Progress card (overall | current app, side by side) =====
        prog = self._card(outer, row=3, column=0, sticky="ew", pady=(0, px(12)))
        prog.columnconfigure(0, weight=1, uniform="prog")
        prog.columnconfigure(2, weight=1, uniform="prog")
        prog.columnconfigure(1, minsize=px(28))

        def bar_block(col, title_style):
            box = ttk.Frame(prog, style="Inner.TFrame")
            box.grid(row=0, column=col, sticky="ew")
            box.columnconfigure(0, weight=1)
            lbl = ttk.Label(box, text="", style=title_style)
            lbl.grid(row=0, column=0, sticky="w")
            val = ttk.Label(box, text="", style="CardValue.TLabel")
            val.grid(row=0, column=1, sticky="e")
            bar = ttk.Progressbar(box, orient="horizontal", mode="determinate", maximum=100,
                                  style="Accent.Horizontal.TProgressbar")
            bar.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(px(8), 0))
            return lbl, val, bar

        self.pb_label, self.pb_value_lbl, self.pb = bar_block(0, "Card.TLabel")
        self.pb_label.configure(text="Ready", font=self.theme.fonts["semibold"])
        self.pb2_label, self.pb2_value_lbl, self.pb2 = bar_block(2, "CardMuted.TLabel")
        self.pb2_label.configure(text="Waiting for a task")
        self._phase = None

        # ===== Log card =====
        self.log_card = self._card(outer, row=4, column=0, sticky="nsew")
        self.log_card.columnconfigure(0, weight=1)
        lh = ttk.Frame(self.log_card, style="Inner.TFrame")
        lh.grid(row=0, column=0, columnspan=2, sticky="ew")
        lh.columnconfigure(0, weight=1)
        ttk.Label(lh, text="Activity log", style="CardH.TLabel").grid(row=0, column=0, sticky="w")
        self.btn_save_log = self._btn(lh, "Export", self.save_export_log, "Small.Card.Ghost.TButton", "save")
        self.btn_save_log.grid(row=0, column=1, sticky="e")
        self.btn_toggle_log = self._btn(lh, "Hide", self.toggle_log, "Small.Card.Ghost.TButton", "chevron_up")
        self.btn_toggle_log.grid(row=0, column=2, sticky="e", padx=(px(4), 0))

        self.log_body = ttk.Frame(self.log_card, style="Inner.TFrame")
        self.log_body.grid(row=1, column=0, columnspan=2, sticky="nsew", pady=(px(10), 0))
        self.log_body.columnconfigure(0, weight=1)
        self.log_body.rowconfigure(0, weight=1)
        self.log_box = self._tk(tk.Text(self.log_body, wrap="none", height=4, font=self.theme.fonts["mono"],
                                        relief="flat", bd=0, padx=px(10), pady=px(8), highlightthickness=0,
                                        undo=False),
                                bg="log_bg", fg="text", insertbackground="text", selectbackground="select",
                                selectforeground="text")
        ys = ttk.Scrollbar(self.log_body, orient="vertical", command=self.log_box.yview,
                           style="Slim.Vertical.TScrollbar")
        xs = ttk.Scrollbar(self.log_body, orient="horizontal", command=self.log_box.xview,
                           style="Slim.Horizontal.TScrollbar")
        self.log_box.configure(yscrollcommand=ys.set, xscrollcommand=xs.set)
        self.log_box.grid(row=0, column=0, sticky="nsew")
        ys.grid(row=0, column=1, sticky="ns", padx=(px(4), 0))
        xs.grid(row=1, column=0, sticky="ew", pady=(px(4), 0))
        self.log_box.bind("<Key>", self._log_readonly)

        self.log_visible = True
        if self.settings.get("log_visible", True) is False:
            self.toggle_log(save=False)
        else:
            outer.rowconfigure(4, weight=2, minsize=self.px(130))

        self._controls_to_disable = [self.btn_check, self.chk_unknown, self.chk_close_running,
                                     self.btn_sel_all, self.btn_sel_none, self.btn_temp, self.btn_about,
                                     self.btn_theme]
        self._style_tree_tags()
        self._apply_button_icons()
        self._refresh_theme_button()
        self.update_counter()

    def _style_tree_tags(self):
        p = self.theme.p
        # Later-configured tags win, so order = priority (lowest first).
        self.tree.tag_configure("odd", background=p["row_alt"])
        self.tree.tag_configure("hover", background=p["row_hover"])
        self.tree.tag_configure("checked", background=p["row_checked"])
        self.tree.tag_configure("active", background=p["active_bg"], foreground=p["accent"])
        self.tree.tag_configure("ok", background=p["ok_bg"], foreground=p["success"])
        self.tree.tag_configure("fail", background=p["fail_bg"], foreground=p["danger"])
        self.tree.tag_configure("skip", background=p["skip_bg"], foreground=p["warn"])
        self.tree.tag_configure("cancel", background=p["cancel_bg"], foreground=p["warn"])
        self.img_unchecked, self.img_checked = self.theme.images["cb"]
        for it in self.all_items:
            self.tree.item(it, image=self.img_checked if it in self.checked_items else self.img_unchecked)
        tb = self.log_box
        tb.tag_configure("time", foreground=p["faint"])
        tb.tag_configure("ok", foreground=p["success"])
        tb.tag_configure("err", foreground=p["danger"])
        tb.tag_configure("warn", foreground=p["warn"])
        tb.tag_configure("muted", foreground=p["muted"])
        tb.tag_configure("head", foreground=p["accent"])

    # ------------------------------------------------------------------
    #  Theme switching & prefs
    # ------------------------------------------------------------------
    def _refresh_theme_button(self):
        dark = self.theme.name == "dark"
        img = self.theme.icon("sun" if dark else "moon", self.theme.p["text"], 16)
        if img:
            self.btn_theme.configure(image=img, text="", compound="image", padding=(self.px(9), self.px(7)))
        else:
            self.btn_theme.configure(text="☀" if dark else "☾", image="", compound="none")

    def toggle_theme(self):
        new = "light" if self.theme.name == "dark" else "dark"
        for b in self._icon_buttons:
            b._icon_img = "unset"
        self.theme.apply(new)
        self._retheme_tk()
        self._style_tree_tags()
        self._apply_button_icons()
        self._refresh_theme_button()
        self._search_refresh_color()
        self._restripe()
        for w in [self.root] + [d for d in self._dialogs if d.winfo_exists()]:
            set_titlebar_theme(w, new == "dark", self.theme.p["bg"])
        self.settings["theme"] = new
        save_settings(self.settings)

    def _prewarm_other_theme(self):
        """Build the other palette's images/elements once, while idle."""
        cur = self.theme.name
        other = "light" if cur == "dark" else "dark"
        try:
            self.theme.apply(other)
            for name in ICONS:
                for col in (PALETTES[other]["text"], PALETTES[other]["on_accent"]):
                    self.theme.icon(name, col, 16)
        finally:
            self.theme.apply(cur)

    def _save_prefs(self):
        self.settings["include_unknown"] = bool(self.include_unknown_var.get())
        self.settings["auto_close"] = bool(self.close_running_var.get())
        self.settings["log_visible"] = bool(self.log_visible)
        save_settings(self.settings)

    # ------------------------------------------------------------------
    #  Thread-safe UI pump
    # ------------------------------------------------------------------
    def post(self, fn, *args):
        """Run fn(*args) on the Tk thread (safe to call from any thread)."""
        self._uiq.put((fn, args))

    def _pump(self):
        try:
            for _ in range(300):
                fn, args = self._uiq.get_nowait()
                try:
                    fn(*args)
                except Exception:
                    traceback.print_exc()
        except queue.Empty:
            pass
        pending = self._pct_pending
        if pending is not None:
            self._pct_pending = None
            self.per_app_update_percent(*pending)
        self._flush_log()
        self.root.after(40, self._pump)

    # ------------------------------------------------------------------
    #  Logging (thread-safe, batched so bursts of output never stall the UI)
    # ------------------------------------------------------------------
    def log(self, text):
        stamp = time.strftime("%H:%M:%S")
        with self._log_lock:
            for line in str(text).splitlines() or [""]:
                self._log_buf.append((stamp, line))

    def _flush_log(self):
        with self._log_lock:
            if not self._log_buf:
                return
            buf, self._log_buf = self._log_buf, []
        tb = self.log_box
        at_end = tb.yview()[1] >= 0.999
        for stamp, line in buf:
            low = line.lower()
            if line.startswith(("✔", "✓")) or "success" in low:
                tag = "ok"
            elif line.startswith(("✕", "❌")) or "error" in low or "failed" in low:
                tag = "err"
            elif line.startswith(("⚠", "[Skip]", "[Stalled]", "[In use]")):
                tag = "warn"
            elif line.startswith(("Updating ", "Starting", "Summary")):
                tag = "head"
            elif line.startswith("["):
                tag = "muted"
            else:
                tag = ""
            tb.insert("end", stamp + "  ", ("time",), line + "\n", (tag,) if tag else ())
        lines = int(tb.index("end-1c").split(".")[0])
        if lines > LOG_MAX_LINES:
            tb.delete("1.0", f"{lines - LOG_MAX_LINES}.0")
        if at_end:
            tb.see("end")

    def _log_readonly(self, e):
        # Allow copy / select-all, block edits.
        if (e.state & 0x4) and e.keysym.lower() in ("c", "a"):
            if e.keysym.lower() == "a":
                self.log_box.tag_add("sel", "1.0", "end")
                return "break"
            return None
        if e.keysym in ("Up", "Down", "Left", "Right", "Prior", "Next", "Home", "End"):
            return None
        return "break"

    # ------------------------------------------------------------------
    #  Dialog helpers
    # ------------------------------------------------------------------
    def _dialog(self, title, modal=True):
        p = self.theme.p
        win = tk.Toplevel(self.root)
        win.withdraw()
        win.title(title)
        win.resizable(False, False)
        win.transient(self.root)
        win.configure(bg=p["bg"])
        if self.window_icon_path:
            try:
                win.iconbitmap(self.window_icon_path)
            except Exception:
                pass
        frame = ttk.Frame(win, padding=(self.px(24), self.px(20)))
        frame.pack(fill="both", expand=True)
        self._dialogs = [d for d in self._dialogs if d.winfo_exists()] + [win]
        win._modal = modal
        return win, frame

    def _show_dialog(self, win):
        self.center_child(win)
        win.deiconify()
        set_titlebar_theme(win, self.theme.name == "dark", self.theme.p["bg"])
        if getattr(win, "_modal", False):
            try:
                win.grab_set()
            except tk.TclError:
                pass
        win.focus_force()

    def center_child(self, tlv):
        tlv.update_idletasks()
        w, h = tlv.winfo_reqwidth(), tlv.winfo_reqheight()
        x = self.root.winfo_rootx() + (self.root.winfo_width() - w) // 2
        y = self.root.winfo_rooty() + (self.root.winfo_height() - h) // 3
        tlv.geometry(f"+{max(0, x)}+{max(0, y)}")

    def center_on_screen(self):
        w, h = self._win_w, self._win_h
        sw, sh = self.root.winfo_screenwidth(), self.root.winfo_screenheight()
        self.root.geometry(f"{w}x{h}+{max(0, (sw - w) // 2)}+{max(0, (sh - h) // 2 - self.px(20))}")

    def show_loading(self, text="Loading..."):
        if self.loading_win:
            return
        win, frame = self._dialog("Please wait")
        self.loading_win = win
        ttk.Label(frame, text=text, font=self.theme.fonts["h2"]).pack(anchor="w", pady=(0, self.px(12)))
        pb = ttk.Progressbar(frame, mode="indeterminate", length=self.px(320),
                             style="Accent.Horizontal.TProgressbar")
        pb.pack(fill="x")
        pb.start(12)
        win.protocol("WM_DELETE_WINDOW", lambda: None)
        self._show_dialog(win)

    def hide_loading(self):
        if self.loading_win:
            try:
                self.loading_win.grab_release()
            except Exception:
                pass
            try:
                self.loading_win.destroy()
            except Exception:
                pass
            self.loading_win = None

    # ------------------------------------------------------------------
    #  Search / filter / sort
    # ------------------------------------------------------------------
    def _search_focus_in(self, _=None):
        if self._search_placeholder:
            self._search_placeholder = False
            self.search.delete(0, "end")
            self._search_refresh_color()

    def _search_focus_out(self, _=None):
        if not self.search.get():
            self._search_placeholder = True
            self.search.insert(0, "Filter apps…")
            self._search_refresh_color()

    def _search_refresh_color(self):
        p = self.theme.p
        self.search.configure(foreground=p["faint"] if self._search_placeholder else p["text"])

    def _filter_text(self):
        return "" if self._search_placeholder else self.filter_var.get().strip().lower()

    def _schedule_filter(self):
        # Debounce: re-filter once typing pauses, not on every keystroke.
        if self._filter_job:
            self.root.after_cancel(self._filter_job)
        self._filter_job = self.root.after(140, self.apply_filter)

    def apply_filter(self):
        self._filter_job = None
        q = self._filter_text()
        for it in self.all_items:
            if not self.tree.exists(it):
                continue
            if not q:
                show = True
            else:
                v = self.tree.item(it, "values")
                show = q in str(v[0]).lower() or q in str(v[1]).lower()
            if show:
                self.tree.move(it, "", "end")
            else:
                self.tree.detach(it)
        self._restripe()
        self.update_counter()

    def _restripe(self):
        for i, it in enumerate(self.tree.get_children("")):
            tags = [t for t in self.tree.item(it, "tags") if t != "odd"]
            if i % 2 == 1:
                tags.insert(0, "odd")
            self.tree.item(it, tags=tags)

    def sort_by(self, col):
        if self.updating or not self.all_items:
            return
        prev_col, prev_desc = self._sort_state
        desc = (not prev_desc) if prev_col == col else False
        self._sort_state = (col, desc)

        def key(it):
            v = self.tree.set(it, col)
            if col in ("Current", "Available"):
                return [(0, int(x), "") if x.isdigit() else (1, 0, x) for x in re.split(r"[.\-]", v.lower())]
            return v.lower()

        self.all_items.sort(key=key, reverse=desc)
        for c in self.COLS:
            arrow = ("  ↓" if desc else "  ↑") if c == col else ""
            self.tree.heading(c, text=self._headings[c] + arrow)
        self.apply_filter()

    # ------------------------------------------------------------------
    #  Tree interaction
    # ------------------------------------------------------------------
    def _on_tree_click(self, e):
        region = self.tree.identify("region", e.x, e.y)
        if region in ("tree", "cell"):
            if self.updating:
                return "break"
            it = self.tree.identify_row(e.y)
            if it:
                self._toggle_row(it)
            return "break"
        if region == "separator" and self.tree.identify_column(e.x) == "#0":
            return "break"  # keep the checkbox column fixed
        return None

    def _on_space(self, _=None):
        it = self.tree.focus() or self._hover_item
        if it and not self.updating:
            self._toggle_row(it)
        return "break"

    def _on_tree_motion(self, e):
        it = self.tree.identify_row(e.y)
        if it != self._hover_item:
            self._set_hover(it or None)

    def _set_hover(self, it):
        old = self._hover_item
        self._hover_item = it
        for item, add in ((old, False), (it, True)):
            if item and self.tree.exists(item):
                tags = [t for t in self.tree.item(item, "tags") if t != "hover"]
                if add:
                    tags.append("hover")
                self.tree.item(item, tags=tags)

    def _on_double_click_header(self, e):
        if self.tree.identify("region", e.x, e.y) != "separator":
            return
        col = self.tree.identify_column(e.x - 1)
        if not col or col == "#0":
            return
        cid = self.COLS[int(col[1:]) - 1]
        f = self.theme.fonts["body"]
        mx = self.theme.fonts["small_semi"].measure(self.tree.heading(cid, "text"))
        for it in self.all_items:
            mx = max(mx, f.measure(self.tree.set(it, cid)))
        self.tree.column(cid, width=mx + self.px(28))

    def _set_checked(self, it, on):
        if on:
            self.checked_items.add(it)
        else:
            self.checked_items.discard(it)
        tags = [t for t in self.tree.item(it, "tags") if t != "checked"]
        if on:
            tags.append("checked")
        self.tree.item(it, image=self.img_checked if on else self.img_unchecked, tags=tags)

    def _toggle_row(self, item):
        if not item or self.updating:
            return
        self._set_checked(item, item not in self.checked_items)
        self.update_counter()

    def select_all(self):
        if self.updating:
            return
        for it in self.tree.get_children(""):  # visible (filtered) rows
            self._set_checked(it, True)
        self.update_counter()

    def select_none(self):
        if self.updating:
            return
        for it in self.all_items:
            self._set_checked(it, False)
        self.update_counter()

    def update_counter(self):  # noqa: C901
        total = len(self.all_items)
        shown = len(self.tree.get_children(""))
        sel = len(self.checked_items)
        if total == 0:
            self.counter_var.set("")
        else:
            txt = f"{total} found  ·  {sel} selected"
            if shown != total:
                txt += f"  ·  {shown} shown"
            self.counter_var.set(txt)
        if not self.updating:
            want_style = "Card.Accent.TButton" if sel else "Card.Secondary.TButton"
            want_text = f"Update selected ({sel})" if sel else "Update selected"
            changes = {}
            if str(self.btn_update.cget("style")) != want_style:
                changes["style"] = want_style
            if str(self.btn_update.cget("text")) != want_text:
                changes["text"] = want_text
            if changes:
                self.btn_update.configure(**changes)
                if "style" in changes:  # icon colour depends on the style
                    self._apply_button_icons()

    def clear_tree(self):
        self.checked_items.clear()
        self.id_to_item.clear()
        self.all_items = []
        self._hover_item = None
        self.tree.delete(*self.tree.get_children(""))

    def _set_result(self, it, key):
        if not self.tree.exists(it):
            return
        text, tag = STATUS[key]
        self.tree.set(it, "Result", text)
        keep = [t for t in self.tree.item(it, "tags") if t in ("odd", "hover", "checked")]
        if tag:
            keep.append(tag)
        self.tree.item(it, tags=keep)

    def _show_row_menu(self, e):
        if self.updating:
            return
        it = self.tree.identify_row(e.y)
        if not it:
            return
        self._menu_item = it
        try:
            self.row_menu.tk_popup(e.x_root, e.y_root)
        finally:
            self.row_menu.grab_release()

    def _show_temp_menu(self):
        b = self.btn_temp
        try:
            self.temp_menu.tk_popup(b.winfo_rootx(), b.winfo_rooty() + b.winfo_height() + self.px(2))
        finally:
            self.temp_menu.grab_release()

    # ------------------------------------------------------------------
    #  Log panel
    # ------------------------------------------------------------------
    def toggle_log(self, save=True):
        if self.log_visible:
            self.log_body.grid_remove()
            self.outer.rowconfigure(4, weight=0, minsize=0)
            self.btn_toggle_log.configure(text="Show")
            self.btn_toggle_log._icon_name = "chevron_down"
            self.log_visible = False
        else:
            self.log_body.grid()
            self.outer.rowconfigure(4, weight=2, minsize=self.px(130))
            self.btn_toggle_log.configure(text="Hide")
            self.btn_toggle_log._icon_name = "chevron_up"
            self.log_visible = True
        self._apply_button_icons()
        if save:
            self._save_prefs()

    def save_export_log(self):
        content = self.log_box.get("1.0", "end-1c")
        if not content.strip():
            messagebox.showinfo("Export log", "The log is empty.", parent=self.root)
            return
        path = filedialog.asksaveasfilename(
            parent=self.root, title="Save log", defaultextension=".txt",
            initialfile=f"WindowsAppUpdater_{time.strftime('%Y%m%d_%H%M%S')}.txt",
            filetypes=[("Text file", "*.txt"), ("Log file", "*.log"), ("All files", "*.*")])
        if not path:
            return
        try:
            with open(path, "w", encoding="utf-8") as f:
                f.write(content)
        except Exception as e:
            messagebox.showerror("Export log", str(e), parent=self.root)
            return
        self.log(f"[Log] Saved to {path}")
        self._reveal_in_explorer(path)

    @staticmethod
    def _reveal_in_explorer(path):
        try:
            if IS_WIN:
                subprocess.Popen(["explorer", "/select,", os.path.normpath(path)])
            else:
                webbrowser.open(os.path.dirname(path))
        except Exception:
            pass

    # ------------------------------------------------------------------
    #  Enable/disable during update
    # ------------------------------------------------------------------
    def _disable_controls_for_update(self):
        self.updating = True
        for w in self._controls_to_disable:
            try:
                w.state(["disabled"])
            except Exception:
                pass
        if isinstance(self.btn_admin, ttk.Button):
            self.btn_admin.state(["disabled"])
        self.search.state(["disabled"])
        self.btn_update.configure(text="Cancel", style="Card.Danger.TButton")
        self.btn_update._icon_name = "cancel"
        self.btn_update.state(["!disabled"])
        self._apply_button_icons()

    def _enable_controls_after_update(self):
        self.updating = False
        for w in self._controls_to_disable:
            try:
                w.state(["!disabled"])
            except Exception:
                pass
        if isinstance(self.btn_admin, ttk.Button):
            self.btn_admin.state(["!disabled"])
        self.search.state(["!disabled"])
        self.btn_update._icon_name = "download"
        self.btn_update.state(["!disabled"])
        self._enable_select_buttons(bool(self.all_items))
        self.update_counter()

    def _enable_select_buttons(self, enable):
        for b in (self.btn_sel_all, self.btn_sel_none):
            b.state(["!disabled"] if enable else ["disabled"])

    # ------------------------------------------------------------------
    #  Progress
    # ------------------------------------------------------------------
    def progress_start(self, total):
        self.pb_total = max(0, int(total))
        self.pb_value = 0
        self.pb.configure(maximum=max(self.pb_total, 1), value=0, style="Accent.Horizontal.TProgressbar")
        self.pb_label.configure(text=f"Updating {self.pb_total} app{'s' if self.pb_total != 1 else ''}")
        self.pb_value_lbl.configure(text=f"0 / {self.pb_total}")
        self._pb2_reset("Preparing…")

    def progress_step(self, inc=1):
        if getattr(self, "pb_total", 0) <= 0:
            return
        self.pb_value = min(self.pb_total, self.pb_value + inc)
        self.pb.configure(value=self.pb_value)
        self.pb_value_lbl.configure(text=f"{self.pb_value} / {self.pb_total}")

    def progress_finish(self, canceled=False, summary=""):
        total = getattr(self, "pb_total", 0)
        if total > 0:
            self.pb.configure(value=total if not canceled else self.pb_value,
                              style="Accent.Horizontal.TProgressbar" if canceled else "Success.Horizontal.TProgressbar")
        self.pb_label.configure(text="Canceled" if canceled else "All done")
        self.pb_value_lbl.configure(text=summary)
        self._pb2_reset("Waiting for a task")

    def _pb2_reset(self, text):
        self._phase = None
        try:
            self.pb2.stop()
        except Exception:
            pass
        self.pb2.configure(mode="determinate", value=0)
        self.pb2_label.configure(text=text)
        self.pb2_value_lbl.configure(text="")

    def per_app_reset(self, name):
        self._phase = "download"
        self._pct_pending = None
        try:
            self.pb2.stop()
        except Exception:
            pass
        self.pb2.configure(mode="determinate", value=0)
        self.pb2_label.configure(text=f"Downloading {name}")
        self.pb2_value_lbl.configure(text="")

    def per_app_begin_install(self, name):
        if self._phase == "install":
            return
        self._phase = "install"
        self.pb2_label.configure(text=f"Installing {name}")
        self.pb2_value_lbl.configure(text="")
        self.pb2.configure(mode="indeterminate")
        self.pb2.start(14)

    def per_app_update_percent(self, pct, name=None):
        if self._phase != "download":
            return
        pct = max(0, min(100, int(pct)))
        self.pb2.configure(value=pct)
        self.pb2_value_lbl.configure(text=f"{pct}%")

    def per_app_finish(self, name):
        self._phase = None
        try:
            self.pb2.stop()
        except Exception:
            pass
        self.pb2.configure(mode="determinate", value=100)
        self.pb2_label.configure(text=f"Finished {name}")
        self.pb2_value_lbl.configure(text="")

    # ------------------------------------------------------------------
    #  Admin / close
    # ------------------------------------------------------------------
    def run_as_admin(self):
        if self.updating:
            return
        if relaunch_as_admin():
            self.root.destroy()

    def on_close(self):
        if self.updating:
            if not messagebox.askyesno("Update in progress",
                                       "An update is still running.\n\nStop it and exit?", parent=self.root):
                return
            self.cancel_requested = True
            kill_process_tree(self.current_proc)
        self.root.destroy()

    # ------------------------------------------------------------------
    #  About
    # ------------------------------------------------------------------
    def show_about(self):
        win, frame = self._dialog("About", modal=False)
        p = self.theme.p
        px = self.px
        card = ttk.Frame(frame, style="Card.TFrame", padding=(px(24), px(20)))
        card.pack(fill="both", expand=True)
        head = ttk.Frame(card, style="Inner.TFrame")
        head.pack(anchor="w")
        ttk.Label(head, text=APP_TITLE, style="CardH.TLabel", font=self.theme.fonts["title"]).pack(side="left")
        ttk.Label(head, text=APP_VERSION_ONLY, style="Pill.TLabel").pack(side="left", padx=(px(10), 0))
        ttk.Label(card, text="A free tool that updates your Windows apps using winget.",
                  style="CardMuted.TLabel").pack(anchor="w", pady=(px(4), px(16)))

        def row(label, value=None, link=None, image=None):
            r = ttk.Frame(card, style="Inner.TFrame")
            r.pack(fill="x", pady=px(3))
            ttk.Label(r, text=label, style="CardMuted.TLabel", width=12).pack(side="left")
            if value is not None:
                lbl = ttk.Label(r, text=value, style="Link.TLabel" if link else "Card.TLabel",
                                cursor="hand2" if link else "")
                lbl.pack(side="left")
                if link:
                    lbl.bind("<Button-1>", lambda e: webbrowser.open(link))
            if image is not None:
                ttk.Label(r, image=image, style="Card.TLabel").pack(side="left", padx=(px(8), 0))

        flag = load_flag_image()
        win._flag = flag
        row("Author", "ilukezippo (BoYaqoub)", image=flag)
        row("Released", DATE_APP)
        row("Contact", CONTACT_EMAIL, link=f"mailto:{CONTACT_EMAIL}")
        row("Website", "github.com/ilukezippo/Windows-App-Updater", link=GITHUB_REPO_PAGE)

        btns = ttk.Frame(card, style="Inner.TFrame")
        btns.pack(fill="x", pady=(px(20), 0))
        donate_img = make_donate_image(self.theme.s)
        win._don = donate_img
        tk.Button(btns, image=donate_img, text="☕  Buy me a coffee", compound="center",
                  font=self.theme.fonts["semibold"], fg="#3b2300", activeforeground="#3b2300",
                  bg=p["surface"], activebackground=p["surface"], bd=0, highlightthickness=0,
                  cursor="hand2", relief="flat", command=lambda: webbrowser.open(DONATE_PAGE)).pack(side="left")
        self._btn(btns, "Close", win.destroy, "Card.Accent.TButton").pack(side="right")
        b = self._btn(btns, "Check for update", self.manual_check_for_update, "Card.Secondary.TButton", "update")
        b.pack(side="right", padx=(0, px(8)))
        self._apply_button_icons()
        win.bind("<Escape>", lambda e: win.destroy())
        self._show_dialog(win)

    # ------------------------------------------------------------------
    #  Self-update (GitHub releases)
    # ------------------------------------------------------------------
    @staticmethod
    def _parse_ver_tuple(v):
        return tuple(int(n) for n in re.findall(r"\d+", v or "")[:4]) or (0,)

    @staticmethod
    def _fetch_latest_release():
        req = urllib.request.Request(GITHUB_API_LATEST, headers={"User-Agent": "Windows-App-Updater"})
        with urllib.request.urlopen(req, timeout=10) as r:
            return json.loads(r.read().decode("utf-8", "replace"))

    def manual_check_for_update(self):
        self.show_loading("Checking for a new version…")

        def worker():
            try:
                data = self._fetch_latest_release()
                tag = str(data.get("tag_name") or data.get("name") or "").strip()
                newer = bool(tag and self._parse_ver_tuple(tag) > self._parse_ver_tuple(APP_VERSION_ONLY))
            except Exception as e:
                self.post(self._update_check_failed, str(e))
                return
            self.post(self._update_check_done, data, tag, newer)

        threading.Thread(target=worker, daemon=True).start()

    def _update_check_failed(self, msg):
        self.hide_loading()
        messagebox.showerror("Update check failed", msg, parent=self.root)

    def _update_check_done(self, data, tag, newer):
        self.hide_loading()
        if newer:
            if messagebox.askyesno("New version available",
                                   f"Version {tag} is available (you have {APP_VERSION_ONLY}).\n\n"
                                   "Download and install it now?", parent=self.root):
                self._download_and_run_latest(data)
        else:
            messagebox.showinfo("You're up to date", f"{APP_VERSION_ONLY} is the latest version.",
                                parent=self.root)

    def check_latest_app_version_async(self):
        def worker():
            try:
                data = self._fetch_latest_release()
                tag = str(data.get("tag_name") or data.get("name") or "").strip()
                if tag and self._parse_ver_tuple(tag) > self._parse_ver_tuple(APP_VERSION_ONLY):
                    self.post(self._prompt_new_version, data, tag)
            except Exception:
                pass

        threading.Thread(target=worker, daemon=True).start()

    def _prompt_new_version(self, data, tag):
        if self.updating:  # never interrupt a running batch
            self.log(f"[Update] {APP_TITLE} {tag} is available — see About › Check for update.")
            return
        if messagebox.askyesno("New version available",
                               f"Version {tag} is available (you have {APP_VERSION_ONLY}).\n\n"
                               "Download and install it now?", parent=self.root):
            self._download_and_run_latest(data)

    def _download_and_run_latest(self, release_json: dict):
        assets = release_json.get("assets") or []
        exe_assets = [a for a in assets if (a.get("name") or "").lower().endswith(".exe")
                      and a.get("browser_download_url")]
        exe_assets.sort(key=lambda a: (0 if "updater" in (a.get("name") or "").lower() else 1,
                                       a.get("name") or ""))
        if not exe_assets or not getattr(sys, "frozen", False) or not IS_WIN:
            msg = ("Self-update only works in the packaged EXE build." if exe_assets
                   else "No .exe was found in the latest release.")
            messagebox.showinfo("Update", msg + "\nOpening the releases page instead.", parent=self.root)
            webbrowser.open(GITHUB_RELEASES_PAGE)
            return
        asset = exe_assets[0]
        url, name = asset["browser_download_url"], os.path.basename(asset.get("name") or "Windows-App-Updater.exe")
        dest = os.path.join(tempfile.gettempdir(), name)
        self.show_loading(f"Downloading {name}…")

        def on_pct(pct):
            self.post(lambda: (self.pb2.configure(value=pct), self.pb2_label.configure(text=f"Downloading {name}"),
                               self.pb2_value_lbl.configure(text=f"{pct}%")))

        def worker():
            try:
                _download_file(url, dest, progress_cb=on_pct)
            except Exception as e:
                self.post(self._update_check_failed, f"Download failed:\n{e}")
                return
            self.post(self._launch_self_update, dest, name)

        threading.Thread(target=worker, daemon=True).start()

    def _launch_self_update(self, new_path, asset_name):
        """Swap the EXE with a PowerShell helper (UTF-8 with BOM, so paths with
        Arabic or other non-ASCII characters survive — a .bat file would mangle them)."""
        self.hide_loading()
        final_path = sys.executable
        target = os.path.join(os.path.dirname(final_path), asset_name)
        script = f"""$ErrorActionPreference = 'SilentlyContinue'
$old = {_ps_quote(final_path)}
$new = {_ps_quote(new_path)}
$target = {_ps_quote(target)}
try {{ Wait-Process -Id {os.getpid()} -Timeout 30 }} catch {{}}
$ok = $false
for ($i = 0; $i -lt 40; $i++) {{
    try {{ Move-Item -LiteralPath $new -Destination $target -Force -ErrorAction Stop; $ok = $true; break }}
    catch {{ Start-Sleep -Milliseconds 750 }}
}}
if ($ok) {{
    if ($old -ne $target) {{
        for ($i = 0; $i -lt 20; $i++) {{
            Remove-Item -LiteralPath $old -Force -ErrorAction SilentlyContinue
            if (-not (Test-Path -LiteralPath $old)) {{ break }}
            Start-Sleep -Milliseconds 750
        }}
    }}
    Start-Process -FilePath $target
}} else {{
    Start-Process -FilePath $new
}}
Remove-Item -LiteralPath $MyInvocation.MyCommand.Path -Force
"""
        script_path = os.path.join(tempfile.gettempdir(), "WindowsAppUpdater_selfupdate.ps1")
        try:
            with open(script_path, "w", encoding="utf-8-sig") as f:
                f.write(script)
        except Exception as e:
            messagebox.showerror("Update", f"Could not prepare the update:\n{e}", parent=self.root)
            return

        args = f'-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "{script_path}"'
        if os.access(os.path.dirname(final_path), os.W_OK):
            try:
                subprocess.Popen(["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-WindowStyle",
                                  "Hidden", "-File", script_path], creationflags=CREATE_NO_WINDOW,
                                 close_fds=True)
            except Exception as e:
                messagebox.showerror("Update", f"Could not start the updater:\n{e}", parent=self.root)
                return
        else:
            if not messagebox.askyesno("Administrator needed",
                                       "The app is installed in a protected folder.\n\n"
                                       "Allow the updater to run as administrator to replace it?",
                                       parent=self.root):
                webbrowser.open(GITHUB_RELEASES_PAGE)
                return
            rc = ctypes.windll.shell32.ShellExecuteW(None, "runas", "powershell.exe", args, None, 0)
            if not rc or int(rc) <= 32:
                messagebox.showinfo("Update", "The update was canceled.", parent=self.root)
                return
        self.root.destroy()

    # ------------------------------------------------------------------
    #  Temp folder
    # ------------------------------------------------------------------
    def _temp_dir(self):
        return tempfile.gettempdir()

    @staticmethod
    def _fmt_bytes(n):
        v = float(n)
        for u in ("B", "KB", "MB", "GB", "TB"):
            if v < 1024 or u == "TB":
                return f"{int(v)} B" if u == "B" else f"{v:.1f} {u}".replace(".0 ", " ")
            v /= 1024.0

    _SNAPSHOT_EXTS = (".exe", ".msi", ".msix", ".msixbundle", ".appx", ".appxbundle", ".zip", ".7z", ".rar",
                      ".cab")

    def _snapshot_temp(self):
        root = os.path.join(self._temp_dir(), "WinGet")
        snap = {}
        if not os.path.isdir(root):
            return snap
        for dirpath, _, files in os.walk(root):
            for fn in files:
                if os.path.splitext(fn)[1].lower() in self._SNAPSHOT_EXTS:
                    p = os.path.join(dirpath, fn)
                    try:
                        st = os.stat(p)
                        if st.st_size >= 1_000_000:
                            snap[p] = st.st_mtime
                    except OSError:
                        pass
        return snap

    @staticmethod
    def _find_new_installer_files(before, after):
        news = [p for p in after if p not in before or after[p] > before.get(p, 0)]
        return sorted(set(news), key=lambda p: after.get(p, 0), reverse=True)

    def _winget_downloads_for_id(self, package_id):
        hits = []
        base = os.path.join(self._temp_dir(), "WinGet")
        if not (package_id and os.path.isdir(base)):
            return hits
        pid_l = package_id.lower()
        for dirpath, _, files in os.walk(base):
            if pid_l in os.path.basename(dirpath).lower():
                for fn in files:
                    if os.path.splitext(fn)[1].lower() in self._SNAPSHOT_EXTS:
                        hits.append(os.path.join(dirpath, fn))

        def mtime(p):
            try:
                return os.path.getmtime(p)
            except OSError:
                return 0

        hits.sort(key=mtime, reverse=True)
        return hits

    def open_temp(self):
        try:
            os.startfile(self._temp_dir())
        except Exception as e:
            messagebox.showerror("Open temp", str(e), parent=self.root)

    def clear_temp_async(self):
        tmp = self._temp_dir()
        if not messagebox.askyesno("Clean temp files",
                                   f"Delete files and folders inside:\n\n{tmp}\n\n"
                                   "Anything currently in use is skipped automatically.\n\nContinue?",
                                   parent=self.root):
            return
        self.show_loading("Cleaning temp files…")
        protect = set()
        mei = getattr(sys, "_MEIPASS", None)
        if mei:
            protect.add(_norm(mei))  # our own unpacked runtime (onefile build)

        def worker():
            files = dirs = skipped = 0
            freed = 0
            details = []
            try:
                entries = list(os.scandir(tmp))
            except Exception as e:
                self.log(f"[Clean temp] Error: {e}")
                entries = []
            for e in entries:
                p = e.path
                if _norm(p) in protect:
                    continue
                try:
                    if e.is_file(follow_symlinks=False):
                        size = e.stat(follow_symlinks=False).st_size
                        os.remove(p)
                        files += 1
                        freed += size
                        details.append(f"  file    {p}")
                    elif e.is_dir(follow_symlinks=False):
                        # Renaming a folder fails on Windows if anything inside is open.
                        # That tells us the whole folder is in use (e.g. another app's
                        # unpacked runtime), so we leave it intact instead of half-deleting it.
                        probe = p + ".wau-del"
                        try:
                            os.rename(p, probe)
                        except OSError:
                            skipped += 1
                            continue
                        n_f, n_b, ok = self._delete_tree(probe)
                        files += n_f
                        freed += n_b
                        if ok:
                            dirs += 1
                            details.append(f"  folder  {p}")
                        else:
                            skipped += 1
                except OSError:
                    skipped += 1
            shown = details[:200]
            if shown:
                self.log("[Clean temp] Deleted:\n" + "\n".join(shown)
                         + (f"\n  … and {len(details) - 200} more" if len(details) > 200 else ""))
            summary = (f"Folders: {dirs}  ·  Files: {files}  ·  Skipped (in use): {skipped}  ·  "
                       f"Freed: {self._fmt_bytes(freed)}")
            self.log(f"[Clean temp] {summary}")
            self.post(self._clear_temp_done, dirs, files, skipped, freed)

        threading.Thread(target=worker, daemon=True).start()

    @staticmethod
    def _delete_tree(path):
        files = freed = 0
        ok = True
        for dp, dns, fns in os.walk(path, topdown=False):
            for fn in fns:
                fp = os.path.join(dp, fn)
                try:
                    size = os.lstat(fp).st_size
                    try:
                        os.chmod(fp, 0o666)
                    except OSError:
                        pass
                    os.remove(fp)
                    files += 1
                    freed += size
                except OSError:
                    ok = False
            for dn in dns:
                try:
                    os.rmdir(os.path.join(dp, dn))
                except OSError:
                    ok = False
        try:
            os.rmdir(path)
        except OSError:
            ok = False
        return files, freed, ok

    def _clear_temp_done(self, dirs, files, skipped, freed):
        self.hide_loading()
        messagebox.showinfo("Clean temp files",
                            f"Space freed: {self._fmt_bytes(freed)}\n\n"
                            f"Folders deleted: {dirs}\nFiles deleted: {files}\nSkipped (in use): {skipped}",
                            parent=self.root)

    def _menu_open_downloads(self):
        it = self._menu_item
        if not it or self.updating:
            return
        pid = self.tree.set(it, "Id")
        files = [p for p in (self.pkg_downloads.get(pid) or []) if os.path.exists(p)]
        files = files or self._winget_downloads_for_id(pid)
        if not files:
            messagebox.showinfo("Downloads", "No downloaded installer found for this app.", parent=self.root)
            return
        # Reveal (don't run!) the newest installer.
        self._reveal_in_explorer(files[0])

    def _menu_delete_downloads(self):
        it = self._menu_item
        if not it or self.updating:
            return
        pid = self.tree.set(it, "Id")
        files = [p for p in (self.pkg_downloads.get(pid) or []) if os.path.exists(p)]
        files = [p for p in (files or self._winget_downloads_for_id(pid)) if os.path.exists(p)]
        if not files:
            messagebox.showinfo("Delete downloads", "No downloaded installer to delete for this app.",
                                parent=self.root)
            return
        if not messagebox.askyesno("Delete downloads", "Delete these file(s)?\n\n" + "\n".join(files),
                                   parent=self.root):
            return
        deleted = freed = 0
        for p in files:
            try:
                size = os.path.getsize(p)
                os.remove(p)
                deleted += 1
                freed += size
                self.log(f"[Delete] {p}")
            except Exception as e:
                self.log(f"[Delete] Could not delete {p}: {e}")
        self.pkg_downloads[pid] = [p for p in self.pkg_downloads.get(pid, []) if os.path.exists(p)]
        self.log(f"[Delete] Deleted {deleted} file(s), freed {self._fmt_bytes(freed)}")

    def _menu_copy_id(self):
        it = self._menu_item
        if it:
            self.root.clipboard_clear()
            self.root.clipboard_append(self.tree.set(it, "Id"))

    # ------------------------------------------------------------------
    #  Check for updates
    # ------------------------------------------------------------------
    def check_for_updates_async(self):
        if self.updating:
            return
        include = bool(self.include_unknown_var.get())
        self.btn_check.state(["disabled"])
        self.show_loading("Checking for updates…")
        self.log("Checking for updates…")

        def worker():
            try:
                pkgs = get_winget_upgrades(include_unknown=include)
            except Exception as e:
                self.post(self._check_failed, str(e))
                return
            self.post(self.populate_tree, pkgs)

        threading.Thread(target=worker, daemon=True).start()

    def _check_failed(self, msg):
        self.hide_loading()
        self.btn_check.state(["!disabled"])
        self.log(f"[winget] Error: {msg}")
        messagebox.showerror("winget error", f"Couldn't get the update list:\n\n{msg}", parent=self.root)

    def populate_tree(self, pkgs):
        self.hide_loading()
        self.clear_tree()
        self.btn_check.state(["!disabled"])
        self._sort_state = (None, False)
        for c in self.COLS:
            self.tree.heading(c, text=self._headings[c])
        if not pkgs:
            self.empty_lbl.configure(text="You're all set ✓\nEvery app is already up to date.")
            self.empty_lbl.place(in_=self.tree, relx=0.5, rely=0.55, anchor="center")
            self.log("✓ No apps need updating.")
            self._enable_select_buttons(False)
            self.update_counter()
            return
        self.empty_lbl.place_forget()
        for p in pkgs:
            it = self.tree.insert("", "end", text="", image=self.img_unchecked,
                                  values=(p["name"], p["id"], p.get("current", ""), p.get("available", ""), ""))
            self.id_to_item[p["id"]] = it
            self.all_items.append(it)
        self.log(f"Found {len(pkgs)} update{'s' if len(pkgs) != 1 else ''}.")
        self.apply_filter()
        self._enable_select_buttons(True)
        self.update_counter()

    # ------------------------------------------------------------------
    #  Close-app prompt (called from the worker thread)
    # ------------------------------------------------------------------
    def ask_close_app(self, name):
        result = {"choice": "no", "remember": False}
        done = threading.Event()

        def _prompt():
            try:
                win, frame = self._dialog("App is running")
                result["win"] = win
                px = self.px
                ttk.Label(frame, text=f"“{name}” is open", font=self.theme.fonts["h2"]).pack(anchor="w")
                ttk.Label(frame, text="It's blocking its own update. Close it now and retry?\n\n"
                                      "⚠ Save your work first — anything unsaved in it will be lost.",
                          style="Sub.TLabel", justify="left", wraplength=px(400)).pack(anchor="w",
                                                                                     pady=(px(6), px(14)))
                remember_var = tk.BooleanVar(value=False)
                box = ttk.Frame(frame, style="Card.TFrame", padding=(px(10), px(6)))
                box.pack(fill="x", pady=(0, px(16)))
                ttk.Checkbutton(box, text="  Remember my choice for the rest of this batch",
                                variable=remember_var, style="Switch.TCheckbutton").pack(anchor="w")
                btns = ttk.Frame(frame)
                btns.pack(fill="x")

                def choose(value):
                    result["choice"] = value
                    result["remember"] = bool(remember_var.get())
                    try:
                        win.grab_release()
                    except Exception:
                        pass
                    win.destroy()
                    done.set()

                self._btn(btns, "Close app & retry", lambda: choose("yes"), "Danger.TButton").pack(side="right")
                self._btn(btns, "Skip this app", lambda: choose("no"), "Secondary.TButton").pack(
                    side="right", padx=(0, px(8)))
                self._apply_button_icons()
                win.protocol("WM_DELETE_WINDOW", lambda: choose("no"))
                win.bind("<Escape>", lambda e: choose("no"))
                self._show_dialog(win)
            except Exception:
                done.set()

        self.post(_prompt)
        while not done.wait(0.25):
            if self.cancel_requested:
                def _close():
                    w = result.get("win")
                    if w is not None and w.winfo_exists():
                        w.destroy()
                self.post(_close)
                break
        return result["choice"], result["remember"]

    # ------------------------------------------------------------------
    #  Update selected (+ Cancel)
    # ------------------------------------------------------------------
    def update_selected_async(self):
        if self.updating:
            self.cancel_requested = True
            self.btn_update.configure(text="Canceling…")
            self.btn_update.state(["disabled"])
            self.log("Cancel requested — stopping the current install…")
            threading.Thread(target=kill_process_tree, args=(self.current_proc,), daemon=True).start()
            return

        # Snapshot everything the worker needs while we're on the Tk thread.
        targets = []
        for it in self.all_items:
            if it in self.checked_items:
                vals = self.tree.item(it, "values")
                pid = vals[1]
                if pid:
                    targets.append({"id": pid, "name": vals[0] or pid, "current": (vals[2] or "").strip(),
                                    "item": it})
        if not targets:
            messagebox.showinfo("Nothing selected", "Tick the apps you want to update first.", parent=self.root)
            return

        auto_close = bool(self.close_running_var.get())
        if auto_close and not messagebox.askyesno(
                "Auto-close is on",
                "If an app blocks its own update it will be force-closed automatically.\n"
                "Any unsaved work in that app will be lost.\n\nContinue?", parent=self.root):
            return
        include_unknown = bool(self.include_unknown_var.get())
        self._remembered_close_choice = None
        self.skip_requested = False
        self._can_skip = len(targets) > 1

        for it in self.all_items:
            self.tree.set(it, "Result", "")
            self.tree.item(it, tags=[t for t in self.tree.item(it, "tags") if t in ("odd", "checked")])
        for t in targets:
            self._set_result(t["item"], "queued")

        self._disable_controls_for_update()
        self.btn_skip.state(["!disabled"] if self._can_skip else ["disabled"])
        self.cancel_requested = False
        self.current_proc = None
        self.log(f"Starting updates for {len(targets)} package(s)…")
        self.progress_start(len(targets))
        threading.Thread(target=self._update_worker, args=(targets, auto_close, include_unknown),
                         daemon=True).start()

    def _update_worker(self, targets, auto_close, include_unknown):
        results = {}
        percent_re = re.compile(r"(\d{1,3})\s*%")
        size_re = re.compile(r"([\d.]+)\s*(KB|MB|GB)\s*/\s*([\d.]+)\s*(KB|MB|GB)", re.I)
        unit = {"KB": 1_000, "MB": 1_000_000, "GB": 1_000_000_000}
        spinner_re = re.compile(r"^[\s\\/|\-\u2580-\u259F\u2500-\u257F]+$")
        progressbar_re = re.compile(r"[\u2580-\u259F\u2500-\u257F]")
        install_re = re.compile(r"starting package install|installing|extracting archive", re.I)

        def run_winget_once(t):
            pid, cur, it, name = t["id"], t["current"], t["item"], t["name"]
            cmd = ["winget", "upgrade", "--id", pid, "--exact", "--accept-package-agreements",
                   "--accept-source-agreements", "--disable-interactivity", "--silent"]
            if include_unknown or not cur or cur.lower() == "unknown" or cur.startswith("<"):
                cmd.append("--include-unknown")
            snap_before = self._snapshot_temp()
            # stderr is merged into stdout: a separate, unread stderr pipe can fill up
            # and freeze winget, which then looked like a "stall".
            proc = subprocess.Popen(cmd, shell=False, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                                    stdin=subprocess.DEVNULL, startupinfo=_hidden_startupinfo(),
                                    creationflags=_no_window_flags())
            self.current_proc = proc
            if self.cancel_requested:  # Cancel pressed while we were starting up
                kill_process_tree(proc)
            line_q = queue.Queue()

            def _reader():
                # winget redraws its progress bar in place with '\r', so split on both
                # '\r' and '\n'. Read in chunks (not 1 char at a time) to keep CPU low.
                dec = codecs.getincrementaldecoder("utf-8")("replace")
                buf = ""
                try:
                    while True:
                        b = proc.stdout.read1(4096) if hasattr(proc.stdout, "read1") else proc.stdout.read(1)
                        if not b:
                            break
                        buf += dec.decode(b)
                        parts = re.split(r"[\r\n]", buf)
                        buf = parts.pop()
                        for part in parts:
                            if part:
                                line_q.put(part)
                    buf += dec.decode(b"", final=True)
                except Exception:
                    pass
                finally:
                    if buf:
                        line_q.put(buf)
                    line_q.put(None)

            reader = threading.Thread(target=_reader, daemon=True)
            reader.start()

            last_pct = -1
            last_activity = time.monotonic()
            stalled = skip_now = False
            while True:
                if self.skip_requested:
                    self.skip_requested = False
                    skip_now = True
                    kill_process_tree(proc)
                    break
                if self.cancel_requested:
                    kill_process_tree(proc)
                    break
                if time.monotonic() - last_activity > STALL_TIMEOUT_SECONDS:
                    stalled = True
                    self.log(f"[Stalled] {pid}: no activity for {STALL_TIMEOUT_SECONDS}s — giving up on it.")
                    kill_process_tree(proc)
                    break
                try:
                    line = line_q.get(timeout=0.3)
                except queue.Empty:
                    continue
                if line is None:
                    break
                ln = line.rstrip()
                if not ln or spinner_re.match(ln):
                    continue
                last_activity = time.monotonic()

                pct = None
                m = None
                for m in percent_re.finditer(ln):
                    pass
                if m:
                    pct = int(m.group(1))
                else:
                    m2 = size_re.search(ln)
                    if m2:
                        try:
                            hv, hu, tv, tu = m2.groups()
                            tot = float(tv) * unit[tu.upper()]
                            if tot > 0:
                                pct = int(round(float(hv) * unit[hu.upper()] * 100 / tot))
                        except Exception:
                            pct = None
                is_progress = pct is not None or bool(progressbar_re.search(ln))
                if pct is not None and pct != last_pct:
                    last_pct = pct
                    self._pct_pending = (pct, name)  # coalesced by the UI pump
                if not is_progress:
                    self.log(ln)
                    if install_re.search(ln):
                        self.post(self.per_app_begin_install, name)

            if self.skip_requested:  # pressed in the final moment
                self.skip_requested = False
                skip_now = True
            reader.join(timeout=5)
            try:
                proc.wait(timeout=15)
            except Exception:
                kill_process_tree(proc)

            new_files = self._find_new_installer_files(snap_before, self._snapshot_temp())
            if new_files:
                lst = self.pkg_downloads.setdefault(pid, [])
                for p in new_files:
                    if p not in lst:
                        lst.append(p)
            if stalled:
                return "stalled", skip_now
            return proc.returncode, skip_now

        def close_running_for(t):
            pid, name = t["id"], t["name"]
            self.log(f"[Close] Looking for running processes of {name}…")
            loc = find_install_folder(pid, name)
            if not loc:
                self.log(f"[Close] Couldn't find where {name} is installed; can't close it safely.")
                return 0
            procs = running_processes_under(loc)
            if not procs:
                self.log(f"[Close] No running processes found in {loc}.")
                return 0
            names = ", ".join(sorted({os.path.basename(p) for _, p in procs}))
            self.log(f"[Close] Closing {name}: {names}")
            killed = kill_pids([p for p, _ in procs])
            time.sleep(1.5)  # let Windows release file locks
            self.log(f"[Close] Closed {killed} process(es).")
            return killed

        for t in targets:
            if self.cancel_requested:
                break
            pid, it, name = t["id"], t["item"], t["name"]
            self.post(self.per_app_reset, name)
            self.post(self._set_result, it, "updating")
            self.post(self.tree.see, it)
            self.log(f"Updating {name}  ({pid})")
            status = "failed"
            try:
                rc, skip_now = run_winget_once(t)
                if (rc != "stalled" and not skip_now and not self.cancel_requested
                        and _normalize_winget_code(rc) in WINGET_IN_USE_SET):
                    do_close = False
                    if auto_close or self._remembered_close_choice == "yes":
                        do_close = True
                    elif self._remembered_close_choice == "no":
                        skip_now = True
                        self.log(f"[Skip] {name}: left running (remembered choice).")
                    else:
                        choice, remember = self.ask_close_app(name)
                        if remember:
                            self._remembered_close_choice = choice
                        if choice == "yes":
                            do_close = True
                        else:
                            skip_now = True
                            self.log(f"[Skip] {name}: left running.")
                    if do_close and not self.cancel_requested and close_running_for(t) > 0:
                        self.log(f"Retrying {name}…")
                        self.post(self.per_app_reset, name)
                        rc, skip_now = run_winget_once(t)

                if rc == "stalled":
                    status = "stalled"
                elif self.cancel_requested:
                    status = "canceled"
                else:
                    status = classify_winget_result(rc, skip_now=skip_now)
                    if status == "inuse":
                        self.log(f"[In use] {name}: still blocked (it may have restarted itself).")
                    elif status == "failed":
                        code = _normalize_winget_code(rc)
                        self.log(f"✕ {name} failed (exit code 0x{code:08X})." if isinstance(code, int)
                                 else f"✕ {name} failed.")
            except Exception as ex:
                status = "canceled" if self.cancel_requested else "failed"
                if status == "failed":
                    self.log(f"Error: {ex}")
            finally:
                results[pid] = RESULT_BUCKET.get(status, "failed")
                if status in ("success", "reboot"):
                    self.log(f"✓ {name} updated" + (" — restart Windows to finish." if status == "reboot" else "."))
                self.post(self._finish_item, it, name, status)
                if self._can_skip and not self.cancel_requested:
                    self.post(lambda: self.btn_skip.state(["!disabled"]))

        self.post(self._batch_done, targets, results)

    def _finish_item(self, it, name, status):
        self._set_result(it, status)
        if status in ("success", "reboot"):
            self._set_checked(it, False)
            self._set_result(it, status)  # re-apply tags after uncheck
        if not self.cancel_requested:
            self.per_app_finish(name)
        self.progress_step(1)
        self.update_counter()

    def _batch_done(self, targets, results):
        canceled = self.cancel_requested
        if canceled:
            for t in targets:
                if self.tree.exists(t["item"]) and self.tree.set(t["item"], "Result") in ("", STATUS["queued"][0],
                                                                                          STATUS["updating"][0]):
                    self._set_result(t["item"], "canceled")
            self.log("■ Canceled by user.")
        vals = list(results.values())
        ok, fail = vals.count("success"), vals.count("failed")
        skip, canc, stall = vals.count("skipped"), vals.count("canceled"), vals.count("stalled")
        summary = f"Summary → {ok} updated · {fail} failed · {skip} skipped · {canc} canceled"
        if stall:
            summary += f" · {stall} stalled"
        self.log(summary)
        if fail == 0 and stall == 0 and not canceled:
            play_success_sound()
        self.cancel_requested = False
        self.skip_requested = False
        self.current_proc = None
        self.btn_skip.state(["disabled"])
        self._enable_controls_after_update()
        short = f"{ok} updated" + (f" · {fail + stall} failed" if fail + stall else "") + (
            f" · {skip} skipped" if skip else "")
        self.progress_finish(canceled=canceled, summary=short)

    def skip_current(self):
        if not self.updating:
            return
        self.skip_requested = True
        self.btn_skip.state(["disabled"])
        self.log("[Skip] Skipping the current app…")


# ===================== main =====================
if __name__ == "__main__":
    enable_dpi_awareness()
    root = tk.Tk()
    app = WingetUpdaterUI(root)
    root.mainloop()
