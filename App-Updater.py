import json
import re
import subprocess
import threading
import tkinter as tk
from tkinter import messagebox, ttk
import tkinter.font as tkfont
import sys
import os
import ctypes
import winsound
import webbrowser
from io import BytesIO
from typing import Optional
from PIL import Image, ImageDraw, ImageFont

# ====================== App Constants ======================
APP_NAME_VERSION = "Windows App Updater v1.1"

# ====================== PyInstaller resource helper ======================
def resource_path(relative_path: str) -> str:
    """Return path to resource whether running from source or PyInstaller EXE."""
    if hasattr(sys, "_MEIPASS"):
        return os.path.join(sys._MEIPASS, relative_path)  # type: ignore[attr-defined]
    return os.path.join(os.path.abspath("."), relative_path)

# ====================== Hide child console windows ======================
CREATE_NO_WINDOW = 0x08000000

def _hidden_startupinfo() -> subprocess.STARTUPINFO:  # type: ignore[name-defined]
    si = subprocess.STARTUPINFO()  # type: ignore[attr-defined]
    si.dwFlags |= subprocess.STARTF_USESHOWWINDOW      # type: ignore[attr-defined]
    si.wShowWindow = 0  # SW_HIDE
    return si

# ====================== Elevation helpers ======================
def is_admin() -> bool:
    try:
        return bool(ctypes.windll.shell32.IsUserAnAdmin())
    except Exception:
        return False

def relaunch_as_admin():
    """Relaunch the program with admin rights."""
    if is_admin():
        return
    if getattr(sys, "frozen", False):
        app = sys.executable
        params = " ".join(f'"{a}"' for a in sys.argv[1:])
        ctypes.windll.shell32.ShellExecuteW(None, "runas", app, params, None, 1)
    else:
        app = sys.executable
        script = os.path.abspath(sys.argv[0])
        params = " ".join([f'"{script}"'] + [f'"{a}"' for a in sys.argv[1:]])
        ctypes.windll.shell32.ShellExecuteW(None, "runas", app, params, None, 1)
    sys.exit(0)

# ====================== Tooltip helper ======================
class ToolTip:
    def __init__(self, widget, text):
        self.widget = widget
        self.text = text
        self.tip_window = None
        widget.bind("<Enter>", self.show_tip)
        widget.bind("<Leave>", self.hide_tip)

    def show_tip(self, _=None):
        if self.tip_window or not self.text:
            return
        x = self.widget.winfo_rootx() + 25
        y = self.widget.winfo_rooty() + self.widget.winfo_height() + 10
        self.tip_window = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(True)
        tw.wm_geometry(f"+{x}+{y}")
        label = tk.Label(
            tw, text=self.text, justify="left",
            background="#ffffe0", relief="solid", borderwidth=1,
            font=("Segoe UI", 9)
        )
        label.pack(ipadx=6, ipady=2)

    def hide_tip(self, _=None):
        if self.tip_window:
            self.tip_window.destroy()
            self.tip_window = None

# ====================== Icon & sound helpers ======================
def set_app_icon(root: tk.Tk) -> Optional[str]:
    ico = resource_path("windows-updater.ico")
    if os.path.exists(ico):
        try:
            root.iconbitmap(ico)
            return ico
        except Exception:
            pass
    return None

def apply_icon_to_toplevel(tlv: tk.Toplevel, icon_path: Optional[str]):
    if not icon_path:
        return
    try:
        tlv.iconbitmap(icon_path)
    except Exception:
        pass

def load_flag_image() -> Optional[tk.PhotoImage]:
    png = resource_path("kuwait.png")
    if os.path.exists(png):
        try:
            return tk.PhotoImage(file=png)
        except Exception:
            pass
    ico = resource_path("kuwait.ico")
    if os.path.exists(ico):
        try:
            im = Image.open(ico)
            if hasattr(im, "n_frames"):
                im.seek(im.n_frames - 1)
            im = im.convert("RGBA")
            max_h = 18
            if im.height > max_h:
                ratio = max_h / float(im.height)
                im = im.resize((max(16, int(im.width * ratio)), max_h), Image.LANCZOS)
            bio = BytesIO()
            im.save(bio, format="PNG")
            bio.seek(0)
            return tk.PhotoImage(data=bio.read())
        except Exception:
            return None
    return None

def play_success_sound():
    wav = resource_path("success.wav")
    if os.path.exists(wav):
        try:
            winsound.PlaySound(wav, winsound.SND_FILENAME | winsound.SND_ASYNC)
            return
        except Exception:
            pass
    try:
        winsound.MessageBeep(winsound.MB_ICONASTERISK)
    except Exception:
        pass

# ====================== Donate image (drawn at runtime) ======================
def make_donate_image(width=160, height=44):
    """Return a glossy orange rounded pill as a Tk PhotoImage (no text)."""
    radius = height // 2
    top = (255, 187, 71)
    mid = (247, 162, 28)
    bot = (225, 140, 22)

    im = Image.new("RGBA", (width, height), (0, 0, 0, 0))
    dr = ImageDraw.Draw(im)

    # Vertical gradient
    for y in range(height):
        if y < height * 0.6:
            t = y / (height * 0.6)
            col = tuple(int(top[i] * (1 - t) + mid[i] * t) for i in range(3)) + (255,)
        else:
            t = (y - height * 0.6) / (height * 0.4)
            col = tuple(int(mid[i] * (1 - t) + bot[i] * t) for i in range(3)) + (255,)
        dr.line([(0, y), (width, y)], fill=col)

    # Rounded mask
    mask = Image.new("L", (width, height), 0)
    ImageDraw.Draw(mask).rounded_rectangle([0, 0, width - 1, height - 1], radius=radius, fill=255)
    im.putalpha(mask)

    # Gloss highlight
    highlight = Image.new("RGBA", (width, height), (255, 255, 255, 0))
    ImageDraw.Draw(highlight).rounded_rectangle(
        [2, 2, width - 3, height // 2], radius=radius - 2, fill=(255, 255, 255, 70)
    )
    im = Image.alpha_composite(im, highlight)

    # Border
    ImageDraw.Draw(im).rounded_rectangle(
        [0.5, 0.5, width - 1.5, height - 1.5], radius=radius, outline=(200, 120, 20, 255), width=2
    )

    bio = BytesIO()
    im.save(bio, format="PNG")
    bio.seek(0)
    return tk.PhotoImage(data=bio.read())


    # Measure text (fallbacks for older Pillow)
    try:
        tw, th = dr.textsize(text, font=font)          # widely supported
    except Exception:
        try:
            bbox = font.getbbox(text)                  # (l,t,r,b)
            tw, th = bbox[2] - bbox[0], bbox[3] - bbox[1]
        except Exception:
            tw, th = (len(text) * font_size * 0.6, font_size)

    tx = int((width - tw) / 2)
    ty = int((height - th) / 2) - 1

    # Draw text with a subtle shadow to pop on orange
    shadow = (0, 0, 0, 80)
    dr.text((tx + 1, ty + 1), text, font=font, fill=shadow)
    dr.text((tx, ty), text, font=font, fill=(15, 52, 98, 255))  # navy

    bio = BytesIO()
    im.save(bio, format="PNG")
    bio.seek(0)
    return tk.PhotoImage(data=bio.read())


# ====================== winget helpers ======================
def run(cmd):
    env = os.environ.copy()
    env["DOTNET_CLI_UI_LANGUAGE"] = "en"
    p = subprocess.run(
        cmd, capture_output=True, text=True, shell=False,
        encoding="utf-8", errors="replace", env=env,
        startupinfo=_hidden_startupinfo(), creationflags=CREATE_NO_WINDOW
    )
    return p.returncode, p.stdout.strip(), p.stderr.strip()

def try_json_parsers(include_unknown: bool):
    base = ["--accept-source-agreements", "--disable-interactivity", "--output", "json"]
    flag = ["--include-unknown"] if include_unknown else []
    attempts = [
        ["winget", "upgrade", *flag, *base],
        ["winget", "list", "--upgrade-available", *base],
        ["winget", "list", "--upgrades", *base],
    ]
    last_err = ""
    for cmd in attempts:
        code, out, err = run(cmd)
        if code == 0 and out:
            try:
                data = json.loads(out)
                return normalize_winget_json(data)
            except Exception as e:
                last_err = f"{err or ''}\nJSON parse error: {e}"
        else:
            last_err = err or "winget returned a non-zero exit code."
    raise RuntimeError(last_err.strip() or "Failed to get JSON from winget.")

def normalize_winget_json(data):
    items = []
    if isinstance(data, list):
        iterable = data
    elif isinstance(data, dict):
        if "Sources" in data:
            iterable = []
            for src in data.get("Sources", []):
                iterable.extend(src.get("Packages", []))
        else:
            iterable = data.get("Packages", [])
    else:
        iterable = []

    for it in iterable:
        name      = it.get("PackageName") or it.get("Name") or ""
        pkg_id    = it.get("PackageIdentifier") or it.get("Id") or ""
        available = it.get("AvailableVersion") or it.get("Available") or ""
        current   = it.get("Version") or it.get("InstalledVersion") or ""
        if name and pkg_id and available:
            items.append({"name": name, "id": pkg_id, "available": available, "current": current})
    return items

def parse_table_upgrade_output(text):
    lines = [ln for ln in text.splitlines() if ln.strip()]
    header_idx = -1
    for i, ln in enumerate(lines):
        if re.search(r"\bName\b", ln) and re.search(r"\bId\b", ln) and re.search(r"\bAvailable\b", ln):
            header_idx = i
            break
    if header_idx < 0 or header_idx + 1 >= len(lines):
        return []

    start = header_idx + 1
    if start < len(lines) and re.match(r"^[\s\-]+$", lines[start].replace(" ", "")):
        start += 1

    items = []
    for ln in lines[start:]:
        if "No applicable updates" in ln:
            return []
        parts = re.split(r"\s{2,}", ln.rstrip())
        if len(parts) < 4:
            continue
        if len(parts) >= 5:
            name, pkg_id = parts[0], parts[1]
            current   = parts[2] if len(parts) > 2 else ""
            available = parts[3] if len(parts) > 3 else ""
        else:
            name, pkg_id = parts[0], parts[1]
            current, available = "", parts[2]
        if name and pkg_id and available and not name.startswith("-"):
            items.append({"name": name, "id": pkg_id, "current": current, "available": available})
    return items

def get_winget_upgrades(include_unknown: bool):
    code, _, _ = run(["winget", "--version"])
    if code != 0:
        raise RuntimeError("winget not found. Install the App Installer from Microsoft Store.")
    try:
        return try_json_parsers(include_unknown)
    except Exception as e_json:
        cmd = ["winget", "upgrade", "--accept-source-agreements", "--disable-interactivity"]
        if include_unknown:
            cmd.insert(2, "--include-unknown")
        code, out, err = run(cmd)
        if code != 0:
            raise RuntimeError((err or str(e_json)).strip())
        parsed = parse_table_upgrade_output(out)
        if parsed:
            return parsed
        raise RuntimeError(str(e_json))

# ====================== Checkbox images (drawn at runtime) ======================
def make_checkbox_images(size: int = 16):
    """Create simple checkbox PNGs at runtime (no external files)."""
    unchecked = tk.PhotoImage(width=size, height=size)
    unchecked.put("white", to=(0, 0, size, size))
    border = "gray20"
    unchecked.put(border, to=(0, 0, size, 1))
    unchecked.put(border, to=(0, size - 1, size, size))
    unchecked.put(border, to=(0, 0, 1, size))
    unchecked.put(border, to=(size - 1, 0, size, size))

    checked = tk.PhotoImage(width=size, height=size)
    checked.tk.call(checked, "copy", unchecked)
    mark = "#2e7d32"
    pts = [(3, size // 2), (4, size // 2 + 1), (5, size // 2 + 2),
           (6, size // 2 + 3), (7, size // 2 + 2),
           (8, size // 2 + 1), (9, size // 2), (10, size // 2 - 1)]
    for (x, y) in pts:
        checked.put(mark, to=(x, y, x + 1, y + 1))
        checked.put(mark, to=(x, y - 1, x + 1, y))
    return unchecked, checked

# ====================== UI Class ======================
class WingetUpdaterUI:
    def __init__(self, root):
        self.root = root
        self.root.title(APP_NAME_VERSION)
        self.root.geometry("1280x900")
        self.root.minsize(1180, 830)

        self.updating = False
        self.cancel_requested = False
        self.current_proc = None
        self.loading_win = None
        self.window_icon_path = set_app_icon(self.root)

        # checkbox images and state store
        self.img_unchecked, self.img_checked = make_checkbox_images(16)
        self.checked_items = set()   # <- single source of truth for row selection

        # ===== Header =====
        header = ttk.Frame(self.root); header.pack(fill="x", pady=(10, 0))
        ttk.Label(header, text=APP_NAME_VERSION, font=("Segoe UI", 18, "bold")).pack(side="left", padx=12)

        right = ttk.Frame(header); right.pack(side="right", padx=12)
        self.btn_admin = ttk.Button(right, text="Run as Admin", command=self.run_as_admin)
        self.btn_admin.pack(side="right")
        if is_admin():
            self.btn_admin.config(text="Running as Admin", state="disabled")
        else:
            ToolTip(self.btn_admin, "Run the app as admin if you like to install all apps silently")

        # Donate button (image) will be created later in signature row

        # ===== Top controls =====
        top = ttk.Frame(self.root); top.pack(fill="x", padx=12, pady=6)

        self.btn_check = ttk.Button(top, text="Check for Updates", command=self.check_for_updates_async)
        self.btn_check.pack(side="left")

        self.include_unknown_var = tk.BooleanVar(value=False)
        self.chk_unknown = ttk.Checkbutton(top, text="Include unknown apps", variable=self.include_unknown_var)
        self.chk_unknown.pack(side="left", padx=(10, 0))

        ttk.Button(top, text="Select All",  command=self.select_all).pack(side="left", padx=(10, 0))
        ttk.Button(top, text="Select None", command=self.select_none).pack(side="left", padx=(6, 0))

        self.btn_update = ttk.Button(top, text="Update Selected", command=self.update_selected_async)
        self.btn_update.pack(side="right")

        # Counter
        self.counter_var = tk.StringVar(value="0 apps found • 0 selected")
        ttk.Label(self.root, textvariable=self.counter_var).pack(anchor="w", padx=12)

        # ===== Tree with both scrollbars =====
        tree_wrap = ttk.Frame(self.root); tree_wrap.pack(fill="both", expand=True, padx=12, pady=(8, 8))

        cols = ("Name", "Id", "Current", "Available")
        self.fixed_cols = cols
        self.tree = ttk.Treeview(tree_wrap, columns=cols, show="tree headings", height=22, selectmode="none")

        # Headings
        self.tree.heading("#0",       text="Select",   anchor="center")
        self.tree.heading("Name",     text="Name",     anchor="w")
        self.tree.heading("Id",       text="Id",       anchor="w")
        self.tree.heading("Current",  text="Current",  anchor="center")
        self.tree.heading("Available",text="Available",anchor="center")

        # Column widths
        # --- Select column: fit header text, centered checkboxes, locked resize ---
        font = tkfont.nametofont("TkDefaultFont")
        text_width = font.measure("Select") + 20  # padding so it doesn’t feel cramped
        self.tree.column("#0",
            width=text_width,
            minwidth=text_width,
            anchor="center",
            stretch=False
        )

        # Other columns start with sensible widths; users can resize them or double-click to auto-fit
        self.tree.column("Name",      width=520, minwidth=140, anchor="w",      stretch=False)
        self.tree.column("Id",        width=560, minwidth=200, anchor="w",      stretch=False)
        self.tree.column("Current",   width=110, minwidth=70,  anchor="center", stretch=False)
        self.tree.column("Available", width=110, minwidth=70,  anchor="center", stretch=False)

        # Lock order of data columns
        self.tree["displaycolumns"] = self.fixed_cols

        # Scrollbars
        ysb = ttk.Scrollbar(tree_wrap, orient="vertical",   command=self.tree.yview)
        xsb = ttk.Scrollbar(tree_wrap, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscroll=ysb.set, xscroll=xsb.set)

        self.tree.grid(row=0, column=0, sticky="nsew")
        ysb.grid(row=0, column=1, sticky="ns")
        xsb.grid(row=1, column=0, sticky="ew")
        tree_wrap.rowconfigure(0, weight=1)
        tree_wrap.columnconfigure(0, weight=1)

        # Mouse handling — toggle reliably on col #0; block header reordering; block resizing for #0; auto-fit on double click
        self._block_header_drag = False
        self._block_resize_select = False

        self.tree.bind("<Button-1>", self._on_mouse_down, add="+")
        self.tree.bind("<B1-Motion>", self._on_mouse_drag, add="+")
        self.tree.bind("<ButtonRelease-1>", self._on_mouse_up, add="+")
        self.tree.bind("<Double-Button-1>", self._on_double_click_header, add="+")

        # ===== Progress bar =====
        pb_wrap = ttk.Frame(self.root); pb_wrap.pack(fill="x", padx=12, pady=(0, 4))
        self.pb_label = ttk.Label(pb_wrap, text="Idle"); self.pb_label.pack(side="left")
        self.pb = ttk.Progressbar(pb_wrap, orient="horizontal", mode="determinate")
        self.pb.pack(fill="x", expand=True, padx=10)

        # ===== Signature (before log) =====
        sig_frame = ttk.Frame(self.root); sig_frame.pack(fill="x", padx=12, pady=(4, 0))
        ttk.Label(sig_frame, text="").pack(side="left", expand=True)

        # Donate button image (keep a reference on self!)
        self.donate_img = make_donate_image(width=160, height=44)   # keep a reference on self
        self.btn_donate = tk.Button(
            sig_frame,
            image=self.donate_img,
            text="Donate",
            compound="center",          # overlay text on the image
            font=("Segoe UI", 11, "bold"),
            fg="#0f3462",               # navy text color
            activeforeground="#0f3462",
            bd=0,
            highlightthickness=0,
            cursor="hand2",
            relief="flat",
            command=self.open_donate_link,
        )
        self.btn_donate.pack(side="right", padx=(8, 0))
        ToolTip(self.btn_donate, "Support development with a small donation")

        # Flag + name to the left of the donate button (leave this as-is)
        self.flag_img = load_flag_image()
        if self.flag_img:
            tk.Label(sig_frame, image=self.flag_img).pack(side="right", padx=(8, 6))
        ttk.Label(sig_frame, text="Made by BoYaqoub - ilukezippo@gmail.com", font=("Segoe UI", 9)).pack(side="right")

        # ===== Log =====
        ttk.Label(self.root, text="Update Log:", font=("Segoe UI", 10, "bold")).pack(anchor="w", padx=12, pady=(6, 2))

        log_wrap = ttk.Frame(self.root); log_wrap.pack(fill="both", expand=False, padx=12, pady=(0, 10))
        self.log_box = tk.Text(log_wrap, height=14, wrap="none", font=("Consolas", 10))
        log_ysb = ttk.Scrollbar(log_wrap, orient="vertical",   command=self.log_box.yview)
        log_xsb = ttk.Scrollbar(log_wrap, orient="horizontal", command=self.log_box.xview)
        self.log_box.configure(yscrollcommand=log_ysb.set, xscrollcommand=log_xsb.set)
        self.log_box.grid(row=0, column=0, sticky="nsew")
        log_ysb.grid(row=0, column=1, sticky="ns")
        log_xsb.grid(row=1, column=0, sticky="ew")
        log_wrap.rowconfigure(0, weight=1)
        log_wrap.columnconfigure(0, weight=1)

        self.root.after(0, self.center_on_screen)

    # ----- mouse handlers: block header reordering; lock select column resize; toggle on #0
    def _on_mouse_down(self, event):
        region = self.tree.identify("region", event.x, event.y)
        if region == "heading":
            self._block_header_drag = True
            return "break"
        if region == "separator":
            # if this separator is the one to the RIGHT of #0, prevent resizing
            col_left = self.tree.identify_column(event.x - 1)  # column to the left of separator
            if col_left == "#0":
                self._block_resize_select = True
                return "break"
            self._block_resize_select = False
            return  # allow resizing for other columns

        col = self.tree.identify_column(event.x)
        if col == "#0":
            item = self.tree.identify_row(event.y)
            if item:
                if item in self.checked_items:
                    self.checked_items.remove(item)
                    self.tree.item(item, image=self.img_unchecked)
                else:
                    self.checked_items.add(item)
                    self.tree.item(item, image=self.img_checked)
                self.update_counter()
                return "break"

        self._block_header_drag = False
        self._block_resize_select = False

    def _on_mouse_drag(self, event):
        if self._block_header_drag or self._block_resize_select:
            return "break"

    def _on_mouse_up(self, event):
        self._block_header_drag = False
        self._block_resize_select = False
        try:
            if tuple(self.tree["displaycolumns"]) != tuple(self.fixed_cols):
                self.tree["displaycolumns"] = self.fixed_cols
        except Exception:
            pass

    # ----- Excel-style auto-fit on double-click separator -----
    def _on_double_click_header(self, event):
        if self.tree.identify("region", event.x, event.y) != "separator":
            return
        col_left = self.tree.identify_column(event.x - 1)
        if not col_left or col_left == "#0":
            return  # don't auto-fit Select column
        self.autofit_column(col_left)

    def autofit_column(self, col_id: str):
        """Resize column to fit its widest content + padding (like Excel)."""
        heading = self.tree.heading(col_id, "text") or ""
        font = tkfont.nametofont("TkDefaultFont")
        pad = 24  # pixels

        max_px = font.measure(heading)
        for item in self.tree.get_children(""):
            val = self.tree.set(item, col_id) or ""
            px = font.measure(val)
            if px > max_px:
                max_px = px

        minw = int(self.tree.column(col_id, "minwidth") or 20)
        new_w = max(minw, max_px + pad)
        self.tree.column(col_id, width=new_w)

    # ----- window centering -----
    def center_on_screen(self):
        self.root.update_idletasks()
        w, h = self.root.winfo_width(), self.root.winfo_height()
        sw, sh = self.root.winfo_screenwidth(), self.root.winfo_screenheight()
        x, y = (sw - w) // 2, (sh - h) // 2
        self.root.geometry(f"{w}x{h}+{x}+{y}")

    # ----- admin button handler -----
    def run_as_admin(self):
        relaunch_as_admin()

    # ----- donation link -----
    def open_donate_link(self):
        # Replace with your preferred donation link:
        webbrowser.open("https://buymeacoffee.com/ilukezippo")

    # ====================== Loading screen ======================
    def show_loading(self, text="Loading..."):
        if self.loading_win:
            return
        self.loading_win = tk.Toplevel(self.root)
        self.loading_win.title("")
        self.loading_win.transient(self.root)
        self.loading_win.grab_set()
        self.loading_win.resizable(False, False)
        apply_icon_to_toplevel(self.loading_win, self.window_icon_path)

        ttk.Label(self.loading_win, text=text, font=("Segoe UI", 12, "bold")).pack(padx=20, pady=(16, 8))
        pb = ttk.Progressbar(self.loading_win, mode="indeterminate", length=280)
        pb.pack(padx=20, pady=(0, 16))
        pb.start(10)
        self.loading_win.update_idletasks()
        x = self.root.winfo_x() + (self.root.winfo_width() - self.loading_win.winfo_width()) // 2
        y = self.root.winfo_y() + (self.root.winfo_height() - self.loading_win.winfo_height()) // 2
        self.loading_win.geometry(f"+{x}+{y}")
        self.loading_win.protocol("WM_DELETE_WINDOW", lambda: None)

    def hide_loading(self):
        if self.loading_win:
            try:
                self.loading_win.grab_release()
            except Exception:
                pass
            self.loading_win.destroy()
            self.loading_win = None

    # ====================== Progress helpers ======================
    def progress_start(self, phase: str, total: int):
        self.pb_phase = phase
        self.pb_total = max(0, int(total))
        self.pb_value = 0
        self.pb.configure(maximum=max(self.pb_total, 1), value=0, mode="determinate")
        self.pb_label.configure(text=f"{phase}: 0/{self.pb_total}")
        self.root.update_idletasks()

    def progress_step(self, inc: int = 1):
        if self.pb_total <= 0:
            return
        self.pb_value = min(self.pb_total, self.pb_value + inc)
        self.pb.configure(value=self.pb_value)
        self.pb_label.configure(text=f"{self.pb_phase}: {self.pb_value}/{self.pb_total}")
        self.root.update_idletasks()

    def progress_finish(self, canceled=False):
        if getattr(self, "pb_total", 0) > 0:
            self.pb.configure(value=self.pb_total)
            suffix = " (canceled)" if canceled else " (done)"
            self.pb_label.configure(text=f"{self.pb_phase}: {self.pb_total}/{self.pb_total}{suffix}")
        else:
            self.pb_label.configure(text="Idle")
        self.root.update_idletasks()

    # ====================== Selection helpers ======================
    def _iter_items(self):
        return self.tree.get_children("")

    def select_all(self):
        for item in self._iter_items():
            self.checked_items.add(item)
            self.tree.item(item, image=self.img_checked)
        self.update_counter()

    def select_none(self):
        for item in self._iter_items():
            self.checked_items.discard(item)
            self.tree.item(item, image=self.img_unchecked)
        self.update_counter()

    def update_counter(self):
        total = len(self.tree.get_children(""))
        selected = len(self.checked_items)
        self.counter_var.set(f"{total} apps found • {selected} selected")

    def clear_tree(self):
        self.checked_items.clear()
        for i in self._iter_items():
            self.tree.delete(i)

    # ====================== Check for updates (async with loading) ======================
    def check_for_updates_async(self):
        include_unknown = bool(self.include_unknown_var.get())
        self.btn_check.config(state="disabled")
        self.show_loading("Checking for updates...")

        def worker():
            try:
                pkgs = get_winget_upgrades(include_unknown=include_unknown)
            except Exception as e:
                self.root.after(0, lambda: (
                    self.hide_loading(),
                    self.btn_check.config(state="normal"),
                    self.counter_var.set("0 apps found • 0 selected"),
                    messagebox.showerror("winget error", f"Failed to query updates:\n{e}"),
                    self.log(f"[winget] {e}")
                ))
                return
            self.root.after(0, lambda: self.populate_tree(pkgs))

        threading.Thread(target=worker, daemon=True).start()

    def populate_tree(self, pkgs):
        self.hide_loading()
        self.clear_tree()
        self.btn_check.config(state="normal")
        if not pkgs:
            self.counter_var.set("0 apps found • 0 selected")
            self.log("No apps need updating.")
            return

        # Keep order as returned by winget (do NOT sort alphabetically)
        for p in pkgs:
            item = self.tree.insert(
                "", "end",
                text="",              # #0 has no text
                image=self.img_unchecked,  # centered via anchor and fixed width
                values=(p["name"], p["id"], p.get("current", ""), p.get("available", "")),
            )
            self.checked_items.discard(item)
        self.update_counter()

    # ====================== Update selected (async + Cancel) ======================
    def update_selected_async(self):
        if getattr(self, "updating", False):
            self.cancel_requested = True
            self.btn_update.config(text="Cancelling...", state="disabled")
            proc = self.current_proc
            if proc and proc.poll() is None:
                try:
                    proc.terminate()
                except Exception:
                    pass
            return

        # Gather selection from our state set
        targets = []
        for item in list(self.checked_items):
            pkg_id  = self.tree.set(item, "Id")
            current = (self.tree.set(item, "Current") or "").strip()
            if pkg_id:
                targets.append((pkg_id, current))

        if not targets:
            messagebox.showinfo("No Selection", "No apps selected for update.")
            return

        self.updating = True
        self.cancel_requested = False
        self.current_proc = None
        self.btn_check.config(state="disabled")
        self.btn_update.config(text="Cancel", state="normal")
        self.log(f"Starting updates for {len(targets)} package(s)...")

        self.progress_start("Updating", len(targets))

        def worker():
            for pkg_id, current in targets:
                if self.cancel_requested:
                    break
                self.root.after(0, lambda pid=pkg_id: self.log(f"Updating {pid} ..."))
                try:
                    cmd = [
                        "winget", "upgrade", "--id", pkg_id,
                        "--accept-package-agreements", "--accept-source-agreements",
                        "--disable-interactivity", "-h"
                    ]
                    if self.include_unknown_var.get() or (not current) or (current.lower() == "unknown"):
                        cmd.insert(2, "--include-unknown")

                    self.current_proc = subprocess.Popen(
                        cmd,
                        shell=False, text=True,
                        stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                        startupinfo=_hidden_startupinfo(),
                        creationflags=CREATE_NO_WINDOW
                    )
                    spinner_re = re.compile(r"^[\s\\/\|\-\r]+$")
                    while True:
                        if self.cancel_requested and self.current_proc and self.current_proc.poll() is None:
                            try:
                                self.current_proc.terminate()
                            except Exception:
                                pass
                        line = self.current_proc.stdout.readline()
                        if not line:
                            break
                        ln = line.rstrip()
                        if ln and not spinner_re.match(ln):
                            self.root.after(0, lambda s=ln: self.log(s))
                    _, err = self.current_proc.communicate()
                    if err:
                        self.root.after(0, lambda e=err: self.log(e.strip()))
                except Exception as ex:
                    self.root.after(0, lambda ex=ex: self.log(f"Error: {ex}"))
                finally:
                    self.root.after(0, lambda pid=pkg_id: self.log(f"✔ Finished {pid}"))
                    self.root.after(0, lambda: self.progress_step(1))

            def done():
                canceled = self.cancel_requested
                if canceled:
                    self.log("Cancelled.")
                else:
                    self.log("All selected updates completed.")
                    play_success_sound()
                self.updating = False
                self.cancel_requested = False
                self.current_proc = None
                self.btn_check.config(state="normal")
                self.btn_update.config(text="Update Selected", state="normal")
                self.progress_finish(canceled=canceled)
            self.root.after(0, done)

        threading.Thread(target=worker, daemon=True).start()

    # ====================== Logging ======================
    def log(self, text: str):
        self.log_box.insert(tk.END, text + "\n")
        self.log_box.see(tk.END)
        self.root.update_idletasks()

# ====================== main ======================
if __name__ == "__main__":
    root = tk.Tk()
    app = WingetUpdaterUI(root)
    root.mainloop()
