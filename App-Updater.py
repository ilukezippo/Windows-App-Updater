import json, re, subprocess, threading, tkinter as tk
from tkinter import ttk, messagebox, filedialog
import tkinter.font as tkfont
import sys, os, ctypes, winsound, webbrowser, tempfile, shutil
from io import BytesIO
from PIL import Image, ImageDraw
import urllib.request



APP_VERSION_ONLY = "v2.2.1"
APP_NAME_VERSION = f"Windows App Updater {APP_VERSION_ONLY}"
DATE_APP = "2025/10/01"
GITHUB_RELEASES_PAGE = "https://github.com/ilukezippo/Windows-App-Updater/releases"
GITHUB_API_LATEST = "https://api.github.com/repos/ilukezippo/Windows-App-Updater/releases/latest"
DONATE_PAGE = "https://buymeacoffee.com/ilukezippo"

WIN_W = 950
WIN_H_FULL = 650
WIN_H_COMPACT = 500
LIST_PIXELS = 240  # << apps list height AND log height

CREATE_NO_WINDOW = 0x08000000


def _hidden_startupinfo():
    si = subprocess.STARTUPINFO()
    si.dwFlags |= subprocess.STARTF_USESHOWWINDOW
    si.wShowWindow = 0
    return si


def resource_path(p):
    return os.path.join(getattr(sys, "_MEIPASS", os.path.abspath(".")), p)


def is_admin():
    try:
        return bool(ctypes.windll.shell32.IsUserAnAdmin())
    except Exception:
        return False


def relaunch_as_admin():
    if is_admin(): return
    if getattr(sys, "frozen", False):
        app = sys.executable
        params = subprocess.list2cmdline(sys.argv[1:])
        ctypes.windll.shell32.ShellExecuteW(None, "runas", app, params, None, 1)
    else:
        app = sys.executable
        script = os.path.abspath(sys.argv[0])
        params = subprocess.list2cmdline([script] + sys.argv[1:])
        ctypes.windll.shell32.ShellExecuteW(None, "runas", app, params, None, 1)
    sys.exit(0)


class ToolTip:
    def __init__(self, w, text):
        self.w = w;
        self.text = text;
        self.tip = None
        w.bind("<Enter>", self.show);
        w.bind("<Leave>", self.hide)

    def show(self, _=None):
        if self.tip or not self.text: return
        x = self.w.winfo_rootx() + 25;
        y = self.w.winfo_rooty() + self.w.winfo_height() + 10
        self.tip = tk.Toplevel(self.w);
        self.tip.wm_overrideredirect(True);
        self.tip.wm_geometry(f"+{x}+{y}")
        tk.Label(self.tip, text=self.text, justify="left", background="#ffffe0",
                 relief="solid", borderwidth=1, font=("Segoe UI", 9)).pack(ipadx=6, ipady=2)

    def hide(self, _=None):
        if self.tip: self.tip.destroy(); self.tip = None


def set_app_icon(root):
    ico = resource_path("windows-updater.ico")
    if os.path.exists(ico):
        try:
            root.iconbitmap(ico); return ico
        except Exception:
            pass
    return None


def apply_icon_to_tlv(tlv, icon):
    if icon:
        try:
            tlv.iconbitmap(icon)
        except Exception:
            pass


def load_flag_image():
    png = resource_path("kuwait.png")
    if os.path.exists(png):
        try:
            return tk.PhotoImage(file=png)
        except Exception:
            pass
    return None


def make_donate_image(w=160, h=44):
    r = h // 2;
    top = (255, 187, 71);
    mid = (247, 162, 28);
    bot = (225, 140, 22)
    im = Image.new("RGBA", (w, h), (0, 0, 0, 0));
    dr = ImageDraw.Draw(im)
    for y in range(h):
        t = y / (h * 0.6) if y < h * 0.6 else (y - h * 0.6) / (h * 0.4)
        c = tuple(int((top[i] if y < h * 0.6 else mid[i]) * (1 - t) + (mid[i] if y < h * 0.6 else bot[i]) * t) for i in
                  range(3)) + (255,)
        dr.line([(0, y), (w, y)], fill=c)
    mask = Image.new("L", (w, h), 0);
    ImageDraw.Draw(mask).rounded_rectangle([0, 0, w - 1, h - 1], radius=r, fill=255);
    im.putalpha(mask)
    hl = Image.new("RGBA", (w, h), (255, 255, 255, 0))
    ImageDraw.Draw(hl).rounded_rectangle([2, 2, w - 3, h // 2], radius=r - 2, fill=(255, 255, 255, 70))
    im = Image.alpha_composite(im, hl)
    ImageDraw.Draw(im).rounded_rectangle([0.5, 0.5, w - 1.5, h - 1.5], radius=r, outline=(200, 120, 20, 255), width=2)
    bio = BytesIO();
    im.save(bio, format="PNG");
    bio.seek(0);
    return tk.PhotoImage(data=bio.read())


def play_success_sound():
    wav = resource_path("success.wav")
    try:
        if os.path.exists(wav):
            winsound.PlaySound(wav, winsound.SND_FILENAME | winsound.SND_ASYNC)
        else:
            winsound.MessageBeep(winsound.MB_ICONASTERISK)
    except Exception:
        pass


def run(cmd):
    env = os.environ.copy();
    env["DOTNET_CLI_UI_LANGUAGE"] = "en"
    p = subprocess.run(cmd, capture_output=True, text=True, shell=False, encoding="utf-8", errors="replace",
                       env=env, startupinfo=_hidden_startupinfo(), creationflags=CREATE_NO_WINDOW)
    return p.returncode, p.stdout.strip(), p.stderr.strip()


def get_upgrade_list(include_unknown=False):
    args = ["winget", "upgrade", "--accept-source-agreements", "--disable-interactivity", "--source", "winget"]
    if include_unknown:
        args.append("--include-unknown")
    code, out, err = run(args)
    if code != 0 or not out:
        raise RuntimeError(err or "winget upgrade failed")
    return parse_table_upgrade_output(out)


def parse_table_upgrade_output(text: str):
    """
    Robust parser for `winget upgrade` table output.
    - Uses header column positions to slice fields (no fragile regex).
    - Handles wrapped rows (long Name continues on the next line).
    """

    lines = text.splitlines()
    items = []

    # 1) find header and compute column starts
    header_idx = None
    for i, ln in enumerate(lines):
        if ln.strip().startswith("Name") and "Id" in ln and "Version" in ln and "Available" in ln:
            header_idx = i
            break
    if header_idx is None:
        return items  # nothing parseable

    header = lines[header_idx]

    # Derive column start indices from header labels
    def col_start(label):
        return header.index(label)

    i_name = col_start("Name")
    i_id = col_start("Id")
    i_ver = col_start("Version")
    i_av = col_start("Available")
    i_src = header.find("Source")
    if i_src == -1:
        # Some builds omit Source; put it at end
        i_src = len(header)

    # 2) iterate data lines after the dashed separator
    i = header_idx + 1
    # Skip the dashed rule if present
    while i < len(lines) and set(lines[i].strip()) <= {"-"}:
        i += 1

    cur_row = None
    while i < len(lines):
        ln = lines[i]
        i += 1
        if not ln.strip():
            continue

        # If the line is shorter than expected columns, pad it
        if len(ln) < i_src:
            ln = ln + " " * (i_src - len(ln))
            # Skip winget summary line like "5 upgrades available."
            if re.match(r"^\d+\s+upgrades?\s+available\.?$", ln.strip(), flags=re.I):
                continue
        # Skip footer line about unknown version numbers
        if re.search(r"have version numbers that cannot be determined", ln, flags=re.I):
            continue

        # Slice by columns
        name = ln[i_name:i_id].rstrip()
        pid = ln[i_id:i_ver].strip()
        curr = ln[i_ver:i_av].strip()
        avail = ln[i_av:i_src].strip()
        src = ln[i_src:].strip() if i_src < len(ln) else ""

        # Clean up version fields: remove things like "(August 2025)"
        curr = re.sub(r"\s*\(.*?\)", "", curr).strip()
        avail = re.sub(r"\s*\(.*?\)", "", avail).strip()
        name = re.sub(r"\s+\d+\s+upgrades?\s+available\.?\s*$", "", name, flags=re.I).strip()

        # Wrapped continuation: no Id/Version/Available, just more Name text
        if not pid and not curr and not avail and name:
            if cur_row is not None:
                cur_row["name"] = (cur_row["name"] + " " + name).strip()
            continue

        # A real row
        if pid and avail:
            cur_row = {"name": name, "id": pid, "current": curr, "available": avail}
            items.append(cur_row)
        else:
            # Sometimes Available is blank for unknowns; still keep the row
            if pid:
                cur_row = {"name": name, "id": pid, "current": curr, "available": avail or ""}
                items.append(cur_row)
            else:
                cur_row = None

    return items


def get_winget_upgrades(include_unknown):
    c, _, _ = run(["winget", "--version"])
    if c != 0: raise RuntimeError("winget not found. Install the App Installer from Microsoft Store.")
    try:
        return get_upgrade_list(include_unknown)
    except Exception as e_json:
        cmd = ["winget", "upgrade", "--accept-source-agreements", "--disable-interactivity"]
        if include_unknown: cmd.insert(2, "--include-unknown")
        c, o, e = run(cmd)
        if c != 0: raise RuntimeError((e or str(e_json)).strip())
        parsed = parse_table_upgrade_output(o)
        if parsed: return parsed
        raise RuntimeError(str(e_json))


def make_checkbox_images(size=16):
    u = tk.PhotoImage(width=size, height=size);
    u.put("white", to=(0, 0, size, size));
    b = "gray20"
    u.put(b, to=(0, 0, size, 1));
    u.put(b, to=(0, size - 1, size, size));
    u.put(b, to=(0, 0, 1, size));
    u.put(b, to=(size - 1, 0, size, size))
    c = tk.PhotoImage(width=size, height=size);
    c.tk.call(c, "copy", u);
    mark = "#2e7d32"
    for (x, y) in [(3, size // 2), (4, size // 2 + 1), (5, size // 2 + 2), (6, size // 2 + 3), (7, size // 2 + 2),
                   (8, size // 2 + 1), (9, size // 2), (10, size // 2 - 1)]:
        c.put(mark, to=(x, y, x + 1, y + 1));
        c.put(mark, to=(x, y - 1, x + 1, y))
    return u, c


def _download_file(url: str, dest_path: str, progress_cb=None):
    """
    Stream download with:
    - .part temp file + atomic replace (prevents half-written exe from being used)
    - Content-Length verification when available (prevents truncated downloads)
    - Quick EXE sanity check (MZ header)
    """
    tmp = dest_path + ".part"

    # Clean leftover .part
    try:
        if os.path.exists(tmp):
            os.remove(tmp)
    except Exception:
        pass

    req = urllib.request.Request(url, headers={"User-Agent": "Windows-App-Updater"})
    with urllib.request.urlopen(req, timeout=60) as r:
        total = int(r.headers.get("Content-Length") or 0)
        done = 0

        with open(tmp, "wb") as f:
            while True:
                chunk = r.read(1024 * 128)  # 128 KB
                if not chunk:
                    break
                f.write(chunk)
                done += len(chunk)
                if progress_cb and total > 0:
                    try:
                        progress_cb(int(done * 100 / total))
                    except Exception:
                        pass

    # ✅ Verify completeness if server provided size
    if total > 0 and done != total:
        try:
            os.remove(tmp)
        except Exception:
            pass
        raise RuntimeError(f"Download incomplete ({done}/{total} bytes). Please retry.")

    # ✅ Quick sanity check for Windows EXE (optional but helpful)
    try:
        with open(tmp, "rb") as f:
            if f.read(2) != b"MZ":
                raise RuntimeError("Downloaded file is not a valid Windows EXE.")
    except Exception:
        try:
            os.remove(tmp)
        except Exception:
            pass
        raise

    # ✅ Atomic move into place (no partial exe ever gets launched)
    os.replace(tmp, dest_path)
    return dest_path

def _check_vc_runtime():
    # These are the most common dependencies needed by python312.dll
    # Security: use absolute System32 path to prevent DLL search order hijacking
    system32 = os.path.join(os.environ.get("WINDIR", r"C:\Windows"), "System32")
    dlls = ["vcruntime140.dll", "vcruntime140_1.dll", "msvcp140.dll"]
    missing = []
    for d in dlls:
        try:
            ctypes.WinDLL(os.path.join(system32, d))
        except OSError:
            missing.append(d)
    return missing



def _sanitize_batch_path(p):
    """Remove characters dangerous in batch scripts to prevent command injection."""
    return p.replace('"', '').replace('%', '').replace('^', '').replace('&', '').replace('|', '').replace('<', '').replace('>', '').replace('`', '')

class WingetUpdaterUI:
    def __init__(self, root):
        self.root = root;
        self.root.title(APP_NAME_VERSION)
        self.root.geometry(f"{WIN_W}x{WIN_H_FULL}");
        self.root.minsize(WIN_W, WIN_H_COMPACT)
        self.style = ttk.Style(self.root)
        self.style.configure("Big.TButton", padding=(14, 8), font=("Segoe UI", 10, "bold"))
        self.style.configure("Big.TCheckbutton", padding=(8, 6), font=("Segoe UI", 10))

        self.updating = False;
        self.cancel_requested = False;
        self.current_proc = None;
        self._state_lock = threading.Lock()
        self.loading_win = None
        self.window_icon_path = set_app_icon(self.root)

        self.img_unchecked, self.img_checked = make_checkbox_images(16)
        self.checked_items = set();
        self.id_to_item = {};
        self.pkg_downloads = {}

        # Skip state
        self.skip_requested = False
        self._can_skip = False
        # ===== Header =====
        header = ttk.Frame(self.root);
        header.pack(fill="x", pady=(10, 0))
        ttk.Label(header, text=APP_NAME_VERSION, font=("Segoe UI", 18, "bold")).pack(side="left", padx=12)
        right = ttk.Frame(header);
        right.pack(side="right", padx=12)
        self.btn_admin = ttk.Button(right, text="Run as Admin", command=self.run_as_admin, style="Big.TButton")
        self.btn_admin.pack(side="right")
        if is_admin():
            self.btn_admin.config(text="Running as Admin", state="disabled")
        else:
            ToolTip(self.btn_admin, "Run the app as admin to install apps silently")

        # ===== Row 1: Check - Include - Update - Clear - Open =====
        row1 = ttk.Frame(self.root);
        row1.pack(fill="x", padx=12, pady=6)
        width_btn = max(len("Check for Updates"), len("Update Selected")) + 2
        self.btn_check = ttk.Button(row1, text="Check for Updates", command=self.check_for_updates_async,
                                    style="Big.TButton", width=width_btn);
        self.btn_check.pack(side="left")
        self.include_unknown_var = tk.BooleanVar(value=False)
        self.chk_unknown = ttk.Checkbutton(row1, text="Include unknown apps", variable=self.include_unknown_var,
                                           style="Big.TCheckbutton");
        self.chk_unknown.pack(side="left", padx=(12, 0))
        self.btn_update = ttk.Button(row1, text="Update Selected", command=self.update_selected_async,
                                     style="Big.TButton", width=width_btn);
        self.btn_update.pack(side="left", padx=(12, 0))
        self.btn_open_temp = ttk.Button(row1, text="Open Temp", command=self.open_temp, style="Big.TButton");
        self.btn_open_temp.pack(side="right", padx=(12, 0))
        self.btn_clear_temp = ttk.Button(row1, text="Clear Temp", command=self.clear_temp_async, style="Big.TButton");
        self.btn_clear_temp.pack(side="right", padx=(12, 0))
        ToolTip(self.btn_clear_temp, "Delete unnecessary temporary installer files downloaded by apps.\n"
                                     "Safe to use - running apps won't be affected.")

        # ===== Row 2: Select All . Select None . Counter . About (right) =====
        row2 = ttk.Frame(self.root);
        row2.pack(fill="x", padx=12, pady=(0, 6))
        self.btn_sel_all = ttk.Button(row2, text="Select All", command=self.select_all, style="Big.TButton",
                                      state="disabled");
        self.btn_sel_all.pack(side="left")
        self.btn_sel_none = ttk.Button(row2, text="Select None", command=self.select_none, style="Big.TButton",
                                       state="disabled");
        self.btn_sel_none.pack(side="left", padx=(6, 0))
        self.counter_var = tk.StringVar(value="0 apps found • 0 selected")
        self.btn_skip = ttk.Button(row2, text="Skip", command=self.skip_current, style="Big.TButton", state="disabled")
        self.btn_skip.pack(side="left", padx=(12, 0))

        ttk.Label(row2, textvariable=self.counter_var).pack(side="left", padx=(12, 0))
        self.btn_about = ttk.Button(row2, text="About", command=self.show_about, style="Big.TButton");
        self.btn_about.pack(side="right")

        # Controls to disable during update
        self._controls_to_disable = [self.btn_check, self.chk_unknown, self.btn_sel_all, self.btn_sel_none,
                                     self.btn_open_temp, self.btn_clear_temp, self.btn_about]

        # ===== Apps list (fixed height) =====
        tree_wrap = ttk.Frame(self.root, height=LIST_PIXELS);
        tree_wrap.pack(fill="x", expand=False, padx=12, pady=(4, 8));
        tree_wrap.pack_propagate(False)
        cols = ("Name", "Id", "Current", "Available", "Result");
        self.fixed_cols = cols
        self.tree = ttk.Treeview(tree_wrap, columns=cols, show="tree headings", height=8, selectmode="none")
        self.tree.heading("#0", text="Select", anchor="center")
        self.tree.heading("Name", text="Name", anchor="w")
        self.tree.heading("Id", text="Id", anchor="w")
        self.tree.heading("Current", text="Current", anchor="center")
        self.tree.heading("Available", text="Available", anchor="center")
        self.tree.heading("Result", text="Result", anchor="center")
        font = tkfont.nametofont("TkDefaultFont");
        selw = max(50, font.measure("Select") + 18)
        self.tree.column("#0", width=selw, minwidth=selw, anchor="center", stretch=False)
        self.tree.column("Name", width=260, minwidth=160, anchor="w", stretch=False)
        self.tree.column("Id", width=340, minwidth=220, anchor="w", stretch=False)
        self.tree.column("Current", width=90, minwidth=70, anchor="center", stretch=False)
        self.tree.column("Available", width=100, minwidth=80, anchor="center", stretch=False)
        self.tree.column("Result", width=120, minwidth=100, anchor="center", stretch=False)
        self._col_caps = {"Name": 320, "Id": 420, "Current": 120, "Available": 140, "Result": 160}
        for k, v in {"ok": "#e8f5e9", "fail": "#ffebee", "skip": "#fff8e1", "cancel": "#fff3e0",
                     "checked": "#e3f2fd"}.items():
            self.tree.tag_configure(k, background=v)
        ysb = ttk.Scrollbar(tree_wrap, orient="vertical", command=self.tree.yview)
        xsb = ttk.Scrollbar(tree_wrap, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscroll=ysb.set, xscroll=xsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew");
        ysb.grid(row=0, column=1, sticky="ns");
        xsb.grid(row=1, column=0, sticky="ew")
        tree_wrap.rowconfigure(0, weight=1);
        tree_wrap.columnconfigure(0, weight=1)
        self.tree.bind("<Button-1>", self._on_mouse_down, add="+")
        self.tree.bind("<B1-Motion>", self._on_mouse_drag, add="+")
        self.tree.bind("<ButtonRelease-1>", self._on_mouse_up, add="+")
        self.tree.bind("<Double-Button-1>", self._on_double_click_header, add="+")
        self.tree.bind("<Configure>", self._on_tree_configure, add="+")
        self.tree["displaycolumns"] = self.fixed_cols

        # Context menu for per-app downloads
        self.row_menu = tk.Menu(self.root, tearoff=0)
        self.row_menu.add_command(label="Open downloaded file(s)", command=self._menu_open_downloads)
        self.row_menu.add_command(label="Delete downloaded file(s)", command=self._menu_delete_downloads)
        self._menu_item = None

        def _show_menu(e, tree=self.tree):
            if self.updating: return
            it = tree.identify_row(e.y)
            if not it: return
            self._menu_item = it;
            tree.selection_set(it)
            try:
                self.row_menu.tk_popup(e.x_root, e.y_root)
            finally:
                self.row_menu.grab_release()

        self.tree.bind("<Button-3>", _show_menu)

        # ===== Progress bars =====
        pbw = ttk.Frame(self.root);
        pbw.pack(fill="x", padx=12, pady=(0, 4))
        self.pb_label = ttk.Label(pbw, text="Update");
        self.pb_label.pack(side="left")
        self.pb = ttk.Progressbar(pbw, orient="horizontal", mode="determinate");
        self.pb.pack(fill="x", expand=True, padx=10)
        pbw2 = ttk.Frame(self.root);
        pbw2.pack(fill="x", padx=12, pady=(0, 8))
        self.pb2_label = ttk.Label(pbw2, text="Download");
        self.pb2_label.pack(side="left")
        self.pb2 = ttk.Progressbar(pbw2, orient="horizontal", mode="determinate", maximum=100, value=0);
        self.pb2.pack(fill="x", expand=True, padx=10)

        # Spinner state
        self._spin_job = None;
        self._spin_frames = ["|", "/", "-", "\\"];
        self._spin_index = 0
        self._spin_base = "Downloading";
        self._spin_name = None;
        self._spin_pct = 0

        # ===== Log (fixed height, equals list) =====
        # Right-side log controls (Hide always visible, Save toggled)
        log_ctrls = ttk.Frame(self.root)
        log_ctrls.pack(side="right", padx=12, pady=(0, 2), anchor="ne")

        self.btn_toggle_log = ttk.Button(log_ctrls, text="Hide Log",
                                         command=self.toggle_log, style="Big.TButton")
        self.btn_toggle_log.pack(fill="x")

        self.btn_save_log = ttk.Button(log_ctrls, text="Save / Export Log",
                                       command=self.save_export_log, style="Big.TButton")
        self.btn_save_log.pack(fill="x", pady=(6, 0))
        # Ensure both buttons have same width
        btn_width = max(len("Hide Log"), len("Save / Export Log")) + 2
        self.btn_toggle_log.config(width=btn_width)
        self.btn_save_log.config(width=btn_width)

        self.log_wrap = ttk.Frame(self.root, height=LIST_PIXELS);
        self.log_wrap.pack(fill="x", expand=False, padx=12, pady=(0, 10));
        self.log_wrap.pack_propagate(False)
        self.log_box = tk.Text(self.log_wrap, wrap="none", font=("Consolas", 10))
        ys = ttk.Scrollbar(self.log_wrap, orient="vertical", command=self.log_box.yview)
        xs = ttk.Scrollbar(self.log_wrap, orient="horizontal", command=self.log_box.xview)
        self.log_box.configure(yscrollcommand=ys.set, xscrollcommand=xs.set)
        self.log_box.grid(row=0, column=0, sticky="nsew");
        ys.grid(row=0, column=1, sticky="ns");
        xs.grid(row=1, column=0, sticky="ew")
        self.log_wrap.rowconfigure(0, weight=1);
        self.log_wrap.columnconfigure(0, weight=1)
        self.log_visible = True

        self.root.after(0, self.center_on_screen)
        self.root.after(1200, self.check_latest_app_version_async)

    def manual_check_for_update(self):
        # Show a tiny loader so the About window stays responsive
        self.show_loading("Checking for updates...")

        def worker():
            try:
                req = urllib.request.Request(GITHUB_API_LATEST, headers={"User-Agent": "Windows-App-Updater"})
                with urllib.request.urlopen(req, timeout=10) as r:
                    data = json.loads(r.read().decode("utf-8", "replace"))
                tag = str(data.get("tag_name") or data.get("name") or "").strip()
                cur = APP_VERSION_ONLY
                newer = bool(tag and self._parse_ver_tuple(tag) > self._parse_ver_tuple(cur))
            except Exception as e:
                self.root.after(0, lambda: (
                    self.hide_loading(),
                    messagebox.showerror("Update check failed", str(e))
                ))
                return

            def after():
                self.hide_loading()
                if newer:
                    if messagebox.askyesno(
                            "New Version Available",
                            f"A newer version {tag} is available.\n\nDownload & install it now?"
                    ):
                        self._download_and_run_latest(data)
                else:
                    messagebox.showinfo("You're up to date", f"Current version {cur} is the latest.")

            self.root.after(0, after)

        threading.Thread(target=worker, daemon=True).start()

    def save_export_log(self):
        try:
            content = self.log_box.get("1.0", "end-1c")
            if not content.strip():
                messagebox.showinfo("Export Log", "Log is empty.")
                return

            path = filedialog.asksaveasfilename(
                title="Save Log",
                defaultextension=".txt",
                initialfile="WindowsAppUpdater_Log.txt",
                filetypes=[("Text file", "*.txt"), ("Log file", "*.log"), ("All files", "*.*")]
            )
            if not path:
                return

            with open(path, "w", encoding="utf-8") as f:
                f.write(content)

            messagebox.showinfo("Export Log", f"Saved to:\n{path}")
            try:
                os.startfile(os.path.dirname(path))  # open folder for convenience
            except Exception:
                pass
        except Exception as e:
            messagebox.showerror("Export Log", str(e))

    def _download_and_run_latest(self, release_json: dict):
        import os, sys, tempfile, threading, subprocess, textwrap, webbrowser
        from tkinter import messagebox

        assets = release_json.get("assets") or []

        INSTALLER_EXTS = (".exe", ".msi", ".msixbundle", ".msix", ".appxbundle", ".appx")

        def _rank(a):
            n = (a.get("name") or "").lower()
            pref = 0 if ("setup" in n or "installer" in n) else 1
            ext_pref = 0 if any(n.endswith(e) for e in INSTALLER_EXTS) else 2
            return (ext_pref, pref, n)

        asset = None
        for cand in sorted(assets, key=_rank):
            n = (cand.get("name") or "").lower()
            if any(n.endswith(e) for e in INSTALLER_EXTS):
                asset = cand
                break

        if not asset:
            messagebox.showinfo(
                "No installer asset",
                "Couldn’t find a .exe/.msi asset in the latest release. Opening releases page."
            )
            webbrowser.open(GITHUB_RELEASES_PAGE)
            return

        url = asset.get("browser_download_url") or ""
        name = os.path.basename(asset.get("name") or "setup.exe")
        name = re.sub(r'[^a-zA-Z0-9._\-]', '_', name)
        if not name:
            name = "setup.exe"
        if not url:
            messagebox.showerror("Download error", "Latest release asset is missing a download URL.")
            return
        if not url.startswith("https://github.com/") and not url.startswith("https://objects.githubusercontent.com/"):
            messagebox.showerror("Download error", "Download URL must be from GitHub (HTTPS).")
            return

        if not name.lower().endswith(".exe"):
            messagebox.showinfo(
                "Installer type",
                "Latest asset isn’t an .exe. Opening releases page instead."
            )
            webbrowser.open(GITHUB_RELEASES_PAGE)
            return

        # ✅ EXE-only updater (never replace .py)
        if not getattr(sys, "frozen", False):
            messagebox.showinfo(
                "Update",
                "Self-update is only supported in the packaged EXE build.\nOpening releases page instead."
            )
            webbrowser.open(GITHUB_RELEASES_PAGE)
            return

        dest = os.path.join(tempfile.gettempdir(), name)

        self.show_loading(f"Downloading {name} ...")

        def worker():
            try:
                def _on_pct(p):
                    self.root.after(0, lambda: (
                        self.pb2.configure(value=p),
                        self.pb2_label.configure(text=f"Downloading: {p}%")
                    ))

                # ✅ Use your verified _download_file() (atomic + size check)
                _download_file(url, dest, progress_cb=_on_pct)

            except Exception as e:
                self.root.after(0, lambda: (
                    self.hide_loading(),
                    messagebox.showerror("Download failed", str(e))
                ))
                return

            missing = _check_vc_runtime()
            if missing:
                messagebox.showerror(
                    "Missing system runtime",
                    "The new build requires Microsoft Visual C++ Runtime.\n\n"
                    f"Missing DLLs: {', '.join(missing)}\n\n"
                    "Install: Microsoft Visual C++ Redistributable 2015–2022 (x64) "
                    "and also (x86) then try update again."
                )
                return

            def _launch():
                self.hide_loading()

                final_path = _sanitize_batch_path(sys.executable)  # current running EXE
                new_path = _sanitize_batch_path(dest)

                script = textwrap.dedent(f"""\
                @echo off
                setlocal

                set "FINAL={final_path}"
                set "NEW={new_path}"

                timeout /t 2 /nobreak >nul

                set /a tries=0

                :waitloop
                del /f /q "%FINAL%" >nul 2>&1
                if exist "%FINAL%" (
                  set /a tries+=1
                  if %tries% GEQ 30 (
                    REM Could not delete (locked/permissions). Run NEW anyway.
                    start "" "%NEW%"
                    goto cleanup
                  )
                  timeout /t 1 /nobreak >nul
                  goto waitloop
                )

                move /y "%NEW%" "%FINAL%" >nul 2>&1
                if not exist "%FINAL%" (
                  REM move failed (permission/AV). Run NEW anyway.
                  start "" "%NEW%"
                  goto cleanup
                )

                start "" "%FINAL%"

                :cleanup
                (del "%~f0" >nul 2>&1)
                exit /b
                """)

                script_path = os.path.join(tempfile.gettempdir(), "update_self.bat")
                with open(script_path, "w", encoding="utf-8") as f:
                    f.write(script)

                install_dir = os.path.dirname(final_path)
                writable = os.access(install_dir, os.W_OK)

                try:
                    if not writable:
                        # 🔥 If installed in a protected folder (e.g. Program Files), ask to elevate
                        if messagebox.askyesno(
                                "Admin required",
                                "The app is installed in a protected folder.\n\n"
                                "Run the updater as Administrator to replace the EXE?"
                        ):
                            # Run BAT elevated
                            subprocess.Popen([
                                "powershell",
                                "-NoProfile",
                                "-ExecutionPolicy", "Bypass",
                                "-Command",
                                f'Start-Process cmd.exe -Verb RunAs -ArgumentList \'/c "{script_path}"\''
                            ])
                        else:
                            webbrowser.open(GITHUB_RELEASES_PAGE)
                            return
                    else:
                        # Normal (no admin needed)
                        CREATE_NO_WINDOW = 0x08000000
                        subprocess.Popen(
                            ["cmd.exe", "/c", script_path],
                            creationflags=CREATE_NO_WINDOW
                        )
                except Exception:
                    # Fallback
                    os.startfile(script_path)

                # Exit so the batch can replace the file
                self.root.destroy()

            self.root.after(0, _launch)

        threading.Thread(target=worker, daemon=True).start()

    # =================== About ===================
    def show_about(self):
        win = tk.Toplevel(self.root);
        win.title("About");
        win.resizable(False, False);
        apply_icon_to_tlv(win, set_app_icon(win))
        frame = ttk.Frame(win, padding=16);
        frame.pack(fill="both", expand=True)
        tk.Label(frame, text="Windows App Updater", font=("Segoe UI", 14, "bold")).pack(pady=(0, 4))
        tk.Label(frame, text="is a freeware Python App based on Windows Winget to update applications",
                 wraplength=520, justify="center").pack(pady=(0, 8))
        tk.Label(frame, text=f"Version {APP_VERSION_ONLY} - {DATE_APP}").pack(pady=(0, 8))
        row = ttk.Frame(frame);
        row.pack()
        tk.Label(row, text="Author: ilukezippo (BoYaqoub)").pack(side="left")
        flag = load_flag_image()
        if flag:
            tk.Label(row, image=flag).pack(side="left", padx=(6, 0))
            win._flag = flag

        # New line for email contact
        email_row = ttk.Frame(frame);
        email_row.pack(pady=(6, 0))
        tk.Label(email_row, text="For any feedback contact: ").pack(side="left")

        email_lbl = tk.Label(email_row, text="ilukezippo@gmail.com",
                             fg="#1a73e8", cursor="hand2", font=("Segoe UI", 9, "underline"))
        email_lbl.pack(side="left")
        email_lbl.bind("<Button-1>", lambda e: webbrowser.open("mailto:ilukezippo@gmail.com"))

        link_row = ttk.Frame(frame);
        link_row.pack(pady=(8, 0))
        tk.Label(link_row, text="Info and Latest Updates at ").pack(side="left")
        link = tk.Label(link_row, text="https://github.com/ilukezippo/Windows-App-Updater",
                        fg="#1a73e8", cursor="hand2", font=("Segoe UI", 9, "underline"))
        link.pack(side="left");
        link.bind("<Button-1>", lambda e: webbrowser.open("https://github.com/ilukezippo/Windows-App-Updater"))
        # Donate INSIDE About
        donate_img = make_donate_image(160, 44);
        win._don = donate_img
        tk.Button(frame, image=donate_img, text="Donate", compound="center",
                  font=("Segoe UI", 11, "bold"), fg="#0f3462", activeforeground="#0f3462",
                  bd=0, highlightthickness=0, cursor="hand2", relief="flat",
                  command=lambda: webbrowser.open(DONATE_PAGE)).pack(pady=(12, 0))
        # Manual "Check for Update" button
        ttk.Button(frame, text="Check for Update", command=self.manual_check_for_update, style="Big.TButton").pack(
            pady=(6, 6))

        ttk.Button(frame, text="Close", command=win.destroy, style="Big.TButton").pack(pady=(10, 0))
        self.center_child(win)

    def center_child(self, tlv):
        tlv.update_idletasks()
        x = self.root.winfo_x() + (self.root.winfo_width() - tlv.winfo_width()) // 2
        y = self.root.winfo_y() + (self.root.winfo_height() - tlv.winfo_height()) // 2
        tlv.geometry(f"+{x}+{y}")

    # ===== GitHub update check =====
    def _parse_ver_tuple(self, v):
        return tuple(int(n) for n in re.findall(r"\d+", v)[:4]) or (0,)

    def check_latest_app_version_async(self):
        def worker():
            try:
                req = urllib.request.Request(GITHUB_API_LATEST, headers={"User-Agent": "Windows-App-Updater"})
                with urllib.request.urlopen(req, timeout=10) as r:
                    data = json.loads(r.read().decode("utf-8", "replace"))
                tag = str(data.get("tag_name") or data.get("name") or "").strip()
                if tag and self._parse_ver_tuple(tag) > self._parse_ver_tuple(APP_VERSION_ONLY):
                    def _ask_and_handle():
                        if messagebox.askyesno(
                                "New Version Available",
                                f"A newer version {tag} is available.\n\nDownload & install it now?"
                        ):
                            # Pass the whole JSON to the downloader/launcher
                            self._download_and_run_latest(data)
                        else:
                            # Optional: still offer to open releases page
                            if messagebox.askyesno("View Releases", "Open the releases page instead?"):
                                webbrowser.open(GITHUB_RELEASES_PAGE)

                    self.root.after(0, _ask_and_handle)

            except Exception:
                pass

        threading.Thread(target=worker, daemon=True).start()

    # ===== Log toggle resizes window =====
    def toggle_log(self):
        if self.log_visible:
            self.log_wrap.forget()
            self.btn_save_log.forget()  # hide only Save button
            self.btn_toggle_log.config(text="Show Log")
            self.root.geometry(f"{WIN_W}x{WIN_H_COMPACT}")
            self.log_visible = False
        else:
            self.log_wrap.pack(fill="x", expand=False, padx=12, pady=(0, 10))
            self.log_wrap.configure(height=LIST_PIXELS)
            self.btn_save_log.pack(fill="x", pady=(6, 0))  # re-show Save button under Hide Log
            self.btn_toggle_log.config(text="Hide Log")
            self.root.geometry(f"{WIN_W}x{WIN_H_FULL}")
            self.log_visible = True
        self.root.update_idletasks()

    # ===== Enable/disable during update =====
    def _disable_controls_for_update(self):
        self.updating = True
        for w in self._controls_to_disable:
            try:
                w.config(state="disabled")
            except Exception:
                pass
        try:
            self.btn_admin.config(state="disabled")
        except Exception:
            pass
        self.btn_update.config(text="Cancel", state="normal")
        try:
            self._tree_prev_state = self.tree.cget("state")
        except Exception:
            self._tree_prev_state = "normal"
        try:
            self.tree.config(state="disabled")
        except Exception:
            pass

    def _enable_controls_after_update(self):
        for w in self._controls_to_disable:
            try:
                w.config(state="normal")
            except Exception:
                pass
        try:
            if is_admin():
                self.btn_admin.config(text="Running as Admin", state="disabled")
            else:
                self.btn_admin.config(text="Run as Admin", state="normal")
        except Exception:
            pass
        try:
            self.tree.config(state=self._tree_prev_state if hasattr(self, "_tree_prev_state") else "normal")
        except Exception:
            pass
        self.btn_update.config(text="Update Selected", state="normal");
        self.updating = False

    # ===== Column clamping =====
    def _on_tree_configure(self, _=None):
        self.root.after_idle(self._fit_columns_to_tree)

    def _fit_columns_to_tree(self):
        try:
            avail = max(0, self.tree.winfo_width() - 2)
        except Exception:
            return
        cols = ["#0", "Name", "Id", "Current", "Available", "Result"]
        widths = {c: int(self.tree.column(c, "width")) for c in cols if c in ("#0",) or c in self.fixed_cols}
        mins = {c: int(self.tree.column(c, "minwidth") or 20) for c in cols if c in ("#0",) or c in self.fixed_cols}
        caps = {"#0": widths.get("#0", 60), **self._col_caps}
        for c in cols:
            if c in widths and widths[c] > caps.get(c, 10_000):
                self.tree.column(c, width=caps[c]);
                widths[c] = caps[c]
        total = sum(widths.get(c, 0) for c in cols if c in widths)
        if total <= avail or avail <= 0: return
        overflow = total - avail
        reducible = {c: max(0, widths[c] - mins.get(c, 20)) for c in cols if c in widths}
        order = ["Id", "Name", "Result", "Available", "Current", "#0"];
        total_red = sum(reducible.get(c, 0) for c in order)
        if total_red == 0: return
        for c in order:
            if overflow <= 0: break
            r = reducible.get(c, 0)
            if r <= 0: continue
            share = int(round(overflow * (r / total_red)));
            share = max(1, min(r, share))
            neww = max(mins.get(c, 20), widths[c] - share);
            self.tree.column(c, width=neww);
            overflow -= (widths[c] - neww)

    # ===== Mouse & selection =====
    def _on_mouse_down(self, e):
        if self.updating: return "break"
        region = self.tree.identify("region", e.x, e.y)
        if region == "heading": self._block_header_drag = True; return "break"
        if region == "separator":
            if self.tree.identify_column(e.x - 1) == "#0": self._block_resize_select = True; return "break"
            self._block_resize_select = False;
            return
        if region in ("tree", "cell"):
            it = self.tree.identify_row(e.y)
            if it: self._toggle_row(it); return "break"
        self._block_header_drag = False;
        self._block_resize_select = False

    def _on_mouse_drag(self, e):
        if self.updating or getattr(self, "_block_header_drag", False) or getattr(self, "_block_resize_select",
                                                                                  False): return "break"

    def _on_mouse_up(self, e):
        if self.updating: return "break"
        self._block_header_drag = False;
        self._block_resize_select = False
        try:
            if tuple(self.tree["displaycolumns"]) != tuple(self.fixed_cols): self.tree[
                "displaycolumns"] = self.fixed_cols
        except Exception:
            pass

    def _on_double_click_header(self, e):
        if self.updating: return "break"
        if self.tree.identify("region", e.x, e.y) != "separator": return
        col = self.tree.identify_column(e.x - 1)
        if not col or col == "#0": return
        self.autofit_column(col);
        self._fit_columns_to_tree()

    def _toggle_row(self, item):
        if not item or self.updating: return
        if item in self.checked_items:
            self.checked_items.remove(item);
            self.tree.item(item, image=self.img_unchecked)
            tags = set(self.tree.item(item, "tags") or ());
            tags.discard("checked");
            self.tree.item(item, tags=tuple(tags))
        else:
            self.checked_items.add(item);
            self.tree.item(item, image=self.img_checked)
            tags = set(self.tree.item(item, "tags") or ());
            tags.add("checked");
            self.tree.item(item, tags=tuple(tags))
        self.update_counter()

    def autofit_column(self, cid):
        head = self.tree.heading(cid, "text") or "";
        f = tkfont.nametofont("TkDefaultFont");
        pad = 24
        mx = f.measure(head)
        for it in self.tree.get_children(""):
            v = self.tree.set(it, cid) or "";
            mx = max(mx, f.measure(v))
        neww = max(int(self.tree.column(cid, "minwidth") or 20), min(mx + pad, self._col_caps.get(cid, 360)))
        self.tree.column(cid, width=neww, stretch=False)

    def autofit_all(self):
        for c in self.fixed_cols:
            if c != "#0": self.autofit_column(c)
        self._fit_columns_to_tree()

    # ===== Positioning =====
    def center_on_screen(self):
        self.root.update_idletasks()
        w, h = self.root.winfo_width(), self.root.winfo_height();
        sw, sh = self.root.winfo_screenwidth(), self.root.winfo_screenheight()
        x, y = (sw - w) // 2, (sh - h) // 2;
        self.root.geometry(f"{w}x{h}+{x}+{y}")

    def run_as_admin(self):
        relaunch_as_admin()

    # ===== Loading mini window =====
    def show_loading(self, text="Loading..."):
        if self.loading_win: return
        w = tk.Toplevel(self.root);
        self.loading_win = w;
        w.transient(self.root);
        w.grab_set();
        w.resizable(False, False);
        apply_icon_to_tlv(w, self.window_icon_path)
        ttk.Label(w, text=text, font=("Segoe UI", 12, "bold")).pack(padx=20, pady=(16, 8))
        pb = ttk.Progressbar(w, mode="indeterminate", length=280);
        pb.pack(padx=20, pady=(0, 16));
        pb.start(10)
        w.update_idletasks()
        x = self.root.winfo_x() + (self.root.winfo_width() - w.winfo_width()) // 2;
        y = self.root.winfo_y() + (self.root.winfo_height() - w.winfo_height()) // 2
        w.geometry(f"+{x}+{y}");
        w.protocol("WM_DELETE_WINDOW", lambda: None)

    def hide_loading(self):
        if self.loading_win:
            try:
                self.loading_win.grab_release()
            except Exception:
                pass
            self.loading_win.destroy();
            self.loading_win = None

    # ===== Progress + spinner =====
    def progress_start(self, phase, total):
        self.pb_total = max(0, int(total));
        self.pb_value = 0
        self.pb.configure(maximum=max(self.pb_total, 1), value=0, mode="determinate")
        self.pb_label.configure(text=f"Updating: 0/{self.pb_total}")
        self.pb2.configure(value=0, maximum=100, mode="determinate")
        self.pb2_label.configure(text="Downloading: 0%");
        self.root.update_idletasks()

    def progress_step(self, inc=1):
        if self.pb_total <= 0: return
        self.pb_value = min(self.pb_total, self.pb_value + inc);
        self.pb.configure(value=self.pb_value)
        self.pb_label.configure(text=f"Updating: {self.pb_value}/{self.pb_total}");
        self.root.update_idletasks()

    def progress_finish(self, canceled=False):
        if getattr(self, "pb_total", 0) > 0:
            self.pb.configure(value=self.pb_total)
            self.pb_label.configure(
                text=f"Update: {self.pb_total}/{self.pb_total}" + (" (canceled)" if canceled else " (done)"))
        else:
            self.pb_label.configure(text="Update")
        self._spinner_stop("Download");
        self.pb2.configure(value=0);
        self.root.update_idletasks()

    def per_app_reset(self, name):
        self.pb2.configure(value=0, maximum=100, mode="determinate"); self._spinner_start("Downloading",
                                                                                          name); self.root.update_idletasks()

    def per_app_update_percent(self, pct, name=None):
        pct = max(0, min(100, int(pct)));
        self.pb2.configure(value=pct)
        if name: self._spin_name = name
        self._spinner_set_pct(pct)
        if pct >= 100: self._spinner_stop(f"Download ({self._spin_name}): 100%")
        self.root.update_idletasks()

    def _spinner_start(self, base, name):
        self._spin_base = base;
        self._spin_name = name;
        self._spin_pct = 0;
        self._spin_index = 0
        if self._spin_job is None: self._spinner_tick()

    def _spinner_tick(self):
        frame = self._spin_frames[self._spin_index % len(self._spin_frames)];
        self._spin_index += 1
        name = f" ({self._spin_name})" if self._spin_name else "";
        pct = f" {self._spin_pct}%" if isinstance(self._spin_pct, int) else ""
        self.pb2_label.configure(text=f"{self._spin_base}{name} {frame}{pct}")
        self._spin_job = self.root.after(150, self._spinner_tick)

    def _spinner_set_pct(self, pct):
        self._spin_pct = max(0, min(100, int(pct)))

    def _spinner_stop(self, final="Download"):
        if self._spin_job is not None:
            try:
                self.root.after_cancel(self._spin_job)
            except Exception:
                pass
            self._spin_job = None
        self.pb2_label.configure(text=final)

    # ===== Selection & temp =====
    def _iter_items(self):
        return self.tree.get_children("")

    def select_all(self):
        if self.updating: return
        for i in self._iter_items():
            self.checked_items.add(i);
            self.tree.item(i, image=self.img_checked)
            tags = set(self.tree.item(i, "tags") or ());
            tags.add("checked");
            self.tree.item(i, tags=tuple(tags))
        self.update_counter()

    def select_none(self):
        if self.updating: return
        for i in self._iter_items():
            self.checked_items.discard(i);
            self.tree.item(i, image=self.img_unchecked)
            tags = set(self.tree.item(i, "tags") or ());
            tags.discard("checked");
            self.tree.item(i, tags=tuple(tags))
        self.update_counter()

    def update_counter(self):
        total = len(self.tree.get_children(""));
        selected = len(self.checked_items)
        self.counter_var.set(f"{total} apps found • {selected} selected")

    def clear_tree(self):
        self.checked_items.clear();
        self.id_to_item.clear()
        for i in self._iter_items(): self.tree.delete(i)

    def _temp_dir(self):
        return tempfile.gettempdir()

    def _fmt_bytes(self, n: int) -> str:
        units = ("B", "KB", "MB", "GB", "TB")
        v = float(n)
        for u in units:
            if v < 1024 or u == "TB":
                if u == "B":
                    return f"{int(v)} B"
                s = f"{v:.2f}".rstrip("0").rstrip(".")
                return f"{s} {u}"
            v /= 1024.0

    def _snapshot_temp(self):
        root = self._temp_dir()
        snap = {}
        exts = (".exe", ".msi", ".msix", ".msixbundle", ".appx", ".appxbundle", ".zip", ".7z", ".rar", ".cab")
        try:
            for dirpath, _, files in os.walk(root):
                for fn in files:
                    if os.path.splitext(fn)[1].lower() in exts:
                        p = os.path.join(dirpath, fn)
                        try:
                            st = os.lstat(p)
                            if st.st_size >= 1_000_000:  # only remember >= ~1MB
                                snap[p] = st.st_mtime
                        except Exception:
                            pass
        except Exception:
            pass
        return snap

    def _find_new_installer_files(self, b, a):
        exts = (".exe", ".msi", ".msix", ".msixbundle", ".appx", ".appxbundle", ".zip", ".7z", ".rar", ".cab")
        news = [p for p in a if (p not in b or a[p] > b.get(p, 0)) and os.path.splitext(p)[1].lower() in exts]
        return sorted(set(news), key=lambda p: a.get(p, 0), reverse=True)

    def _winget_downloads_for_id(self, package_id: str):
        """Fallback: scan %TEMP%\\WinGet for installers belonging to this package id."""
        hits = []
        base = os.path.join(self._temp_dir(), "WinGet")
        if not (package_id and os.path.isdir(base)):
            return hits
        pid_l = package_id.lower()
        exts = (".exe", ".msi", ".msix", ".msixbundle", ".appx", ".appxbundle", ".zip", ".7z", ".rar", ".cab")
        for dirpath, _, files in os.walk(base):
            # folders are typically: <base>\<id>.<version>\Installer-*.exe
            if pid_l in os.path.basename(dirpath).lower():
                for fn in files:
                    if os.path.splitext(fn)[1].lower() in exts:
                        hits.append(os.path.join(dirpath, fn))
        hits.sort(key=lambda p: os.path.getmtime(p), reverse=True)
        return hits

    def open_temp(self):
        try:
            os.startfile(self._temp_dir())
        except Exception as e:
            messagebox.showerror("Open Temp", str(e))

    def clear_temp_async(self):
        tmp = self._temp_dir()
        if not messagebox.askyesno("Clear Temp",
                                   f"This will delete files and folders inside:\n\n{tmp}\n\nItems in use will be skipped.\n\nProceed?"):
            return
        self.show_loading("Clearing Temp...")

        def worker():
            fdel = ddel = freed = 0

            def onerr(func, p, exc):
                try:
                    if os.path.islink(p): return
                    os.chmod(p, 0o666); func(p)
                except Exception:
                    pass

            try:
                for e in os.scandir(tmp):
                    p = e.path
                    try:
                        if e.is_file(follow_symlinks=False):
                            try:
                                freed += os.path.getsize(p)
                            except:
                                pass
                            try:
                                os.remove(p); fdel += 1; self.root.after(0, lambda
                                    s=f"[Clear Temp] Deleted file: {p}": self.log(s))
                            except Exception:
                                pass
                        elif e.is_dir(follow_symlinks=False):
                            size = 0
                            for dp, _, fns in os.walk(p):
                                for fn in fns:
                                    fp = os.path.join(dp, fn)
                                    try:
                                        size += os.path.getsize(fp)
                                    except:
                                        pass
                            try:
                                shutil.rmtree(p, onerror=onerr); ddel += 1; freed += size; self.root.after(0, lambda
                                    s=f"[Clear Temp] Deleted folder: {p}": self.log(s))
                            except Exception:
                                pass
                    except Exception:
                        pass
            except Exception as e:
                self.root.after(0, lambda: self.log(f"[Clear Temp] Error: {e}"))

            def done():
                self.hide_loading()
                self.log(f"[Clear Temp] Folders: {ddel} | Files: {fdel} | Freed: {self._fmt_bytes(freed)}")
                messagebox.showinfo("Clear Temp",
                                    f"Folders deleted: {ddel}\nFiles deleted: {fdel}\nSpace freed: {self._fmt_bytes(freed)}")

            self.root.after(0, done)

        threading.Thread(target=worker, daemon=True).start()

    def _menu_open_downloads(self):
        it = self._menu_item
        if not it or self.updating:
            return
        pid = self.tree.set(it, "Id")
        files = list(self.pkg_downloads.get(pid) or [])
        if not files:
            files = self._winget_downloads_for_id(pid)
        if not files:
            messagebox.showinfo("Downloads", "No downloaded files recorded for this app.")
            return
        for p in files:
            try:
                os.startfile(p if os.path.exists(p) else os.path.dirname(p) or self._temp_dir())
            except Exception:
                pass

    def _menu_delete_downloads(self):
        it = self._menu_item
        if not it or self.updating:
            return
        pid = self.tree.set(it, "Id")
        files = [p for p in (self.pkg_downloads.get(pid) or []) if os.path.exists(p)]
        if not files:
            # try WinGet cache if we never tracked it
            files = self._winget_downloads_for_id(pid)
        files = [p for p in files if os.path.exists(p)]
        if not files:
            messagebox.showinfo("Delete Downloads", "No downloaded files to delete for this app.")
            return
        if not messagebox.askyesno("Delete Downloads",
                                   "Delete the downloaded file(s) for this app?\n\n" + "\n".join(files)):
            return
        deleted = freed = 0
        for p in files:
            try:
                try:
                    freed += os.path.getsize(p)
                except Exception:
                    pass
                os.remove(p)
                deleted += 1
                self.log(f"[Delete] {p}")
            except Exception as e:
                self.log(f"[Delete] Could not delete {p}: {e}")
        # prune memory
        self.pkg_downloads[pid] = [p for p in self.pkg_downloads.get(pid, []) if os.path.exists(p)]
        self.log(f"[Delete] Deleted {deleted} file(s), freed {self._fmt_bytes(freed)}")

    # ===== Check for updates & populate =====
    def check_for_updates_async(self):
        include = bool(self.include_unknown_var.get())
        self.btn_check.config(state="disabled");
        self.show_loading("Checking for updates...")

        def worker():
            try:
                pkgs = get_winget_upgrades(include_unknown=include)
            except Exception as e:
                self.root.after(0, lambda: (self.hide_loading(), self.btn_check.config(state="normal"),
                                            self.counter_var.set("0 apps found • 0 selected"),
                                            messagebox.showerror("winget error", f"Failed to query updates:\n{e}"),
                                            self.log(f"[winget] {e}")))
                return
            self.root.after(0, lambda: self.populate_tree(pkgs))

        threading.Thread(target=worker, daemon=True).start()

    def populate_tree(self, pkgs):
        self.hide_loading();
        self.clear_tree();
        self.btn_check.config(state="normal")
        if not pkgs:
            self.counter_var.set("0 apps found • 0 selected")
            self.log("No apps need updating.")
            self._enable_select_buttons(False)
            return
        for p in pkgs:
            it = self.tree.insert("", "end", text="", image=self.img_unchecked,
                                  values=(p["name"], p["id"], p.get("current", ""), p.get("available", ""), ""))
            self.id_to_item[p["id"]] = it;
            self.checked_items.discard(it)
        self.update_counter();
        self.autofit_all();
        self._enable_select_buttons(True)

    # ===== Update selected (+ Cancel) =====
    def update_selected_async(self):
        if self.updating:
            self.cancel_requested = True;
            self.btn_update.config(text="Cancelling...", state="disabled")
            if self.current_proc and self.current_proc.poll() is None:
                try:
                    self.current_proc.terminate()
                except Exception:
                    pass
            return

        # Build targets in the current UI order (top → bottom)
        targets = []
        for it in self.tree.get_children(""):
            if it in self.checked_items:
                pid = self.tree.set(it, "Id")
                cur = (self.tree.set(it, "Current") or "").strip()
                if pid:
                    targets.append((pid, cur, it))
        if not targets: messagebox.showinfo("No Selection", "No apps selected for update."); return

        # Enable Skip only during updates and only if more than one target
        self.skip_requested = False
        self._can_skip = len(targets) > 1
        try:
            self.btn_skip.config(state=("normal" if self._can_skip else "disabled"))
        except Exception:
            pass
        # Clear old results before starting new update
        for it in self.tree.get_children(""):
            self.tree.set(it, "Result", "")
            # also remove any old tags (fail/ok/skip/cancel)
            self.tree.item(it, tags=())

        self._disable_controls_for_update()
        self.cancel_requested = False;
        self.current_proc = None
        self.log(f"Starting updates for {len(targets)} package(s)...")
        results = {}
        self.progress_start("Updating", len(targets))

        def worker():
            percent_re = re.compile(r"(\d{1,3})\s*%")
            size_re = re.compile(r"([\d\.]+)\s*(KB|MB|GB)\s*/\s*([\d\.]+)\s*(KB|MB|GB)", re.I)
            unit = {"KB": 1_000, "MB": 1_000_000, "GB": 1_000_000_000}
            spinner_re = re.compile(r"^[\s\\/\|\-\r\u2580-\u259F\u2500-\u257F]+$")

            for pid, cur, it in targets:
                if self.cancel_requested: break
                skip_now = False
                name = self.tree.set(it, "Name") or pid
                self.root.after(0, lambda n=name: self.per_app_reset(n))
                self.root.after(0, lambda p=pid: self.log(f"Updating {p} ..."))
                try:
                    cmd = ["winget", "upgrade", "--id", pid, "--accept-package-agreements",
                           "--accept-source-agreements", "--disable-interactivity", "-h"]
                    if self.include_unknown_var.get() or (not cur) or (cur.lower() == "unknown"): cmd.insert(2,
                                                                                                             "--include-unknown")
                    env = os.environ.copy();
                    env["DOTNET_CLI_UI_LANGUAGE"] = "en"
                    snap_before = self._snapshot_temp()
                    self.current_proc = subprocess.Popen(cmd, shell=False, text=True, stdout=subprocess.PIPE,
                                                         stderr=subprocess.PIPE,
                                                         encoding="utf-8", errors="replace", env=env,
                                                         startupinfo=_hidden_startupinfo(),
                                                         creationflags=CREATE_NO_WINDOW)
                    captured = [];
                    last = -1
                    while True:
                        if self.skip_requested and self.current_proc and self.current_proc.poll() is None:
                            self.skip_requested = False
                            skip_now = True
                            try:
                                self.current_proc.terminate()
                            except Exception:
                                pass
                            break
                        if self.cancel_requested and self.current_proc and self.current_proc.poll() is None:
                            try:
                                self.current_proc.terminate()
                            except Exception:
                                pass

                            def mark_canceled(item=it):
                                self.tree.set(item, "Result", "❌ Canceled");
                                self.tree.item(item, tags=("cancel",));
                                self._spinner_stop("Download")

                            self.root.after(0, mark_canceled);
                            break
                        line = self.current_proc.stdout.readline()
                        if not line: break
                        ln = line.rstrip()
                        if not ln or spinner_re.match(ln): continue
                        captured.append(ln);
                        self.root.after(0, lambda s=ln: self.log(s))
                        m = None
                        for m in percent_re.finditer(ln): pass
                        if m:
                            try:
                                pct = int(m.group(1))
                                if pct != last: last = pct; self.root.after(0, lambda p=pct,
                                                                                      n=name: self.per_app_update_percent(
                                    p, n))
                                continue
                            except Exception:
                                pass
                        m2 = size_re.search(ln)
                        if m2:
                            try:
                                hv, hu, tv, tu = m2.groups()
                                have = float(hv) * unit[hu.upper()];
                                tot = float(tv) * unit[tu.upper()]
                                if tot > 0:
                                    pct = max(0, min(100, int(round((have / tot) * 100))))
                                    if pct != last: last = pct; self.root.after(0, lambda p=pct,
                                                                                          n=name: self.per_app_update_percent(
                                        p, n))
                            except Exception:
                                pass
                    if self.current_proc and self.current_proc.poll() is None:
                        out, err = self.current_proc.communicate()
                    else:
                        out = err = ""
                        try:
                            out = (
                                        self.current_proc.stdout.read() or "") if self.current_proc and self.current_proc.stdout else ""
                            err = (
                                        self.current_proc.stderr.read() or "") if self.current_proc and self.current_proc.stderr else ""
                        except Exception:
                            pass
                    if err: self.root.after(0, lambda e=err: self.log(e.strip()))
                    snap_after = self._snapshot_temp()
                    new_files = self._find_new_installer_files(snap_before, snap_after)
                    if new_files:
                        self.pkg_downloads.setdefault(pid, [])
                        seen = set(self.pkg_downloads[pid])
                        for p in new_files:
                            if p not in seen: self.pkg_downloads[pid].append(p)
                        self.root.after(0,
                                        lambda nf=new_files, pp=pid: self.log(f"[Downloads] {pp}: " + " | ".join(nf)))
                    if self.cancel_requested and (
                            self.current_proc is None or self.current_proc.returncode is not None):
                        results[pid] = "canceled"
                    else:
                        rc = self.current_proc.returncode if self.current_proc else 1
                        joined = "\n".join(captured + ([out] if out else []))
                        low = (joined + "\n" + (err or "")).lower()
                        if skip_now:
                            st, tag = "⏭ Skipped", "skip";
                            results[pid] = "skipped"
                        elif rc != 0 or "failed" in low or "error" in low or "0x" in low:
                            st, tag = "❌ Failed", "fail";
                            results[pid] = "failed"
                        elif "no applicable update" in low or "no packages found" in low:
                            st, tag = "⏭ No update", "skip";
                            results[pid] = "skipped"
                        else:
                            st, tag = "✅ Success", "ok";
                            results[pid] = "success"
                            if self.pkg_downloads.get(pid):
                                st = "✅ Success (downloads)"

                        def apply(item=it, st=st, tg=tag):
                            if self.tree.set(item, "Result") in ("", None):
                                self.tree.set(item, "Result", st);
                                self.tree.item(item, tags=(tg,))

                        self.root.after(0, apply)
                except Exception as ex:
                    if self.cancel_requested:
                        results[pid] = "canceled"

                        def apply_canceled(item=it):
                            self.tree.set(item, "Result", "❌ Canceled");
                            self.tree.item(item, tags=("cancel",));
                            self._spinner_stop("Download")

                        self.root.after(0, apply_canceled)
                    else:
                        results[pid] = "failed";
                        self.root.after(0, lambda e=ex: self.log(f"Error: {e}"))

                        def apply_err(item=it):
                            self.tree.set(item, "Result", "❌ Failed")
                            self.tree.item(item, tags=("fail",))

                        self.root.after(0, apply_err)
                finally:
                    self.root.after(0, lambda p=pid: self.log(f"✔ Finished {p}"))
                    if not self.cancel_requested: self.root.after(0, lambda n=name: self.per_app_update_percent(100, n))
                    self.root.after(0, lambda: self.progress_step(1))

                    # Re-enable Skip for the next item if allowed
                    if self._can_skip:
                        self.root.after(0, lambda: self.btn_skip.config(state="normal"))

            def done():
                canceled = self.cancel_requested
                if canceled:
                    for _, _, it in targets:
                        if not self.tree.set(it, "Result"):
                            self.tree.set(it, "Result", "❌ Canceled");
                            self.tree.item(it, tags=("cancel",))
                    self.log("Cancelled by user.")
                ok = sum(1 for s in results.values() if s == "success")
                fail = sum(1 for s in results.values() if s == "failed")
                skip = sum(1 for s in results.values() if s == "skipped")
                canc = sum(1 for s in results.values() if s == "canceled")
                self.log(f"Summary → ✅ {ok} Success • ❌ {fail} Failed • ⏭ {skip} Skipped • ❌ {canc} Canceled")
                if fail == 0 and not canceled: play_success_sound()
                self.cancel_requested = False;
                self.current_proc = None
                try:
                    self.btn_skip.config(state="disabled")
                except Exception:
                    pass
                self._enable_controls_after_update();
                self.progress_finish(canceled=canceled)

            self.root.after(0, done)

        threading.Thread(target=worker, daemon=True).start()

    def _enable_select_buttons(self, enable: bool):
        state = "normal" if enable else "disabled"
        try:
            self.btn_sel_all.config(state=state)
        except Exception:
            pass
        try:
            self.btn_sel_none.config(state=state)
        except Exception:
            pass

    def skip_current(self):
        """Skip only the current app during an ongoing batch."""
        if not getattr(self, "updating", False):
            return
        self.skip_requested = True
        try:
            self.btn_skip.config(state="disabled")
        except Exception:
            pass
        self.log("[Skip] Skip requested for current app...")

    # ===== Logging =====
    def log(self, text):
        self.log_box.insert(tk.END, text + "\n")
        line_count = int(self.log_box.index("end-1c").split(".")[0])
        if line_count > 5000:
            self.log_box.delete("1.0", f"{line_count - 5000 + 1}.0")
        self.log_box.see(tk.END)
        self.root.update_idletasks()


# ===================== main =====================
if __name__ == "__main__":
    root = tk.Tk()
    app = WingetUpdaterUI(root)
    root.mainloop()