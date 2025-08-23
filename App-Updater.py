import json
import re
import subprocess
import threading
import tkinter as tk
from tkinter import messagebox, ttk
import sys
import os
import ctypes
import winsound
from io import BytesIO
from typing import Optional

# ====================== App Constants ======================
APP_NAME_VERSION = "Windows App Updater v1.0"
UNCHECKED = "☐"
CHECKED = "☑"

# ====================== Elevation helpers ======================
def is_admin() -> bool:
    try:
        return bool(ctypes.windll.shell32.IsUserAnAdmin())
    except Exception:
        return False

def relaunch_as_admin():
    """Relaunch the current program elevated via UAC and exit the non-elevated instance."""
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

    def show_tip(self, event=None):
        if self.tip_window or not self.text:
            return
        # position near the widget
        x = self.widget.winfo_rootx() + 25
        y = self.widget.winfo_rooty() + self.widget.winfo_height() + 10
        self.tip_window = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(True)
        tw.wm_geometry(f"+{x}+{y}")
        label = tk.Label(
            tw,
            text=self.text,
            justify="left",
            background="#ffffe0",
            relief="solid",
            borderwidth=1,
            font=("Segoe UI", 9)
        )
        label.pack(ipadx=6, ipady=2)

    def hide_tip(self, event=None):
        if self.tip_window:
            self.tip_window.destroy()
            self.tip_window = None

# ====================== Icon & sound helpers ======================
def set_app_icon(root: tk.Tk) -> Optional[str]:
    """
    Set the window icon from windows-updater.ico / app.ico if available.
    Returns the path used (or None) so we can apply the same icon to child windows.
    """
    here = os.path.dirname(os.path.abspath(sys.argv[0]))
    candidates = [
        os.path.join(here, "windows-updater.ico"),
        os.path.join(here, "app.ico"),
        os.path.join(here, "assets", "windows-updater.ico"),
    ]
    for ico in candidates:
        if os.path.exists(ico):
            try:
                root.iconbitmap(ico)
                return ico
            except Exception:
                pass
    return None  # keep default if load fails

def apply_icon_to_toplevel(tlv: tk.Toplevel, icon_path: Optional[str]):
    if not icon_path:
        return
    try:
        tlv.iconbitmap(icon_path)
    except Exception:
        pass

def load_flag_image() -> Optional[tk.PhotoImage]:
    """
    Load 'kuwait.png' (preferred) or convert 'kuwait.ico' in-memory using Pillow if available.
    Returns a PhotoImage or None.
    """
    here = os.path.dirname(os.path.abspath(sys.argv[0]))
    pngs = [os.path.join(here, "kuwait.png"), os.path.join(here, "assets", "kuwait.png")]
    icos = [os.path.join(here, "kuwait.ico"), os.path.join(here, "assets", "kuwait.ico")]

    for p in pngs:
        if os.path.exists(p):
            try:
                return tk.PhotoImage(file=p)
            except Exception:
                pass

    for ico in icos:
        if os.path.exists(ico):
            try:
                from PIL import Image
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
    """Play success jingle. If 'success.wav' exists (or assets/success.wav), play it; else use system chime."""
    here = os.path.dirname(os.path.abspath(sys.argv[0]))
    candidates = [
        os.path.join(here, "success.wav"),
        os.path.join(here, "assets", "success.wav"),
    ]
    for wav in candidates:
        if os.path.exists(wav):
            try:
                winsound.PlaySound(wav, winsound.SND_FILENAME | winsound.SND_ASYNC)
                return
            except Exception:
                break
    try:
        winsound.MessageBeep(winsound.MB_ICONASTERISK)
    except Exception:
        pass

# ====================== winget helpers ======================
def run(cmd):
    """Run a process and return (code, stdout, stderr) decoded as UTF-8 (safe for winget Unicode)."""
    env = os.environ.copy()
    env["DOTNET_CLI_UI_LANGUAGE"] = "en"  # English output for dotnet entries
    p = subprocess.run(
        cmd,
        capture_output=True,
        text=True,
        shell=False,
        encoding="utf-8",
        errors="replace",
        env=env,
    )
    return p.returncode, p.stdout.strip(), p.stderr.strip()

def try_json_parsers(include_unknown: bool):
    """Try multiple winget JSON outputs (varies by winget versions)."""
    base = ["--accept-source-agreements", "--disable-interactivity", "--output", "json"]
    flag = ["--include-unknown"] if include_unknown else []
    attempts = [
        ["winget", "upgrade", *flag, *base],
        ["winget", "list", "--upgrade-available", *base],   # older phrasing
        ["winget", "list", "--upgrades", *base],            # newer phrasing
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
    """Normalize winget JSON shapes into list of dicts: name/id/current/available."""
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
        name = it.get("PackageName") or it.get("Name") or ""
        pkg_id = it.get("PackageIdentifier") or it.get("Id") or ""
        available = it.get("AvailableVersion") or it.get("Available") or ""
        current = it.get("Version") or it.get("InstalledVersion") or ""
        if name and pkg_id and available:
            items.append({"name": name, "id": pkg_id, "available": available, "current": current})
    return items

def parse_table_upgrade_output(text):
    """Fallback parser for table output of 'winget upgrade'."""
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
            current = parts[2] if len(parts) > 2 else ""
            available = parts[3] if len(parts) > 3 else ""
        else:
            name, pkg_id = parts[0], parts[1]
            current, available = "", parts[2]
        if name and pkg_id and available and not name.startswith("-"):
            items.append({"name": name, "id": pkg_id, "current": current, "available": available})
    return items

def get_winget_upgrades(include_unknown: bool):
    """Robust getter: try JSON; fallback to table; raise if all fail."""
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

# ====================== UI Class ======================
class WingetUpdaterUI:
    def __init__(self, root):
        self.root = root
        self.root.title(APP_NAME_VERSION)
        self.root.geometry("1280x900")     # larger so the log is clearly visible
        self.root.minsize(1180, 830)

        self.updating = False
        self.cancel_requested = False
        self.current_proc = None
        self.loading_win = None
        self.window_icon_path = set_app_icon(self.root)

        # ===== Header =====
        header = ttk.Frame(self.root)
        header.pack(fill="x", pady=(10, 0))
        ttk.Label(header, text=APP_NAME_VERSION, font=("Segoe UI", 18, "bold")).pack(side="left", padx=12)

        right = ttk.Frame(header)
        right.pack(side="right", padx=12)
        self.btn_admin = ttk.Button(right, text="Run as Admin", command=self.run_as_admin)
        self.btn_admin.pack(side="right")
        if is_admin():
            self.btn_admin.config(text="Running as Admin", state="disabled")
        else:
            ToolTip(self.btn_admin, "Run the app as admin if you like to install all apps silently")

        # ===== Top controls =====
        top = ttk.Frame(self.root)
        top.pack(fill="x", padx=12, pady=6)

        self.btn_check = ttk.Button(top, text="Check for Updates", command=self.check_for_updates_async)
        self.btn_check.pack(side="left")

        # Include unknown toggle
        self.include_unknown_var = tk.BooleanVar(value=False)
        self.chk_unknown = ttk.Checkbutton(top, text="Include unknown apps", variable=self.include_unknown_var)
        self.chk_unknown.pack(side="left", padx=(10, 0))

        ttk.Button(top, text="Select All", command=self.select_all).pack(side="left", padx=(10, 0))
        ttk.Button(top, text="Select None", command=self.select_none).pack(side="left", padx=(6, 0))

        self.btn_update = ttk.Button(top, text="Update Selected", command=self.update_selected_async)
        self.btn_update.pack(side="right")

        # Counter
        self.counter_var = tk.StringVar(value="0 apps found • 0 selected")
        ttk.Label(self.root, textvariable=self.counter_var).pack(anchor="w", padx=12)

        # ===== Tree with both scrollbars =====
        tree_wrap = ttk.Frame(self.root)
        tree_wrap.pack(fill="both", expand=True, padx=12, pady=(8, 8))

        cols = ("Select", "Name", "Id", "Current", "Available")
        self.tree = ttk.Treeview(tree_wrap, columns=cols, show="headings", height=22)
        self.tree.heading("Select", text="Select", anchor="center")
        self.tree.heading("Name", text="Name", anchor="w")
        self.tree.heading("Id", text="Id", anchor="w")
        self.tree.heading("Current", text="Current", anchor="center")
        self.tree.heading("Available", text="Available", anchor="center")

        self.tree.column("Select", width=72, anchor="center", stretch=False)
        self.tree.column("Name", width=520, anchor="w", stretch=True)
        self.tree.column("Id", width=560, anchor="w", stretch=True)
        self.tree.column("Current", width=100, anchor="center", stretch=False)
        self.tree.column("Available", width=100, anchor="center", stretch=False)

        ysb = ttk.Scrollbar(tree_wrap, orient="vertical", command=self.tree.yview)
        xsb = ttk.Scrollbar(tree_wrap, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscroll=ysb.set, xscroll=xsb.set)

        self.tree.grid(row=0, column=0, sticky="nsew")
        ysb.grid(row=0, column=1, sticky="ns")
        xsb.grid(row=1, column=0, sticky="ew")
        tree_wrap.rowconfigure(0, weight=1)
        tree_wrap.columnconfigure(0, weight=1)

        self.tree.bind("<Button-1>", self.on_tree_click)

        # ===== YIFY-style single progress bar + label =====
        pb_wrap = ttk.Frame(self.root)
        pb_wrap.pack(fill="x", padx=12, pady=(0, 4))
        self.pb_label = ttk.Label(pb_wrap, text="Idle")
        self.pb_label.pack(side="left")
        self.pb = ttk.Progressbar(pb_wrap, orient="horizontal", mode="determinate")
        self.pb.pack(fill="x", expand=True, padx=10)

        # ===== Signature above log =====
        sig_frame = ttk.Frame(self.root)
        sig_frame.pack(fill="x", padx=12, pady=(4, 0))
        ttk.Label(sig_frame, text="").pack(side="left", expand=True)  # spacer
        self.flag_img = load_flag_image()
        if self.flag_img:
            tk.Label(sig_frame, image=self.flag_img).pack(side="right", padx=(6, 0))
        ttk.Label(sig_frame, text="Made by BoYaqoub - ilukezippo@gmail.com", font=("Segoe UI", 9)).pack(side="right")

        # ===== Log label + Log with both scrollbars =====
        ttk.Label(self.root, text="Update Log:", font=("Segoe UI", 10, "bold")).pack(anchor="w", padx=12, pady=(6, 2))

        log_wrap = ttk.Frame(self.root)
        log_wrap.pack(fill="both", expand=False, padx=12, pady=(0, 10))

        self.log_box = tk.Text(log_wrap, height=14, wrap="none", font=("Consolas", 10))
        log_ysb = ttk.Scrollbar(log_wrap, orient="vertical", command=self.log_box.yview)
        log_xsb = ttk.Scrollbar(log_wrap, orient="horizontal", command=self.log_box.xview)
        self.log_box.configure(yscrollcommand=log_ysb.set, xscrollcommand=log_xsb.set)

        self.log_box.grid(row=0, column=0, sticky="nsew")
        log_ysb.grid(row=0, column=1, sticky="ns")
        log_xsb.grid(row=1, column=0, sticky="ew")
        log_wrap.rowconfigure(0, weight=1)
        log_wrap.columnconfigure(0, weight=1)

        self.root.after(0, self.center_on_screen)

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

    # ====================== Progress helpers (YIFY-style) ======================
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

    # ====================== Events & helpers ======================
    def on_tree_click(self, event):
        if self.tree.identify("region", event.x, event.y) != "cell":
            return
        if self.tree.identify_column(event.x) != "#1":
            return
        item = self.tree.identify_row(event.y)
        if not item:
            return
        v = self.tree.set(item, "Select")
        self.tree.set(item, "Select", CHECKED if v != CHECKED else UNCHECKED)
        self.update_counter()

    def select_all(self):
        for item in self.tree.get_children():
            self.tree.set(item, "Select", CHECKED)
        self.update_counter()

    def select_none(self):
        for item in self.tree.get_children():
            self.tree.set(item, "Select", UNCHECKED)
        self.update_counter()

    def update_counter(self):
        total = selected = 0
        for item in self.tree.get_children():
            total += 1
            if self.tree.set(item, "Select") == CHECKED:
                selected += 1
        self.counter_var.set(f"{total} apps found • {selected} selected")

    def clear_tree(self):
        for i in self.tree.get_children():
            self.tree.delete(i)
        # no external item map; we read values directly from the tree

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

        # Flat list (sorted by name)
        for p in sorted(pkgs, key=lambda x: x["name"].lower()):
            self.tree.insert(
                "", "end",
                values=(UNCHECKED, p["name"], p["id"], p.get("current", ""), p.get("available", "")),
            )
        self.update_counter()

    # ====================== Update selected (async + Cancel) ======================
    def update_selected_async(self):
        # If currently updating, treat as Cancel
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

        # Gather selection
        ids = []
        for item in self.tree.get_children():
            if self.tree.set(item, "Select") == CHECKED:
                ids.append(self.tree.set(item, "Id"))
        if not ids:
            messagebox.showinfo("No Selection", "No apps selected for update.")
            return

        # Switch to updating mode
        self.updating = True
        self.cancel_requested = False
        self.current_proc = None
        self.btn_check.config(state="disabled")
        self.btn_update.config(text="Cancel", state="normal")
        self.log(f"Starting updates for {len(ids)} package(s)...")

        # YIFY-style progress
        self.progress_start("Updating", len(ids))

        def worker():
            CREATE_NO_WINDOW = 0x08000000  # hide console window on Windows
            for pkg_id in ids:
                if self.cancel_requested:
                    break
                self.root.after(0, lambda pid=pkg_id: self.log(f"Updating {pid} ..."))
                try:
                    self.current_proc = subprocess.Popen(
                        ["winget", "upgrade", "--id", pkg_id,
                         "--accept-package-agreements", "--accept-source-agreements",
                         "--disable-interactivity", "-h"],
                        shell=False,
                        stdout=subprocess.PIPE,
                        stderr=subprocess.PIPE,
                        text=True,
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

            # wrap-up on UI thread
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
