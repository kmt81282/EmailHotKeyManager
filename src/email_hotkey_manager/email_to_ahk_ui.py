import json
import subprocess
import sys
import time
from pathlib import Path

import tkinter as tk
from tkinter import filedialog, ttk, messagebox

import pyautogui
import keyboard  # global hook on Windows


# =========================
# Imports (new src/ package layout + backward compat)
# =========================
try:
    from . import email_to_ahk
except Exception:
    import email_to_ahk  # type: ignore

try:
    from .version import VERSION
except Exception:
    VERSION = "1.1.1"


# =========================
# App storage
# =========================
APP_DIR = Path.home() / ".email_hotkey_generator"
AHK_DIR = APP_DIR / "ahk"

APP_DIR.mkdir(exist_ok=True)
AHK_DIR.mkdir(exist_ok=True)

# Track running AutoHotkey processes per profile name
running_ahk: dict[str, subprocess.Popen] = {}

# Active hotkey state (global listener sends to whichever hotkey window is active)
armed_hotkey: str | None = None
active_status_setter = None  # callable(str, color)
active_armed_ui_setter = None  # callable(hotkey|None)


# =========================
# Helpers
# =========================
def friendly_file_type(path: str) -> str:
    ext = Path(path).suffix.lower()
    if ext == ".csv":
        return "CSV"
    if ext == ".txt":
        return "TXT"
    if ext == ".eml":
        return "EML"
    return "—"


def manager_flash(msg: str):
    manager_msg_var.set(msg)
    root.after(4500, lambda: manager_msg_var.set(""))


def get_ahk_runtime_exe() -> Path | None:
    """
    We run AutoHotkey explicitly so we can terminate it when the hotkey window closes.
    This keeps "bundled runtime" strategy while enabling lifecycle control.
    """
    # PyInstaller onefile/onedir
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        p = Path(sys._MEIPASS) / "ahk_runtime" / "AutoHotkey.exe"
        if p.exists():
            return p

    # Dev repo layout (assets/)
    p = Path.cwd() / "assets" / "ahk_runtime" / "AutoHotkey.exe"
    if p.exists():
        return p

    # Fallback: maybe present in repo root (older layouts)
    p = Path.cwd() / "ahk_runtime" / "AutoHotkey.exe"
    if p.exists():
        return p

    return None


def ensure_ahk_running(profile_name: str, ahk_script_path: Path, set_status):
    """
    Start AutoHotkey process for a given script if not already running.
    We start via AutoHotkey.exe directly to allow termination.
    """
    if profile_name in running_ahk and running_ahk[profile_name].poll() is None:
        return

    ahk_exe = get_ahk_runtime_exe()
    if ahk_exe is None:
        set_status(
            "AutoHotkey runtime not found (expected ./ahk_runtime/AutoHotkey.exe).",
            "red",
        )
        raise FileNotFoundError("AutoHotkey runtime not found.")

    # Start AHK script
    proc = subprocess.Popen(
        [str(ahk_exe), str(ahk_script_path)],
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
        stdin=subprocess.DEVNULL,
        creationflags=subprocess.CREATE_NO_WINDOW if sys.platform.startswith("win") else 0,
    )
    running_ahk[profile_name] = proc
    set_status(f"AutoHotkey started for profile: {profile_name}", "blue")


def stop_ahk(profile_name: str):
    proc = running_ahk.get(profile_name)
    if not proc:
        return
    try:
        if proc.poll() is None:
            proc.terminate()
            # Give it a moment; if stubborn, kill
            for _ in range(15):
                if proc.poll() is not None:
                    break
                time.sleep(0.05)
            if proc.poll() is None:
                proc.kill()
    finally:
        running_ahk.pop(profile_name, None)


def send_hotkey(hotkey: str):
    """
    Send Ctrl+Shift+<key> safely using pyautogui.
    """
    parts = hotkey.lower().split("+")
    modifiers = parts[:-1]
    key = parts[-1]

    for m in modifiers:
        pyautogui.keyDown(m)
        time.sleep(0.02)

    pyautogui.press(key)

    for m in reversed(modifiers):
        time.sleep(0.02)
        pyautogui.keyUp(m)


def arm_hotkey(hotkey: str, set_status):
    global armed_hotkey
    armed_hotkey = hotkey
    if active_armed_ui_setter:
        active_armed_ui_setter(hotkey)
    set_status("Armed. Switch to target field and press F8.", "blue")


def clear_armed(set_status=None):
    global armed_hotkey
    armed_hotkey = None
    if active_armed_ui_setter:
        active_armed_ui_setter(None)
    if set_status:
        set_status("Ready. Choose a hotkey and click ARM.", "black")


def on_trigger_pressed(_event):
    """
    Global F8 handler. Sends armed hotkey and suppresses F8.
    """
    global armed_hotkey

    if armed_hotkey:
        try:
            send_hotkey(armed_hotkey)
            if active_status_setter:
                active_status_setter(f"Sent {armed_hotkey}", "green")
        finally:
            # Disarm after use
            if active_armed_ui_setter:
                active_armed_ui_setter(None)
            armed_hotkey = None
        return False  # suppress F8


# Register global listener ONCE
keyboard.on_press_key("f8", on_trigger_pressed)


# =========================
# Scrollable Frame
# =========================
class ScrollableFrame(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent)

        canvas = tk.Canvas(self, highlightthickness=0)
        scrollbar = ttk.Scrollbar(self, orient="vertical", command=canvas.yview)
        self.inner = ttk.Frame(canvas)

        window = canvas.create_window((0, 0), window=self.inner, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        self.inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind("<Configure>", lambda e: canvas.itemconfigure(window, width=e.width))
        canvas.bind_all("<MouseWheel>", lambda e: canvas.yview_scroll(int(-e.delta / 120), "units"))


# =========================
# Profile metadata (non-breaking)
# =========================
def meta_path_for(ahk_path: Path) -> Path:
    return ahk_path.with_suffix(".meta.json")


def load_meta(ahk_path: Path) -> dict:
    p = meta_path_for(ahk_path)
    if p.exists():
        try:
            return json.loads(p.read_text(encoding="utf-8"))
        except Exception:
            return {}
    return {}


def save_meta(ahk_path: Path, meta: dict):
    p = meta_path_for(ahk_path)
    p.write_text(json.dumps(meta, indent=2), encoding="utf-8")


# =========================
# Profiles
# =========================
def load_profiles():
    profiles = []
    for ahk in AHK_DIR.glob("*.ahk"):
        manifest = ahk.with_suffix(".json")
        if manifest.exists():
            # Count hotkeys
            hotkey_count = None
            try:
                data = json.loads(manifest.read_text(encoding="utf-8"))
                hotkey_count = len(data)
            except Exception:
                hotkey_count = None

            meta = load_meta(ahk)
            source_type = meta.get("source_type", "—")

            profiles.append({
                "name": ahk.stem,
                "ahk": ahk,
                "manifest": manifest,
                "source_type": source_type,
                "hotkey_count": hotkey_count,
            })
    # Stable ordering
    profiles.sort(key=lambda p: p["name"].lower())
    return profiles


def delete_profile(profile, row_frame):
    name = profile["name"]
    try:
        stop_ahk(name)

        profile["ahk"].unlink(missing_ok=True)
        profile["manifest"].unlink(missing_ok=True)
        meta_path_for(profile["ahk"]).unlink(missing_ok=True)

        row_frame.destroy()
        manager_flash(f"Deleted profile: {name}")
    except Exception as e:
        messagebox.showerror("Delete failed", f"Failed to delete {name}:\n\n{e}")


# =========================
# Hotkey Window
# =========================
def open_hotkey_window(profile):
    """
    Hotkey window is the execution surface:
    - Shows runtime status bar
    - Owns AHK lifecycle (closing window terminates AHK)
    - Highlights the armed hotkey row, disables other ARM buttons
    """
    window = tk.Toplevel(root)
    window.title(f"Hotkeys – {profile['name']}")
    window.geometry("520x600")
    window.resizable(False, False)

    # Top instructions
    header = ttk.Frame(window, padding=(14, 12))
    header.pack(fill="x")

    ttk.Label(header, text=f"Profile: {profile['name']}", font=("Segoe UI", 12, "bold")).pack(anchor="w")
    ttk.Label(
        header,
        text="Click ARM → focus your target field → press F8",
        foreground="#444",
    ).pack(anchor="w", pady=(4, 0))

    # Hotkeys list
    list_frame = ttk.Frame(window, padding=(14, 8))
    list_frame.pack(fill="both", expand=True)

    hotkey_scroll = ScrollableFrame(list_frame)
    hotkey_scroll.pack(fill="both", expand=True)

    # Status bar (moved here)
    status_var = tk.StringVar(value="Ready. Choose a hotkey and click ARM.")
    status = tk.Label(window, textvariable=status_var, anchor="w", relief=tk.SUNKEN, padx=10)
    status.pack(side=tk.BOTTOM, fill=tk.X)

    def set_status(msg, color="black"):
        status.config(fg=color)
        status_var.set(msg)
        status.update_idletasks()

    # Start AHK process now (execution window owns lifecycle)
    try:
        ensure_ahk_running(profile["name"], profile["ahk"], set_status)
    except Exception:
        # If runtime missing, still allow viewing manifest; but ARM should fail to be useful
        pass

    # Load manifest
    try:
        manifest = json.loads(profile["manifest"].read_text(encoding="utf-8"))
    except Exception as e:
        set_status(f"Failed to load manifest: {e}", "red")
        manifest = {}

    # Track UI row widgets to support highlight/disable
    row_widgets = {}  # hotkey -> {"frame": tk.Frame, "btn": tk.Button}

    # Use tk widgets (easier to highlight background than ttk)
    for hotkey, info in manifest.items():
        row = tk.Frame(hotkey_scroll.inner, bd=1, relief="flat")
        row.pack(fill="x", pady=4)

        label = tk.Label(row, text=info.get("label", hotkey), anchor="w")
        label.pack(side="left", fill="x", expand=True, padx=(6, 6), pady=4)

        btn = tk.Button(
            row,
            text="ARM",
            width=8,
            command=lambda h=hotkey: arm_hotkey(h, set_status),
        )
        btn.pack(side="right", padx=(6, 0), pady=4)

        row_widgets[hotkey] = {"frame": row, "btn": btn, "label": label}

    def set_armed_ui(armed: str | None):
        # Highlight armed row & disable other ARM buttons
        for hk, w in row_widgets.items():
            is_armed = (armed is not None and hk == armed)
            if is_armed:
                w["frame"].config(bg="#e9f2ff")
                w["label"].config(bg="#e9f2ff")
                w["btn"].config(state="normal")
            else:
                w["frame"].config(bg=window.cget("bg"))
                w["label"].config(bg=window.cget("bg"))
                w["btn"].config(state="disabled" if armed else "normal")

    # Make this window the active execution surface
    global active_status_setter, active_armed_ui_setter
    active_status_setter = set_status
    active_armed_ui_setter = set_armed_ui
    set_armed_ui(armed_hotkey)

    def on_close():
        # If this window is active, clear global execution callbacks & disarm
        global active_status_setter, active_armed_ui_setter
        if active_status_setter == set_status:
            active_status_setter = None
        if active_armed_ui_setter == set_armed_ui:
            active_armed_ui_setter = None

        clear_armed()

        # Terminate AHK for this profile (requested behavior)
        stop_ahk(profile["name"])
        window.destroy()

    window.protocol("WM_DELETE_WINDOW", on_close)


# =========================
# Browse + Generate
# =========================
def browse_input():
    path = filedialog.askopenfilename(
        title="Select Input File",
        filetypes=[
            ("Supported Files", "*.csv *.txt *.eml"),
            ("CSV Files", "*.csv"),
            ("Text Files", "*.txt"),
            ("Email Files", "*.eml"),
        ],
    )
    if path:
        input_entry.delete(0, tk.END)
        input_entry.insert(0, path)
        input_type_var.set(f"Input type: {friendly_file_type(path)}")
        manager_flash(f"Selected {friendly_file_type(path)} input.")


def run_generator():
    input_path = input_entry.get().strip()
    profile_name = profile_entry.get().strip()

    if not Path(input_path).is_file():
        messagebox.showerror("Invalid input", "Input file not found.")
        return

    if not profile_name:
        messagebox.showerror("Missing profile name", "Profile name is required.")
        return

    safe = "".join(c for c in profile_name if c.isalnum() or c in " _-").strip()
    if not safe:
        messagebox.showerror("Invalid profile name", "Profile name must contain letters/numbers.")
        return

    ahk_path = AHK_DIR / f"{safe}.ahk"
    manifest_path = ahk_path.with_suffix(".json")

    try:
        email_to_ahk.run(input_path, ahk_path)

        # Compute hotkey count from manifest
        hotkey_count = None
        try:
            if manifest_path.exists():
                data = json.loads(manifest_path.read_text(encoding="utf-8"))
                hotkey_count = len(data)
        except Exception:
            hotkey_count = None

        # Save small meta file for UI display (non-breaking)
        save_meta(ahk_path, {
            "source_type": friendly_file_type(input_path),
            "source_path": input_path,
            "version": VERSION,
        })

        profile = {
            "name": safe,
            "ahk": ahk_path,
            "manifest": manifest_path,
            "source_type": friendly_file_type(input_path),
            "hotkey_count": hotkey_count,
        }

        # Refresh list (simple, reliable)
        rebuild_profiles_list()
        manager_flash(
            f"Profile '{safe}' created ({profile['source_type']}"
            + (f", {hotkey_count} hotkeys)." if hotkey_count is not None else ").")
        )

        input_entry.delete(0, tk.END)
        profile_entry.delete(0, tk.END)
        input_type_var.set("Input type: —")

    except Exception as e:
        messagebox.showerror("Generation failed", str(e))


# =========================
# Manager UI
# =========================
root = tk.Tk()
root.title(f"Email Hotkey Manager v{VERSION}")
root.geometry("820x640")
root.resizable(False, False)

# Styles (simple hierarchy)
style = ttk.Style()
try:
    style.theme_use("clam")
except Exception:
    pass

# Header
header = ttk.Frame(root, padding=(18, 16))
header.pack(fill="x")

ttk.Label(header, text="Email Hotkey Manager", font=("Segoe UI", 16, "bold")).pack(anchor="w")
ttk.Label(
    header,
    text="Generate domain-based AutoHotkey shortcuts for fast, safe email pasting.",
    foreground="#444",
).pack(anchor="w", pady=(4, 0))

# Main content container
content = ttk.Frame(root, padding=(18, 10))
content.pack(fill="both", expand=True)

# Section: Input
input_card = ttk.LabelFrame(content, text="Input", padding=(14, 12))
input_card.pack(fill="x", pady=(0, 12))

ttk.Label(input_card, text="Select a file (.csv, .txt, or .eml):").grid(row=0, column=0, sticky="w")

row = ttk.Frame(input_card)
row.grid(row=1, column=0, sticky="ew", pady=(8, 0))
row.columnconfigure(0, weight=1)

input_entry = ttk.Entry(row)
input_entry.grid(row=0, column=0, sticky="ew")

ttk.Button(row, text="Browse…", command=browse_input).grid(row=0, column=1, padx=(10, 0))

input_type_var = tk.StringVar(value="Input type: —")
ttk.Label(input_card, textvariable=input_type_var, foreground="#444").grid(row=2, column=0, sticky="w", pady=(8, 0))

# Section: Profile generation
gen_card = ttk.LabelFrame(content, text="Create Profile", padding=(14, 12))
gen_card.pack(fill="x", pady=(0, 12))

ttk.Label(gen_card, text="Profile name:").grid(row=0, column=0, sticky="w")
profile_entry = ttk.Entry(gen_card, width=40)
profile_entry.grid(row=1, column=0, sticky="w", pady=(6, 0))

ttk.Button(
    gen_card,
    text="Generate Hotkeys",
    command=run_generator
).grid(row=1, column=1, padx=(14, 0), pady=(6, 0), sticky="w")

manager_msg_var = tk.StringVar(value="")
ttk.Label(gen_card, textvariable=manager_msg_var, foreground="#0a5").grid(row=2, column=0, columnspan=2, sticky="w", pady=(10, 0))

# Section: Profiles list
profiles_card = ttk.LabelFrame(content, text="Saved Profiles", padding=(14, 10))
profiles_card.pack(fill="both", expand=True)

profiles_scroll = ScrollableFrame(profiles_card)
profiles_scroll.pack(fill="both", expand=True)


def profile_row_text(p: dict) -> str:
    t = p.get("source_type", "—")
    c = p.get("hotkey_count", None)
    if isinstance(c, int):
        return f"{p['name']}   •   {t}   •   {c} hotkeys"
    return f"{p['name']}   •   {t}"


def rebuild_profiles_list():
    # Clear
    for w in profiles_scroll.inner.winfo_children():
        w.destroy()

    for p in load_profiles():
        row = ttk.Frame(profiles_scroll.inner)
        row.pack(fill="x", padx=4, pady=4)

        btn = ttk.Button(
            row,
            text=profile_row_text(p),
            command=lambda prof=p: open_hotkey_window(prof)
        )
        btn.pack(side="left", fill="x", expand=True)

        ttk.Button(
            row,
            text="Delete",
            width=10,
            command=lambda prof=p, rf=row: delete_profile(prof, rf)
        ).pack(side="right", padx=(10, 0))


rebuild_profiles_list()

root.mainloop()
