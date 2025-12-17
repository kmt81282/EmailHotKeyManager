import tkinter as tk
from tkinter import filedialog, ttk
import subprocess
import json
from pathlib import Path
import pyautogui
import keyboard   # ✅ correct library
import time

import email_to_ahk


# =========================
# App storage
# =========================

APP_DIR = Path.home() / ".email_hotkey_generator"
AHK_DIR = APP_DIR / "ahk"

APP_DIR.mkdir(exist_ok=True)
AHK_DIR.mkdir(exist_ok=True)

running_ahk = {}
armed_hotkey = None


# =========================
# Helpers
# =========================

def set_status(msg, color="black"):
    status_label.config(text=msg, fg=color)
    status_label.update_idletasks()


def ensure_ahk_running(profile):
    name = profile["name"]
    if name not in running_ahk:
        subprocess.Popen(["explorer.exe", str(profile["ahk"])], shell=False)
        running_ahk[name] = True
        set_status(f"Started profile: {name}", "blue")

def send_hotkey(hotkey):
    parts = hotkey.lower().split("+")
    modifiers = parts[:-1]
    key = parts[-1]

    # Press modifiers explicitly
    for m in modifiers:
        pyautogui.keyDown(m)
        time.sleep(0.02)

    # Press main key
    pyautogui.press(key)

    # Release modifiers (reverse order)
    for m in reversed(modifiers):
        time.sleep(0.02)
        pyautogui.keyUp(m)

    set_status(f"Sent {hotkey}", "green")


# =========================
# Arm & Trigger logic (FIXED)
# =========================

def arm_hotkey(hotkey):
    global armed_hotkey
    armed_hotkey = hotkey
    set_status(
        "Hotkey armed — switch to target field and press F8",
        "blue"
    )


def on_trigger_pressed(event):
    global armed_hotkey

    if armed_hotkey:
        send_hotkey(armed_hotkey)
        armed_hotkey = None
        return False   # suppress F8 completely


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
# Profiles
# =========================

def load_profiles():
    profiles = []
    for ahk in AHK_DIR.glob("*.ahk"):
        manifest = ahk.with_suffix(".json")
        if manifest.exists():
            profiles.append({
                "name": ahk.stem,
                "ahk": ahk,
                "manifest": manifest
            })
    return profiles


def delete_profile(profile, row_frame):
    name = profile["name"]
    try:
        running_ahk.pop(name, None)
        profile["ahk"].unlink(missing_ok=True)
        profile["manifest"].unlink(missing_ok=True)
        row_frame.destroy()
        set_status(f"Deleted profile: {name}", "green")
    except Exception as e:
        set_status(f"Failed to delete {name}: {e}", "red")


# =========================
# Browse
# =========================

def browse_input():
    path = filedialog.askopenfilename(
        title="Select CSV or TXT File",
        filetypes=[("CSV Files", "*.csv"), ("Text Files", "*.txt")]
    )
    if path:
        input_entry.delete(0, tk.END)
        input_entry.insert(0, path)
        set_status("Input file selected.", "blue")


# =========================
# Generator
# =========================

def run_generator():
    input_path = input_entry.get().strip()
    profile_name = profile_entry.get().strip()

    if not Path(input_path).is_file():
        set_status("Invalid input file.", "red")
        return

    if not profile_name:
        set_status("Profile name required.", "red")
        return

    safe = "".join(c for c in profile_name if c.isalnum() or c in " _-").strip()
    ahk_path = AHK_DIR / f"{safe}.ahk"

    try:
        email_to_ahk.run(input_path, ahk_path)
        profile = {
            "name": safe,
            "ahk": ahk_path,
            "manifest": ahk_path.with_suffix(".json")
        }
        add_profile_row(profile)
        set_status(f"Profile '{safe}' created.", "green")
        input_entry.delete(0, tk.END)
        profile_entry.delete(0, tk.END)
    except Exception as e:
        set_status(str(e), "red")


# =========================
# Hotkey Window
# =========================

def open_hotkey_window(profile):
    ensure_ahk_running(profile)

    window = tk.Toplevel(root)
    window.title(f"Hotkeys – {profile['name']}")
    window.geometry("460x520")
    window.resizable(False, False)

    ttk.Label(
        window,
        text="Click ARM → switch to target field → press F8",
        justify="center"
    ).pack(pady=10)

    frame = ScrollableFrame(window)
    frame.pack(fill="both", expand=True, padx=10, pady=10)

    manifest = json.loads(profile["manifest"].read_text(encoding="utf-8"))

    for hotkey, info in manifest.items():
        row = ttk.Frame(frame.inner)
        row.pack(fill="x", pady=4)

        ttk.Label(row, text=info["label"]).pack(side="left", fill="x", expand=True)

        ttk.Button(
            row,
            text="Arm",
            width=8,
            command=lambda h=hotkey: arm_hotkey(h)
        ).pack(side="right")


# =========================
# UI
# =========================

root = tk.Tk()
root.title("Email → Hotkey Manager")
root.geometry("740x580")
root.resizable(False, False)

tk.Label(root, text="Input CSV or TXT file:").pack(pady=(12, 0))

frame_input = tk.Frame(root)
frame_input.pack(pady=5)

input_entry = tk.Entry(frame_input, width=65)
input_entry.pack(side=tk.LEFT, padx=(0, 6))

tk.Button(frame_input, text="Browse…", command=browse_input).pack(side=tk.LEFT)

tk.Label(root, text="Profile name:").pack(pady=(10, 0))
profile_entry = tk.Entry(root, width=40)
profile_entry.pack(pady=5)

tk.Button(
    root,
    text="Generate Hotkeys",
    width=22,
    height=2,
    command=run_generator
).pack(pady=10)

profiles_frame = ttk.LabelFrame(root, text="Saved Profiles")
profiles_frame.pack(fill="both", expand=True, padx=10, pady=(5, 10))

scroll_profiles = ScrollableFrame(profiles_frame)
scroll_profiles.pack(fill="both", expand=True)


def add_profile_row(profile):
    row = ttk.Frame(scroll_profiles.inner)
    row.pack(fill="x", padx=6, pady=3)

    ttk.Button(
        row,
        text=profile["name"],
        command=lambda p=profile: open_hotkey_window(p)
    ).pack(side="left", fill="x", expand=True)

    ttk.Button(
        row,
        text="Delete",
        width=8,
        command=lambda p=profile, r=row: delete_profile(p, r)
    ).pack(side="right", padx=(6, 0))


for profile in load_profiles():
    add_profile_row(profile)

status_label = tk.Label(
    root,
    text="Ready.",
    anchor="w",
    relief=tk.SUNKEN,
    padx=10
)
status_label.pack(side=tk.BOTTOM, fill=tk.X)

root.mainloop()
