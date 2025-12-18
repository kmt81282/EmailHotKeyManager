# Email Hotkey Manager

**Email Hotkey Manager** is a Windows desktop application that generates and manages AutoHotkey shortcuts for quickly pasting email addresses.

It is designed for speed, reliability, and zero setup:
- No Python required (for end users)
- No AutoHotkey installation required
- Single portable EXE

---

## ✨ Features

- Extract email addresses from **CSV, TXT, or EML** files
- Automatically group emails by domain
- Generate deterministic `Ctrl + Shift + A–Z` hotkeys
- Manage multiple **saved profiles**
- Launch and control hotkeys from a clean GUI
- **F8 Arm & Trigger workflow** (focus-safe, no misfires)
- Bundled AutoHotkey runtime (no external dependencies)
- One-click delete of profiles

---

## 🖥️ Requirements (End Users)

- **Windows 10 or Windows 11 (64-bit)**
- No Python required
- No AutoHotkey installation required

---

🚀 Getting Started (End Users)

1. Download EmailHotkeyManager.zip from the Releases page
2. Right-click the ZIP and choose Extract All
3. Open the extracted EmailHotkeyManager folder
4. Double-click EmailHotkeyManager.exe to launch

ℹ️ Important:
Run the EXE from inside the extracted folder. Do not move the EXE out of the folder.

📌 Optional: Run from Desktop or Start Menu

If you’d like to launch the app from another location (such as the Desktop or Start menu):
1. Right-click EmailHotkeyManager.exe
2. Choose Create shortcut
3. Move the shortcut wherever you like
Use the shortcut to launch the app.
Do not move the original EXE out of the folder.

## 📄 Supported Input Files

- **.csv** — Email addresses in any column
- **.txt** — Plain text email lists
- **.eml** — Saved email messages
  - Extracts from headers: From, To, Cc, Bcc
  - Extracts from body: text/plain and text/html

---

## 🧭 How to Use

### 1️⃣ Generate a Profile
1. Click **Browse** and select a CSV, TXT, or EML file
2. Enter a **Profile Name**
3. Click **Generate Hotkeys**

The profile is saved automatically.

---

### 2️⃣ Open a Profile
- Click a saved profile from the list
- A new window opens showing available hotkeys

---

### 3️⃣ Use Hotkeys (Arm & Trigger)

This app uses a **safe Arm & Trigger workflow** to avoid focus issues.

1. Click **ARM** next to the desired email group
2. Click into the target field (Word, Outlook, browser, etc.)
3. Press **F8**
4. Emails paste instantly

> ℹ️ The **F8 key is fully suppressed** during triggering, so no extra characters are inserted.

---

## ⌨️ Hotkey Behavior

- Hotkeys are assigned sequentially:
  - `Ctrl + Shift + A`
  - `Ctrl + Shift + B`
  - …
- Reserved hotkeys:
  - `Ctrl + Shift + Z` → All emails
  - `Ctrl + Shift + R` → Sequential email entry
  - `Ctrl + Shift + X` → Exit AutoHotkey script

---

## 🗂️ Profiles

- Profiles are stored locally on your machine
- Each profile contains:
  - A generated `.ahk` script
  - A hotkey manifest
- Profiles can be deleted at any time from the UI

---

## 🔐 Security & Trust Notes

- Uses AutoHotkey for automation
- Uses a global hotkey listener (`F8`)
- Antivirus software may show a warning on first run  
  (common for automation tools)
- Code signing can be added if required

---

## 🛠️ Built With

- Python
- Tkinter
- AutoHotkey v1 (bundled runtime)
- PyInstaller

---

## 👩‍💻 Development (Optional)

> This section is only for developers building from source.  
> End users do **not** need Python.

```bat
pip install -r requirements.txt
set PYTHONPATH=src
python -m email_hotkey_manager.email_to_ahk_ui
Build EXE
pyinstaller --clean specs/email_to_ahk_ui.spec

📦 Distribution

This application is distributed as a single portable EXE via GitHub Releases.

No installer. No registry changes.

👤 Author

Marc Turner

📄 License

This project is provided for internal or personal use.
Add a license file if you plan to distribute publicly.
