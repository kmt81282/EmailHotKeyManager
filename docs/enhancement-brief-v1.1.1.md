# Email Hotkey Manager — v1.1.1 Enhancement Brief

## Project Overview
Email Hotkey Manager is a Windows desktop application that:
- Extracts email addresses from files
- Groups them by domain
- Generates deterministic AutoHotkey (AHK v1) scripts
- Uses a Python UI to manage profiles and trigger hotkeys safely
- Ships as a single EXE with a bundled AutoHotkey runtime

Current release: v1.0.0  
Next release target: v1.1.1

---

## Current Architecture (Summary)

### Core
- `email_to_ahk.py`
  - Extracts emails from input files
  - Groups by domain
  - Assigns Ctrl+Shift+A–Z hotkeys (sequential, collision-free)
  - Generates:
    - `.ahk` script (headless)
    - `.json` manifest (hotkey → label mapping)

### UI
- `email_to_ahk_ui.py`
  - Tkinter-based GUI
  - Profile manager (create, open, delete)
  - Hotkey window per profile
  - **Arm & Trigger model**
    - User clicks ARM
    - Focuses target app
    - Presses **F8**
    - Python sends Ctrl+Shift hotkey
  - Uses bundled AutoHotkey runtime
  - Packaged via PyInstaller into a single EXE

---

## v1.1.1 Enhancement Goal

### Primary Feature
Add support for **`.eml` email files** as an input source.

Users should be able to:
- Select an `.eml` file
- Extract email addresses from:
  - Headers (From, To, Cc, Bcc)
  - Body (plain text and HTML)
- Generate hotkeys exactly the same way as CSV/TXT

No behavior changes for existing file types.

---

## Technical Design (Approved Direction)

### Parsing Strategy
- Use Python standard library only:
  - `email.parser.BytesParser`
  - `email.policy.default`
- No Outlook, COM, or external dependencies
- Safe for PyInstaller bundling

### Extraction Rules
- Extract emails from:
  - Headers: From / To / Cc / Bcc
  - Body parts: `text/plain`, `text/html`
- Use the same email regex as CSV/TXT
- Normalize to lowercase
- De-duplicate results

---

## Planned Code Changes

### `email_to_ahk.py`
- Extend `extract_emails_from_file()` to support `.eml`
- Add helper:
  - `extract_emails_from_eml(file_path)`
- No changes to:
  - Hotkey assignment
  - AHK generation
  - Manifest format

### `email_to_ahk_ui.py`
- Update file picker to allow:
  - `.csv`
  - `.txt`
  - `.eml`
- UI enhancement opportunities (open for redesign in v1.1.1):
  - Clearer input file type indication
  - Optional file-type icon or label
  - Improved status messaging for extraction results

---

## Versioning Plan

- Version: **v1.1.1**
- Update:
  - `version.txt`
  - Git tag: `v1.1.1`
- Release notes:
  - “Added support for .eml email files”
  - “Extracts addresses from headers and body”
  - “No additional setup required”

---

## Constraints / Non-Goals

- No breaking changes
- No change to Arm & Trigger model
- No change to AutoHotkey runtime strategy
- No new third-party Python dependencies

---

## Next Steps for New Chat

Focus areas:
1. Implement `.eml` extraction in `email_to_ahk.py`
2. Update UI file picker + UX polish
3. Optional UI refinements specific to email files
4. Update README for `.eml` support
5. Prepare v1.1.1 release

