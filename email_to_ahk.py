import re
import csv
import json
from collections import defaultdict
from pathlib import Path
import string


# ===============================
# Constants
# ===============================

EMAIL_PATTERN = r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}'
ENCODINGS = ['utf-8', 'utf-16', 'iso-8859-1']

# Reserved shortcuts
RESERVED_KEYS = {"Z", "R", "X"}
AVAILABLE_KEYS = [k for k in string.ascii_uppercase if k not in RESERVED_KEYS]


# ===============================
# Email extraction
# ===============================

def extract_emails_from_file(file_path):
    emails = set()
    file_path = Path(file_path)

    if not file_path.exists():
        raise FileNotFoundError("Input file not found.")

    if file_path.suffix.lower() == ".csv":
        for encoding in ENCODINGS:
            try:
                with open(file_path, newline="", encoding=encoding) as csvfile:
                    reader = csv.reader(csvfile)
                    for row in reader:
                        for cell in row:
                            emails.update(
                                e.lower() for e in re.findall(EMAIL_PATTERN, cell)
                            )
                break
            except (UnicodeDecodeError, csv.Error):
                continue

    elif file_path.suffix.lower() == ".txt":
        for encoding in ENCODINGS:
            try:
                content = file_path.read_text(encoding=encoding)
                emails.update(
                    e.lower() for e in re.findall(EMAIL_PATTERN, content)
                )
                break
            except UnicodeDecodeError:
                continue
    else:
        raise ValueError("Unsupported file type. Use CSV or TXT.")

    emails.discard("marc.turner@ukg.com")
    return sorted(emails)


# ===============================
# Grouping
# ===============================

def group_emails_by_domain(emails):
    domain_groups = defaultdict(list)
    for email in emails:
        domain = email.split("@", 1)[1]
        domain_groups[domain].append(email)
    return domain_groups


# ===============================
# AHK + Manifest generation
# ===============================

def create_ahk_and_manifest(domain_groups, ahk_path):
    ahk_path = Path(ahk_path)
    manifest_path = ahk_path.with_suffix(".json")

    manifest = {}
    key_iter = iter(AVAILABLE_KEYS)

    with open(ahk_path, "w", encoding="utf-8") as ahk:

        # ---- Per-domain hotkeys (A → Z sequential) ----
        for domain, emails in domain_groups.items():
            try:
                key = next(key_iter)
            except StopIteration:
                raise RuntimeError("Too many domains (max 23 supported).")

            email_string = ", ".join(emails) + ","

            ahk.write(f"^+{key}::\n")
            ahk.write(f"Clipboard := \"{email_string}\"\n")
            ahk.write("ClipWait\n")
            ahk.write("SendInput, ^v\n")
            ahk.write("Return\n\n")

            manifest[f"Ctrl+Shift+{key}"] = {
                "label": domain,
                "type": "domain"
            }

        # ---- All Emails ----
        all_emails = ", ".join(
            email for emails in domain_groups.values() for email in emails
        ) + ","

        ahk.write("^+Z::\n")
        ahk.write(f"Clipboard := \"{all_emails}\"\n")
        ahk.write("ClipWait\n")
        ahk.write("SendInput, ^v\n")
        ahk.write("Return\n\n")

        manifest["Ctrl+Shift+Z"] = {
            "label": "All Emails",
            "type": "all"
        }

        # ---- Sequential ----
        ahk.write("^+R::\n")
        for email in all_emails.split(","):
            if email:
                ahk.write(f"Clipboard := \"{email}\"\n")
                ahk.write("ClipWait\n")
                ahk.write("SendInput, ^v\n")
                ahk.write("Sleep, 2000\n")
                ahk.write("Send, {Enter}\n")
        ahk.write("Return\n\n")

        manifest["Ctrl+Shift+R"] = {
            "label": "Sequential Emails",
            "type": "sequential"
        }

        # ---- Exit ----
        ahk.write("^+X::ExitApp\n")

    manifest_path.write_text(
        json.dumps(manifest, indent=2),
        encoding="utf-8"
    )


# ===============================
# UI entry point
# ===============================

def run(input_file_path, ahk_output_path):
    emails = extract_emails_from_file(input_file_path)
    if not emails:
        raise ValueError("No emails found.")

    domain_groups = group_emails_by_domain(emails)
    create_ahk_and_manifest(domain_groups, ahk_output_path)


# ===============================
# CLI support
# ===============================

def main():
    input_path = input("Enter CSV or TXT file path: ").strip('"')
    output_path = input("Enter output .ahk file path: ").strip('"')
    run(input_path, output_path)
    print("AHK + manifest created successfully.")


if __name__ == "__main__":
    main()
