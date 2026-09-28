#!/usr/bin/env python3
"""
Fingertips Automation - MyDocsOnline Importer

What this script does
----------------------
1. Fetches ONLY genuine top-level MyDocsOnline upload notifications.
2. Keeps the existing local raw .eml cache.
3. Ignores forwarded/replied emails that merely quote an old MDO notification.
4. Parses cached genuine MDO notifications into a preview CSV.
5. Creates ONE PREVIEW ROW PER UPLOADED FILE.
6. Keeps rows from the same notification tied together with Upload Group ID.
7. Extracts:
   - email/upload date
   - job date from filename where possible
   - MDO folder/path
   - uploader name/email
   - filename
   - description
   - minutes/speakers where present
   - rate/turnaround text where present
   - explicit numeric client rate where present
   - Source ID
   - validation/status fields
8. Writes ignored cached follow-up/forward messages to a separate CSV.

This parser produces structured cached MDO job data for the monthly spreadsheet generator.

Folder structure
----------------
email_cache/
    mydocsonline/
        index.csv
        raw/
            *.eml
        parsed_jobs_preview.csv
        ignored_followups.csv

Credentials
-----------
PowerShell:
    $env:FINGERTIPS_EMAIL="cbennett@fingertipstyping.co.uk"
    $env:FINGERTIPS_EMAIL_PASSWORD="your-password-here"

Run normally (fetch new direct MDO emails, then parse cache):
    python mydocsonline_import.py

Parse the existing cache only, without connecting to email:
    python mydocsonline_import.py --parse-only
"""

from __future__ import annotations

import argparse
import csv
import email
import hashlib
import imaplib
import os
import re
from datetime import datetime
from email.header import decode_header, make_header
from email.message import Message
from email.utils import parsedate_to_datetime, parseaddr
from pathlib import Path

from openpyxl import Workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.table import Table, TableStyleInfo

IMAP_HOST = "mailcl3.extendcp.co.uk"
IMAP_PORT = 993
MAILBOX = "INBOX"
MYDOCS_ADDRESS = "donotreply@mydocsonline.com"

CACHE_ROOT = Path("email_cache") / "mydocsonline"
RAW_DIR = CACHE_ROOT / "raw"
INDEX_PATH = CACHE_ROOT / "index.csv"
PREVIEW_PATH = CACHE_ROOT / "parsed_jobs_preview.csv"
IGNORED_PATH = CACHE_ROOT / "ignored_followups.csv"

# First-folder aliases can be extended later or replaced by client_config.xlsx.
CLIENT_ALIASES = {
    "eoe hr": "East of England",
    "snodland medical centre": "Snodland Medical Centre",
}

AUDIO_EXTENSIONS = {
    ".mp3", ".m4a", ".wav", ".wma", ".aac", ".flac", ".dss", ".ds2",
}
DOCUMENT_EXTENSIONS = {
    ".doc", ".docx", ".pdf", ".rtf", ".odt", ".txt", ".ppt", ".pptx",
}
KNOWN_FILE_EXTENSIONS = AUDIO_EXTENSIONS | DOCUMENT_EXTENSIONS

INDEX_COLUMNS = [
    "Source ID", "Message ID", "IMAP UID", "Date", "Parsed Date",
    "From", "From Email", "Subject", "Cached File", "Detected By",
]

PREVIEW_COLUMNS = [
    "Date",
    "Date Source",
    "Email Date",
    "Subject",
    "Client",
    "MDO Root Folder",
    "MDO Path",
    "Uploader",
    "Uploader Email",
    "Description",
    "Uploaded Filename",
    "File Number",
    "Files In Upload",
    "Client Rate",
    "Rate / Turnaround",
    "Mins",
    "Speakers",
    "Other info",
    "Proofing",
    "Typist",
    "Typist Rate",
    "Source ID",
    "Parent Source ID",
    "Upload Group ID",
    "Message ID",
    "Validation Issues",
    "Status",
]

IGNORED_COLUMNS = [
    "Email Date", "From", "From Email", "Subject", "Message ID",
    "Cached File", "Reason",
]


def decode_mime(value: str | None) -> str:
    if not value:
        return ""
    try:
        return str(make_header(decode_header(value)))
    except Exception:
        return value


def sender_address(sender_header: str) -> str:
    return parseaddr(sender_header)[1].strip().lower()


def parse_email_date(msg: Message) -> datetime | None:
    raw = msg.get("Date")
    if not raw:
        return None
    try:
        return parsedate_to_datetime(raw)
    except Exception:
        return None


def extract_text(msg: Message) -> str:
    """Extract readable body text, preferring text/plain."""
    if msg.is_multipart():
        plain_parts: list[str] = []
        html_parts: list[str] = []

        for part in msg.walk():
            disposition = str(part.get("Content-Disposition", "")).lower()
            if "attachment" in disposition:
                continue

            content_type = part.get_content_type()
            if content_type not in {"text/plain", "text/html"}:
                continue

            try:
                payload = part.get_payload(decode=True)
                if payload is None:
                    continue
                charset = part.get_content_charset() or "utf-8"
                text = payload.decode(charset, errors="replace")
            except Exception:
                continue

            if content_type == "text/plain":
                plain_parts.append(text)
            else:
                html_parts.append(text)

        text = "\n".join(plain_parts if plain_parts else html_parts)
    else:
        try:
            payload = msg.get_payload(decode=True)
            if payload is None:
                return ""
            charset = msg.get_content_charset() or "utf-8"
            text = payload.decode(charset, errors="replace")
        except Exception:
            return ""

    if "<html" in text.lower() or "<body" in text.lower():
        text = re.sub(r"<style.*?</style>", " ", text, flags=re.I | re.S)
        text = re.sub(r"<script.*?</script>", " ", text, flags=re.I | re.S)
        text = re.sub(r"<br\s*/?>", "\n", text, flags=re.I)
        text = re.sub(r"</p\s*>", "\n", text, flags=re.I)
        text = re.sub(r"<[^>]+>", " ", text)

    text = text.replace("\r\n", "\n").replace("\r", "\n")
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n[ \t]+", "\n", text)
    return text.strip()


def is_direct_mdo_upload(msg: Message) -> bool:
    sender = sender_address(decode_mime(msg.get("From")))
    subject = decode_mime(msg.get("Subject")).strip().lower()

    return (
        sender == MYDOCS_ADDRESS
        and subject.startswith("upload to \\")
    )


def contains_quoted_mdo_notification(msg: Message) -> bool:
    if is_direct_mdo_upload(msg):
        return False

    body = extract_text(msg).lower()
    return (
        "my docs online" in body
        or "donotreply@mydocsonline.com" in body
    ) and "upload complete to \\" in body


def make_parent_source_id(msg: Message, fallback: str) -> tuple[str, str]:
    message_id = decode_mime(msg.get("Message-ID")).strip()
    if message_id:
        return f"EMAIL:{message_id}", message_id
    return f"EMAIL:{fallback}", ""


def stable_group_id(parent_source_id: str) -> str:
    digest = hashlib.sha256(parent_source_id.encode("utf-8")).hexdigest()[:16]
    return f"MDO:{digest}"


def cache_filename(source_id: str, parsed_date: datetime | None) -> str:
    digest = hashlib.sha256(source_id.encode("utf-8")).hexdigest()[:20]
    date_prefix = parsed_date.strftime("%Y-%m-%d") if parsed_date else "unknown-date"
    return f"{date_prefix}_{digest}.eml"


def load_existing_source_ids() -> set[str]:
    if not INDEX_PATH.exists():
        return set()

    existing: set[str] = set()
    with INDEX_PATH.open("r", encoding="utf-8-sig", newline="") as f:
        for row in csv.DictReader(f):
            source_id = (row.get("Source ID") or "").strip()
            if source_id:
                existing.add(source_id)
    return existing


def append_index_rows(rows: list[dict]) -> None:
    if not rows:
        return

    exists = INDEX_PATH.exists()
    with INDEX_PATH.open("a", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=INDEX_COLUMNS)
        if not exists:
            writer.writeheader()
        writer.writerows(rows)


def fetch_new_direct_mdo() -> None:
    username = os.getenv("FINGERTIPS_EMAIL")
    password = os.getenv("FINGERTIPS_EMAIL_PASSWORD")

    if not username or not password:
        raise SystemExit(
            "Missing credentials.\n\n"
            "Either set FINGERTIPS_EMAIL / FINGERTIPS_EMAIL_PASSWORD,\n"
            "or run with --parse-only to parse the existing cache."
        )

    RAW_DIR.mkdir(parents=True, exist_ok=True)
    existing = load_existing_source_ids()
    index_rows: list[dict] = []

    scanned = 0
    direct_mdo = 0
    new_cached = 0

    print(f"Connecting securely to {IMAP_HOST}:{IMAP_PORT} ...")

    with imaplib.IMAP4_SSL(IMAP_HOST, IMAP_PORT) as mail:
        mail.login(username, password)
        mail.select(MAILBOX, readonly=True)

        status, data = mail.uid("search", None, "ALL")
        if status != "OK":
            raise RuntimeError("Could not list INBOX messages.")

        uids = data[0].split()
        print(f"INBOX contains {len(uids)} message(s).")
        print("Fetching genuine top-level MyDocsOnline uploads only...\n")

        for idx, uid_bytes in enumerate(uids, start=1):
            scanned += 1
            uid = uid_bytes.decode(errors="replace")

            status, msg_data = mail.uid("fetch", uid_bytes, "(RFC822)")
            if status != "OK":
                continue

            raw = None
            for item in msg_data:
                if isinstance(item, tuple) and len(item) >= 2:
                    raw = item[1]
                    break
            if raw is None:
                continue

            msg = email.message_from_bytes(raw)

            if not is_direct_mdo_upload(msg):
                if idx % 100 == 0:
                    print(
                        f"  Checked {idx}/{len(uids)} | "
                        f"direct MDO: {direct_mdo} | new: {new_cached}",
                        end="\r",
                    )
                continue

            direct_mdo += 1
            source_id, message_id = make_parent_source_id(msg, f"UID:{uid}")

            if source_id in existing:
                continue

            parsed_date = parse_email_date(msg)
            filename = cache_filename(source_id, parsed_date)
            output_path = RAW_DIR / filename
            output_path.write_bytes(raw)

            sender = decode_mime(msg.get("From"))
            index_rows.append({
                "Source ID": source_id,
                "Message ID": message_id,
                "IMAP UID": uid,
                "Date": decode_mime(msg.get("Date")),
                "Parsed Date": parsed_date.isoformat() if parsed_date else "",
                "From": sender,
                "From Email": sender_address(sender),
                "Subject": decode_mime(msg.get("Subject")),
                "Cached File": str(output_path),
                "Detected By": "direct-mdo-upload",
            })

            existing.add(source_id)
            new_cached += 1

            if idx % 100 == 0:
                print(
                    f"  Checked {idx}/{len(uids)} | "
                    f"direct MDO: {direct_mdo} | new: {new_cached}",
                    end="\r",
                )

        mail.logout()

    append_index_rows(index_rows)
    print(" " * 90, end="\r")
    print(
        f"Fetch complete: scanned {scanned}, "
        f"found {direct_mdo} direct MDO upload(s), "
        f"cached {new_cached} new email(s).\n"
    )


def split_nonempty_lines(text: str) -> list[str]:
    return [line.strip() for line in text.splitlines() if line.strip()]


def extract_mdo_path(subject: str, body: str) -> str:
    # Subject is the most reliable source.
    match = re.search(r"Upload to\s+(\\[^\r\n]+)", subject, flags=re.I)
    if match:
        return match.group(1).strip()

    match = re.search(r"Upload complete to\s+(\\[^\r\n]+)", body, flags=re.I)
    return match.group(1).strip() if match else ""


def path_parts(path: str) -> list[str]:
    return [part.strip() for part in path.split("\\") if part.strip()]


def client_from_path(path: str, uploader: str) -> tuple[str, str]:
    parts = path_parts(path)
    root = parts[0] if parts else ""

    if root.lower() == "customer upload":
        # No paid client folder: uploader/client name is usually second folder
        # or the form's First Name, Last Name field.
        client = parts[1] if len(parts) > 1 else uploader
        return client, root

    client = CLIENT_ALIASES.get(root.lower(), root)
    return client, root


def extract_label_value(lines: list[str], label: str) -> str:
    label_lower = label.lower()

    for i, line in enumerate(lines):
        if line.lower() == label_lower:
            if i + 1 < len(lines):
                return lines[i + 1].strip()
            return ""

        # Occasionally the value may be on the same line after ":".
        if line.lower().startswith(label_lower + ":"):
            return line.split(":", 1)[1].strip()

    return ""


def extract_uploader(body: str) -> tuple[str, str]:
    lines = split_nonempty_lines(body)

    # Customer Upload form.
    full_name = extract_label_value(lines, "First Name, Last Name")
    email_address = extract_label_value(lines, "Email Address")
    if full_name or email_address:
        return full_name, email_address

    # Paid-folder format.
    uploader = ""
    uploader_email = ""

    for line in lines:
        m = re.match(r"^From:\s*(.+)$", line, flags=re.I)
        if m and "my docs online" not in m.group(1).lower():
            uploader = m.group(1).strip()
            break

    for line in lines:
        m = re.match(r"^From Email:\s*(.+)$", line, flags=re.I)
        if m:
            uploader_email = m.group(1).strip()
            break

    return uploader, uploader_email


def extract_rate_turnaround(body: str) -> str:
    lines = split_nonempty_lines(body)
    return extract_label_value(lines, "Rate and Turnaround Time")


def explicit_numeric_rate(rate_text: str) -> str:
    if not rate_text:
        return ""
    m = re.search(r"£\s*(\d+(?:\.\d+)?)", rate_text)
    return m.group(1) if m else ""


def extract_minutes_speakers(body: str) -> tuple[str, str]:
    lines = split_nonempty_lines(body)
    detail = extract_label_value(lines, "Length of audio and amount of speakers")

    # Also inspect full body because not every MDO template uses the labels.
    target = detail or body

    minutes = ""
    speakers = ""

    m = re.search(r"\b(\d+(?:\.\d+)?)\s*minutes?\b", target, flags=re.I)
    if m:
        minutes = m.group(1)

    s = re.search(r"\b(\d+)\s*speakers?\b", target, flags=re.I)
    if s:
        speakers = s.group(1)

    return minutes, speakers


def extract_other_info(body: str) -> str:
    lines = split_nonempty_lines(body)
    return extract_label_value(
        lines,
        "Provide: Type of transcription required and any other information"
    )


def looks_like_uploaded_filename(line: str) -> bool:
    cleaned = line.strip().rstrip(",").strip()
    if not cleaned:
        return False

    # Remove optional "(12,345 bytes)" suffix.
    cleaned = re.sub(r"\s+\([\d,]+\s+bytes\)\s*$", "", cleaned, flags=re.I)

    suffix = Path(cleaned).suffix.lower()
    return suffix in KNOWN_FILE_EXTENSIONS


def clean_uploaded_filename(line: str) -> str:
    line = line.strip().rstrip(",").strip()
    line = re.sub(r"\s+\([\d,]+\s+bytes\)\s*$", "", line, flags=re.I)
    return line.strip()


def extract_filenames(body: str) -> list[str]:
    lines = split_nonempty_lines(body)
    filenames: list[str] = []

    for line in lines:
        if looks_like_uploaded_filename(line):
            filename = clean_uploaded_filename(line)
            if filename not in filenames:
                filenames.append(filename)

    return filenames


def parse_2digit_year(value: str) -> int:
    year = int(value)
    return 2000 + year if year < 70 else 1900 + year


def extract_date_from_filename(filename: str) -> datetime | None:
    name = Path(filename).stem

    # YYYY-MM-DD / YYYY_MM_DD / YYYY MM DD
    m = re.search(
        r"(?<!\d)(20\d{2})[-_. ](0?[1-9]|1[0-2])[-_. ]([0-2]?\d|3[01])(?!\d)",
        name,
    )
    if m:
        try:
            return datetime(int(m.group(1)), int(m.group(2)), int(m.group(3)))
        except ValueError:
            pass

    # DD.MM.YY / DD-MM-YY / DD_MM_YY
    m = re.search(
        r"(?<!\d)([0-2]?\d|3[01])[./_-](0?[1-9]|1[0-2])[./_-](\d{2,4})(?!\d)",
        name,
    )
    if m:
        try:
            year = int(m.group(3))
            if year < 100:
                year = parse_2digit_year(m.group(3))
            return datetime(year, int(m.group(2)), int(m.group(1)))
        except ValueError:
            pass

    return None


def description_from_filename(filename: str) -> str:
    stem = Path(filename).stem

    # Remove a leading ISO-like date.
    stem = re.sub(
        r"^\s*20\d{2}[-_. ](?:0?[1-9]|1[0-2])[-_. ](?:[0-2]?\d|3[01])[-_. ]*",
        "",
        stem,
    )

    # Remove a leading DD.MM.YY-like date.
    stem = re.sub(
        r"^\s*(?:[0-2]?\d|3[01])[./_-](?:0?[1-9]|1[0-2])[./_-]\d{2,4}[-_. ]*",
        "",
        stem,
    )

    stem = stem.replace("_", " ")
    stem = re.sub(r"\s+", " ", stem).strip(" -_.")
    return stem or Path(filename).stem


def choose_job_date(filename: str, email_date: datetime | None) -> tuple[str, str]:
    file_date = extract_date_from_filename(filename)
    if file_date:
        return file_date.strftime("%d/%m/%Y"), "Filename"

    if email_date:
        return email_date.strftime("%d/%m/%Y"), "Email"

    return "", ""


def validation_issues(row: dict) -> list[str]:
    issues: list[str] = []

    if not row["Date"]:
        issues.append("Missing Date")
    if not row["Client"]:
        issues.append("Missing Client")
    if not row["Description"]:
        issues.append("Missing Description")
    if not row["Mins"]:
        issues.append("Missing Mins")
    if not row["Client Rate"]:
        issues.append("Missing Client Rate")

    return issues


def parse_direct_mdo_message(msg: Message, cache_name: str) -> list[dict]:
    subject = decode_mime(msg.get("Subject"))
    body = extract_text(msg)
    email_date = parse_email_date(msg)

    parent_source_id, message_id = make_parent_source_id(msg, cache_name)
    group_id = stable_group_id(parent_source_id)

    mdo_path = extract_mdo_path(subject, body)
    uploader, uploader_email = extract_uploader(body)
    client, root_folder = client_from_path(mdo_path, uploader)

    rate_turnaround = extract_rate_turnaround(body)
    client_rate = explicit_numeric_rate(rate_turnaround)
    minutes, speakers = extract_minutes_speakers(body)
    other_info = extract_other_info(body)

    filenames = extract_filenames(body)

    # Keep a visible row even if a malformed notification has no parseable file.
    if not filenames:
        filenames = [""]

    rows: list[dict] = []

    for file_index, filename in enumerate(filenames, start=1):
        job_date, date_source = choose_job_date(filename, email_date)
        description = description_from_filename(filename) if filename else ""

        file_source_id = (
            f"{parent_source_id}#FILE:{file_index}"
            if len(filenames) > 1
            else parent_source_id
        )

        row = {
            "Date": job_date,
            "Date Source": date_source,
            "Email Date": email_date.strftime("%d/%m/%Y") if email_date else "",
            "Subject": subject.replace("\r", " ").replace("\n", " ").strip(),
            "Client": client,
            "MDO Root Folder": root_folder,
            "MDO Path": mdo_path,
            "Uploader": uploader,
            "Uploader Email": uploader_email,
            "Description": description,
            "Uploaded Filename": filename,
            "File Number": file_index,
            "Files In Upload": len(filenames),
            "Client Rate": client_rate,
            "Rate / Turnaround": rate_turnaround,
            "Mins": minutes,
            "Speakers": speakers,
            "Other info": other_info,
            "Proofing": "",
            "Typist": "",
            "Typist Rate": "",
            "Source ID": file_source_id,
            "Parent Source ID": parent_source_id,
            "Upload Group ID": group_id,
            "Message ID": message_id,
            "Validation Issues": "",
            "Status": "",
        }

        issues = validation_issues(row)
        row["Validation Issues"] = "; ".join(issues)
        row["Status"] = "Needs Review" if issues else "Ready"
        rows.append(row)

    return rows


def write_csv(path: Path, columns: list[str], rows: list[dict]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=columns, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(rows)


def parse_cache() -> None:
    RAW_DIR.mkdir(parents=True, exist_ok=True)

    preview_rows: list[dict] = []
    ignored_rows: list[dict] = []
    direct_count = 0
    followup_count = 0
    unrelated_count = 0

    eml_files = sorted(RAW_DIR.glob("*.eml"))

    print(f"Parsing {len(eml_files)} cached .eml file(s)...")

    for eml_path in eml_files:
        try:
            msg = email.message_from_bytes(eml_path.read_bytes())
        except Exception as exc:
            ignored_rows.append({
                "Email Date": "",
                "From": "",
                "From Email": "",
                "Subject": "",
                "Message ID": "",
                "Cached File": str(eml_path),
                "Reason": f"Could not parse .eml: {exc}",
            })
            continue

        if is_direct_mdo_upload(msg):
            direct_count += 1
            preview_rows.extend(parse_direct_mdo_message(msg, eml_path.name))
            continue

        if contains_quoted_mdo_notification(msg):
            followup_count += 1
            parsed_date = parse_email_date(msg)
            sender = decode_mime(msg.get("From"))
            ignored_rows.append({
                "Email Date": parsed_date.strftime("%d/%m/%Y") if parsed_date else "",
                "From": sender,
                "From Email": sender_address(sender),
                "Subject": decode_mime(msg.get("Subject")),
                "Message ID": decode_mime(msg.get("Message-ID")),
                "Cached File": str(eml_path),
                "Reason": "Follow-up/forward quoting an older MDO notification; no new job created",
            })
        else:
            unrelated_count += 1
            parsed_date = parse_email_date(msg)
            sender = decode_mime(msg.get("From"))
            ignored_rows.append({
                "Email Date": parsed_date.strftime("%d/%m/%Y") if parsed_date else "",
                "From": sender,
                "From Email": sender_address(sender),
                "Subject": decode_mime(msg.get("Subject")),
                "Message ID": decode_mime(msg.get("Message-ID")),
                "Cached File": str(eml_path),
                "Reason": "Cached by older logic but is not a direct MDO upload",
            })

    # Stable ordering makes comparison between parser versions easy.
    preview_rows.sort(
        key=lambda row: (
            row["Email Date"],
            row["Parent Source ID"],
            int(row["File Number"]),
        )
    )

    write_csv(PREVIEW_PATH, PREVIEW_COLUMNS, preview_rows)
    write_csv(IGNORED_PATH, IGNORED_COLUMNS, ignored_rows)

    ready = sum(1 for row in preview_rows if row["Status"] == "Ready")
    needs_review = len(preview_rows) - ready

    print()
    print("=" * 80)
    print("MYDOCSONLINE PARSE SUMMARY")
    print("=" * 80)
    print(f"Cached .eml files:          {len(eml_files)}")
    print(f"Direct MDO upload emails:   {direct_count}")
    print(f"Follow-ups ignored:         {followup_count}")
    print(f"Other cached items ignored: {unrelated_count}")
    print(f"Preview job rows produced:  {len(preview_rows)}")
    print(f"Ready rows:                 {ready}")
    print(f"Needs Review rows:          {needs_review}")
    print()
    print(f"Preview CSV:  {PREVIEW_PATH.resolve()}")
    print(f"Ignored CSV:  {IGNORED_PATH.resolve()}")



# ---------------------------------------------------------------------------
# Excel working-sheet generation
# ---------------------------------------------------------------------------

EXCEL_HEADERS = [
    "Date",
    "Client",
    "Typist",
    "Description",
    "Client Rate",
    "Typist Rate",
    "Mins",
    "Other info",
    "Proofing",
    "Typist total",
    "Sub total",
    "VAT",
    "Total",
    "Source ID",
    "Validation Issues",
    "Status",
]


def parse_ddmmyyyy(value: str):
    value = (value or "").strip()
    if not value:
        return None
    try:
        return datetime.strptime(value, "%d/%m/%Y")
    except ValueError:
        return None


def safe_float(value: str):
    value = (value or "").strip()
    if not value:
        return None
    try:
        return float(value)
    except ValueError:
        return None


def excel_other_info(row: dict) -> str:
    parts = ["MDO upload"]

    uploader = (row.get("Uploader") or "").strip()
    if uploader:
        parts.append(f"Uploader: {uploader}")

    turnaround = (row.get("Rate / Turnaround") or "").strip()
    if turnaround:
        parts.append(f"Rate/turnaround: {turnaround}")

    speakers = (row.get("Speakers") or "").strip()
    if speakers:
        parts.append(f"Speakers: {speakers}")

    filename = (row.get("Uploaded Filename") or "").strip()
    if filename:
        parts.append(f"File: {filename}")

    parsed_job_date = (row.get("Date") or "").strip()
    email_date = (row.get("Email Date") or "").strip()

    # The workbook's Date column uses the upload/email date because the
    # user chooses the invoice month by emails received in that month.
    # If the filename contains an older/different date, preserve it here.
    if parsed_job_date and email_date and parsed_job_date != email_date:
        parts.append(f"Filename date: {parsed_job_date}")

    extra = (row.get("Other info") or "").strip()
    if extra:
        parts.append(extra)

    return " | ".join(parts)


def load_parsed_preview_rows() -> list[dict]:
    if not PREVIEW_PATH.exists():
        raise FileNotFoundError(
            f"Parsed preview not found: {PREVIEW_PATH}\n"
            "Run the parser first."
        )

    with PREVIEW_PATH.open("r", encoding="utf-8-sig", newline="") as f:
        return list(csv.DictReader(f))


def filter_new_work_rows(rows: list[dict]) -> tuple[list[dict], int]:
    """
    Remove MDO notifications that are clearly uploads into a completed-work
    folder. These are delivery/return events, not new client jobs.
    """
    kept: list[dict] = []
    excluded = 0

    for row in rows:
        path = (row.get("MDO Path") or "").lower()

        if "transcriptions done" in path:
            excluded += 1
            continue

        kept.append(row)

    return kept, excluded


def filter_rows_for_month(rows: list[dict], month: str | None) -> list[dict]:
    """
    Filter by MDO/email upload month, not by a date embedded in the filename.

    month format: YYYY-MM
    """
    if not month:
        return rows

    try:
        datetime.strptime(month, "%Y-%m")
    except ValueError as exc:
        raise ValueError("--month must use YYYY-MM, e.g. 2026-08") from exc

    filtered: list[dict] = []

    for row in rows:
        email_date = parse_ddmmyyyy(row.get("Email Date", ""))
        if email_date and email_date.strftime("%Y-%m") == month:
            filtered.append(row)

    return filtered


def build_excel_workbook(rows: list[dict], output_path: Path, month: str | None) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Jobs"

    title_month = (
        datetime.strptime(month, "%Y-%m").strftime("%B %Y")
        if month
        else "All Cached MDO Jobs"
    )

    # Title.
    ws.merge_cells("A1:P1")
    ws["A1"] = f"MDO Jobs - {title_month}"
    ws["A1"].font = Font(bold=True, color="FFFFFF", size=16)
    ws["A1"].fill = PatternFill("solid", fgColor="1F4E78")
    ws["A1"].alignment = Alignment(vertical="center")
    ws.row_dimensions[1].height = 26

    # User-facing note.
    ws.merge_cells("A2:P2")
    ws["A2"] = (
        f"{len(rows)} MyDocsOnline job row(s). "
        "Fill missing values; Validation Issues and Status update automatically."
    )
    ws["A2"].font = Font(italic=True, color="1F1F1F")
    ws["A2"].fill = PatternFill("solid", fgColor="D9EAF7")
    ws["A2"].alignment = Alignment(wrap_text=True, vertical="center")
    ws.row_dimensions[2].height = 34

    header_row = 4

    for col_num, header in enumerate(EXCEL_HEADERS, start=1):
        cell = ws.cell(header_row, col_num, header)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="4472C4")
        cell.alignment = Alignment(
            horizontal="center",
            vertical="center",
            wrap_text=True,
        )

    ws.row_dimensions[header_row].height = 30

    first_data_row = 5

    for row_num, row in enumerate(rows, start=first_data_row):
        # Date = email/upload date for month-based importing.
        ws.cell(row_num, 1, parse_ddmmyyyy(row.get("Email Date", "")))
        ws.cell(row_num, 2, (row.get("Client") or "").strip())
        ws.cell(row_num, 3, (row.get("Typist") or "").strip())
        ws.cell(row_num, 4, (row.get("Description") or "").strip())
        ws.cell(row_num, 5, safe_float(row.get("Client Rate", "")))
        ws.cell(row_num, 6, safe_float(row.get("Typist Rate", "")))
        ws.cell(row_num, 7, safe_float(row.get("Mins", "")))
        ws.cell(row_num, 8, excel_other_info(row))
        ws.cell(row_num, 9, (row.get("Proofing") or "").strip())

        # Formula-driven calculated columns.
        ws.cell(
            row_num,
            10,
            f'=IF(OR(F{row_num}="",G{row_num}=""),"",ROUND(F{row_num}*G{row_num},2))',
        )
        ws.cell(
            row_num,
            11,
            f'=IF(OR(E{row_num}="",G{row_num}=""),"",ROUND(E{row_num}*G{row_num},2))',
        )
        ws.cell(
            row_num,
            12,
            f'=IF(K{row_num}="","",ROUND(K{row_num}*20%,2))',
        )
        ws.cell(
            row_num,
            13,
            f'=IF(K{row_num}="","",ROUND(K{row_num}+L{row_num},2))',
        )

        ws.cell(row_num, 14, (row.get("Source ID") or "").strip())

        # Keep this simple for the client: no separate Parse Status column.
        ws.cell(
            row_num,
            15,
            (
                f'=TEXTJOIN("; ",TRUE,'
                f'IF(A{row_num}="","Missing Date",""),'
                f'IF(B{row_num}="","Missing Client",""),'
                f'IF(D{row_num}="","Missing Description",""),'
                f'IF(E{row_num}="","Missing Client Rate",""),'
                f'IF(G{row_num}="","Missing Minutes",""))'
            ),
        )
        ws.cell(
            row_num,
            16,
            f'=IF(O{row_num}="","Ready","Needs Review")',
        )

    last_data_row = max(first_data_row, first_data_row + len(rows) - 1)

    # Formats.
    for row_num in range(first_data_row, last_data_row + 1):
        ws.cell(row_num, 1).number_format = "dd/mm/yyyy"

        for col_num in (5, 6, 10, 11, 12, 13):
            ws.cell(row_num, col_num).number_format = '£0.00'

        ws.cell(row_num, 7).number_format = "0.00"

        for col_num in range(1, 17):
            ws.cell(row_num, col_num).alignment = Alignment(
                vertical="top",
                wrap_text=col_num in (4, 8, 14, 15),
            )

    # Highlight fields normally completed/reviewed by the user.
    editable_fill = PatternFill("solid", fgColor="FFFDF2")
    for col_num in (3, 5, 6, 7, 9):
        for row_num in range(first_data_row, last_data_row + 1):
            ws.cell(row_num, col_num).fill = editable_fill

    # Conditional formatting for Status.
    if rows:
        needs_review_fill = PatternFill("solid", fgColor="FFF2CC")
        ready_fill = PatternFill("solid", fgColor="E2F0D9")

        ws.conditional_formatting.add(
            f"P{first_data_row}:P{last_data_row}",
            FormulaRule(
                formula=[f'P{first_data_row}="Needs Review"'],
                fill=needs_review_fill,
                font=Font(color="9C6500", bold=True),
            ),
        )
        ws.conditional_formatting.add(
            f"P{first_data_row}:P{last_data_row}",
            FormulaRule(
                formula=[f'P{first_data_row}="Ready"'],
                fill=ready_fill,
                font=Font(color="375623", bold=True),
            ),
        )

    widths = {
        "A": 12,
        "B": 22,
        "C": 18,
        "D": 38,
        "E": 13,
        "F": 13,
        "G": 10,
        "H": 42,
        "I": 12,
        "J": 13,
        "K": 13,
        "L": 11,
        "M": 13,
        "N": 34,
        "O": 30,
        "P": 16,
    }

    for column, width in widths.items():
        ws.column_dimensions[column].width = width

    ws.freeze_panes = "A5"

    if rows:
        table_ref = f"A4:P{last_data_row}"
        table = Table(displayName="MDOJobsTable", ref=table_ref)
        table.tableStyleInfo = TableStyleInfo(
            name="TableStyleMedium2",
            showFirstColumn=False,
            showLastColumn=False,
            showRowStripes=True,
            showColumnStripes=False,
        )
        ws.add_table(table)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)


def generate_excel(month: str | None, output: str | None) -> Path:
    parsed_rows = load_parsed_preview_rows()
    new_work_rows, completed_excluded = filter_new_work_rows(parsed_rows)
    selected_rows = filter_rows_for_month(new_work_rows, month)

    if output:
        output_path = Path(output)
    elif month:
        pretty = datetime.strptime(month, "%Y-%m").strftime("%B_%Y")
        output_path = Path(f"MDO_Jobs_{pretty}.xlsx")
    else:
        output_path = Path("MDO_Jobs_All_Cached.xlsx")

    build_excel_workbook(selected_rows, output_path, month)

    print()
    print("=" * 80)
    print("EXCEL GENERATION SUMMARY")
    print("=" * 80)
    print(f"Parsed MDO rows:             {len(parsed_rows)}")
    print(f"Completed-work rows skipped: {completed_excluded}")
    print(f"Rows written to Excel:       {len(selected_rows)}")
    if month:
        print(f"Selected upload month:       {month}")
    print(f"Excel file:                  {output_path.resolve()}")

    return output_path



def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--parse-only",
        action="store_true",
        help="Do not connect to IMAP; parse the existing local cache only.",
    )
    parser.add_argument(
        "--month",
        help="Only write jobs whose MDO/email upload date is in YYYY-MM, e.g. 2026-08.",
    )
    parser.add_argument(
        "--output",
        help="Optional output .xlsx path/name.",
    )
    args = parser.parse_args()

    if not args.parse_only:
        fetch_new_direct_mdo()

    # Always rebuild parsed preview from the current cache.
    parse_cache()

    # Then generate the working Excel workbook automatically.
    generate_excel(args.month, args.output)


if __name__ == "__main__":
    main()
