#!/usr/bin/env python3
"""
Fingertips Automation - Email Cache

Purpose
-------
Build a reusable local development dataset from the mailbox.

The script:
1. Connects to the Fingertips INBOX in READ-ONLY mode.
2. Caches each unseen email as a raw .eml file.
3. Maintains an index.csv with useful metadata/body previews.
4. Generates a representative review_sample.csv of about 75 emails.

The sample deliberately mixes:
- direct MyDocsOnline uploads
- emails with attachments
- work-like subjects/bodies
- likely marketing/supplier noise
- emails from frequent senders
- random remaining emails

This is NOT the final classifier.  The sample categories are only there to
help us inspect a useful cross-section of real mailbox traffic.

Folder structure
----------------
email_cache/
    all/
        index.csv
        review_sample.csv
        raw/
            *.eml

Credentials
-----------
PowerShell:
    $env:FINGERTIPS_EMAIL="cbennett@fingertipstyping.co.uk"
    $env:FINGERTIPS_EMAIL_PASSWORD="your-password-here"

Run:
    python email_cache.py

Rebuild the sample from the existing cache/index without connecting to IMAP:
    python email_cache.py --sample-only

Optional sample size:
    python email_cache.py --sample-size 100
"""

from __future__ import annotations

import argparse
import csv
import email
import hashlib
import html
import imaplib
import os
import random
import re
from collections import Counter
from datetime import datetime
from email.header import decode_header, make_header
from email.message import Message
from email.utils import parsedate_to_datetime, parseaddr
from pathlib import Path

IMAP_HOST = "mailcl3.extendcp.co.uk"
IMAP_PORT = 993
MAILBOX = "INBOX"

CACHE_ROOT = Path("email_cache") / "all"
RAW_DIR = CACHE_ROOT / "raw"
INDEX_PATH = CACHE_ROOT / "index.csv"
SAMPLE_PATH = CACHE_ROOT / "review_sample.csv"

DEFAULT_SAMPLE_SIZE = 75
RANDOM_SEED = 42

INDEX_COLUMNS = [
    "Source ID",
    "Message ID",
    "IMAP UID",
    "Date",
    "Parsed Date",
    "From",
    "From Email",
    "From Domain",
    "Subject",
    "Attachments",
    "Attachment Count",
    "Body Preview",
    "Cached File",
]

SAMPLE_COLUMNS = [
    "Date",
    "From",
    "From Email",
    "Subject",
    "Attachments",
    "Body Preview",
    "Sample Reason",
    "Actual Label",
    "Notes",
    "Source ID",
    "Cached File",
]

AUDIO_EXTENSIONS = {
    ".wav", ".mp3", ".m4a", ".wma", ".aac", ".flac", ".dss", ".ds2"
}

DOCUMENT_EXTENSIONS = {
    ".doc", ".docx", ".rtf", ".odt", ".txt", ".pdf", ".ppt", ".pptx"
}

WORK_EXTENSIONS = AUDIO_EXTENSIONS | DOCUMENT_EXTENSIONS

WORK_TERMS = [
    "transcription",
    "dictation",
    "audio",
    "typing",
    "typed",
    "work completed",
    "completed typing",
    "file attached",
    "please find attached",
    "work attached",
    "minutes",
    "mins",
    "speakers",
    "verbatim",
    "proofread",
    "proof reading",
    "customer upload",
    "upload to \\",
]

MARKETING_TERMS = [
    "unsubscribe",
    "newsletter",
    "special offer",
    "promotion",
    "marketing",
    "sale",
    "discount",
    "renewal",
    "invoice due",
    "payment reminder",
    "webinar",
    "subscribe",
]

MYDOCS_ADDRESS = "donotreply@mydocsonline.com"


def decode_mime(value: str | None) -> str:
    if not value:
        return ""
    try:
        return str(make_header(decode_header(value)))
    except Exception:
        return value


def sender_address(sender_header: str) -> str:
    return parseaddr(sender_header)[1].strip().lower()


def sender_domain(address: str) -> str:
    if "@" not in address:
        return ""
    return address.rsplit("@", 1)[1].lower()


def parse_email_date(msg: Message) -> datetime | None:
    raw = msg.get("Date")
    if not raw:
        return None
    try:
        return parsedate_to_datetime(raw)
    except Exception:
        return None


def extract_body_text(msg: Message) -> str:
    """
    Extract readable message text while ignoring attachments.
    Prefer text/plain; use HTML only when no plain-text version exists.
    """
    plain_parts: list[str] = []
    html_parts: list[str] = []

    parts = msg.walk() if msg.is_multipart() else [msg]

    for part in parts:
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

    if not plain_parts:
        text = re.sub(r"<style.*?</style>", " ", text, flags=re.I | re.S)
        text = re.sub(r"<script.*?</script>", " ", text, flags=re.I | re.S)
        text = re.sub(r"<br\s*/?>", "\n", text, flags=re.I)
        text = re.sub(r"</p\s*>", "\n", text, flags=re.I)
        text = re.sub(r"<[^>]+>", " ", text)
        text = html.unescape(text)

    text = text.replace("\r\n", "\n").replace("\r", "\n")
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n[ \t]+", "\n", text)
    text = re.sub(r"\n{3,}", "\n\n", text)

    return text.strip()


def attachment_names(msg: Message) -> list[str]:
    names: list[str] = []

    for part in msg.walk():
        filename = part.get_filename()
        if filename:
            decoded = decode_mime(filename).strip()
            if decoded:
                names.append(decoded)

    return names


def make_source_id(msg: Message, uid: str) -> tuple[str, str]:
    message_id = decode_mime(msg.get("Message-ID")).strip()

    if message_id:
        return f"EMAIL:{message_id}", message_id

    return f"EMAIL:UID:{uid}", ""


def cache_filename(source_id: str, parsed_date: datetime | None) -> str:
    digest = hashlib.sha256(source_id.encode("utf-8")).hexdigest()[:20]
    date_prefix = parsed_date.strftime("%Y-%m-%d") if parsed_date else "unknown-date"
    return f"{date_prefix}_{digest}.eml"


def read_index() -> list[dict]:
    if not INDEX_PATH.exists():
        return []

    with INDEX_PATH.open("r", encoding="utf-8-sig", newline="") as f:
        return list(csv.DictReader(f))


def write_index(rows: list[dict]) -> None:
    CACHE_ROOT.mkdir(parents=True, exist_ok=True)

    # Stable chronological-ish ordering makes the file easier to inspect.
    rows = sorted(
        rows,
        key=lambda r: (
            r.get("Parsed Date", ""),
            r.get("IMAP UID", ""),
            r.get("Source ID", ""),
        ),
    )

    with INDEX_PATH.open("w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=INDEX_COLUMNS)
        writer.writeheader()
        writer.writerows(rows)


def fetch_and_cache_all() -> None:
    username = os.getenv("FINGERTIPS_EMAIL")
    password = os.getenv("FINGERTIPS_EMAIL_PASSWORD")

    if not username or not password:
        raise SystemExit(
            "Missing credentials.\n\n"
            "In PowerShell, run:\n"
            '$env:FINGERTIPS_EMAIL="cbennett@fingertipstyping.co.uk"\n'
            '$env:FINGERTIPS_EMAIL_PASSWORD="your-password-here"\n'
        )

    RAW_DIR.mkdir(parents=True, exist_ok=True)

    existing_rows = read_index()
    existing_by_uid = {
        row["IMAP UID"]: row
        for row in existing_rows
        if (row.get("IMAP UID") or "").strip()
    }

    print(f"Connecting securely to {IMAP_HOST}:{IMAP_PORT} ...")
    print(f"Local cache: {CACHE_ROOT.resolve()}")
    print(f"Already indexed: {len(existing_rows)} email(s)\n")

    new_rows: list[dict] = []
    failed = 0

    with imaplib.IMAP4_SSL(IMAP_HOST, IMAP_PORT) as mail:
        mail.login(username, password)
        mail.select(MAILBOX, readonly=True)

        status, data = mail.uid("search", None, "ALL")
        if status != "OK":
            raise RuntimeError("Could not list INBOX messages.")

        uids = data[0].split()
        total = len(uids)

        unseen_uids = [
            uid_bytes
            for uid_bytes in uids
            if uid_bytes.decode(errors="replace") not in existing_by_uid
        ]

        print(f"INBOX currently contains {total} message(s).")
        print(f"Unseen by local cache:       {len(unseen_uids)} message(s).\n")

        for index, uid_bytes in enumerate(unseen_uids, start=1):
            uid = uid_bytes.decode(errors="replace")

            status, msg_data = mail.uid("fetch", uid_bytes, "(RFC822)")
            if status != "OK":
                failed += 1
                continue

            raw_email = None

            for item in msg_data:
                if isinstance(item, tuple) and len(item) >= 2:
                    raw_email = item[1]
                    break

            if raw_email is None:
                failed += 1
                continue

            try:
                msg = email.message_from_bytes(raw_email)

                source_id, message_id = make_source_id(msg, uid)
                parsed_date = parse_email_date(msg)

                sender = decode_mime(msg.get("From"))
                from_email = sender_address(sender)
                subject = decode_mime(msg.get("Subject"))

                attachments = attachment_names(msg)
                body = extract_body_text(msg)
                body_preview = body[:1200]

                if len(body) > 1200:
                    body_preview += " ... [truncated]"

                filename = cache_filename(source_id, parsed_date)
                cached_path = RAW_DIR / filename

                # Preserve exact source message bytes.
                cached_path.write_bytes(raw_email)

                new_rows.append(
                    {
                        "Source ID": source_id,
                        "Message ID": message_id,
                        "IMAP UID": uid,
                        "Date": decode_mime(msg.get("Date")),
                        "Parsed Date": (
                            parsed_date.isoformat()
                            if parsed_date is not None
                            else ""
                        ),
                        "From": sender,
                        "From Email": from_email,
                        "From Domain": sender_domain(from_email),
                        "Subject": subject,
                        "Attachments": " | ".join(attachments),
                        "Attachment Count": str(len(attachments)),
                        "Body Preview": body_preview,
                        "Cached File": str(cached_path),
                    }
                )

            except Exception as exc:
                failed += 1
                print(f"\nCould not process UID {uid}: {exc}")
                continue

            if index % 50 == 0 or index == len(unseen_uids):
                print(
                    f"  Cached {index}/{len(unseen_uids)} unseen email(s)",
                    end="\r",
                )

        mail.logout()

    combined = existing_rows + new_rows
    write_index(combined)

    print(" " * 80, end="\r")
    print("=" * 80)
    print("EMAIL CACHE SUMMARY")
    print("=" * 80)
    print(f"Server emails visible: {total}")
    print(f"Already cached:        {len(existing_rows)}")
    print(f"Newly cached:          {len(new_rows)}")
    print(f"Failures:              {failed}")
    print(f"Total local index:     {len(combined)}")
    print(f"Index:                 {INDEX_PATH.resolve()}")
    print(f"Raw .eml files:        {RAW_DIR.resolve()}")
    print()


def lower_blob(row: dict) -> str:
    return " ".join(
        [
            row.get("Subject", ""),
            row.get("Body Preview", ""),
            row.get("Attachments", ""),
            row.get("From", ""),
        ]
    ).lower()


def has_work_attachment(row: dict) -> bool:
    attachments = row.get("Attachments", "")

    if not attachments:
        return False

    for raw_name in attachments.split("|"):
        name = raw_name.strip().lower()

        for extension in WORK_EXTENSIONS:
            if name.endswith(extension):
                return True

    return False


def looks_work_like(row: dict) -> bool:
    blob = lower_blob(row)
    return any(term in blob for term in WORK_TERMS)


def looks_marketing_or_supplier(row: dict) -> bool:
    blob = lower_blob(row)
    return any(term in blob for term in MARKETING_TERMS)


def is_direct_mdo(row: dict) -> bool:
    sender = (row.get("From Email") or "").lower()
    subject = (row.get("Subject") or "").lower().strip()

    return sender == MYDOCS_ADDRESS and subject.startswith("upload to \\")


def choose_from_bucket(
    selected: dict[str, tuple[dict, set[str]]],
    bucket_rows: list[dict],
    reason: str,
    target_count: int,
    rng: random.Random,
) -> None:
    """
    Add unique rows from one sampling bucket.

    selected maps Source ID -> (row, set of sample reasons).
    If a row is selected by multiple buckets, keep one row and record every reason.
    """
    if target_count <= 0 or not bucket_rows:
        return

    candidates = bucket_rows[:]
    rng.shuffle(candidates)

    added_for_this_bucket = 0

    for row in candidates:
        source_id = row.get("Source ID", "")

        if not source_id:
            continue

        if source_id in selected:
            selected[source_id][1].add(reason)
            continue

        selected[source_id] = (row, {reason})
        added_for_this_bucket += 1

        if added_for_this_bucket >= target_count:
            break


def build_representative_sample(sample_size: int) -> None:
    rows = read_index()

    if not rows:
        raise SystemExit(
            f"No cache index found at {INDEX_PATH}.\n"
            "Run without --sample-only first."
        )

    rng = random.Random(RANDOM_SEED)

    # Sender frequency is calculated across the full cached mailbox.
    sender_counts = Counter(
        (row.get("From Email") or row.get("From") or "").lower()
        for row in rows
    )

    frequent_sender_rows = [
        row
        for row in rows
        if sender_counts[
            (row.get("From Email") or row.get("From") or "").lower()
        ] >= 5
    ]

    buckets = {
        "Direct MDO upload": [row for row in rows if is_direct_mdo(row)],
        "Has work-like attachment": [row for row in rows if has_work_attachment(row)],
        "Work-like terms": [row for row in rows if looks_work_like(row)],
        "Likely marketing/supplier": [
            row for row in rows if looks_marketing_or_supplier(row)
        ],
        "Frequent sender": frequent_sender_rows,
        "No attachment": [
            row
            for row in rows
            if int(row.get("Attachment Count") or 0) == 0
        ],
    }

    # Approximate allocation. Duplicates across buckets are collapsed.
    # Remaining slots are filled randomly from the whole mailbox.
    allocations = {
        "Direct MDO upload": max(5, round(sample_size * 0.12)),
        "Has work-like attachment": max(10, round(sample_size * 0.20)),
        "Work-like terms": max(10, round(sample_size * 0.20)),
        "Likely marketing/supplier": max(8, round(sample_size * 0.15)),
        "Frequent sender": max(8, round(sample_size * 0.15)),
        "No attachment": max(8, round(sample_size * 0.12)),
    }

    selected: dict[str, tuple[dict, set[str]]] = {}

    for reason, bucket_rows in buckets.items():
        choose_from_bucket(
            selected,
            bucket_rows,
            reason,
            allocations[reason],
            rng,
        )

    # Fill any remaining slots from emails not yet chosen.
    remainder = [
        row
        for row in rows
        if row.get("Source ID", "") not in selected
    ]
    rng.shuffle(remainder)

    for row in remainder:
        if len(selected) >= sample_size:
            break

        source_id = row.get("Source ID", "")
        if source_id:
            selected[source_id] = (row, {"Random remainder"})

    # If overlap made the result larger than requested, trim deterministically.
    selected_items = list(selected.values())
    rng.shuffle(selected_items)
    selected_items = selected_items[:sample_size]

    # Sort final workbook-like CSV chronologically for easy reading.
    selected_items.sort(
        key=lambda item: (
            item[0].get("Parsed Date", ""),
            item[0].get("From Email", ""),
            item[0].get("Subject", ""),
        )
    )

    output_rows: list[dict] = []

    for row, reasons in selected_items:
        output_rows.append(
            {
                "Date": row.get("Parsed Date", ""),
                "From": row.get("From", ""),
                "From Email": row.get("From Email", ""),
                "Subject": row.get("Subject", ""),
                "Attachments": row.get("Attachments", ""),
                "Body Preview": row.get("Body Preview", ""),
                "Sample Reason": "; ".join(sorted(reasons)),
                "Actual Label": "",
                "Notes": "",
                "Source ID": row.get("Source ID", ""),
                "Cached File": row.get("Cached File", ""),
            }
        )

    CACHE_ROOT.mkdir(parents=True, exist_ok=True)

    with SAMPLE_PATH.open("w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=SAMPLE_COLUMNS)
        writer.writeheader()
        writer.writerows(output_rows)

    print("=" * 80)
    print("REPRESENTATIVE SAMPLE SUMMARY")
    print("=" * 80)
    print(f"Cached emails available: {len(rows)}")
    print(f"Requested sample size:   {sample_size}")
    print(f"Sample rows produced:    {len(output_rows)}")
    print(f"Review file:             {SAMPLE_PATH.resolve()}")
    print()
    print("Sampling buckets in the mailbox:")
    for reason, bucket_rows in buckets.items():
        print(f"  {reason:<28} {len(bucket_rows):>5}")
    print()
    print(
        "Open review_sample.csv in Excel and inspect it. "
        "Actual Label and Notes are intentionally blank."
    )


def main() -> None:
    parser = argparse.ArgumentParser()

    parser.add_argument(
        "--sample-only",
        action="store_true",
        help="Do not connect to IMAP; rebuild the review sample from the current index.",
    )

    parser.add_argument(
        "--sample-size",
        type=int,
        default=DEFAULT_SAMPLE_SIZE,
        help=f"Number of emails in review sample (default: {DEFAULT_SAMPLE_SIZE}).",
    )

    args = parser.parse_args()

    if args.sample_size < 10:
        raise SystemExit("--sample-size must be at least 10.")

    if not args.sample_only:
        fetch_and_cache_all()

    build_representative_sample(args.sample_size)


if __name__ == "__main__":
    main()
