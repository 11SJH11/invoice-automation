#!/usr/bin/env python3
"""
Fingertips Mailbox History Diagnostic

Purpose
-------
Inspect the INBOX without any work-email filtering and report:

- Total messages visible to IMAP
- Oldest visible email Date: header
- Newest visible email Date: header
- Number of emails per calendar day
- Any gaps of 2 or more consecutive days with zero emails

This script is READ ONLY.
It does not move/delete/mark emails as read.
"""

from __future__ import annotations

import email
import imaplib
import os
from collections import Counter
from datetime import date, timedelta
from email.utils import parsedate_to_datetime

IMAP_HOST = "mailcl3.extendcp.co.uk"
IMAP_PORT = 993
MAILBOX = "INBOX"

# Only print this many daily rows around the oldest/newest edge.
# Set to None if you want every single day printed.
PRINT_ALL_DAYS = True


def parse_date_header(raw_headers: bytes):
    try:
        msg = email.message_from_bytes(raw_headers)
        raw_date = msg.get("Date")
        if not raw_date:
            return None
        dt = parsedate_to_datetime(raw_date)
        if dt is None:
            return None
        return dt.date()
    except Exception:
        return None


def date_range(start: date, end: date):
    current = start
    while current <= end:
        yield current
        current += timedelta(days=1)


def main():
    username = os.getenv("FINGERTIPS_EMAIL")
    password = os.getenv("FINGERTIPS_EMAIL_PASSWORD")

    if not username or not password:
        raise SystemExit(
            "Missing credentials.\n\n"
            "In PowerShell, run:\n"
            '$env:FINGERTIPS_EMAIL="cbennett@fingertipstyping.co.uk"\n'
            '$env:FINGERTIPS_EMAIL_PASSWORD="your-password-here"\n'
        )

    print(f"Connecting securely to {IMAP_HOST}:{IMAP_PORT} ...")

    counts = Counter()
    parsed_count = 0
    failed_count = 0

    with imaplib.IMAP4_SSL(IMAP_HOST, IMAP_PORT) as mail:
        mail.login(username, password)
        mail.select(MAILBOX, readonly=True)

        status, data = mail.uid("search", None, "ALL")
        if status != "OK":
            raise RuntimeError("Could not list INBOX messages.")

        uids = data[0].split()
        print(f"INBOX contains {len(uids)} message(s).")
        print("Reading lightweight Date headers...\n")

        for idx, uid in enumerate(uids, start=1):
            status, msg_data = mail.uid("fetch", uid, "(RFC822.HEADER)")
            if status != "OK":
                failed_count += 1
                continue

            raw_headers = None
            for item in msg_data:
                if isinstance(item, tuple) and len(item) >= 2:
                    raw_headers = item[1]
                    break

            if not raw_headers:
                failed_count += 1
                continue

            d = parse_date_header(raw_headers)
            if d is None:
                failed_count += 1
                continue

            counts[d] += 1
            parsed_count += 1

            if idx % 100 == 0 or idx == len(uids):
                print(
                    f"  Checked {idx}/{len(uids)} messages; "
                    f"parsed dates: {parsed_count}",
                    end="\r",
                )

        mail.logout()

    print(" " * 80, end="\r")

    if not counts:
        print("No usable Date headers were found.")
        return

    oldest = min(counts)
    newest = max(counts)

    print("=" * 80)
    print("MAILBOX DATE HISTORY")
    print("=" * 80)
    print(f"Total IMAP messages:       {len(uids)}")
    print(f"Dates successfully parsed: {parsed_count}")
    print(f"Dates failed/missing:      {failed_count}")
    print(f"Oldest email Date:         {oldest:%A %d %B %Y}")
    print(f"Newest email Date:         {newest:%A %d %B %Y}")
    print(f"Calendar span:             {(newest - oldest).days + 1} day(s)")
    print()

    print("=" * 80)
    print("EMAIL COUNT BY DAY")
    print("=" * 80)

    for d in date_range(oldest, newest):
        count = counts.get(d, 0)
        marker = " <-- ZERO" if count == 0 else ""
        print(f"{d:%a %d %b %Y}: {count:4d}{marker}")

    print()
    print("=" * 80)
    print("GAPS OF 2+ CONSECUTIVE DAYS WITH ZERO EMAILS")
    print("=" * 80)

    gaps = []
    gap_start = None
    previous = None

    for d in date_range(oldest, newest):
        if counts.get(d, 0) == 0:
            if gap_start is None:
                gap_start = d
            previous = d
        else:
            if gap_start is not None:
                length = (previous - gap_start).days + 1
                if length >= 2:
                    gaps.append((gap_start, previous, length))
                gap_start = None
                previous = None

    if gap_start is not None:
        length = (previous - gap_start).days + 1
        if length >= 2:
            gaps.append((gap_start, previous, length))

    if not gaps:
        print("No gaps of 2 or more consecutive zero-email days were found.")
    else:
        for start, end, length in gaps:
            if start == end:
                print(f"{start:%d %b %Y}: {length} day")
            else:
                print(
                    f"{start:%d %b %Y} -> {end:%d %b %Y}: "
                    f"{length} consecutive days"
                )

    print()
    print("This script is read-only. No mailbox messages were changed.")


if __name__ == "__main__":
    main()
