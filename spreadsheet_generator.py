#!/usr/bin/env python3
"""
Fingertips Automation - Master Spreadsheet Generator

Purpose
-------
Combine the work-detection pieces already built into one monthly Excel sheet.

Inputs:
    parsed_jobs_preview.csv
        Job-level MyDocsOnline parse output from mydocsonline_import_v14/v15.

    email_cache/all/classified_emails.csv
        Email classifier output.

    sender_registry.csv
        Optional sender names/roles.

Output:
    generated_spreadsheets/Fingertips_Jobs_<Month>_<Year>.xlsx

Current behaviour
---------------------
1. Selects MDO jobs by EMAIL/UPLOAD month.
2. Excludes MDO uploads into "Transcriptions done" folders.
3. Converts NEW_CLIENT_WORK emails into job rows.
   - One audio attachment = one job row.
   - Multiple audio attachments = one row per audio file.
4. Deduplicates exact Source IDs.
5. Conservatively matches TYPIST_RETURN emails to jobs using:
   A. exact job references (strongest), then
   B. strong filename/description similarity.
6. When a return is confidently matched, it can fill:
   - Typist
   - Mins
   - Speakers retained as internal parsing evidence
7. Leaves unknown Client Rate / Typist Rate blank.
8. Adds Excel formulas for:
   - Typist total
   - Sub total
   - VAT
   - Total
   - Validation Issues
   - Status
9. Infers first-pass Client Rate / Typist Rate from the supplied price lists.
10. Supports exact client-specific overrides through rate_overrides.csv.
11. Retains inferred-rate evidence internally for audit/debugging.
12. Recognises speaker forms such as "2ppl", "3 people", "4 speakers".
13. Uses collision-safe turnaround parsing so "24 hours" is not read as "4 hour".
14. Does NOT overwrite source CSV files.

This is intentionally conservative.  A missing value should become Needs Review
rather than being guessed.

Run:
    python spreadsheet_generator.py --month 2026-08

Optional explicit paths:
    python spreadsheet_generator.py ^
        --month 2026-08 ^
        --mdo parsed_jobs_preview.csv ^
        --classified email_cache/all/classified_emails.csv ^
        --registry sender_registry.csv

Dependency:
    python -m pip install openpyxl
"""

from __future__ import annotations

import argparse
import csv
import os
import re
from datetime import datetime
from difflib import SequenceMatcher
from pathlib import Path

from openpyxl import Workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.table import Table, TableStyleInfo


DEFAULT_MDO = Path("email_cache") / "mydocsonline" / "parsed_jobs_preview.csv"
DEFAULT_CLASSIFIED = Path("email_cache") / "all" / "classified_emails.csv"
DEFAULT_REGISTRY = Path("sender_registry.csv")
DEFAULT_OUTPUT_DIR = Path("generated_spreadsheets")
DEFAULT_RATE_OVERRIDES = Path("rate_overrides.csv")
DEFAULT_CLIENT_ALIASES = Path("client_aliases.csv")

AUDIO_EXTENSIONS = {
    ".wav", ".mp3", ".m4a", ".wma", ".aac", ".flac", ".dss", ".ds2"
}

JOB_COLUMNS = [
    "Date",
    "Client",
    "Typist",
    "Description",
    "Client Rate",
    "Typist Rate",
    "Mins",
    "Email Subject",
    "Proofing",
    "Typist total",
    "Sub total",
    "VAT",
    "Total",
    "Source ID",
    "Validation Issues",
    "Status",
]

CLIENT_DOMAIN_MAP = {
    "acornsurveyors.com": "Acorn Surveyors",
    "eastofengland.coop": "East of England",
    "tpsmanagement.uk": "TPS Management",
}

# Normal words that do not help filename matching.
MATCH_STOP_WORDS = {
    "the", "and", "for", "to", "of", "a", "an", "re", "fw", "fwd",
    "typing", "transcript", "transcription", "completed", "attached",
    "file", "work", "audio", "dictation", "meeting"
}

JOB_REF_RE = re.compile(r"\b(?:A\d{5,}|DS\d{5,}|KH-\d+)\b", re.I)


def read_csv(path: Path) -> list[dict]:
    if not path.exists():
        raise SystemExit(f"Cannot find input file: {path}")

    with path.open("r", encoding="utf-8-sig", newline="") as f:
        return list(csv.DictReader(f))


def parse_month(value: str) -> tuple[int, int]:
    try:
        dt = datetime.strptime(value, "%Y-%m")
    except ValueError:
        raise SystemExit("--month must be in YYYY-MM format, for example 2026-08")
    return dt.year, dt.month


def parse_date(value: str | None) -> datetime | None:
    text = (value or "").strip()
    if not text:
        return None

    formats = (
        "%d/%m/%Y",
        "%Y-%m-%d",
        "%Y-%m-%dT%H:%M:%S%z",
        "%Y-%m-%dT%H:%M:%S",
    )

    for fmt in formats:
        try:
            return datetime.strptime(text, fmt)
        except ValueError:
            continue

    # ISO timestamps may contain fractional seconds.
    try:
        return datetime.fromisoformat(text)
    except Exception:
        return None


def in_month(value: str | None, year: int, month: int) -> bool:
    dt = parse_date(value)
    return bool(dt and dt.year == year and dt.month == month)


def display_date(value: str | None) -> str:
    dt = parse_date(value)
    return dt.strftime("%d/%m/%Y") if dt else ""


def filename_extension(name: str) -> str:
    return os.path.splitext((name or "").strip())[1].lower()


def split_pipe(value: str | None) -> list[str]:
    return [x.strip() for x in (value or "").split("|") if x.strip()]


def clean_filename_description(filename: str) -> str:
    stem = os.path.splitext(os.path.basename(filename))[0]

    # Remove common leading dates:
    stem = re.sub(
        r"^\s*(?:20\d{2})[- _](?:0?[1-9]|1[0-2])[- _](?:0?[1-9]|[12]\d|3[01])[- _]*",
        "",
        stem,
    )
    stem = re.sub(
        r"^\s*(?:0?[1-9]|[12]\d|3[01])[- _.](?:0?[1-9]|1[0-2])[- _.](?:20\d{2})[- _]*",
        "",
        stem,
    )

    # Remove a few common recorder extensions that can appear inside a stem.
    stem = re.sub(r"\.(?:m4a|mp3|wav|wma|ds2|dss)$", "", stem, flags=re.I)
    stem = re.sub(r"[_]+", " ", stem)
    stem = re.sub(r"\s+", " ", stem).strip(" -_")
    return stem or filename


def normalise_match_text(value: str) -> str:
    text = clean_filename_description(value).lower()
    text = re.sub(r"[^a-z0-9]+", " ", text)
    tokens = [
        token
        for token in text.split()
        if token not in MATCH_STOP_WORDS and (len(token) > 1 or token.isdigit())
    ]
    return " ".join(tokens)


def extract_refs(*values: str) -> set[str]:
    refs: set[str] = set()
    for value in values:
        for match in JOB_REF_RE.finditer(value or ""):
            refs.add(match.group(0).upper())
    return refs


def registry_maps(rows: list[dict]) -> tuple[dict[str, str], dict[str, str]]:
    names: dict[str, str] = {}
    roles: dict[str, str] = {}

    for row in rows:
        if (row.get("Active") or "").strip().lower() not in {"yes", "y", "true", "1"}:
            continue

        email_address = (row.get("Email Address") or "").strip().lower()
        if not email_address:
            continue

        names[email_address] = (row.get("Name / Organisation") or "").strip()
        roles[email_address] = (row.get("Role") or "").strip().upper()

    return names, roles


def infer_person_name(from_header: str, from_email: str) -> str:
    # Prefer human-looking display name from "Name <address>".
    display = re.sub(r"<[^>]+>", "", from_header or "").strip().strip('"')
    if display and "@" not in display and len(display) <= 80:
        return display

    local = (from_email or "").split("@", 1)[0]
    local = re.sub(r"[._\-]+", " ", local)
    local = re.sub(r"\d+", "", local)
    local = re.sub(r"\s+", " ", local).strip()
    return local.title()


def infer_client(row: dict, registry_names: dict[str, str]) -> str:
    sender = (row.get("From Email") or "").strip().lower()
    domain = sender.rsplit("@", 1)[1] if "@" in sender else ""

    # Exact registered client organisation.
    registered_name = registry_names.get(sender, "")
    if registered_name:
        return registered_name

    if domain in CLIENT_DOMAIN_MAP:
        return CLIENT_DOMAIN_MAP[domain]

    subject = row.get("Subject") or ""
    lower_subject = subject.lower()

    if "acorn surveyors" in lower_subject:
        return "Acorn Surveyors"
    if "east of england" in lower_subject:
        return "East of England"

    # Conservative fallback: derive organisation-ish text from domain.
    if domain:
        root = domain.split(".")[0]
        root = re.sub(r"[-_]+", " ", root).strip()
        return root.title()

    return ""


def build_mdo_jobs(mdo_rows: list[dict], year: int, month: int) -> list[dict]:
    jobs: list[dict] = []

    for row in mdo_rows:
        # Invoice month is based on when the job/upload email arrived.
        if not in_month(row.get("Email Date"), year, month):
            continue

        path = (row.get("MDO Path") or "").lower()
        if "transcriptions done" in path:
            continue

        info_bits = []

        date_source = (row.get("Date Source") or "").strip()
        filename_date = (row.get("Date") or "").strip()
        email_date = (row.get("Email Date") or "").strip()

        if date_source.lower() == "filename" and filename_date and filename_date != email_date:
            info_bits.append(f"Filename date: {filename_date}")

        if row.get("Speakers"):
            info_bits.append(f"Speakers: {row['Speakers']}")

        if row.get("Rate / Turnaround"):
            info_bits.append(f"Rate/turnaround: {row['Rate / Turnaround']}")

        if row.get("Other info"):
            info_bits.append(row["Other info"].strip())

        uploaded = (row.get("Uploaded Filename") or "").strip()
        if uploaded:
            info_bits.append(f"Source file: {uploaded}")

        jobs.append(
            {
                "Date": display_date(row.get("Email Date")),
                "Client": (row.get("Client") or "").strip(),
                "Typist": (row.get("Typist") or "").strip(),
                "Description": (row.get("Description") or clean_filename_description(uploaded)).strip(),
                "Client Rate": (row.get("Client Rate") or "").strip(),
                "Typist Rate": (row.get("Typist Rate") or "").strip(),
                "Mins": (row.get("Mins") or "").strip(),
                "Email Subject": (row.get("Subject") or "").replace("\r", " ").replace("\n", " ").strip(),
                "Other info": "; ".join(bit for bit in info_bits if bit),
                "Proofing": (row.get("Proofing") or "").strip(),
                "Source ID": (row.get("Source ID") or "").strip(),
                "_source_type": "MDO",
                "_source_file": uploaded,
                "_refs": extract_refs(
                    row.get("Description") or "",
                    uploaded,
                    row.get("Other info") or "",
                ),
            }
        )

    return jobs


def build_direct_jobs(
    classified_rows: list[dict],
    registry_names: dict[str, str],
    year: int,
    month: int,
) -> list[dict]:
    jobs: list[dict] = []

    for row in classified_rows:
        if row.get("Classification") != "NEW_CLIENT_WORK":
            continue
        if not in_month(row.get("Date"), year, month):
            continue

        attachments = split_pipe(row.get("Matched Work Files"))
        audio_files = [a for a in attachments if filename_extension(a) in AUDIO_EXTENSIONS]

        # If classifier marked new work but failed to expose the audio list,
        # create one conservative row from the subject rather than silently losing it.
        units = audio_files or [""]

        for index, audio_file in enumerate(units, start=1):
            parent_source = (row.get("Source ID") or "").strip()
            source_id = (
                f"{parent_source}#FILE:{index}"
                if len(units) > 1
                else parent_source
            )

            subject = (row.get("Subject") or "").replace("\r", " ").replace("\n", " ")
            description = (
                clean_filename_description(audio_file)
                if audio_file
                else subject.strip()
            )

            refs = extract_refs(
                row.get("Matched Job Reference") or "",
                subject,
                audio_file,
            )

            info_bits = ["Source: Direct email"]
            if audio_file:
                info_bits.append(f"Source file: {audio_file}")
            if refs:
                info_bits.append("Reference: " + ", ".join(sorted(refs)))

            jobs.append(
                {
                    "Date": display_date(row.get("Date")),
                    "Client": infer_client(row, registry_names),
                    "Typist": "",
                    "Description": description,
                    "Client Rate": "",
                    "Typist Rate": "",
                    "Mins": (row.get("Matched Minutes") or "").split("|")[0].strip(),
                    "Email Subject": subject.strip(),
                    "Other info": "; ".join(info_bits),
                    "Proofing": "",
                    "Source ID": source_id,
                    "_source_type": "DIRECT_EMAIL",
                    "_source_file": audio_file,
                    "_refs": refs,
                }
            )

    return jobs


def job_match_score(job: dict, returned_row: dict) -> tuple[float, str]:
    return_refs = extract_refs(
        returned_row.get("Matched Job Reference") or "",
        returned_row.get("Subject") or "",
        returned_row.get("Matched Work Files") or "",
    )
    job_refs = set(job.get("_refs") or set())

    shared_refs = job_refs & return_refs
    if shared_refs:
        return 1.0, "reference " + ", ".join(sorted(shared_refs))

    return_files = split_pipe(returned_row.get("Matched Work Files"))
    if not return_files:
        return 0.0, ""

    job_texts = [
        normalise_match_text(job.get("Description") or ""),
        normalise_match_text(job.get("_source_file") or ""),
    ]
    job_texts = [x for x in job_texts if x]

    best = 0.0
    best_name = ""

    for returned_file in return_files:
        returned_text = normalise_match_text(returned_file)
        if not returned_text:
            continue

        returned_numbers = {
            token for token in returned_text.split() if token.isdigit()
        }

        for job_text in job_texts:
            job_numbers = {
                token for token in job_text.split() if token.isdigit()
            }

            # "Part 1" must not match "Part 2".  A combined return such as
            # "Part 1 and Part 2" may still identify the typist for each part,
            # but metadata from that combined return is handled separately.
            if job_numbers and returned_numbers and not (job_numbers & returned_numbers):
                continue

            ratio = SequenceMatcher(None, job_text, returned_text).ratio()

            a = set(job_text.split())
            b = set(returned_text.split())
            overlap = len(a & b) / max(1, len(a | b))

            score = max(ratio, overlap)

            # Require at least two meaningful shared tokens for a filename
            # similarity match unless the strings are extremely close.
            shared_token_count = len(a & b)
            if shared_token_count < 2 and ratio < 0.90:
                score = 0.0

            if score > best:
                best = score
                best_name = returned_file

    return best, f"filename {best_name}" if best_name else ""


def return_metadata_safe_for_job(job: dict, returned: dict) -> bool:
    """
    Minutes/speakers are only copied when the return identifies one job
    unambiguously.  This prevents a combined "Pt 1 and Pt 2 - 8 mins" return
    from placing 8 minutes onto BOTH rows.
    """
    return_refs = extract_refs(
        returned.get("Matched Job Reference") or "",
        returned.get("Subject") or "",
        returned.get("Matched Work Files") or "",
    )
    job_refs = set(job.get("_refs") or set())

    if job_refs & return_refs:
        return True

    job_texts = {
        normalise_match_text(job.get("Description") or ""),
        normalise_match_text(job.get("_source_file") or ""),
    }
    job_texts.discard("")

    for returned_file in split_pipe(returned.get("Matched Work Files")):
        returned_text = normalise_match_text(returned_file)
        if returned_text and returned_text in job_texts:
            return True

    return False


def enrich_from_typist_returns(
    jobs: list[dict],
    classified_rows: list[dict],
    registry_names: dict[str, str],
    year: int,
    month: int,
) -> tuple[int, int]:
    """
    Pick at most ONE typist-return match per job.

    Multiple return emails can quote/resend the same completed transcript, so
    attaching every matching return to the same job would be misleading.
    We compare all returns, collapse duplicate evidence from the same sender,
    and only apply a match when the best sender is clearly ahead.
    """
    matched = 0
    ambiguous = 0

    returns = [
        row for row in classified_rows
        if row.get("Classification") == "TYPIST_RETURN"
        and in_month(row.get("Date"), year, month)
    ]

    for job in jobs:
        # Best candidate per sender. Duplicate/resend emails from one typist
        # should not create artificial ambiguity.
        best_by_sender: dict[str, tuple[float, dict, str]] = {}

        for returned in returns:
            score, why = job_match_score(job, returned)
            if score < 0.82:
                continue

            sender = (returned.get("From Email") or "").strip().lower()
            sender_key = sender or (returned.get("From") or "").strip().lower()

            previous = best_by_sender.get(sender_key)

            # Prefer higher score. If equal, prefer the return containing
            # minutes/speaker metadata because it is more useful.
            metadata_count = int(bool(split_pipe(returned.get("Matched Minutes"))))
            metadata_count += int(bool(split_pipe(returned.get("Matched Speakers"))))

            if previous is None:
                best_by_sender[sender_key] = (score, returned, why)
            else:
                prev_score, prev_row, _ = previous
                prev_metadata = int(bool(split_pipe(prev_row.get("Matched Minutes"))))
                prev_metadata += int(bool(split_pipe(prev_row.get("Matched Speakers"))))

                if score > prev_score or (
                    abs(score - prev_score) < 0.0001
                    and metadata_count > prev_metadata
                ):
                    best_by_sender[sender_key] = (score, returned, why)

        candidates = sorted(
            best_by_sender.values(),
            key=lambda item: item[0],
            reverse=True,
        )

        if not candidates:
            continue

        # Different typists with nearly equal matches means we cannot safely
        # know which return belongs to this job.
        if len(candidates) > 1 and candidates[1][0] >= candidates[0][0] - 0.03:
            ambiguous += 1
            continue

        score, returned, why = candidates[0]

        sender = (returned.get("From Email") or "").strip().lower()
        typist = registry_names.get(sender) or infer_person_name(
            returned.get("From") or "",
            sender,
        )

        if not job.get("Typist") and typist:
            job["Typist"] = typist

        metadata_safe = return_metadata_safe_for_job(job, returned)

        returned_minutes = split_pipe(returned.get("Matched Minutes"))
        if metadata_safe and not job.get("Mins") and len(returned_minutes) == 1:
            job["Mins"] = returned_minutes[0]

        speakers = split_pipe(returned.get("Matched Speakers"))
        info_bits = [job.get("Other info", "")]

        if (
            metadata_safe
            and len(speakers) == 1
            and "Speakers:" not in job.get("Other info", "")
        ):
            info_bits.append(f"Speakers: {speakers[0]}")

        if not metadata_safe and (returned_minutes or speakers):
            info_bits.append(
                "Return contains mins/speakers but they were not auto-applied "
                "because the return may cover multiple jobs"
            )

        info_bits.append(
            f"Matched typist return: {returned.get('Subject', '').strip()} "
            f"({why}, score {score:.2f})"
        )
        job["Other info"] = "; ".join(x for x in info_bits if x)

        matched += 1

    return matched, ambiguous


def _job_date_for_dedupe(job: dict) -> datetime | None:
    return parse_date(job.get("Date"))


def _job_quality(job: dict) -> int:
    score = 0
    extension = filename_extension(job.get("_source_file") or "")

    if extension in AUDIO_EXTENSIONS:
        score += 10
    if (job.get("Mins") or "").strip():
        score += 3
    if (job.get("Typist") or "").strip():
        score += 2
    return score


def _jobs_are_semantic_duplicates(left: dict, right: dict) -> bool:
    left_client = normalise_match_text(left.get("Client") or "")
    right_client = normalise_match_text(right.get("Client") or "")
    left_desc = normalise_match_text(left.get("Description") or "")
    right_desc = normalise_match_text(right.get("Description") or "")

    if not left_desc or left_client != right_client or left_desc != right_desc:
        return False

    left_refs = set(left.get("_refs") or set())
    right_refs = set(right.get("_refs") or set())
    if left_refs and right_refs and left_refs & right_refs:
        return True

    left_file = normalise_match_text(left.get("_source_file") or "")
    right_file = normalise_match_text(right.get("_source_file") or "")
    if left_file and right_file and left_file == right_file:
        return True

    left_date = _job_date_for_dedupe(left)
    right_date = _job_date_for_dedupe(right)
    return bool(
        left_date
        and right_date
        and abs((left_date.date() - right_date.date()).days) <= 2
    )


def _prefer_duplicate(left: dict, right: dict) -> dict:
    left_score = _job_quality(left)
    right_score = _job_quality(right)

    if right_score > left_score:
        return right
    if left_score > right_score:
        return left

    left_date = _job_date_for_dedupe(left)
    right_date = _job_date_for_dedupe(right)
    if left_date and right_date and right_date > left_date:
        return right

    return left


def deduplicate_jobs(jobs: list[dict]) -> tuple[list[dict], int]:
    exact_seen: set[str] = set()
    first_pass: list[dict] = []
    duplicates = 0

    for job in jobs:
        source_id = (job.get("Source ID") or "").strip()

        if source_id and source_id in exact_seen:
            duplicates += 1
            continue

        if source_id:
            exact_seen.add(source_id)

        first_pass.append(job)

    output: list[dict] = []

    for job in first_pass:
        duplicate_index = None

        for idx, existing in enumerate(output):
            if _jobs_are_semantic_duplicates(existing, job):
                duplicate_index = idx
                break

        if duplicate_index is None:
            output.append(job)
            continue

        output[duplicate_index] = _prefer_duplicate(output[duplicate_index], job)
        duplicates += 1

    return output, duplicates


CLIENT_ALIAS_COLUMNS = [
    "Detected Client",
    "Invoice Client",
    "Notes",
    "Active",
]


def create_client_aliases_if_missing(path: Path) -> None:
    if path.exists():
        return

    with path.open("w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=CLIENT_ALIAS_COLUMNS)
        writer.writeheader()

    print(f"Created optional client alias file: {path.resolve()}")


def load_client_aliases(path: Path) -> dict[str, str]:
    create_client_aliases_if_missing(path)
    aliases: dict[str, str] = {}

    with path.open("r", encoding="utf-8-sig", newline="") as f:
        for row in csv.DictReader(f):
            active = (row.get("Active") or "").strip().lower()
            if active not in {"yes", "y", "true", "1"}:
                continue

            detected = (row.get("Detected Client") or "").strip()
            invoice_client = (row.get("Invoice Client") or "").strip()

            if detected and invoice_client:
                aliases[detected.lower()] = invoice_client

    return aliases


def apply_client_aliases(jobs: list[dict], aliases: dict[str, str]) -> int:
    changed = 0

    for job in jobs:
        detected = (job.get("Client") or "").strip()
        replacement = aliases.get(detected.lower())

        if replacement and replacement != detected:
            job["Client"] = replacement
            info = (job.get("Other info") or "").strip()
            note = f"Client alias: {detected} -> {replacement}"
            job["Other info"] = "; ".join(x for x in (info, note) if x)
            changed += 1

    return changed



RATE_OVERRIDE_COLUMNS = [
    "Client",
    "Client Rate",
    "Typist Rate",
    "Notes",
    "Active",
]

DISCUSSION_TERMS = {
    "interview",
    "meeting",
    "mtg",
    "hearing",
    "discussion",
    "disciplinary",
    "investigation",
    "appeal",
    "grievance",
}

MEDICAL_TERMS = {
    "medical",
    "clinic",
    "clinical",
    "patient",
    "doctor",
    "gp",
    "hospital",
}

SAME_DAY_TERMS = {
    "same day",
    "sameday",
}

OVERNIGHT_TERMS = {
    "overnight",
}

FOUR_HOUR_PATTERNS = [
    re.compile(r"(?<!\d)4\s*(?:-|\s)?hours?\b", re.I),
    re.compile(r"\bfour\s*(?:-|\s)?hours?\b", re.I),
]

WEEKEND_TERMS = {
    "weekend",
}

THREE_FIVE_DAY_TERMS = {
    "3-5 day",
    "3 to 5 day",
    "3 day",
    "4 day",
    "5 day",
}

ONE_WEEK_TERMS = {
    "one week",
    "1 week",
    "7 day",
    "7-day",
}

VERBATIM_TERMS = {
    "verbatim",
}


def create_rate_overrides_if_missing(path: Path) -> None:
    """
    Optional client-specific rate overrides.

    Example:
        Acorn Surveyors,1.25,0.65,Confirmed bespoke client rate,YES

    Any active exact-client override wins over the general 2009 rate rules.
    """
    if path.exists():
        return

    with path.open("w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=RATE_OVERRIDE_COLUMNS)
        writer.writeheader()

    print(f"Created optional rate override file: {path.resolve()}")


def load_rate_overrides(path: Path) -> dict[str, dict]:
    create_rate_overrides_if_missing(path)

    overrides: dict[str, dict] = {}

    with path.open("r", encoding="utf-8-sig", newline="") as f:
        for row in csv.DictReader(f):
            active = (row.get("Active") or "").strip().lower()
            if active not in {"yes", "y", "true", "1"}:
                continue

            client = (row.get("Client") or "").strip()
            if not client:
                continue

            overrides[client.lower()] = row

    return overrides


def _contains_any_rate_term(text: str, terms: set[str]) -> bool:
    text = text.lower()
    return any(term in text for term in terms)



def _matches_any_pattern(text: str, patterns: list[re.Pattern]) -> bool:
    return any(pattern.search(text) for pattern in patterns)


def _turnaround_flags(text: str) -> dict[str, bool]:
    """
    Parse turnaround cues without substring collisions.

    In particular, "24 hours" must NOT match the "4 hour" rule.
    """
    lower_text = text.lower()

    return {
        "same_day": _contains_any_rate_term(lower_text, SAME_DAY_TERMS),
        "overnight": _contains_any_rate_term(lower_text, OVERNIGHT_TERMS),
        "four_hour": _matches_any_pattern(lower_text, FOUR_HOUR_PATTERNS),
        "weekend": _contains_any_rate_term(lower_text, WEEKEND_TERMS),
        "three_five_day": _contains_any_rate_term(lower_text, THREE_FIVE_DAY_TERMS),
        "one_week": _contains_any_rate_term(lower_text, ONE_WEEK_TERMS),
    }


def _speaker_count(job: dict) -> int | None:
    """
    Prefer evidence belonging to the CURRENT job.

    MDO notifications can contain multiple uploaded filenames and some parsed
    free-text fields may mention a sibling file.  Therefore Description and
    the current source filename are checked before broad Other info.
    """
    patterns = [
        re.compile(r"\bSpeakers:\s*(\d{1,2})\b", re.I),
        re.compile(r"\b(\d{1,2})\s*speakers?\b", re.I),
        re.compile(r"\b(\d{1,2})\s*people\b", re.I),
        re.compile(r"\b(\d{1,2})\s*persons?\b", re.I),
        re.compile(r"\b(\d{1,2})\s*ppl\b", re.I),
        re.compile(r"\b(\d{1,2})ppl\b", re.I),
        re.compile(r"\b(\d{1,2})\s*person\b", re.I),
        re.compile(r"\bone[\s-]*to[\s-]*one\b", re.I),
        re.compile(r"\b1[\s-]*to[\s-]*1\b", re.I),
    ]

    evidence_sources = [
        job.get("Description", ""),
        job.get("_source_file", ""),
        job.get("Other info", ""),
    ]

    for text in evidence_sources:
        text = text or ""
        for pattern in patterns:
            match = pattern.search(text)
            if match:
                if match.lastindex:
                    try:
                        return int(match.group(1))
                    except ValueError:
                        pass
                else:
                    return 2

    return None


def _general_rate_inference(job: dict) -> tuple[str, str, str]:
    """
    Returns:
        (client_rate, typist_rate, explanation)

    Source rules encoded from supplied April 2009 price lists.

    CLIENT:
      One-person: standard 1.30, same-day 1.50, weekend 1.60,
                  overnight 1.50, four-hour 1.70.
      Medical one-person: +0.20.
      Interviews/discussions: 1-2 speakers 1.40, 3-4 1.50,
                              5-6 1.60, then +0.10 / two speakers.
      Interview verbatim: +0.20.
      Interview medical: +0.30.
      3-5 day turnaround: -0.05.
      One week: -0.10.

    TYPIST:
      One-person: standard 0.65, same-day 0.75,
                  overnight 0.75, four-hour 0.85.
      Medical one-person: +0.10.
      Interviews/discussions: 1-2 0.70, 3-4 0.75,
                              5-6 0.80, 7-8 0.85, etc.
      Interview verbatim: +0.10.
      Interview medical: +0.15.

    Important:
      No typist weekend rate is stated in the supplied sheet.
      We therefore leave it blank rather than inventing one.
    """
    client = (job.get("Client") or "").strip()
    description = (job.get("Description") or "").strip()
    info = (job.get("Other info") or "").strip()
    combined = f"{client} {description} {info}".lower()

    speakers = _speaker_count(job)

    is_medical = (
        "snodland medical centre" in client.lower()
        or _contains_any_rate_term(combined, MEDICAL_TERMS)
    )

    is_discussion = (
        speakers is not None and speakers >= 2
    ) or _contains_any_rate_term(combined, DISCUSSION_TERMS)

    is_verbatim = _contains_any_rate_term(combined, VERBATIM_TERMS)

    turnaround = _turnaround_flags(combined)
    same_day = turnaround["same_day"]
    overnight = turnaround["overnight"]
    four_hour = turnaround["four_hour"]
    weekend = turnaround["weekend"]
    three_five_day = turnaround["three_five_day"]
    one_week = turnaround["one_week"]

    notes: list[str] = []
    client_rate: float | None = None
    typist_rate: float | None = None

    if is_discussion:
        # Speaker count unknown: use the lowest discussion tier as an
        # explicit baseline and flag it for checking.
        effective_speakers = speakers if speakers is not None else 2

        if effective_speakers <= 2:
            client_rate = 1.40
            typist_rate = 0.70
        elif effective_speakers <= 4:
            client_rate = 1.50
            typist_rate = 0.75
        elif effective_speakers <= 6:
            client_rate = 1.60
            typist_rate = 0.80
        else:
            # Both supplied lists rise 10p client / 5p typist for each
            # additional pair after 5-6 / 7-8 pattern.
            extra_pairs = max(0, (effective_speakers - 5) // 2)
            client_rate = 1.60 + (0.10 * extra_pairs)
            typist_rate = 0.80 + (0.05 * extra_pairs)

        if speakers is None:
            notes.append(
                "discussion/interview inferred; speaker count unknown, "
                "using 1-to-1 baseline"
            )
        else:
            notes.append(f"discussion/interview; {speakers} speaker(s)")

        if is_medical:
            client_rate += 0.30
            typist_rate += 0.15
            notes.append("medical interview modifier")

        if is_verbatim:
            client_rate += 0.20
            typist_rate += 0.10
            notes.append("verbatim modifier")

        # The supplied list gives long-turnaround discounts generally.
        if one_week:
            client_rate -= 0.10
            notes.append("one-week client discount")
        elif three_five_day:
            client_rate -= 0.05
            notes.append("3-5 day client discount")

        # No explicit same-day/overnight/4-hour interview premium is stated.
        if same_day or overnight or four_hour or weekend:
            notes.append(
                "turnaround cue present but supplied list does not state "
                "an interview-specific turnaround premium"
            )

    else:
        # One-person dictation baseline.
        if four_hour:
            client_rate = 1.70
            typist_rate = 0.85
            notes.append("one-person 4-hour")
        elif weekend:
            client_rate = 1.60
            typist_rate = None
            notes.append(
                "one-person weekend; client rate stated, "
                "typist weekend rate not supplied"
            )
        elif same_day:
            client_rate = 1.50
            typist_rate = 0.75
            notes.append("one-person same-day")
        elif overnight:
            client_rate = 1.50
            typist_rate = 0.75
            notes.append("one-person overnight")
        else:
            client_rate = 1.30
            typist_rate = 0.65
            notes.append("one-person standard 24-hour baseline")

            if one_week:
                client_rate -= 0.10
                notes.append("one-week client discount")
            elif three_five_day:
                client_rate -= 0.05
                notes.append("3-5 day client discount")

        if is_medical:
            client_rate += 0.20
            if typist_rate is not None:
                typist_rate += 0.10
            notes.append("medical one-person modifier")

    def money(value: float | None) -> str:
        if value is None:
            return ""
        return f"{value:.2f}"

    return money(client_rate), money(typist_rate), "; ".join(notes)


def apply_rate_inference(
    jobs: list[dict],
    overrides: dict[str, dict],
) -> tuple[int, int, int]:
    """
    Fill missing Client Rate / Typist Rate.

    Existing explicit values are preserved.
    Client-specific overrides beat general inference.
    Every automatically filled rate retains an internal audit note.
    """
    client_filled = 0
    typist_filled = 0
    override_used = 0

    for job in jobs:
        existing_client_rate = str(job.get("Client Rate") or "").strip()
        existing_typist_rate = str(job.get("Typist Rate") or "").strip()

        if existing_client_rate and existing_typist_rate:
            continue

        client_key = (job.get("Client") or "").strip().lower()
        override = overrides.get(client_key)

        inferred_client = ""
        inferred_typist = ""
        explanation = ""

        if override:
            inferred_client = (override.get("Client Rate") or "").strip()
            inferred_typist = (override.get("Typist Rate") or "").strip()
            override_note = (override.get("Notes") or "").strip()
            explanation = "client-specific rate override"
            if override_note:
                explanation += f": {override_note}"
            override_used += 1
        else:
            inferred_client, inferred_typist, explanation = _general_rate_inference(job)

        filled_any = False

        if not existing_client_rate and inferred_client:
            job["Client Rate"] = inferred_client
            client_filled += 1
            filled_any = True

        if not existing_typist_rate and inferred_typist:
            job["Typist Rate"] = inferred_typist
            typist_filled += 1
            filled_any = True

        if filled_any:
            current = (job.get("Other info") or "").strip()
            rate_note = (
                f"RATE INFERRED - CHECK: "
                f"Client £{job.get('Client Rate') or '?'} / "
                f"Typist £{job.get('Typist Rate') or '?'} per min; "
                f"{explanation}"
            )
            job["Other info"] = "; ".join(
                piece for piece in [current, rate_note] if piece
            )

    return client_filled, typist_filled, override_used


def excel_formula_validation(row_num: int) -> str:
    """Show only genuinely missing data in the user-facing validation column."""
    return (
        f'=TEXTJOIN("; ",TRUE,'
        f'IF(A{row_num}="","Missing Date",""),'
        f'IF(B{row_num}="","Missing Client",""),'
        f'IF(D{row_num}="","Missing Description",""),'
        f'IF(E{row_num}="","Missing Client Rate",""),'
        f'IF(G{row_num}="","Missing Mins",""))'
    )


def excel_formula_status(row_num: int) -> str:
    return f'=IF(O{row_num}="","Ready","Needs Review")'


def build_workbook(jobs: list[dict], output: Path, year: int, month: int) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Jobs"

    month_name = datetime(year, month, 1).strftime("%B %Y")

    ws.merge_cells("A1:R1")
    ws["A1"] = f"Fingertips Jobs - {month_name}"
    ws["A1"].font = Font(size=16, bold=True, color="FFFFFF")
    ws["A1"].fill = PatternFill("solid", fgColor="1F4E78")
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 26

    ws["A2"] = "Generated from cached MDO + classified email data. Review all rows and complete any blank required fields before invoicing."
    ws.merge_cells("A2:R2")
    ws["A2"].font = Font(italic=True, color="666666")

    header_row = 4

    workbook_columns = JOB_COLUMNS + ["Source File", "Internal Info"]

    for col_index, header in enumerate(workbook_columns, start=1):
        cell = ws.cell(header_row, col_index, header)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="4472C4")
        cell.alignment = Alignment(horizontal="center", vertical="center")

    jobs = sorted(
        jobs,
        key=lambda j: (
            parse_date(j.get("Date")) or datetime(1900, 1, 1),
            (j.get("Client") or "").lower(),
            (j.get("Description") or "").lower(),
        )
    )

    first_data_row = header_row + 1

    for offset, job in enumerate(jobs):
        r = first_data_row + offset

        values = [
            job.get("Date", ""),
            job.get("Client", ""),
            job.get("Typist", ""),
            job.get("Description", ""),
            job.get("Client Rate", ""),
            job.get("Typist Rate", ""),
            job.get("Mins", ""),
            job.get("Email Subject", ""),
            job.get("Proofing", ""),
        ]

        for c, value in enumerate(values, start=1):
            ws.cell(r, c, value)

        ws.cell(r, 10, f'=IF(OR(F{r}="",G{r}=""),"",F{r}*G{r})')
        ws.cell(r, 11, f'=IF(OR(E{r}="",G{r}=""),"",E{r}*G{r})')
        ws.cell(r, 12, f'=IF(K{r}="","",K{r}*20%)')
        ws.cell(r, 13, f'=IF(K{r}="","",K{r}+L{r})')
        ws.cell(r, 14, job.get("Source ID", ""))
        ws.cell(r, 15, excel_formula_validation(r))
        ws.cell(r, 16, excel_formula_status(r))
        ws.cell(r, 17, job.get("_source_file", ""))
        ws.cell(r, 18, job.get("Other info", ""))

    last_row = max(first_data_row, first_data_row + len(jobs) - 1)

    if jobs:
        table = Table(displayName="JobsTable", ref=f"A{header_row}:R{last_row}")
        table.tableStyleInfo = TableStyleInfo(
            name="TableStyleMedium2",
            showFirstColumn=False,
            showLastColumn=False,
            showRowStripes=True,
            showColumnStripes=False,
        )
        ws.add_table(table)

    # Editable/input columns get a pale fill.
    editable_fill = PatternFill("solid", fgColor="FFF2CC")
    for col in ["B", "C", "D", "E", "F", "G", "H", "I"]:
        for r in range(first_data_row, last_row + 1):
            ws[f"{col}{r}"].fill = editable_fill

    # Formula / calculated columns.
    formula_fill = PatternFill("solid", fgColor="E2F0D9")
    for col in ["J", "K", "L", "M", "O", "P"]:
        for r in range(first_data_row, last_row + 1):
            ws[f"{col}{r}"].fill = formula_fill

    # Highlight Needs Review status.
    if jobs:
        ws.conditional_formatting.add(
            f"P{first_data_row}:P{last_row}",
            FormulaRule(
                formula=[f'$P{first_data_row}="Needs Review"'],
                fill=PatternFill("solid", fgColor="F4CCCC"),
            ),
        )

    # Number formats.
    for r in range(first_data_row, last_row + 1):
        for c in [5, 6, 10, 11, 12, 13]:
            ws.cell(r, c).number_format = '£0.00'
        ws.cell(r, 7).number_format = '0.00'

    widths = {
        "A": 13,
        "B": 24,
        "C": 22,
        "D": 42,
        "E": 13,
        "F": 13,
        "G": 10,
        "H": 48,
        "I": 14,
        "J": 14,
        "K": 14,
        "L": 12,
        "M": 14,
        "N": 46,
        "O": 34,
        "P": 16,
        "Q": 32,
        "R": 70,
    }

    for col, width in widths.items():
        ws.column_dimensions[col].width = width

    # Preserve technical evidence for automation/debugging without cluttering
    # the spreadsheet the client reviews.
    ws.column_dimensions["Q"].hidden = True
    ws.column_dimensions["R"].hidden = True

    for row in ws.iter_rows(min_row=header_row, max_row=last_row):
        for cell in row:
            cell.alignment = Alignment(
                vertical="top",
                wrap_text=True,
            )

    ws.freeze_panes = "A5"
    ws.auto_filter.ref = f"A{header_row}:R{last_row}"

    output.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output)


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--month", required=True, help="Invoice month as YYYY-MM")
    parser.add_argument("--mdo", default=str(DEFAULT_MDO))
    parser.add_argument("--classified", default=str(DEFAULT_CLASSIFIED))
    parser.add_argument("--registry", default=str(DEFAULT_REGISTRY))
    parser.add_argument("--rate-overrides", default=str(DEFAULT_RATE_OVERRIDES))
    parser.add_argument("--client-aliases", default=str(DEFAULT_CLIENT_ALIASES))
    parser.add_argument("--output", default="")
    args = parser.parse_args()

    year, month = parse_month(args.month)

    mdo_path = Path(args.mdo)
    classified_path = Path(args.classified)
    registry_path = Path(args.registry)
    rate_overrides_path = Path(args.rate_overrides)
    client_aliases_path = Path(args.client_aliases)

    mdo_rows = read_csv(mdo_path)
    classified_rows = read_csv(classified_path)
    registry_rows = read_csv(registry_path) if registry_path.exists() else []

    registry_names, _ = registry_maps(registry_rows)
    rate_overrides = load_rate_overrides(rate_overrides_path)
    client_aliases = load_client_aliases(client_aliases_path)

    mdo_jobs = build_mdo_jobs(mdo_rows, year, month)
    direct_jobs = build_direct_jobs(
        classified_rows,
        registry_names,
        year,
        month,
    )

    jobs = mdo_jobs + direct_jobs
    alias_changes = apply_client_aliases(jobs, client_aliases)
    jobs, duplicate_count = deduplicate_jobs(jobs)

    matched_returns, ambiguous_returns = enrich_from_typist_returns(
        jobs,
        classified_rows,
        registry_names,
        year,
        month,
    )

    rate_client_filled, rate_typist_filled, rate_overrides_used = apply_rate_inference(
        jobs,
        rate_overrides,
    )

    if args.output:
        output = Path(args.output)
    else:
        month_label = datetime(year, month, 1).strftime("%B_%Y")
        output = (
            DEFAULT_OUTPUT_DIR /
            f"Fingertips_Jobs_{month_label}.xlsx"
        )

    build_workbook(jobs, output, year, month)

    missing_rate = sum(1 for j in jobs if not str(j.get("Client Rate", "")).strip())
    missing_mins = sum(1 for j in jobs if not str(j.get("Mins", "")).strip())
    with_typist = sum(1 for j in jobs if str(j.get("Typist", "")).strip())

    print("=" * 80)
    print("FINGERTIPS MASTER SPREADSHEET")
    print("=" * 80)
    print(f"Month:                       {year:04d}-{month:02d}")
    print(f"MDO job rows included:       {len(mdo_jobs)}")
    print(f"Direct-email jobs included:  {len(direct_jobs)}")
    print(f"Exact duplicate rows removed:{duplicate_count:>8}")
    print(f"Final job rows:               {len(jobs)}")
    print()
    print(f"Typist returns matched:       {matched_returns}")
    print(f"Ambiguous returns skipped:    {ambiguous_returns}")
    print(f"Rows with Typist filled:      {with_typist}")
    print(f"Rows missing Mins:            {missing_mins}")
    print(f"Rows missing Client Rate:     {missing_rate}")
    print()
    print(f"Client rates auto-filled:     {rate_client_filled}")
    print(f"Typist rates auto-filled:     {rate_typist_filled}")
    print(f"Client overrides used:        {rate_overrides_used}")
    print(f"Rate overrides file:          {rate_overrides_path.resolve()}")
    print()
    print(f"Output: {output.resolve()}")
    print()
    print(
        "IMPORTANT: Review Needs Review rows before invoice generation. "
        "The generator deliberately leaves uncertain values blank rather than guessing."
    )


if __name__ == "__main__":
    main()
