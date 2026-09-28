#!/usr/bin/env python3
"""
Fingertips Automation - Email Classifier

Deterministic classifier for the cached mailbox.

It does not modify invoices or the master jobs workbook.
It reads:
    email_cache/all/index.csv

It writes:
    email_cache/all/classified_emails.csv
    email_cache/all/classification_summary.csv
    email_cache/all/sender_analysis.csv
    email_cache/all/unsure_analysis.csv
    sender_registry.csv   (persistent/editable)

Classification categories:
    NEW_JOB_MDO
    NEW_CLIENT_WORK
    TYPIST_RETURN
    TYPIST_AVAILABILITY
    WORK_FOLLOWUP
    IGNORE
    UNSURE

Design goals:
- Prefer transparent/auditable rules over opaque guesses.
- Sender role has more weight than generic keywords.
- Never silently turn an ambiguous email into a job.
- Ambiguous cases go to UNSURE for later review.
- Keep raw .eml cache immutable.
"""

from __future__ import annotations

import argparse
import csv
import re
from collections import Counter
from pathlib import Path
from typing import Iterable

CACHE_ROOT = Path("email_cache") / "all"
INDEX_PATH = CACHE_ROOT / "index.csv"
OUTPUT_PATH = CACHE_ROOT / "classified_emails.csv"
SUMMARY_PATH = CACHE_ROOT / "classification_summary.csv"
REGISTRY_PATH = Path("sender_registry.csv")
SENDER_ANALYSIS_PATH = CACHE_ROOT / "sender_analysis.csv"
UNSURE_ANALYSIS_PATH = CACHE_ROOT / "unsure_analysis.csv"

CLASSIFICATIONS = {
    "NEW_JOB_MDO",
    "NEW_CLIENT_WORK",
    "TYPIST_RETURN",
    "TYPIST_AVAILABILITY",
    "WORK_FOLLOWUP",
    "IGNORE",
    "UNSURE",
}

REGISTRY_ROLES = {"CLIENT", "TYPIST", "MDO", "IGNORE", "INTERNAL", "UNKNOWN"}

OUTPUT_COLUMNS = [
    "Date",
    "From",
    "From Email",
    "Subject",
    "Attachments",
    "Classification",
    "Confidence",
    "Reason",
    "Sender Role",
    "Matched Minutes",
    "Matched Speakers",
    "Matched Job Reference",
    "Matched Work Files",
    "Source ID",
    "Cached File",
]

REGISTRY_COLUMNS = [
    "Email Address",
    "Domain",
    "Role",
    "Name / Organisation",
    "Notes",
    "Active",
]

# ---------------------------------------------------------------------------
# Seed registry
#
# These are ONLY addresses directly evidenced by the review sample / existing
# project knowledge.  The file is editable: add senders as we learn the mailbox.
# ---------------------------------------------------------------------------

SEED_REGISTRY = [
    # MDO
    {
        "Email Address": "donotreply@mydocsonline.com",
        "Domain": "mydocsonline.com",
        "Role": "MDO",
        "Name / Organisation": "MyDocsOnline",
        "Notes": "Automated upload notifications",
        "Active": "YES",
    },

    # Known client example from sample
    {
        "Email Address": "julie.kelly@acornsurveyors.com",
        "Domain": "acornsurveyors.com",
        "Role": "CLIENT",
        "Name / Organisation": "Acorn Surveyors",
        "Notes": "New dictation sender",
        "Active": "YES",
    },
    {
        "Email Address": "colleaguerelationsteam@eastofengland.coop",
        "Domain": "eastofengland.coop",
        "Role": "CLIENT",
        "Name / Organisation": "East of England Co-op",
        "Notes": "Client / HR work-related correspondence",
        "Active": "YES",
    },

    # Typists evidenced in the sample
    {
        "Email Address": "lindajval@gmail.com",
        "Domain": "gmail.com",
        "Role": "TYPIST",
        "Name / Organisation": "Linda Valentine",
        "Notes": "",
        "Active": "YES",
    },
    {
        "Email Address": "sandrapressler67@gmail.com",
        "Domain": "gmail.com",
        "Role": "TYPIST",
        "Name / Organisation": "Sandra Pressler",
        "Notes": "",
        "Active": "YES",
    },
    {
        "Email Address": "muindijames618@gmail.com",
        "Domain": "gmail.com",
        "Role": "TYPIST",
        "Name / Organisation": "James Muindi",
        "Notes": "",
        "Active": "YES",
    },
    {
        "Email Address": "patriciaoosthuizen@yahoo.com",
        "Domain": "yahoo.com",
        "Role": "TYPIST",
        "Name / Organisation": "Patricia / Pat Grindrod",
        "Notes": "",
        "Active": "YES",
    },
    {
        "Email Address": "tina@tina-taylor.co.uk",
        "Domain": "tina-taylor.co.uk",
        "Role": "TYPIST",
        "Name / Organisation": "Tina Taylor",
        "Notes": "",
        "Active": "YES",
    },
    {
        "Email Address": "stefanghita.8776@gmail.com",
        "Domain": "gmail.com",
        "Role": "TYPIST",
        "Name / Organisation": "Stefan Ghita",
        "Notes": "Sample included typist invoice",
        "Active": "YES",
    },
    {
        "Email Address": "jean.mac.blackburn@gmail.com",
        "Domain": "gmail.com",
        "Role": "TYPIST",
        "Name / Organisation": "Jean Blackburn",
        "Notes": "Sample included typist invoice / work detail",
        "Active": "YES",
    },
    {
        "Email Address": "keithmuchoki001@gmail.com",
        "Domain": "gmail.com",
        "Role": "TYPIST",
        "Name / Organisation": "Keith Muchoki",
        "Notes": "",
        "Active": "YES",
    },
    {
        "Email Address": "taabumagdalene@gmail.com",
        "Domain": "gmail.com",
        "Role": "TYPIST",
        "Name / Organisation": "Magdalene Taabu Mweu",
        "Notes": "",
        "Active": "YES",
    },

    # Obvious non-work/system senders from the sample
    {
        "Email Address": "messages-noreply@linkedin.com",
        "Domain": "linkedin.com",
        "Role": "IGNORE",
        "Name / Organisation": "LinkedIn",
        "Notes": "",
        "Active": "YES",
    },
    {
        "Email Address": "notifications@stripe.com",
        "Domain": "stripe.com",
        "Role": "IGNORE",
        "Name / Organisation": "Stripe",
        "Notes": "Supplier/platform notification",
        "Active": "YES",
    },
    {
        "Email Address": "newsletter@email.monsoon.co.uk",
        "Domain": "email.monsoon.co.uk",
        "Role": "IGNORE",
        "Name / Organisation": "Monsoon",
        "Notes": "Newsletter",
        "Active": "YES",
    },
    {
        "Email Address": "no-reply@info.ee.co.uk",
        "Domain": "info.ee.co.uk",
        "Role": "IGNORE",
        "Name / Organisation": "EE",
        "Notes": "Account/security messages",
        "Active": "YES",
    },
    {
        "Email Address": "ebay@ebay.com",
        "Domain": "ebay.com",
        "Role": "IGNORE",
        "Name / Organisation": "eBay",
        "Notes": "",
        "Active": "YES",
    },

    # Internal mailbox
    {
        "Email Address": "bennett.fingertips@gmail.com",
        "Domain": "gmail.com",
        "Role": "INTERNAL",
        "Name / Organisation": "Fingertips",
        "Notes": "Internal/self-sent secondary account",
        "Active": "YES",
    },
    {
        "Email Address": "cbennett@fingertipstyping.co.uk",
        "Domain": "fingertipstyping.co.uk",
        "Role": "INTERNAL",
        "Name / Organisation": "Fingertips",
        "Notes": "Internal/self-sent message",
        "Active": "YES",
    },
]

AUDIO_EXTENSIONS = {
    ".wav", ".mp3", ".m4a", ".wma", ".aac", ".flac", ".dss", ".ds2"
}

RETURN_DOCUMENT_EXTENSIONS = {
    ".doc", ".docx", ".rtf", ".odt", ".txt"
}

INVOICE_EXTENSIONS = {
    ".doc", ".docx", ".pdf", ".xls", ".xlsx"
}

NEW_JOB_TERMS = [
    "new dictation",
    "new audio",
    "please transcribe",
    "for transcription",
    "to transcribe",
    "please type",
    "typing required",
    "new work",
    "dictation",
]

RETURN_TERMS = [
    "completed typing",
    "completed transcript",
    "completed transcription",
    "transcript attached",
    "please find attached the transcript",
    "attached the transcript",
    "file attached",
    "files attached",
    "herewith the following",
    "here it is",
    "completed co-op typing",
    "completed coop typing",
]

AVAILABILITY_TERMS = [
    "i am available for typing",
    "i'm available for typing",
    "im available for typing",
    "i am available for work",
    "i'm available for work",
    "im available for work",
    "available for typing",
    "available for work",
    "available to work",
    "availability for work",
    "available for more typing",
    "available for more work",
    "available should any jobs",
    "available should any work",
    "available this morning",
    "available this afternoon",
    "available this evening",
    "available today",
    "free for work",
    "free for typing",
    "still free",
    "free again",
    "do you have any typing",
    "do you have typing",
    "any typing available",
    "any work available",
    "checking in to let you know that i am available",
]

AVAILABILITY_SUBJECT_PATTERNS = [
    re.compile(r"^\s*(?:re:\s*)?(?:i\s+am\s+)?available(?:\s+for\s+(?:work|typing))?\s*[.!-]*\s*$", re.I),
    re.compile(r"^\s*(?:re:\s*)?availability(?:\s+for\s+(?:work|typing))?\s*[.!-]*\s*$", re.I),
    re.compile(r"^\s*(?:re:\s*)?available\s+(?:today|this\s+morning|this\s+afternoon|this\s+evening|to\s+work)\s*[.!-]*\s*$", re.I),
    re.compile(r"^\s*(?:re:\s*)?(?:free\s+for\s+(?:work|typing)|still\s+free|free\s+again)\s*[.!-]*\s*$", re.I),
]

FOLLOWUP_TERMS = [
    "reuploaded",
    "re-uploaded",
    "did you receive",
    "confirm receipt",
    "any update",
    "when will",
    "when is",
    "query",
    "question",
    "missing audio",
    "cannot open",
    "can't open",
    "could you resend",
    "please resend",
    "just a thought",
    "just a though",
]

INVOICE_TERMS = [
    "invoice",
    "my invoice",
    "invoice attached",
]

STRONG_MARKETING_TERMS = [
    "unsubscribe",
    "newsletter",
    "special offer",
    "sale",
    "discount",
    "paid social",
    "google ads",
    "meta ads",
    "tiktok ads",
    "seo",
    "digital marketing",
    "website redesign",
    "free audit",
    "webinar",
    "promotion",
    "shop now",
    "save ",
]

SECURITY_ACCOUNT_TERMS = [
    "one-time code",
    "one time code",
    "passcode",
    "verification code",
    "log in code",
    "login code",
    "sign-in",
    "sign in",
    "new device is using your account",
    "new sign-in",
    "password reset",
    "set your password",
    "verify your account",
]

SUPPLIER_TERMS = [
    "tax invoice",
    "renew your expired",
    "renewal",
    "payment reminder",
    "quotation attached",
    "estimated quotation",
]

# Domains that are plainly consumer/newsletter/platform traffic in this sample.
IGNORE_DOMAIN_SUFFIXES = {
    "linkedin.com",
    "winefamly.com",
    "monsoon.co.uk",
    "moonpig.com",
    "theaa.com",
    "ebay.com",
    "stripe.com",
    "secureserver.net",
    "discovery-campervans.com.au",
    "fabshophop.com",
    "meekestrade.com",
    "paypal.co.uk",
    "anoiadiari.cat",
    "email.sportsdirect.com",
    "newlook.com",
    "debenhamsoutlet.com",
    "mobiles.co.uk",
    "hobbycraft.co.uk",
    "official.pandora.net",
    "mandmdirect.com",
    "grandprixradio.nl",
    "skiddle.com",
    "morleysbexleyheath.co.uk",
    "morrisons.com",
    "vouchercodes.co.uk",
    "risingmedia.com",
    "riverisland.com",
    "santander.co.uk",
    "emails1.santander.co.uk",
    "asda.com",
    "nextemail.co.uk",
    "runthrough.co.uk",
    "vercel.com",
    "evri.com",
    "jkattire.co.uk",
    "mainlinemenswear.co.uk",
    "messaging.nowtv.com",
    "databricks.com",
    "uk-safetymanagement.co.uk",
    "leisurecentre.com",
    "marketing.workbooks.com",
    "argos.co.uk",
    "lta-tennis.com",
    "soldierfield.com",
    "aosom.co.uk",
}

# Work filename / subject patterns
MINUTES_PATTERNS = [
    re.compile(r"\b(\d{1,4})\s*(?:mins?|minutes?)\b", re.I),
    re.compile(r"\[\s*m\s*(\d{1,4})\s*\]", re.I),
]

SPEAKERS_PATTERNS = [
    re.compile(r"\b(\d{1,2})\s*speakers?\b", re.I),
    re.compile(r"\[\s*s\s*(\d{1,2})\s*\]", re.I),
]

JOB_REFERENCE_PATTERNS = [
    re.compile(r"\bA\d{5,}\b", re.I),      # e.g. A024481
    re.compile(r"\bDS\d{5,}\b", re.I),     # e.g. DS900334
    re.compile(r"\bKH-\d+\b", re.I),       # e.g. KH-10806
]


def normalise(value: str | None) -> str:
    return (value or "").strip()


def lower(value: str | None) -> str:
    return normalise(value).lower()


def split_attachments(raw: str | None) -> list[str]:
    text = normalise(raw)
    if not text:
        return []
    return [piece.strip() for piece in text.split("|") if piece.strip()]


def file_extension(filename: str) -> str:
    name = filename.lower().strip()
    dot = name.rfind(".")
    return name[dot:] if dot >= 0 else ""


def contains_any(text: str, terms: Iterable[str]) -> bool:
    text = text.lower()
    return any(term in text for term in terms)



def looks_like_typist_availability(subject: str, body: str) -> bool:
    """
    Strict availability detector.

    The previous version treated broad 'available' wording as evidence and
    could misclassify marketing emails.  v19 only accepts:
    - a concise availability-style subject, or
    - an explicit typing/work availability phrase in subject/body.
    """
    subject = normalise(subject)
    body = normalise(body)
    joined = f"{subject}\n{body}".lower()

    if any(pattern.match(subject) for pattern in AVAILABILITY_SUBJECT_PATTERNS):
        return True

    return contains_any(joined, AVAILABILITY_TERMS)


def create_registry_if_missing() -> None:
    if REGISTRY_PATH.exists():
        return

    CACHE_ROOT.mkdir(parents=True, exist_ok=True)
    with REGISTRY_PATH.open("w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=REGISTRY_COLUMNS)
        writer.writeheader()
        writer.writerows(SEED_REGISTRY)

    print(f"Created editable sender registry: {REGISTRY_PATH}")


def load_registry() -> tuple[dict[str, dict], dict[str, list[dict]]]:
    create_registry_if_missing()

    by_email: dict[str, dict] = {}
    by_domain: dict[str, list[dict]] = {}

    with REGISTRY_PATH.open("r", encoding="utf-8-sig", newline="") as f:
        for row in csv.DictReader(f):
            if lower(row.get("Active")) not in {"yes", "y", "true", "1"}:
                continue

            role = normalise(row.get("Role")).upper()
            if role not in REGISTRY_ROLES:
                continue

            address = lower(row.get("Email Address"))
            domain = lower(row.get("Domain"))

            if address:
                by_email[address] = row

            if domain:
                by_domain.setdefault(domain, []).append(row)

    return by_email, by_domain


def resolve_sender_role(
    from_email: str,
    from_domain: str,
    by_email: dict[str, dict],
    by_domain: dict[str, list[dict]],
) -> tuple[str, str]:
    """
    Exact email match always wins.

    Domain role is used only when every active registry row for the domain
    agrees on the same role. This prevents gmail.com from accidentally
    becoming a "typist domain".
    """
    email_key = lower(from_email)
    domain_key = lower(from_domain)

    if email_key in by_email:
        row = by_email[email_key]
        return normalise(row.get("Role")).upper(), "exact sender registry"

    matches = by_domain.get(domain_key, [])
    roles = {normalise(row.get("Role")).upper() for row in matches}

    if len(roles) == 1 and domain_key not in {"gmail.com", "yahoo.com", "outlook.com", "hotmail.com"}:
        return next(iter(roles)), "domain registry"

    return "UNKNOWN", ""


def extract_numbers(text: str, patterns: list[re.Pattern]) -> list[str]:
    found: list[str] = []
    for pattern in patterns:
        for match in pattern.finditer(text):
            value = match.group(1)
            if value not in found:
                found.append(value)
    return found


def extract_job_references(text: str) -> list[str]:
    refs: list[str] = []
    for pattern in JOB_REFERENCE_PATTERNS:
        for match in pattern.finditer(text):
            value = match.group(0).upper()
            if value not in refs:
                refs.append(value)
    return refs


def work_files(attachments: list[str]) -> list[str]:
    return [
        name
        for name in attachments
        if file_extension(name) in AUDIO_EXTENSIONS | RETURN_DOCUMENT_EXTENSIONS
    ]


def has_audio_attachment(attachments: list[str]) -> bool:
    return any(file_extension(name) in AUDIO_EXTENSIONS for name in attachments)


def has_return_document(attachments: list[str]) -> bool:
    return any(file_extension(name) in RETURN_DOCUMENT_EXTENSIONS for name in attachments)


def has_invoice_attachment(attachments: list[str]) -> bool:
    return any(file_extension(name) in INVOICE_EXTENSIONS for name in attachments)


def domain_is_obvious_ignore(domain: str) -> bool:
    domain = lower(domain)
    return any(
        domain == suffix or domain.endswith("." + suffix)
        for suffix in IGNORE_DOMAIN_SUFFIXES
    )



def build_sender_profiles(rows: list[dict]) -> dict[str, dict]:
    """
    Learn conservative sender behaviour from the cached mailbox itself.

    This is intentionally NOT machine learning.  It is aggregate evidence:
    repeated availability messages and repeated returned-document behaviour
    are strong evidence that a sender is a typist.

    We do not automatically infer clients from generic work-like language,
    because the cost of falsely creating a job is higher.
    """
    profiles: dict[str, dict] = {}

    for row in rows:
        sender = lower(row.get("From Email"))
        if not sender:
            continue

        p = profiles.setdefault(
            sender,
            {
                "Email Count": 0,
                "Availability Count": 0,
                "Return Evidence Count": 0,
                "Returned Document Count": 0,
                "Audio Attachment Count": 0,
                "New Job Evidence Count": 0,
                "Invoice Count": 0,
                "Inferred Role": "",
                "Inference Reason": "",
            },
        )

        subject = normalise(row.get("Subject"))
        body = normalise(row.get("Body Preview"))
        joined = f"{subject}\n{body}".lower()
        attachments = split_attachments(row.get("Attachments"))

        p["Email Count"] += 1

        if looks_like_typist_availability(subject, body):
            p["Availability Count"] += 1

        returned_doc = has_return_document(attachments)
        return_language = contains_any(joined, RETURN_TERMS)
        minutes_found = bool(extract_numbers(joined, MINUTES_PATTERNS))

        if returned_doc:
            p["Returned Document Count"] += 1

        if returned_doc and (return_language or minutes_found):
            p["Return Evidence Count"] += 1

        if has_audio_attachment(attachments):
            p["Audio Attachment Count"] += 1

        if has_audio_attachment(attachments) and contains_any(joined, NEW_JOB_TERMS):
            p["New Job Evidence Count"] += 1

        if contains_any(joined, INVOICE_TERMS) and has_invoice_attachment(attachments):
            p["Invoice Count"] += 1

    for sender, p in profiles.items():
        availability = p["Availability Count"]
        returns = p["Return Evidence Count"]
        returned_docs = p["Returned Document Count"]

        # High-confidence typist inference:
        # - repeated availability messages, OR
        # - repeated convincing returns, OR
        # - both availability and return evidence.
        if availability >= 3:
            p["Inferred Role"] = "TYPIST"
            p["Inference Reason"] = (
                f"mailbox behaviour: {availability} availability messages"
            )
        elif returns >= 2:
            p["Inferred Role"] = "TYPIST"
            p["Inference Reason"] = (
                f"mailbox behaviour: {returns} returned-document messages"
            )
        elif availability >= 1 and (returns >= 1 or returned_docs >= 2):
            p["Inferred Role"] = "TYPIST"
            p["Inference Reason"] = (
                "mailbox behaviour: availability + returned-document evidence"
            )

    return profiles


def write_sender_analysis(
    profiles: dict[str, dict],
    by_email: dict[str, dict],
) -> None:
    columns = [
        "From Email",
        "Email Count",
        "Availability Count",
        "Return Evidence Count",
        "Returned Document Count",
        "Audio Attachment Count",
        "New Job Evidence Count",
        "Invoice Count",
        "Registry Role",
        "Inferred Role",
        "Inference Reason",
    ]

    rows = []
    for sender, p in profiles.items():
        registry_role = ""
        if sender in by_email:
            registry_role = normalise(by_email[sender].get("Role")).upper()

        rows.append(
            {
                "From Email": sender,
                "Email Count": p["Email Count"],
                "Availability Count": p["Availability Count"],
                "Return Evidence Count": p["Return Evidence Count"],
                "Returned Document Count": p["Returned Document Count"],
                "Audio Attachment Count": p["Audio Attachment Count"],
                "New Job Evidence Count": p["New Job Evidence Count"],
                "Invoice Count": p["Invoice Count"],
                "Registry Role": registry_role,
                "Inferred Role": p["Inferred Role"],
                "Inference Reason": p["Inference Reason"],
            }
        )

    rows.sort(
        key=lambda r: (
            r["Registry Role"] != "",
            r["Inferred Role"] != "TYPIST",
            -int(r["Email Count"]),
            r["From Email"],
        )
    )

    with SENDER_ANALYSIS_PATH.open("w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=columns)
        writer.writeheader()
        writer.writerows(rows)


def classify(row: dict, by_email: dict[str, dict], by_domain: dict[str, list[dict]], sender_profiles: dict[str, dict]) -> dict:
    sender = lower(row.get("From Email"))
    domain = lower(row.get("From Domain"))
    subject = normalise(row.get("Subject"))
    body = normalise(row.get("Body Preview"))
    attachments = split_attachments(row.get("Attachments"))

    joined = f"{subject}\n{body}"
    joined_lower = joined.lower()

    role, role_reason = resolve_sender_role(sender, domain, by_email, by_domain)

    if role == "UNKNOWN":
        profile = sender_profiles.get(sender, {})
        inferred_role = profile.get("Inferred Role", "")
        if inferred_role:
            role = inferred_role
            role_reason = profile.get("Inference Reason", "mailbox behaviour")

    audio_attached = has_audio_attachment(attachments)
    return_doc_attached = has_return_document(attachments)
    invoice_attached = has_invoice_attachment(attachments)

    availability = looks_like_typist_availability(subject, body)
    return_language = contains_any(joined_lower, RETURN_TERMS)
    new_job_language = contains_any(joined_lower, NEW_JOB_TERMS)
    followup_language = contains_any(joined_lower, FOLLOWUP_TERMS)
    invoice_language = contains_any(joined_lower, INVOICE_TERMS)
    marketing_language = contains_any(joined_lower, STRONG_MARKETING_TERMS)
    account_language = contains_any(joined_lower, SECURITY_ACCOUNT_TERMS)
    supplier_language = contains_any(joined_lower, SUPPLIER_TERMS)

    classification = "UNSURE"
    confidence = 0.50
    reasons: list[str] = []

    # ------------------------------------------------------------------
    # 1. Exact MyDocsOnline new upload rule
    # ------------------------------------------------------------------
    if (
        role == "MDO"
        and subject.lower().startswith("upload to \\")
        and "upload complete to \\" in body.lower()
    ):
        classification = "NEW_JOB_MDO"
        confidence = 0.99
        reasons.append("known MDO sender + direct 'Upload to \\' notification")

    # ------------------------------------------------------------------
    # 2. Explicit ignore registry
    # ------------------------------------------------------------------
    elif sender == "support@mydocsonline.com":
        classification = "WORK_FOLLOWUP"
        confidence = 0.96
        reasons.append("MyDocsOnline support correspondence; operational follow-up, not a new job")

    elif role == "IGNORE":
        classification = "IGNORE"
        confidence = 0.99
        reasons.append(role_reason or "known ignore sender")

    # ------------------------------------------------------------------
    # 3. Typist rules
    # ------------------------------------------------------------------
    elif role == "TYPIST":
        # Typist invoices are operational/accounting mail, not job returns.
        if invoice_language and invoice_attached and not return_language:
            classification = "IGNORE"
            confidence = 0.96
            reasons.append(f"typist ({role_reason}) + invoice language/attachment")

        elif return_doc_attached and (return_language or extract_numbers(joined, MINUTES_PATTERNS)):
            classification = "TYPIST_RETURN"
            confidence = 0.98
            reasons.append(f"typist ({role_reason}) + returned document + completion/minutes cue")

        elif return_doc_attached and not availability:
            classification = "TYPIST_RETURN"
            confidence = 0.91
            reasons.append(f"typist ({role_reason}) + returned document attachment")

        elif availability and not return_doc_attached:
            classification = "TYPIST_AVAILABILITY"
            confidence = 0.98
            reasons.append(f"typist ({role_reason}) + availability language")

        elif availability and return_doc_attached:
            # Example: "transcript attached, available for more"
            classification = "TYPIST_RETURN"
            confidence = 0.97
            reasons.append(f"typist ({role_reason}) + returned document; availability is secondary")

        else:
            classification = "WORK_FOLLOWUP"
            confidence = 0.82
            reasons.append(f"typist ({role_reason}) but no clear new return/availability pattern")

    # ------------------------------------------------------------------
    # 4. Known client rules
    # ------------------------------------------------------------------
    elif role == "CLIENT":
        if audio_attached:
            classification = "NEW_CLIENT_WORK"
            confidence = 0.98
            reasons.append("known client + audio attachment")

        elif new_job_language and attachments:
            classification = "NEW_CLIENT_WORK"
            confidence = 0.93
            reasons.append("known client + new-work language + attachment")

        elif followup_language or subject.lower().startswith(("re:", "fw:", "fwd:")):
            classification = "WORK_FOLLOWUP"
            confidence = 0.92
            reasons.append("known client + follow-up/reply pattern")

        else:
            classification = "WORK_FOLLOWUP"
            confidence = 0.80
            reasons.append("known client, but no strong new-job evidence")

    # ------------------------------------------------------------------
    # 5. Internal messages: don't create jobs without explicit evidence
    # ------------------------------------------------------------------
    elif role == "INTERNAL":
        classification = "IGNORE"
        confidence = 0.93
        reasons.append("internal/self-sent message")

    # ------------------------------------------------------------------
    # 6. Unknown sender, strong new-work evidence
    # ------------------------------------------------------------------
    elif audio_attached and new_job_language:
        classification = "NEW_CLIENT_WORK"
        confidence = 0.84
        reasons.append("unknown sender + audio attachment + strong new-work language")

    elif audio_attached and not marketing_language and not supplier_language:
        classification = "UNSURE"
        confidence = 0.62
        reasons.append("audio attachment from unknown sender; requires sender review")

    # ------------------------------------------------------------------
    # 7. Unknown sender, likely returned typing
    # ------------------------------------------------------------------
    elif return_doc_attached and return_language:
        classification = "UNSURE"
        confidence = 0.70
        reasons.append("document + return language, but sender is not yet known as typist")

    # ------------------------------------------------------------------
    # 8. Strong ignore patterns for unknown senders
    # ------------------------------------------------------------------
    elif domain_is_obvious_ignore(domain):
        classification = "IGNORE"
        confidence = 0.96
        reasons.append("obvious non-work platform/newsletter domain")

    elif account_language:
        classification = "IGNORE"
        confidence = 0.96
        reasons.append("account/security notification")

    elif marketing_language and not attachments:
        classification = "IGNORE"
        confidence = 0.93
        reasons.append("strong marketing/newsletter language")

    elif supplier_language and not audio_attached:
        classification = "IGNORE"
        confidence = 0.86
        reasons.append("supplier/accounting/quotation pattern")

    # ------------------------------------------------------------------
    # 9. Unknown availability messages are not safe to classify as typists
    # ------------------------------------------------------------------
    elif availability:
        classification = "UNSURE"
        confidence = 0.66
        reasons.append("availability language from unknown sender")

    # ------------------------------------------------------------------
    # 10. Replies / work-like language with insufficient evidence
    # ------------------------------------------------------------------
    elif followup_language or subject.lower().startswith(("re:", "fw:", "fwd:")):
        classification = "UNSURE"
        confidence = 0.58
        reasons.append("possible follow-up, sender role unknown")

    elif new_job_language:
        classification = "UNSURE"
        confidence = 0.58
        reasons.append("work-like language, but insufficient sender/attachment evidence")

    else:
        classification = "UNSURE"
        confidence = 0.50
        reasons.append("no reliable rule matched")

    minutes = extract_numbers(joined, MINUTES_PATTERNS)
    speakers = extract_numbers(joined, SPEAKERS_PATTERNS)
    refs = extract_job_references(joined + "\n" + " ".join(attachments))
    matched_files = work_files(attachments)

    return {
        "Date": row.get("Parsed Date", row.get("Date", "")),
        "From": row.get("From", ""),
        "From Email": row.get("From Email", ""),
        "Subject": row.get("Subject", ""),
        "Attachments": row.get("Attachments", ""),
        "Classification": classification,
        "Confidence": f"{confidence:.2f}",
        "Reason": "; ".join(reasons),
        "Sender Role": role,
        "Matched Minutes": " | ".join(minutes),
        "Matched Speakers": " | ".join(speakers),
        "Matched Job Reference": " | ".join(refs),
        "Matched Work Files": " | ".join(matched_files),
        "Source ID": row.get("Source ID", ""),
        "Cached File": row.get("Cached File", ""),
    }


def read_index() -> list[dict]:
    if not INDEX_PATH.exists():
        raise SystemExit(
            f"Cannot find {INDEX_PATH}\n"
            "Run email_cache.py first."
        )

    with INDEX_PATH.open("r", encoding="utf-8-sig", newline="") as f:
        return list(csv.DictReader(f))


def write_classified(rows: list[dict]) -> None:
    CACHE_ROOT.mkdir(parents=True, exist_ok=True)

    with OUTPUT_PATH.open("w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=OUTPUT_COLUMNS)
        writer.writeheader()
        writer.writerows(rows)


def write_summary(rows: list[dict]) -> Counter:
    counts = Counter(row["Classification"] for row in rows)

    with SUMMARY_PATH.open("w", encoding="utf-8-sig", newline="") as f:
        writer = csv.writer(f)
        writer.writerow(["Classification", "Count"])
        for classification in [
            "NEW_JOB_MDO",
            "NEW_CLIENT_WORK",
            "TYPIST_RETURN",
            "TYPIST_AVAILABILITY",
            "WORK_FOLLOWUP",
            "IGNORE",
            "UNSURE",
        ]:
            writer.writerow([classification, counts.get(classification, 0)])

    return counts



def _common_subject_shape(subject: str) -> str:
    """
    Normalise subjects enough to group repeated conversation patterns without
    destroying the useful wording.
    """
    subject = normalise(subject)
    subject = re.sub(r"^(?:re|fw|fwd):\s*", "", subject, flags=re.I)
    subject = re.sub(r"\s+", " ", subject).strip()
    return subject[:120]


def write_unsure_analysis(rows: list[dict]) -> None:
    """
    Group remaining UNSURE emails by sender so review is driven by highest
    impact senders, not by hundreds of individual messages.
    """
    groups: dict[str, dict] = {}

    for row in rows:
        if row.get("Classification") != "UNSURE":
            continue

        sender = lower(row.get("From Email")) or lower(row.get("From")) or "(unknown sender)"
        g = groups.setdefault(
            sender,
            {
                "From Email": sender,
                "Unsure Count": 0,
                "Subjects": Counter(),
                "Reasons": Counter(),
                "Attachment Examples": [],
                "Source ID Examples": [],
            },
        )

        g["Unsure Count"] += 1
        subject_shape = _common_subject_shape(row.get("Subject", ""))
        if subject_shape:
            g["Subjects"][subject_shape] += 1

        reason = normalise(row.get("Reason"))
        if reason:
            g["Reasons"][reason] += 1

        attachments = normalise(row.get("Attachments"))
        if attachments and attachments not in g["Attachment Examples"] and len(g["Attachment Examples"]) < 3:
            g["Attachment Examples"].append(attachments)

        source_id = normalise(row.get("Source ID"))
        if source_id and len(g["Source ID Examples"]) < 3:
            g["Source ID Examples"].append(source_id)

    output_rows = []

    for sender, g in groups.items():
        common_subjects = [
            f"{subject} ({count})"
            for subject, count in g["Subjects"].most_common(5)
        ]
        common_reasons = [
            f"{reason} ({count})"
            for reason, count in g["Reasons"].most_common(3)
        ]

        output_rows.append(
            {
                "From Email": sender,
                "Unsure Count": g["Unsure Count"],
                "Common Subjects": " | ".join(common_subjects),
                "Common Reasons": " | ".join(common_reasons),
                "Attachment Examples": " | ".join(g["Attachment Examples"]),
                "Source ID Examples": " | ".join(g["Source ID Examples"]),
            }
        )

    output_rows.sort(key=lambda r: (-int(r["Unsure Count"]), r["From Email"]))

    columns = [
        "From Email",
        "Unsure Count",
        "Common Subjects",
        "Common Reasons",
        "Attachment Examples",
        "Source ID Examples",
    ]

    with UNSURE_ANALYSIS_PATH.open("w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=columns)
        writer.writeheader()
        writer.writerows(output_rows)


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--input",
        default=str(INDEX_PATH),
        help="Reserved for future use; the classifier expects the standard cache index.",
    )
    args = parser.parse_args()

    rows = read_index()
    by_email, by_domain = load_registry()
    sender_profiles = build_sender_profiles(rows)
    write_sender_analysis(sender_profiles, by_email)

    classified = [
        classify(row, by_email, by_domain, sender_profiles)
        for row in rows
    ]

    write_classified(classified)
    counts = write_summary(classified)
    write_unsure_analysis(classified)

    print("=" * 80)
    print("FINGERTIPS EMAIL CLASSIFIER")
    print("=" * 80)
    print(f"Emails classified: {len(classified)}")
    print()

    for classification in [
        "NEW_JOB_MDO",
        "NEW_CLIENT_WORK",
        "TYPIST_RETURN",
        "TYPIST_AVAILABILITY",
        "WORK_FOLLOWUP",
        "IGNORE",
        "UNSURE",
    ]:
        print(f"{classification:<24} {counts.get(classification, 0):>5}")

    print()
    print(f"Classifier output: {OUTPUT_PATH.resolve()}")
    print(f"Summary:           {SUMMARY_PATH.resolve()}")
    print(f"Sender registry:   {REGISTRY_PATH.resolve()}")
    print(f"Sender analysis:   {SENDER_ANALYSIS_PATH.resolve()}")
    print(f"UNSURE analysis:   {UNSURE_ANALYSIS_PATH.resolve()}")
    print()
    print(
        "Important: The classifier does not alter your jobs spreadsheet. "
        "Review the classifications first, especially UNSURE and NEW_CLIENT_WORK."
    )


if __name__ == "__main__":
    main()
