import os
import re
import json
import time
import ssl
import smtplib
import imaplib
import hashlib
import logging
import html
import requests
import email

from datetime import datetime, timedelta
from zoneinfo import ZoneInfo
from email.message import EmailMessage
from email.header import decode_header
from email.utils import formataddr, make_msgid, parseaddr



# =========================================================
# CONFIGURATION
# =========================================================

GOOGLE_SHEET_ID = os.getenv("GOOGLE_SHEET_ID", "").strip()
GOOGLE_SERVICE_ACCOUNT_JSON = os.getenv(
    "GOOGLE_SERVICE_ACCOUNT_JSON", ""
).strip()
MAILBOXES_JSON = os.getenv("MAILBOXES_JSON", "").strip()
# Signature secrets are raw HTML. Do not strip or rewrite their contents.
HAMZA_SIGNATURE = os.getenv("HAMZA_SIGNATURE", "")
ARSLAN_SIGNATURE = os.getenv("ARSLAN_SIGNATURE", "")
JAMES_SIGNATURE = os.getenv("JAMES_SIGNATURE", "")

GEMINI_API_KEY = os.getenv("GEMINI_API_KEY", "").strip()
OPENROUTER_API_KEY = os.getenv("OPENROUTER_API_KEY", "").strip()
OLLAMA_API_KEY = os.getenv("OLLAMA_API_KEY", "").strip()

GEMINI_MODEL = os.getenv("GEMINI_MODEL", "auto").strip()
OPENROUTER_MODEL = os.getenv(
    "OPENROUTER_MODEL", "openrouter/auto"
).strip()
OLLAMA_MODEL = os.getenv("OLLAMA_MODEL", "auto").strip()

GEMINI_URL = (
    "https://generativelanguage.googleapis.com/v1beta/models/"
)
GEMINI_MODELS_URL = (
    "https://generativelanguage.googleapis.com/v1beta/models"
)
OPENROUTER_URL = (
    "https://openrouter.ai/api/v1/chat/completions"
)
OLLAMA_URL = "https://ollama.com/api/chat"
OLLAMA_TAGS_URL = "https://ollama.com/api/tags"

AI_CONNECT_TIMEOUT = 5
AI_READ_TIMEOUT = 20
AI_RETRY_DELAYS = (2,)

EASTERN_TZ = ZoneInfo("America/New_York")

DAILY_LIMIT_PER_MAILBOX = 10

# Sending allowed from SEND_START_HOUR up to (not including) SEND_END_HOUR,
# Eastern time, except during the pause window.
# The time window is only enforced when env ENFORCE_SEND_WINDOW=1 is set.
SEND_START_HOUR = 8
SEND_END_HOUR = 16
PAUSE_START_HOUR = 13
PAUSE_END_HOUR = 14

# Keep the legacy BEST_PRICE reply template out of regular cold outreach.
LEGACY_BEST_PRICE_TEMPLATE_ID = "BEST_PRICE"

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s"
)

logger = logging.getLogger(__name__)
_ACTIVE_RUN_LOG_WS = None

_GEMINI_DISCOVERED_MODEL = None
_OLLAMA_DISCOVERED = None


# =========================================================
# GOOGLE SHEETS
# =========================================================

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets"
]

HEADERS = {
    "Websites": [
        "Website", "Email", "Status", "Task ID",
        "Created At", "Last Sent At", "Mailbox",
        "Reply Status", "Notes"
    ],
    "Templates": [
        "Template ID", "Subject", "Body", "Active"
    ],
    "History": [
        "Timestamp", "Date", "Email", "Website",
        "Mailbox", "Status", "Message ID", "Details",
        "Template"
    ],
    "Replies": [
        "Received At", "From", "Subject",
        "Message ID", "Classification", "Snippet"
    ],
    "Dashboard": [
        "Metric", "Value"
    ],
    "Suppression": [
        "Email", "Reason", "Added At"
    ],
    "Run Log": [
        "Timestamp", "Level", "Event", "Details"
    ],
    "Archive": [
        "Timestamp", "Date", "Email", "Website",
        "Mailbox", "Status", "Message ID", "Details",
        "Template"
    ],
    "Run Log Archive": [
        "Timestamp", "Level", "Event", "Details"
    ]
}


def column_letter(number):
    letters = ""
    while number > 0:
        number, remainder = divmod(number - 1, 26)
        letters = chr(65 + remainder) + letters
    return letters


def sheets_retry(func, *args, **kwargs):
    delays = (2, 5, 10)
    last_exc = None
    for attempt in range(len(delays) + 1):
        try:
            return func(*args, **kwargs)
        except Exception as exc:
            last_exc = exc
            if attempt < len(delays):
                time.sleep(delays[attempt])
    raise last_exc


def connect_google_sheets():
    import gspread
    from google.oauth2.service_account import Credentials

    service_info = json.loads(GOOGLE_SERVICE_ACCOUNT_JSON)
    credentials = Credentials.from_service_account_info(
        service_info,
        scopes=SCOPES
    )
    client = gspread.authorize(credentials)
    return client.open_by_key(GOOGLE_SHEET_ID)


def setup_tabs(spreadsheet):
    existing = {
        ws.title: ws
        for ws in spreadsheet.worksheets()
    }

    for name, headers in HEADERS.items():
        end_column = column_letter(len(headers))

        if name not in existing:
            ws = spreadsheet.add_worksheet(
                title=name,
                rows=1000,
                cols=max(len(headers), 10)
            )
            sheets_retry(
                ws.update,
                range_name=f"A1:{end_column}1",
                values=[headers]
            )
            continue

        ws = existing[name]
        current = ws.row_values(1)

        if not current:
            sheets_retry(
                ws.update,
                range_name=f"A1:{end_column}1",
                values=[headers]
            )
            continue

        if (
            name in ("History", "Archive")
            and current[:8] == headers[:8]
            and len(current) == 8
        ):
            sheets_retry(ws.update_cell, 1, 9, "Template")
            continue

        if len(current) < len(headers) and current == headers[:len(current)]:
            added = headers[len(current):]
            sheets_retry(
                ws.update,
                range_name=f"{column_letter(len(current) + 1)}1:{end_column}1",
                values=[added]
            )
            continue

        if current[:len(headers)] != headers:
            raise RuntimeError(
                f"Header mismatch in worksheet: {name}. "
                f"Expected: {headers}. Found: {current}"
            )


def get_worksheet(spreadsheet, name):
    return spreadsheet.worksheet(name)


def get_rows(worksheet):
    return sheets_retry(worksheet.get_all_records)


def timestamp():
    return datetime.now(EASTERN_TZ).strftime(
        "%Y-%m-%d %H:%M:%S %z"
    )


def now_local():
    return datetime.now(EASTERN_TZ)


def write_run_log(ws, level, event, details=""):
    try:
        sheets_retry(
            ws.append_row,
            [
                timestamp(),
                level,
                event,
                str(details)[:4000]
            ]
        )
    except Exception:
        logger.exception("Failed to write Run Log.")


def pending_task_count(rows):
    return sum(
        1 for row in rows
        if str(row.get("Status", "")).strip().upper() == "PENDING"
    )


def within_send_window():
    if os.getenv("ENFORCE_SEND_WINDOW", "").strip().lower() not in ("1", "true", "yes"):
        return True
    hour = now_local().hour
    if hour < SEND_START_HOUR or hour >= SEND_END_HOUR:
        return False
    if PAUSE_START_HOUR <= hour < PAUSE_END_HOUR:
        return False
    return True


# =========================================================
# GENERAL HELPERS & AI
# =========================================================

def valid_email(address):
    if not address:
        return False
    pattern = (
        r"^[A-Za-z0-9._%+-]+@"
        r"[A-Za-z0-9.-]+\.[A-Za-z]{2,}$"
    )
    return bool(re.fullmatch(pattern, address.strip()))


def normalize_email(address):
    return (address or "").strip().lower()


def decode_text(value):
    if not value:
        return ""
    result = []
    for part, encoding in decode_header(value):
        if isinstance(part, bytes):
            try:
                result.append(
                    part.decode(
                        encoding or "utf-8",
                        errors="replace"
                    )
                )
            except LookupError:
                result.append(part.decode("utf-8", errors="replace"))
        else:
            result.append(part)
    return "".join(result)


def fallback_clean_website_name(website):
    if not website:
        return ""
    name = str(website).strip()
    name = re.sub(r"^https?://", "", name, flags=re.IGNORECASE)
    name = re.sub(r"^www\.", "", name, flags=re.IGNORECASE)
    name = re.split(r"[/\?#]", name, maxsplit=1)[0]
    suffixes = (
        ".co.uk", ".org.uk", ".com.au", ".net.au", ".co.nz",
        ".com.pk", ".co.in", ".com", ".net", ".org", ".io",
        ".biz", ".info", ".edu", ".gov", ".pk", ".uk", ".us",
        ".ca", ".au", ".de", ".fr", ".in"
    )
    lower_name = name.lower()
    for suffix in suffixes:
        if lower_name.endswith(suffix):
            name = name[:-len(suffix)]
            break
    if "." in name:
        name = name.split(".")[-1]
    name = re.sub(r"[-_.]+", " ", name)
    name = re.sub(r"([a-z])([A-Z])", r"\1 \2", name)
    name = re.sub(r"([A-Z])([A-Z][a-z])", r"\1 \2", name)
    name = re.sub(r"\s+", " ", name).strip()
    return name.title()


def _resolve_gemini_model():
    global _GEMINI_DISCOVERED_MODEL
    if GEMINI_MODEL and GEMINI_MODEL.lower() != "auto":
        return GEMINI_MODEL
    if _GEMINI_DISCOVERED_MODEL:
        return _GEMINI_DISCOVERED_MODEL
    try:
        resp = requests.get(
            GEMINI_MODELS_URL,
            params={"key": GEMINI_API_KEY},
            timeout=(AI_CONNECT_TIMEOUT, AI_READ_TIMEOUT)
        )
        resp.raise_for_status()
        models = resp.json().get("models", [])
        candidates = [
            m.get("name", "").replace("models/", "")
            for m in models
            if "generateContent" in m.get("supportedGenerationMethods", [])
        ]
        for preferred in ("gemini-2.5-flash", "gemini-2.0-flash", "gemini-1.5-flash"):
            for c in candidates:
                if preferred in c:
                    _GEMINI_DISCOVERED_MODEL = c
                    return _GEMINI_DISCOVERED_MODEL
        if candidates:
            _GEMINI_DISCOVERED_MODEL = candidates[0]
            return _GEMINI_DISCOVERED_MODEL
    except Exception:
        pass
    _GEMINI_DISCOVERED_MODEL = "gemini-2.0-flash"
    return _GEMINI_DISCOVERED_MODEL


def _resolve_ollama_model():
    global _OLLAMA_DISCOVERED
    if OLLAMA_MODEL and OLLAMA_MODEL.lower() != "auto":
        return OLLAMA_MODEL
    if _OLLAMA_DISCOVERED:
        return _OLLAMA_DISCOVERED
    try:
        headers = {"Authorization": f"Bearer {OLLAMA_API_KEY}"} if OLLAMA_API_KEY else {}
        resp = requests.get(
            OLLAMA_TAGS_URL,
            headers=headers,
            timeout=(AI_CONNECT_TIMEOUT, AI_READ_TIMEOUT)
        )
        resp.raise_for_status()
        models = [m.get("name", "") for m in resp.json().get("models", []) if m.get("name")]
        if models:
            _OLLAMA_DISCOVERED = models[0]
            return _OLLAMA_DISCOVERED
    except Exception:
        pass
    _OLLAMA_DISCOVERED = "llama3.2"
    return _OLLAMA_DISCOVERED


def _call_gemini_json(prompt):
    if not GEMINI_API_KEY:
        raise RuntimeError("GEMINI_API_KEY not set")
    model = _resolve_gemini_model()
    url = f"{GEMINI_URL}{model}:generateContent"
    payload = {
        "contents": [{"parts": [{"text": prompt}]}],
        "generationConfig": {
            "temperature": 0.1,
            "responseMimeType": "application/json"
        }
    }
    resp = requests.post(
        url,
        params={"key": GEMINI_API_KEY},
        json=payload,
        timeout=(AI_CONNECT_TIMEOUT, AI_READ_TIMEOUT)
    )
    resp.raise_for_status()
    data = resp.json()
    text = data["candidates"][0]["content"]["parts"][0]["text"]
    return json.loads(text)


def _call_openrouter_json(prompt):
    if not OPENROUTER_API_KEY:
        raise RuntimeError("OPENROUTER_API_KEY not set")
    headers = {
        "Authorization": f"Bearer {OPENROUTER_API_KEY}",
        "Content-Type": "application/json"
    }
    payload = {
        "model": OPENROUTER_MODEL or "openrouter/auto",
        "messages": [{"role": "user", "content": prompt}],
        "temperature": 0.1,
        "response_format": {"type": "json_object"}
    }
    resp = requests.post(
        OPENROUTER_URL,
        headers=headers,
        json=payload,
        timeout=(AI_CONNECT_TIMEOUT, AI_READ_TIMEOUT)
    )
    resp.raise_for_status()
    content = resp.json()["choices"][0]["message"]["content"]
    return json.loads(content)


def _call_ollama_json(prompt):
    if not OLLAMA_API_KEY:
        raise RuntimeError("OLLAMA_API_KEY not set")
    model = _resolve_ollama_model()
    headers = {
        "Authorization": f"Bearer {OLLAMA_API_KEY}",
        "Content-Type": "application/json"
    }
    payload = {
        "model": model,
        "messages": [{"role": "user", "content": prompt}],
        "stream": False,
        "format": "json",
        "options": {"temperature": 0.1}
    }
    resp = requests.post(
        OLLAMA_URL,
        headers=headers,
        json=payload,
        timeout=(AI_CONNECT_TIMEOUT, AI_READ_TIMEOUT)
    )
    resp.raise_for_status()
    content = resp.json()["message"]["content"]
    return json.loads(content)


def ai_generate_json(prompt, task_name):
    providers = []
    if GEMINI_API_KEY:
        providers.append(("gemini", _call_gemini_json))
    if OPENROUTER_API_KEY:
        providers.append(("openrouter", _call_openrouter_json))
    if OLLAMA_API_KEY:
        providers.append(("ollama", _call_ollama_json))

    last_err = None
    for provider_name, fn in providers:
        for attempt in range(len(AI_RETRY_DELAYS) + 1):
            try:
                return fn(prompt)
            except Exception as exc:
                last_err = exc
                if attempt < len(AI_RETRY_DELAYS):
                    time.sleep(AI_RETRY_DELAYS[attempt])
    raise RuntimeError(f"All AI providers failed for {task_name}") from last_err


def ai_website_name(website):
    fallback = fallback_clean_website_name(website)
    if not website:
        return {"name": fallback, "sure": False}
    prompt = (
        "You format website domains as the human-readable name of the "
        "person, publisher, or brand being contacted for a professional email greeting.\n"
        "RULES:\n"
        "1. Remove URL parts (https://, www., paths, domain suffixes).\n"
        "2. Split joined words into separate words when clear.\n"
        "3. Use proper title case.\n"
        "4. Set 'sure' to false if the domain is ambiguous or random characters. Otherwise true.\n"
        "5. Return ONLY valid JSON: {\"website_name\": \"...\", \"sure\": true}\n"
        f"Website domain: {website}"
    )
    try:
        data = ai_generate_json(prompt, "website_name")
        name = str(data.get("website_name", "")).strip()
        sure = bool(data.get("sure", True))
        if not name or len(name) > 80:
            return {"name": fallback, "sure": False}
        return {"name": name, "sure": sure}
    except Exception:
        raw_host = re.sub(r"^https?://(www\.)?", "", str(website).strip(), flags=re.I).split("/")[0].split(".")[0]
        has_hints = bool(re.search(r"[-_]|[A-Z]", raw_host)) or len(raw_host) <= 8
        return {"name": fallback, "sure": has_hints}


def _formatting_tokens(value):
    return re.findall(r"\w+|[^\w\s]", value or "", flags=re.UNICODE)


def ai_format_template(template):
    original_subject = template["subject"]
    original_body = template["body"]
    prompt = (
        "Improve email layout using only whitespace, line breaks, and blank lines between paragraphs. "
        "Keep greeting on its own line with a blank line after it. "
        "Preserve every single word, placeholder, punctuation mark, capitalization, and order identically. "
        "Return ONLY JSON: {\"subject\": \"...\", \"body\": \"...\"}\n\n"
        f"SUBJECT:\n{original_subject}\n\nBODY:\n{original_body}"
    )
    try:
        data = ai_generate_json(prompt, "template_formatting")
        formatted_subject = str(data.get("subject", ""))
        formatted_body = str(data.get("body", ""))
        if (
            _formatting_tokens(formatted_subject) != _formatting_tokens(original_subject)
            or _formatting_tokens(formatted_body) != _formatting_tokens(original_body)
        ):
            return template
        return {
            **template,
            "subject": formatted_subject.strip(),
            "body": formatted_body.strip()
        }
    except Exception:
        return template


# =========================================================
# MAILBOXES & WARMUP LIMITS
# =========================================================

def load_mailboxes():
    raw_mailboxes = json.loads(MAILBOXES_JSON)
    if not isinstance(raw_mailboxes, list):
        raise ValueError("MAILBOXES_JSON must be a JSON list.")
    active_mailboxes = []
    for mailbox in raw_mailboxes:
        if not isinstance(mailbox, dict):
            raise ValueError("Each mailbox must be a JSON object.")
        if mailbox.get("active", True) is not True:
            continue
        for field in ("email", "password", "smtp_host", "smtp_port"):
            if not mailbox.get(field):
                raise ValueError(f"Missing '{field}' in an active mailbox.")
        mailbox["email"] = mailbox["email"].strip().lower()
        if not valid_email(mailbox["email"]):
            raise ValueError(f"Invalid mailbox email: {mailbox['email']}")
        mailbox.setdefault("name", mailbox["email"])
        mailbox.setdefault("smtp_port", 587)
        if not mailbox.get("imap_host"):
            smtp_host_lower = str(mailbox["smtp_host"]).strip().lower()
            if smtp_host_lower == "smtp.hostinger.com":
                mailbox["imap_host"] = "imap.hostinger.com"
            else:
                mailbox["imap_host"] = mailbox["smtp_host"]
        mailbox.setdefault("imap_port", 993)
        active_mailboxes.append(mailbox)
    if not active_mailboxes:
        raise RuntimeError("No active mailboxes configured.")

    # Keep rotation deterministic even if MAILBOXES_JSON is reordered.
    canonical_order = {
        "hamza.khan@westernarsal.com": 0,
        "arslan.lodhi@westernarsal.com": 1,
        "james.williams@westernarsal.com": 2,
    }
    active_mailboxes.sort(
        key=lambda item: canonical_order.get(normalize_email(item["email"]), 99)
    )
    if len(active_mailboxes) > 3:
        raise RuntimeError("This setup supports exactly these three mailboxes at most.")
    return active_mailboxes


def get_mailbox_daily_limit(mailbox, history_rows):
    # Hard cap each active mailbox at 10 sent emails per day.
    max_limit = min(int(mailbox.get("daily_limit", DAILY_LIMIT_PER_MAILBOX)), DAILY_LIMIT_PER_MAILBOX)
    mb_email = normalize_email(mailbox["email"])
    first_date = None
    created_at_str = str(mailbox.get("created_at", "")).strip()
    if created_at_str:
        try:
            first_date = datetime.fromisoformat(created_at_str).date()
        except ValueError:
            pass
    if first_date is None:
        for row in history_rows:
            if normalize_email(row.get("Mailbox", "")) == mb_email and str(row.get("Status", "")).upper() == "SENT":
                dt_str = str(row.get("Date", "")).strip()
                try:
                    d = datetime.strptime(dt_str, "%Y-%m-%d").date()
                    if first_date is None or d < first_date:
                        first_date = d
                except ValueError:
                    continue
    days_active = 0 if first_date is None else max(0, (now_local().date() - first_date).days)
    if days_active <= 2:
        ramped = 2
    elif days_active <= 6:
        ramped = 4
    elif days_active <= 13:
        ramped = 6
    elif days_active <= 20:
        ramped = 8
    else:
        ramped = max_limit
    return min(ramped, max_limit)


def is_mailbox_cooled_down(history_rows, mailbox_email, cooldown_hours=4):
    mb_email = normalize_email(mailbox_email)
    consecutive_failures = 0
    latest_fail_dt = None
    for row in reversed(history_rows):
        if normalize_email(row.get("Mailbox", "")) != mb_email:
            continue
        status = str(row.get("Status", "")).strip().upper()
        if status == "SENT":
            break
        if status == "FAILED":
            consecutive_failures += 1
            if latest_fail_dt is None:
                raw_ts = str(row.get("Timestamp", "")).strip()
                try:
                    latest_fail_dt = datetime.strptime(raw_ts, "%Y-%m-%d %H:%M:%S %z").astimezone(EASTERN_TZ)
                except ValueError:
                    try:
                        latest_fail_dt = datetime.strptime(raw_ts, "%Y-%m-%d %H:%M:%S").replace(tzinfo=EASTERN_TZ)
                    except ValueError:
                        pass
            if consecutive_failures >= 3:
                break
    if consecutive_failures >= 3 and latest_fail_dt is not None:
        if now_local() - latest_fail_dt < timedelta(hours=cooldown_hours):
            return False, consecutive_failures, latest_fail_dt
    return True, consecutive_failures, latest_fail_dt


def get_ordered_candidate_mailboxes(mailboxes, history_rows):
    last_mailbox = None
    for row in reversed(history_rows):
        if str(row.get("Status", "")).strip().upper() == "SENT":
            last_mailbox = normalize_email(row.get("Mailbox", ""))
            break
    addresses = [normalize_email(m["email"]) for m in mailboxes]
    if not last_mailbox or last_mailbox not in addresses:
        return list(mailboxes)
    start_idx = (addresses.index(last_mailbox) + 1) % len(mailboxes)
    return mailboxes[start_idx:] + mailboxes[:start_idx]


def daily_sent_count(history_rows, mailbox_email, website_rows=None):
    """Count confirmed and uncertain sends conservatively against the daily cap."""
    today = now_local().date().isoformat()
    mb_email = normalize_email(mailbox_email)
    counted_recipients = set()
    for row in history_rows:
        if (
            str(row.get("Date", "")).strip() == today
            and normalize_email(row.get("Mailbox", "")) == mb_email
            and str(row.get("Status", "")).strip().upper() in ("SENT", "SEND_REVIEW")
        ):
            address = normalize_email(row.get("Email", ""))
            counted_recipients.add(address or f"history-row-{len(counted_recipients)}")
    for row in website_rows or []:
        if (
            str(row.get("Status", "")).strip().upper() == "SENT"
            and normalize_email(row.get("Mailbox", "")) == mb_email
            and normalize_email(row.get("Last Sent At", ""))[:10] == today
        ):
            address = normalize_email(row.get("Email", ""))
            if address:
                counted_recipients.add(address)
    return len(counted_recipients)



# =========================================================
# INCOMING REPLY DETECTION & TRACKING
# =========================================================

EMAIL_FINDER = re.compile(r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}")

AUTO_SUBJECT_TERMS = (
    "out of office", "automatic reply", "auto-reply", "autoreply",
    "auto reply", "away from my desk", "vacation reply",
)
NOREPLY_PREFIXES = ("noreply@", "no-reply@", "donotreply@", "do-not-reply@")

# A stronger status never gets overwritten by a weaker one.
_REPLY_STATUS_PRIORITY = {
    "": 0,
    "AUTOREPLY": 1,
    "REPLY": 2,
    "BOUNCE": 3,
    "UNSUBSCRIBE": 3,
}


def is_auto_responder(message):
    """True for out-of-office / automatic mails (RFC 3834 and common headers)."""
    auto_submitted = str(message.get("Auto-Submitted", "")).strip().lower()
    if auto_submitted and auto_submitted != "no":
        return True
    precedence = str(message.get("Precedence", "")).strip().lower()
    if precedence in ("bulk", "junk", "auto_reply", "list"):
        return True
    if message.get("X-Autoreply") or message.get("X-Autorespond"):
        return True
    return False


def _looks_automatic(subject, sender):
    subject_l = (subject or "").lower()
    if any(term in subject_l for term in AUTO_SUBJECT_TERMS):
        return True
    return (sender or "").lower().startswith(NOREPLY_PREFIXES)


def classify_reply(subject, body, from_address="", auto_responder=False):
    text = f"{subject} {body}".lower()
    sender = (from_address or "").lower()
    if sender.startswith(("mailer-daemon@", "postmaster@")):
        return "BOUNCE"
    if any(term in text for term in [
        "mailbox not found", "user unknown", "address rejected",
        "delivery failed", "undelivered mail", "delivery status notification"
    ]):
        return "BOUNCE"
    if auto_responder or _looks_automatic(subject, sender):
        return "AUTOREPLY"
    if any(term in text for term in [
        "unsubscribe", "remove me", "do not contact", "stop emailing", "opt out"
    ]):
        return "UNSUBSCRIBE"
    return "REPLY"


def extract_body(message):
    parts = []
    if message.is_multipart():
        for part in message.walk():
            if (
                part.get_content_type() == "text/plain"
                and part.get_content_disposition() != "attachment"
            ):
                payload = part.get_payload(decode=True)
                if payload:
                    parts.append(
                        payload.decode(
                            part.get_content_charset() or "utf-8",
                            errors="replace"
                        )
                    )
    else:
        payload = message.get_payload(decode=True)
        if payload:
            parts.append(
                payload.decode(
                    message.get_content_charset() or "utf-8",
                    errors="replace"
                )
            )
    return "\n".join(parts).strip()


def extract_bounced_recipients(message, body, known_recipients):
    """Extract only structured failed-recipient fields from a DSN message."""
    candidates = set()

    def add_addresses(value):
        for address in EMAIL_FINDER.findall(str(value or "")):
            normalized = normalize_email(address)
            if normalized in known_recipients:
                candidates.add(normalized)

    add_addresses(message.get("X-Failed-Recipients", ""))
    add_addresses(message.get("Final-Recipient", ""))
    add_addresses(message.get("Original-Recipient", ""))
    # Some providers flatten DSN fields into plain-text body lines. Accept only
    # explicit Final-Recipient / Original-Recipient fields, never arbitrary addresses.
    for match in re.findall(
        r"(?im)^\s*(?:Final|Original)-Recipient\s*:\s*(?:rfc822\s*;\s*)?([^\s<>;,]+@[^\s<>;,]+)",
        body or "",
    ):
        add_addresses(match)

    if message.is_multipart():
        for part in message.walk():
            if part.get_content_type() == "message/delivery-status":
                payload = part.get_payload()
                blocks = payload if isinstance(payload, list) else [payload]
                for block in blocks:
                    if not hasattr(block, "get"):
                        continue
                    add_addresses(block.get("Final-Recipient", ""))
                    add_addresses(block.get("Original-Recipient", ""))

    return candidates


def _message_id_tokens(value):
    """Extract normalized RFC Message-ID tokens from In-Reply-To/References."""
    if not value:
        return set()
    found = re.findall(r"<[^<>]+>", str(value))
    if not found:
        found = [part for part in str(value).split() if "@" in part]
    return {item.strip().lower() for item in found if item.strip()}


def read_replies(mailboxes, replies_ws, suppression_ws, websites_ws, history_ws, archive_ws, run_log_ws):
    """Record incoming mail; only mark REPLY when it matches an outbound thread."""
    existing_ids = {
        str(row.get("Message ID", "")).strip().lower()
        for row in get_rows(replies_ws)
        if row.get("Message ID")
    }

    website_rows = get_rows(websites_ws)
    rows_by_email = {}
    status_by_row = {}
    for idx, website_row in enumerate(website_rows, start=2):
        address = normalize_email(website_row.get("Email", ""))
        if address:
            rows_by_email.setdefault(address, []).append(idx)
        status_by_row[idx] = str(website_row.get("Reply Status", "")).strip().upper()

    known_recipients = set(rows_by_email)
    suppressed_now = get_suppressed_emails(suppression_ws)

    # Only successful outbound sends can establish a reply thread.
    sent_message_to_recipient = {}
    outbound_history = get_rows(history_ws) + (get_rows(archive_ws) if archive_ws is not None else [])
    for history_row in outbound_history:
        if str(history_row.get("Status", "")).strip().upper() != "SENT":
            continue
        recipient = normalize_email(history_row.get("Email", ""))
        if recipient not in known_recipients:
            continue
        for token in _message_id_tokens(history_row.get("Message ID", "")):
            sent_message_to_recipient[token] = recipient

    for mailbox in mailboxes:
        if not mailbox.get("imap_host"):
            continue
        connection = None
        try:
            connection = imaplib.IMAP4_SSL(
                mailbox["imap_host"], int(mailbox.get("imap_port", 993)),
                ssl_context=ssl.create_default_context()
            )
            connection.login(mailbox["email"], mailbox["password"])
            connection.select("INBOX")
            status, data = connection.search(None, "UNSEEN")
            if status != "OK":
                continue

            for number in data[0].split():
                status, message_data = connection.fetch(number, "(RFC822)")
                if status != "OK" or not message_data or not isinstance(message_data[0], tuple):
                    continue

                message = email.message_from_bytes(message_data[0][1])
                message_id = str(message.get("Message-ID", "")).strip()
                message_id_key = message_id.lower()
                from_address = normalize_email(parseaddr(message.get("From", ""))[1])
                subject = decode_text(message.get("Subject", ""))
                body = extract_body(message)
                initial_classification = classify_reply(
                    subject, body, from_address, is_auto_responder(message)
                )
                snippet = body[:1200].replace("\n", " ")

                reference_ids = (
                    _message_id_tokens(message.get("In-Reply-To", ""))
                    | _message_id_tokens(message.get("References", ""))
                )
                matched_recipients = {
                    sent_message_to_recipient[token]
                    for token in reference_ids
                    if token in sent_message_to_recipient
                }

                if initial_classification == "BOUNCE":
                    classification = "BOUNCE"
                    targets = extract_bounced_recipients(message, body, known_recipients)
                elif initial_classification == "UNSUBSCRIBE":
                    # Honor opt-outs even if the sender uses an alternate address in a known thread.
                    classification = "UNSUBSCRIBE"
                    targets = matched_recipients or ({from_address} if from_address in known_recipients else set())
                elif initial_classification == "AUTOREPLY" and matched_recipients:
                    classification = "AUTOREPLY"
                    targets = matched_recipients
                elif initial_classification == "AUTOREPLY" and from_address in known_recipients:
                    classification = "AUTOREPLY"
                    targets = {from_address}
                elif matched_recipients:
                    # Thread headers are stronger evidence than the visible From address;
                    # contacts sometimes reply from an alias or a shared mailbox.
                    classification = initial_classification
                    targets = matched_recipients
                else:
                    # Do not treat unrelated/old mail from a contact as a reply.
                    classification = "UNMATCHED"
                    targets = set()

                if not message_id_key or message_id_key not in existing_ids:
                    sheets_retry(
                        replies_ws.append_row,
                        [timestamp(), from_address, subject, message_id, classification, snippet]
                    )
                    if message_id_key:
                        existing_ids.add(message_id_key)

                for target in targets:
                    if classification in ("UNSUBSCRIBE", "BOUNCE") and target not in suppressed_now:
                        add_to_suppression(suppression_ws, target, classification)
                        suppressed_now.add(target)
                    for sheet_row in rows_by_email.get(target, []):
                        current = status_by_row.get(sheet_row, "")
                        if _REPLY_STATUS_PRIORITY.get(classification, 0) >= _REPLY_STATUS_PRIORITY.get(current, 0):
                            sheets_retry(websites_ws.update_cell, sheet_row, 8, classification)
                            status_by_row[sheet_row] = classification

                connection.store(number, "+FLAGS", "\\Seen")
                if classification == "UNMATCHED" and run_log_ws:
                    write_run_log(
                        run_log_ws, "INFO", "INCOMING_UNMATCHED",
                        f"Mailbox={mailbox['email']} | From={from_address} | Message-ID={message_id or 'missing'}"
                    )

        except Exception as exc:
            logger.exception("Reading replies failed for %s", mailbox.get("email"))
            if run_log_ws:
                write_run_log(
                    run_log_ws, "ERROR", "IMAP_FAILED",
                    f"{mailbox.get('email')}: {type(exc).__name__}: {exc}"
                )
        finally:
            if connection is not None:
                try:
                    connection.logout()
                except Exception:
                    pass



# =========================================================
# REPAIR, TASK DISCOVERY & ARCHIVING
# =========================================================

def repair_stuck_sending_rows(websites_ws, history_ws, run_log_ws):
    """Resolve stale SENDING rows conservatively to prevent accidental duplicates."""
    values = sheets_retry(websites_ws.get_all_values)
    if not values or len(values) < 2:
        return 0
    headers = [str(v).strip() for v in values[0]]
    h_map = {name: idx for idx, name in enumerate(headers)}
    required = ("Email", "Status", "Task ID", "Created At", "Last Sent At", "Mailbox")
    if any(col not in h_map for col in required):
        return 0

    # Use the latest History event per recipient, not any historical SENT event.
    latest_by_email = {}
    for history_row in get_rows(history_ws):
        recipient = normalize_email(history_row.get("Email", ""))
        if recipient:
            latest_by_email[recipient] = history_row

    repaired = 0
    for sheet_row, row in enumerate(values[1:], start=2):
        def get_col(col):
            idx = h_map[col]
            return str(row[idx]).strip() if idx < len(row) and row[idx] is not None else ""
        if get_col("Status").upper() != "SENDING":
            continue

        recipient = normalize_email(get_col("Email"))
        latest = latest_by_email.get(recipient, {})
        latest_status = str(latest.get("Status", "")).strip().upper()
        if latest_status == "SENT":
            new_status = "SENT"
            last_sent = str(latest.get("Timestamp", "")).strip() or timestamp()
            mailbox = normalize_email(latest.get("Mailbox", "")) or get_col("Mailbox")
            values_to_write = [new_status, get_col("Task ID"), get_col("Created At"), last_sent, mailbox]
        else:
            # No definitive SENT record: SMTP may have accepted before a crash.
            # Hold for manual verification rather than automatically sending again.
            values_to_write = ["SEND_REVIEW", get_col("Task ID"), get_col("Created At"), get_col("Last Sent At"), get_col("Mailbox")]
            write_run_log(run_log_ws, "WARNING", "SEND_REVIEW_REQUIRED", f"Recipient={recipient} has a stale SENDING row without a definitive SENT history entry; verify mailbox delivery before resetting to PENDING.")

        sheets_retry(websites_ws.update, range_name=f"C{sheet_row}:G{sheet_row}", values=[values_to_write])
        repaired += 1

    if repaired:
        write_run_log(run_log_ws, "WARNING", "STUCK_SENDING_REPAIRED", f"Reviewed stale rows={repaired}")
    return repaired


def archive_old_logs(spreadsheet, history_ws, run_log_ws, retention_days=30):
    cutoff = now_local().date() - timedelta(days=retention_days)
    try:
        archive_ws = get_worksheet(spreadsheet, "Archive")
        all_hist = sheets_retry(history_ws.get_all_values)
        if len(all_hist) > 1:
            headers = all_hist[0]
            keep_rows, move_rows = [], []
            for row in all_hist[1:]:
                date_str = str(row[1]).strip() if len(row) > 1 else ""
                try:
                    if datetime.strptime(date_str, "%Y-%m-%d").date() < cutoff:
                        move_rows.append((row + [""] * 9)[:9])
                    else:
                        keep_rows.append(row)
                except ValueError:
                    keep_rows.append(row)
            if move_rows:
                sheets_retry(archive_ws.append_rows, move_rows)
                sheets_retry(history_ws.clear)
                sheets_retry(history_ws.update, range_name=f"A1:{column_letter(len(headers))}1", values=[headers])
                if keep_rows:
                    sheets_retry(history_ws.append_rows, keep_rows)
    except Exception as exc:
        logger.exception("History archiving failed.")
        write_run_log(run_log_ws, "ERROR", "ARCHIVE_HISTORY_FAILED", repr(exc))

    try:
        run_archive_ws = get_worksheet(spreadsheet, "Run Log Archive")
        all_logs = sheets_retry(run_log_ws.get_all_values)
        if len(all_logs) > 1:
            headers = all_logs[0]
            keep_logs, move_logs = [], []
            for row in all_logs[1:]:
                try:
                    if datetime.strptime(str(row[0]).strip()[:10], "%Y-%m-%d").date() < cutoff:
                        move_logs.append((row + [""] * 4)[:4])
                    else:
                        keep_logs.append(row)
                except ValueError:
                    keep_logs.append(row)
            if move_logs:
                sheets_retry(run_archive_ws.append_rows, move_logs)
                sheets_retry(run_log_ws.clear)
                sheets_retry(run_log_ws.update, range_name="A1:D1", values=[headers])
                if keep_logs:
                    sheets_retry(run_log_ws.append_rows, keep_logs)
    except Exception:
        logger.exception("Run Log archiving failed.")


def discover_new_tasks(websites_ws, run_log_ws=None):
    values = sheets_retry(websites_ws.get_all_values)
    if not values:
        return 0
    headers = [str(value).strip() for value in values[0]]
    header_map = {name: idx for idx, name in enumerate(headers)}
    if any(name not in header_map for name in ["Website", "Email", "Status", "Task ID", "Created At"]):
        return 0

    def cell(row, column):
        idx = header_map[column]
        return str(row[idx]).strip() if idx < len(row) and row[idx] is not None else ""

    created_count = 0
    for sheet_row, row in enumerate(values[1:], start=2):
        website = cell(row, "Website")
        raw_email = cell(row, "Email")
        recipient = normalize_email(raw_email)
        status = cell(row, "Status")
        if not website or not recipient or not valid_email(recipient) or status:
            continue
        task_id = hashlib.sha256(f"{website}|{recipient}".encode("utf-8")).hexdigest()[:16]
        sheets_retry(
            websites_ws.update,
            range_name=f"C{sheet_row}:E{sheet_row}",
            values=[["PENDING", task_id, timestamp()]]
        )
        created_count += 1
    if run_log_ws and created_count > 0:
        write_run_log(run_log_ws, "INFO", "TASK_DISCOVERY", f"Created PENDING={created_count}")
    return created_count


# =========================================================
# TEMPLATES, SUPPRESSION & SENDING
# =========================================================

def get_active_templates(templates_ws):
    rows = get_rows(templates_ws)
    active = []
    for row in rows:
        template_id = str(row.get("Template ID", "")).strip()
        subject = str(row.get("Subject", "")).strip()
        body = str(row.get("Body", "")).strip()
        # Keep the old BEST_PRICE reply template out of regular outreach.
        if template_id.lower() == LEGACY_BEST_PRICE_TEMPLATE_ID.lower():
            continue
        if str(row.get("Active", "")).strip().lower() in ("true", "yes", "1", "active"):
            if template_id and subject and body:
                active.append({"id": template_id, "subject": subject, "body": body})
    return active


def select_template(templates, task_number):
    if not templates:
        return None
    return templates[task_number % len(templates)]


def get_suppressed_emails(suppression_ws):
    rows = get_rows(suppression_ws)
    return {normalize_email(row.get("Email", "")) for row in rows if row.get("Email")}


def get_contacted_emails(history_rows):
    return {
        normalize_email(row.get("Email", ""))
        for row in history_rows
        if str(row.get("Status", "")).strip().upper() == "SENT" and row.get("Email")
    }


def add_to_suppression(suppression_ws, address, reason):
    address = normalize_email(address)
    if not address:
        return
    if address not in get_suppressed_emails(suppression_ws):
        sheets_retry(suppression_ws.append_row, [address, reason, timestamp()])


def is_bounce_rate_too_high(history_rows, suppression_ws):
    recent_sends = [r for r in history_rows if str(r.get("Status", "")).strip().upper() == "SENT"][-50:]
    if len(recent_sends) < 10:
        return False, 0.0
    bounced = {normalize_email(r.get("Email", "")) for r in get_rows(suppression_ws) if str(r.get("Reason", "")).upper() == "BOUNCE"}
    recent_bounces = sum(1 for r in recent_sends if normalize_email(r.get("Email", "")) in bounced)
    rate = recent_bounces / len(recent_sends)
    return rate > 0.15, rate


def append_history(history_ws, recipient, website, mailbox, status, message_id="", details="", template_id=""):
    sheets_retry(
        history_ws.append_row,
        [timestamp(), now_local().date().isoformat(), recipient, website, mailbox, status, message_id, details, template_id]
    )


def _parse_signature(sig_str):
    """Return the raw HTML secret unchanged; retain legacy JSON compatibility."""
    if sig_str is None or not str(sig_str).strip():
        return ""
    raw = str(sig_str)
    try:
        parsed = json.loads(raw)
        if isinstance(parsed, dict) and isinstance(parsed.get("html"), str):
            return parsed["html"]
        if isinstance(parsed, str):
            return parsed
    except json.JSONDecodeError:
        pass
    return raw


def get_signature_html(mailbox, mailboxes):
    # Explicit mapping prevents a reordered MAILBOXES_JSON from assigning
    # James's signature to another sender (or vice versa).
    signature_by_email = {
        "hamza.khan@westernarsal.com": (HAMZA_SIGNATURE, "HAMZA_SIGNATURE"),
        "arslan.lodhi@westernarsal.com": (ARSLAN_SIGNATURE, "ARSLAN_SIGNATURE"),
        "james.williams@westernarsal.com": (JAMES_SIGNATURE, "JAMES_SIGNATURE"),
    }
    mb_email = normalize_email(mailbox.get("email"))
    if mb_email not in signature_by_email:
        raise RuntimeError(f"No signature mapping configured for mailbox {mb_email}.")
    raw_signature, secret_name = signature_by_email[mb_email]
    signature = _parse_signature(raw_signature)
    if not signature or not signature.strip():
        raise RuntimeError(
            f"Signature is empty (check GitHub Secret {secret_name}). Email was not sent."
        )
    return signature


def remove_body_signoff(body):
    body = body or ""
    lines = body.replace("\r\n", "\n").replace("\r", "\n").split("\n")
    for index, line in enumerate(lines):
        if re.fullmatch(r"\s*best regards\s*,?\s*", line, flags=re.IGNORECASE):
            return "\n".join(lines[:index]).rstrip()
    return body.strip()


def html_signature_to_text(signature_html):
    text = re.sub(r"(?i)<\s*br\s*/?\s*>", "\n", signature_html)
    text = re.sub(r"(?i)</\s*(p|div|tr|li|h[1-6])\s*>", "\n", text)
    text = re.sub(r"<[^>]+>", "", text)
    return re.sub(r"\n[ \t]*\n[ \t]*\n+", "\n\n", html.unescape(text)).strip()


def _sent_folder_candidates(connection):
    """Return likely Sent folders, preferring the IMAP folder marked \\Sent."""
    status, folders = connection.list()
    if status != "OK" or not folders:
        raise RuntimeError("Could not list IMAP folders to locate the Sent folder.")

    flagged = []
    named = []
    for raw_line in folders:
        if not raw_line:
            continue
        line = raw_line.decode("utf-8", errors="replace") if isinstance(raw_line, bytes) else str(raw_line)

        # LIST responses normally end with a quoted mailbox name.
        match = re.search(r'"((?:[^"\\\\]|\\\\.)*)"\s*$', line)
        if match:
            folder = match.group(1).replace(r'\\"', '"').replace(r"\\\\", "\\")
        else:
            # Fallback for servers that return an unquoted mailbox name.
            parts = line.rsplit(" ", 1)
            folder = parts[-1].strip().strip('"') if parts else ""

        if not folder or folder.upper() == "NIL":
            continue

        if re.search(r'\\Sent\b', line, flags=re.IGNORECASE):
            flagged.append(folder)
        if "sent" in folder.lower():
            named.append(folder)

    candidates = []
    for folder in flagged + named + [
        "Sent", "INBOX.Sent", "Sent Items", "INBOX.Sent Items",
        "Sent Messages", "INBOX.Sent Messages"
    ]:
        if folder and folder not in candidates:
            candidates.append(folder)
    return candidates


def save_sent_copy_to_imap(mailbox, message):
    """Save a copy in the mailbox's IMAP Sent folder after SMTP accepts it."""
    imap_host = mailbox.get("imap_host")
    if not imap_host:
        smtp_host = str(mailbox.get("smtp_host", "")).strip().lower()
        imap_host = "imap.hostinger.com" if smtp_host == "smtp.hostinger.com" else smtp_host

    imap_port = int(mailbox.get("imap_port", 993))
    connection = None
    try:
        connection = imaplib.IMAP4_SSL(
            imap_host,
            imap_port,
            ssl_context=ssl.create_default_context(),
        )
        connection.login(mailbox["email"], mailbox["password"])

        candidates = _sent_folder_candidates(connection)
        raw_message = message.as_bytes()
        errors = []

        for folder in candidates:
            try:
                status, response = connection.append(
                    folder,
                    r"(\Seen)",
                    imaplib.Time2Internaldate(time.time()),
                    raw_message,
                )
                if status == "OK":
                    return folder
                errors.append(f"{folder}: {status} {response}")
            except Exception as exc:
                errors.append(f"{folder}: {type(exc).__name__}: {exc}")

        raise RuntimeError(
            "Could not append the sent message to any detected Sent folder. "
            + " | ".join(errors[:6])
        )
    finally:
        if connection is not None:
            try:
                connection.logout()
            except Exception:
                pass


class SMTPDeliveryUncertain(RuntimeError):
    """SMTP connection failed at a point where recipient acceptance is unknown."""


def _send_message_with_uncertainty_guard(server, message):
    try:
        return server.send_message(message)
    except (smtplib.SMTPRecipientsRefused, smtplib.SMTPDataError):
        # The SMTP server explicitly rejected the recipient/message.
        raise
    except smtplib.SMTPResponseException as exc:
        # An explicit 4xx/5xx response is a rejection, not an accepted send.
        if getattr(exc, "smtp_code", 550) >= 400:
            raise
        raise SMTPDeliveryUncertain(str(exc)) from exc
    except Exception as exc:
        # A dropped connection/timeout after DATA can leave delivery ambiguous.
        raise SMTPDeliveryUncertain(
            f"SMTP outcome is uncertain ({type(exc).__name__}: {exc})"
        ) from exc


def send_email(mailbox, recipient, subject, body, mailboxes):
    signature_html = get_signature_html(mailbox, mailboxes)
    clean_body = remove_body_signoff(body)
    message = EmailMessage()
    message["From"] = formataddr((mailbox.get("name", "Western Arsal"), mailbox["email"]))
    message["To"] = recipient
    message["Subject"] = subject
    message["Message-ID"] = make_msgid()

    # Keep the plain-text alternative limited to the authored body. The raw HTML
    # signature secret is appended once, unchanged, in the HTML alternative below.
    message.set_content(clean_body.rstrip())

    paragraphs = re.split(r"\n[ \t]*\n+", clean_body.replace("\r\n", "\n").strip()) if clean_body else []
    body_parts = []
    for paragraph in paragraphs:
        escaped = html.escape(paragraph).replace("\n", "<br>\n")
        body_parts.append(f'<p style="margin:0 0 16px 0;">{escaped}</p>')
    body_html = "".join(body_parts)
    html_content = (
        '<!doctype html><html><body><div style="font-family:Arial,sans-serif;'
        'font-size:14px;line-height:1.6;color:#333;">'
        f'{body_html}<div style="margin-top:24px;">{signature_html}</div>'
        '</div></body></html>'
    )
    message.add_alternative(html_content, subtype="html")

    context = ssl.create_default_context()
    smtp_host, smtp_port = mailbox["smtp_host"], int(mailbox["smtp_port"])
    server = None
    send_attempted = False
    accepted = False
    try:
        # Connection, TLS and authentication failures happen before DATA and are safe to fail over.
        if smtp_port == 465:
            server = smtplib.SMTP_SSL(smtp_host, smtp_port, context=context, timeout=30)
        else:
            server = smtplib.SMTP(smtp_host, smtp_port, timeout=30)
            server.ehlo()
            server.starttls(context=context)
            server.ehlo()
        server.login(mailbox["email"], mailbox["password"])

        send_attempted = True
        _send_message_with_uncertainty_guard(server, message)
        accepted = True
    except SMTPDeliveryUncertain:
        raise
    except (smtplib.SMTPRecipientsRefused, smtplib.SMTPDataError, smtplib.SMTPResponseException):
        # Explicit SMTP rejection: this server did not accept the message.
        raise
    except Exception as exc:
        if send_attempted and not accepted:
            raise SMTPDeliveryUncertain(
                f"SMTP outcome uncertain ({type(exc).__name__}: {exc})"
            ) from exc
        raise
    finally:
        if server is not None:
            try:
                # Once send_message returned, QUIT/close failures must never trigger failover.
                server.quit()
            except Exception:
                try:
                    server.close()
                except Exception:
                    pass
                if accepted:
                    logger.warning("SMTP_QUIT_FAILED after accepted send; not retrying via another mailbox.")

    # SMTP accepted the email. A Sent-folder copy failure is non-fatal.
    try:
        sent_folder = save_sent_copy_to_imap(mailbox, message)
        logger.info("SENT_COPY_SAVED | mailbox=%s | folder=%s | message_id=%s", mailbox["email"], sent_folder, message["Message-ID"])
        if _ACTIVE_RUN_LOG_WS is not None:
            write_run_log(_ACTIVE_RUN_LOG_WS, "INFO", "SENT_COPY_SAVED", f"{mailbox['email']} | folder={sent_folder} | Message-ID={message['Message-ID']}")
    except Exception as exc:
        logger.exception("SENT_COPY_SAVE_FAILED | SMTP accepted email to %s via %s: %s: %s", recipient, mailbox["email"], type(exc).__name__, exc)
        if _ACTIVE_RUN_LOG_WS is not None:
            write_run_log(_ACTIVE_RUN_LOG_WS, "WARNING", "SENT_COPY_SAVE_FAILED", f"SMTP accepted email to {recipient} via {mailbox['email']}, but IMAP Sent copy failed: {type(exc).__name__}: {exc}")

    return message["Message-ID"]

def process_sends(spreadsheet, mailboxes, websites_ws, templates_ws, history_ws, suppression_ws, run_log_ws):
    history_rows = get_rows(history_ws)
    too_many_bounces, bounce_rate = is_bounce_rate_too_high(history_rows, suppression_ws)
    if too_many_bounces:
        write_run_log(run_log_ws, "WARNING", "BOUNCE_GUARD_STOP", f"Bounce rate {bounce_rate:.1%} (>15%)")
        return

    templates = get_active_templates(templates_ws)
    if not templates:
        return

    rows = get_rows(websites_ws)
    suppressed = get_suppressed_emails(suppression_ws)
    archive_rows = get_rows(get_worksheet(spreadsheet, "Archive"))
    contacted = get_contacted_emails(history_rows + archive_rows)
    ordered_mailboxes = get_ordered_candidate_mailboxes(mailboxes, history_rows)
    eligible_mailboxes = [
        mb for mb in ordered_mailboxes
        if daily_sent_count(history_rows, mb["email"], rows) < get_mailbox_daily_limit(mb, history_rows)
        and is_mailbox_cooled_down(history_rows, mb["email"])[0]
    ]
    if not eligible_mailboxes:
        cooldowns = [
            f"{mb['email']}={is_mailbox_cooled_down(history_rows, mb['email'])[1]} consecutive failures"
            for mb in ordered_mailboxes if not is_mailbox_cooled_down(history_rows, mb['email'])[0]
        ]
        write_run_log(run_log_ws, "INFO", "NO_ELIGIBLE_MAILBOX", "All mailboxes are at their daily limit or in cooldown. " + "; ".join(cooldowns))
        return

    for index, row in enumerate(rows):
        website = str(row.get("Website", "")).strip()
        recipient = normalize_email(row.get("Email", ""))
        status = str(row.get("Status", "")).strip().upper()
        sheet_row = index + 2
        if status != "PENDING":
            continue
        if not website or not valid_email(recipient):
            sheets_retry(websites_ws.update_cell, sheet_row, 3, "INVALID")
            continue
        if recipient in suppressed:
            sheets_retry(websites_ws.update_cell, sheet_row, 3, "SUPPRESSED")
            continue
        if recipient in contacted or any(
            normalize_email(item.get("Email", "")) == recipient
            and str(item.get("Status", "")).strip().upper() == "SENT"
            for item in rows
        ):
            sheets_retry(websites_ws.update_cell, sheet_row, 3, "ALREADY_CONTACTED")
            continue

        name_info = ai_website_name(website)
        if not name_info["sure"]:
            existing_notes = str(row.get("Notes", "")).strip()
            note_addition = f"Suggested: {name_info['name']}"
            new_notes = f"{existing_notes} | {note_addition}".strip(" |") if existing_notes else note_addition
            sheets_retry(
                websites_ws.update,
                range_name=f"C{sheet_row}:I{sheet_row}",
                values=[["NAME_UNSURE", row.get("Task ID", ""), row.get("Created At", ""), row.get("Last Sent At", ""), row.get("Mailbox", ""), row.get("Reply Status", ""), new_notes]]
            )
            write_run_log(run_log_ws, "WARNING", "NAME_UNSURE", f"Row {sheet_row}: {website} -> {name_info['name']}")
            continue

        template = ai_format_template(select_template(templates, index))
        template_id = template["id"]
        subject = template["subject"].replace("{website_name}", name_info["name"])
        body = template["body"].replace("{website_name}", name_info["name"])

        history_rows = get_rows(history_ws)
        eligible_mailboxes = [
            mb for mb in get_ordered_candidate_mailboxes(mailboxes, history_rows)
            if daily_sent_count(history_rows, mb["email"], rows) < get_mailbox_daily_limit(mb, history_rows)
            and is_mailbox_cooled_down(history_rows, mb["email"])[0]
        ]
        if not eligible_mailboxes:
            write_run_log(run_log_ws, "INFO", "NO_ELIGIBLE_MAILBOX", "All mailboxes are at their daily limit or in cooldown.")
            return

        last_exc = None
        sent_successfully = False
        for attempt_num, mailbox in enumerate(eligible_mailboxes, start=1):
            mailbox_email = normalize_email(mailbox["email"])
            # Persist the pre-send marker before SMTP. If this fails, do not send.
            append_history(history_ws, recipient, website, mailbox_email, "SENDING", details=f"Attempt {attempt_num}", template_id=template_id)
            sheets_retry(
                websites_ws.update,
                range_name=f"C{sheet_row}:G{sheet_row}",
                values=[["SENDING", row.get("Task ID", ""), row.get("Created At", ""), "", mailbox_email]]
            )
            try:
                message_id = send_email(mailbox, recipient, subject, body, mailboxes)
            except SMTPDeliveryUncertain as exc:
                # Do not fail over: the previous SMTP server may already have accepted it.
                logger.error("SMTP outcome uncertain via %s for %s: %s", mailbox_email, recipient, exc)
                try:
                    append_history(history_ws, recipient, website, mailbox_email, "SEND_REVIEW", details=f"SMTP outcome uncertain: {exc}", template_id=template_id)
                except Exception:
                    logger.exception("Could not write SEND_REVIEW history for %s", recipient)
                try:
                    sheets_retry(websites_ws.update, range_name=f"C{sheet_row}:G{sheet_row}", values=[["SEND_REVIEW", row.get("Task ID", ""), row.get("Created At", ""), row.get("Last Sent At", ""), mailbox_email]])
                except Exception:
                    logger.exception("Could not mark row SEND_REVIEW for %s", recipient)
                write_run_log(run_log_ws, "WARNING", "SMTP_OUTCOME_UNCERTAIN", f"Recipient={recipient} | Mailbox={mailbox_email} | verify delivery before retrying")
                return
            except Exception as exc:
                last_exc = exc
                logger.warning("SMTP send failed via %s: %s", mailbox_email, exc)
                try:
                    append_history(history_ws, recipient, website, mailbox_email, "FAILED", details=f"Attempt {attempt_num}: {type(exc).__name__}: {exc}", template_id=template_id)
                except Exception:
                    logger.exception("Could not write FAILED history for %s", mailbox_email)
                    write_run_log(run_log_ws, "ERROR", "HISTORY_WRITE_FAILED", f"FAILED row not saved for {recipient} via {mailbox_email}")
                if attempt_num < len(eligible_mailboxes):
                    time.sleep(2)
                continue

            # SMTP accepted the message. From here onward, NEVER try another mailbox.
            sent_successfully = True
            try:
                append_history(history_ws, recipient, website, mailbox_email, "SENT", message_id, "SMTP accepted", template_id=template_id)
            except Exception as exc:
                logger.exception("SMTP accepted email but SENT history write failed for %s", recipient)
                write_run_log(run_log_ws, "ERROR", "SENT_HISTORY_WRITE_FAILED", f"Recipient={recipient} | Mailbox={mailbox_email} | Message-ID={message_id} | {type(exc).__name__}: {exc}")

            try:
                sheets_retry(
                    websites_ws.update,
                    range_name=f"C{sheet_row}:G{sheet_row}",
                    values=[["SENT", row.get("Task ID", ""), row.get("Created At", ""), timestamp(), mailbox_email]]
                )
            except Exception as exc:
                logger.exception("SMTP accepted email but Websites status update failed for %s", recipient)
                write_run_log(run_log_ws, "ERROR", "SENT_STATUS_UPDATE_FAILED", f"Recipient={recipient} | Mailbox={mailbox_email} | Message-ID={message_id} | {type(exc).__name__}: {exc}")
            contacted.add(recipient)
            break

        if sent_successfully:
            return

        # All SMTP attempts failed. Record the final outcome without masking the error.
        try:
            sheets_retry(
                websites_ws.update,
                range_name=f"C{sheet_row}:G{sheet_row}",
                values=[["FAILED", row.get("Task ID", ""), row.get("Created At", ""), row.get("Last Sent At", ""), eligible_mailboxes[-1]["email"]]]
            )
        except Exception:
            logger.exception("Could not update Websites row to FAILED for %s", recipient)
        write_run_log(run_log_ws, "ERROR", "SEND_FAILED", f"{recipient}: {type(last_exc).__name__ if last_exc else 'UnknownError'}: {last_exc}")
        return



# =========================================================
# DASHBOARD
# =========================================================

def update_dashboard(dashboard_ws, history_ws, websites_ws, mailboxes):
    history = get_rows(history_ws)
    websites = get_rows(websites_ws)
    counts = {"SENT": 0, "FAILED": 0, "SENDING": 0}
    today = now_local().date().isoformat()
    for row in history:
        if str(row.get("Date", "")) == today:
            status = str(row.get("Status", "")).upper()
            if status in counts:
                counts[status] += 1
    pending = sum(1 for row in websites if str(row.get("Status", "")).strip().upper() == "PENDING")
    name_unsure = sum(1 for row in websites if str(row.get("Status", "")).strip().upper() == "NAME_UNSURE")
    send_review = sum(1 for row in websites if str(row.get("Status", "")).strip().upper() == "SEND_REVIEW")

    dashboard_data = [
        ["Metric", "Value"],
        ["Eastern Date", today],
        ["Active Mailboxes", len(mailboxes)],
        ["Max Daily Limit Per Mailbox", DAILY_LIMIT_PER_MAILBOX],
        ["Today's Sent", counts["SENT"]],
        ["Today's Failed", counts["FAILED"]],
        ["Currently Sending", counts["SENDING"]],
        ["Pending Websites", pending],
        ["Name Unsure Review", name_unsure],
        ["Send Review Required", send_review]
    ]
    sheets_retry(dashboard_ws.clear)
    sheets_retry(dashboard_ws.update, range_name=f"A1:B{len(dashboard_data)}", values=dashboard_data)


# =========================================================
# MAIN
# =========================================================

def main():
    logger.info("Email automation started.")
    spreadsheet = connect_google_sheets()
    setup_tabs(spreadsheet)

    websites_ws = get_worksheet(spreadsheet, "Websites")
    templates_ws = get_worksheet(spreadsheet, "Templates")
    history_ws = get_worksheet(spreadsheet, "History")
    archive_ws = get_worksheet(spreadsheet, "Archive")
    replies_ws = get_worksheet(spreadsheet, "Replies")
    dashboard_ws = get_worksheet(spreadsheet, "Dashboard")
    suppression_ws = get_worksheet(spreadsheet, "Suppression")
    run_log_ws = get_worksheet(spreadsheet, "Run Log")
    global _ACTIVE_RUN_LOG_WS
    _ACTIVE_RUN_LOG_WS = run_log_ws

    write_run_log(run_log_ws, "INFO", "RUN_STARTED", f"Spreadsheet: {spreadsheet.title}")
    mailboxes = load_mailboxes()

    repair_stuck_sending_rows(websites_ws, history_ws, run_log_ws)
    archive_old_logs(spreadsheet, history_ws, run_log_ws, retention_days=30)
    discover_new_tasks(websites_ws, run_log_ws)
    read_replies(mailboxes, replies_ws, suppression_ws, websites_ws, history_ws, archive_ws, run_log_ws)

    can_send = True
    if not HAMZA_SIGNATURE.strip():
        write_run_log(run_log_ws, "ERROR", "SIGNATURE_MISSING", "HAMZA_SIGNATURE is empty. Sending skipped.")
        can_send = False
    if len(mailboxes) > 1 and not ARSLAN_SIGNATURE.strip():
        write_run_log(run_log_ws, "ERROR", "SIGNATURE_MISSING", "ARSLAN_SIGNATURE is empty. Sending skipped.")
        can_send = False
    if len(mailboxes) > 2 and not JAMES_SIGNATURE.strip():
        write_run_log(run_log_ws, "ERROR", "SIGNATURE_MISSING", "JAMES_SIGNATURE is empty. Sending skipped for the third mailbox configuration.")
        can_send = False
    if len(mailboxes) > 3:
        write_run_log(run_log_ws, "ERROR", "UNSUPPORTED_MAILBOX_COUNT", "This signature mapping supports a maximum of three mailboxes.")
        can_send = False

    if can_send and not within_send_window():
        write_run_log(run_log_ws, "INFO", "OUTSIDE_SEND_WINDOW", f"Eastern hour={now_local().hour}. Sending skipped.")
        can_send = False

    if can_send:
        process_sends(spreadsheet, mailboxes, websites_ws, templates_ws, history_ws, suppression_ws, run_log_ws)

    update_dashboard(dashboard_ws, history_ws, websites_ws, mailboxes)
    logger.info("Email automation finished.")


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        logger.exception("FATAL: email automation stopped: %s", exc)
        raise
