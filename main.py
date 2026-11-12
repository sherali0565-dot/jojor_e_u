import os
import re
import json
import time
import ssl
import smtplib
import imaplib
import email
import hashlib
import logging
import html
import random
import requests

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
HAMZA_SIGNATURE = os.getenv("HAMZA_SIGNATURE", "").strip()

GEMINI_API_KEY = os.getenv("GEMINI_API_KEY", "").strip()
OPENROUTER_API_KEY = os.getenv("OPENROUTER_API_KEY", "").strip()
OLLAMA_API_KEY = os.getenv("OLLAMA_API_KEY", "").strip()

# "auto" means the code chooses a medium-tier model by itself.
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

# Sending window (Eastern time). This is only a check,
# not a timer. The 5 minute rhythm comes from the YAML.
SEND_START_HOUR = 8
SEND_END_HOUR = 16
PAUSE_START_HOUR = 13
PAUSE_END_HOUR = 14

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s"
)

logger = logging.getLogger(__name__)

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
    """Retry Google Sheets API operations up to 3 times on transient failures."""
    delays = (2, 5, 10)
    last_exc = None
    for attempt in range(len(delays) + 1):
        try:
            return func(*args, **kwargs)
        except Exception as exc:
            last_exc = exc
            if attempt < len(delays):
                logger.warning(
                    "Google Sheets write/read failed (attempt %d/%d): %s; retrying in %ds",
                    attempt + 1, len(delays) + 1, type(exc).__name__, delays[attempt]
                )
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
            logger.info("Created worksheet: %s", name)
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

        # Old History/Archive tab without the Template column:
        # add the new header automatically.
        if (
            name in ("History", "Archive")
            and current[:8] == headers[:8]
            and len(current) == 8
        ):
            sheets_retry(ws.update_cell, 1, 9, "Template")
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
        logger.exception("Failed to write Run Log after retries.")


def pending_task_count(rows):
    return sum(
        1 for row in rows
        if str(row.get("Status", "")).strip().upper() == "PENDING"
    )


# =========================================================
# GENERAL HELPERS
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
            result.append(
                part.decode(
                    encoding or "utf-8",
                    errors="replace"
                )
            )
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


# =========================================================
# AI PROVIDERS (GEMINI -> OPENROUTER -> OLLAMA FALLBACK)
# =========================================================

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
    except Exception as exc:
        logger.warning("Gemini model auto-discovery failed: %s", type(exc).__name__)

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
    except Exception as exc:
        logger.warning("Ollama model auto-discovery failed: %s", type(exc).__name__)

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
                logger.warning(
                    "AI task=%s provider=%s attempt=%d failed: %s",
                    task_name, provider_name, attempt + 1, type(exc).__name__
                )
                if attempt < len(AI_RETRY_DELAYS):
                    time.sleep(AI_RETRY_DELAYS[attempt])

    raise RuntimeError(f"All AI providers failed for {task_name}") from last_err


def ai_website_name(website):
    """Return {'name': str, 'sure': bool}. Marks unsure when domain words cannot be reliably split."""
    fallback = fallback_clean_website_name(website)
    if not website:
        return {"name": fallback, "sure": False}

    prompt = (
        "You format website domains as the human-readable name of the "
        "person, publisher, or brand being contacted for a professional email greeting.\n"
        "RULES:\n"
        "1. Remove URL parts (https://, www., paths, domain suffixes).\n"
        "2. Split joined words into separate words when clear (e.g., jameswilliams.com -> James Williams).\n"
        "3. Use proper title case.\n"
        "4. Set 'sure' to false if the domain is ambiguous, cryptic, random characters, "
        "or you are not confident how words split. Otherwise set 'sure' to true.\n"
        "5. Return ONLY valid JSON: {\"website_name\": \"...\", \"sure\": true}\n"
        f"Website domain: {website}"
    )

    try:
        data = ai_generate_json(prompt, "website_name")
        name = str(data.get("website_name", "")).strip()
        sure = bool(data.get("sure", True))
        if (
            not name
            or len(name) > 80
            or not re.fullmatch(r"[A-Za-z0-9]+(?:[ '-][A-Za-z0-9]+)*", name)
        ):
            return {"name": fallback, "sure": False}
        return {"name": name, "sure": sure}
    except Exception:
        # If AI is down and domain has no clear separators/CamelCase, mark unsure
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
# MAILBOX CONFIGURATION & WARMUP LIMITS
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
        mailbox.setdefault("imap_host", mailbox["smtp_host"])
        mailbox.setdefault("imap_port", 993)

        active_mailboxes.append(mailbox)

    if not active_mailboxes:
        raise RuntimeError("No active mailboxes configured.")

    addresses = [m["email"] for m in active_mailboxes]
    if len(addresses) != len(set(addresses)):
        raise RuntimeError("Duplicate mailbox addresses found.")

    return active_mailboxes


def get_mailbox_daily_limit(mailbox, history_rows):
    """Gradually ramp daily limit for new mailboxes based on days since first send or created_at."""
    max_limit = int(mailbox.get("daily_limit", DAILY_LIMIT_PER_MAILBOX))
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

    if first_date is None:
        days_active = 0
    else:
        days_active = max(0, (now_local().date() - first_date).days)

    # Gradual ramp-up schedule:
    # Days 0-2: 2/day | Days 3-6: 4/day | Days 7-13: 6/day | Days 14-20: 8/day | Day 21+: full limit (10)
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
    """Pause a mailbox for cooldown_hours if its last 3 attempts were consecutive FAILED."""
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
        elapsed = now_local() - latest_fail_dt
        if elapsed < timedelta(hours=cooldown_hours):
            return False, consecutive_failures, latest_fail_dt

    return True, consecutive_failures, latest_fail_dt


def get_ordered_candidate_mailboxes(mailboxes, history_rows):
    """Return mailboxes rotated starting after the last SENT mailbox."""
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


def daily_sent_count(history_rows, mailbox_email):
    today = now_local().date().isoformat()
    mb_email = normalize_email(mailbox_email)
    return sum(
        1 for row in history_rows
        if str(row.get("Date", "")).strip() == today
        and normalize_email(row.get("Mailbox", "")) == mb_email
        and str(row.get("Status", "")).strip().upper() == "SENT"
    )


# =========================================================
# STUCK ROW REPAIR, TASK DISCOVERY & ARCHIVING
# =========================================================

def repair_stuck_sending_rows(websites_ws, history_ws, run_log_ws):
    """Recover rows stuck in SENDING from a crashed run by checking History."""
    values = sheets_retry(websites_ws.get_all_values)
    if not values or len(values) < 2:
        return 0

    headers = [str(v).strip() for v in values[0]]
    h_map = {name: idx for idx, name in enumerate(headers)}
    required = ("Email", "Status", "Task ID", "Created At", "Last Sent At", "Mailbox")
    if any(col not in h_map for col in required):
        return 0

    history_rows = get_rows(history_ws)
    sent_by_email = {}
    for h in history_rows:
        em = normalize_email(h.get("Email", ""))
        st = str(h.get("Status", "")).strip().upper()
        if em and st == "SENT":
            sent_by_email[em] = {
                "timestamp": str(h.get("Timestamp", "")).strip(),
                "mailbox": normalize_email(h.get("Mailbox", ""))
            }

    repaired = 0
    for sheet_row, row in enumerate(values[1:], start=2):
        def get_col(col):
            idx = h_map[col]
            return str(row[idx]).strip() if idx < len(row) and row[idx] is not None else ""

        if get_col("Status").upper() != "SENDING":
            continue

        recipient = normalize_email(get_col("Email"))
        if recipient in sent_by_email:
            info = sent_by_email[recipient]
            sheets_retry(
                websites_ws.update,
                range_name=f"C{sheet_row}:G{sheet_row}",
                values=[[
                    "SENT",
                    get_col("Task ID"),
                    get_col("Created At"),
                    info["timestamp"] or timestamp(),
                    info["mailbox"] or get_col("Mailbox")
                ]]
            )
            logger.info("Repaired stuck SENDING row %d (%s) -> SENT", sheet_row, recipient)
        else:
            sheets_retry(
                websites_ws.update,
                range_name=f"C{sheet_row}:G{sheet_row}",
                values=[[
                    "PENDING",
                    get_col("Task ID"),
                    get_col("Created At"),
                    get_col("Last Sent At"),
                    ""
                ]]
            )
            logger.info("Repaired stuck SENDING row %d (%s) -> PENDING", sheet_row, recipient)
        repaired += 1

    if repaired:
        write_run_log(run_log_ws, "WARNING", "STUCK_SENDING_REPAIRED", f"Repaired rows={repaired}")
    return repaired


def archive_old_logs(spreadsheet, history_ws, run_log_ws, retention_days=30):
    """Move History and Run Log rows older than retention_days into Archive tabs."""
    cutoff = now_local().date() - timedelta(days=retention_days)

    # 1. Archive History
    try:
        archive_ws = get_worksheet(spreadsheet, "Archive")
        all_hist = sheets_retry(history_ws.get_all_values)
        if len(all_hist) > 1:
            headers = all_hist[0]
            keep_rows = []
            move_rows = []
            for row in all_hist[1:]:
                date_str = str(row[1]).strip() if len(row) > 1 else ""
                try:
                    row_date = datetime.strptime(date_str, "%Y-%m-%d").date()
                    if row_date < cutoff:
                        padded = row + [""] * max(0, 9 - len(row))
                        move_rows.append(padded[:9])
                    else:
                        keep_rows.append(row)
                except ValueError:
                    keep_rows.append(row)

            if move_rows:
                sheets_retry(archive_ws.append_rows, move_rows)
                sheets_retry(history_ws.clear)
                end_col = column_letter(len(headers))
                sheets_retry(history_ws.update, range_name=f"A1:{end_col}1", values=[headers])
                if keep_rows:
                    sheets_retry(history_ws.append_rows, keep_rows)
                logger.info("Archived %d old History row(s).", len(move_rows))
    except Exception:
        logger.exception("Failed archiving old History rows.")

    # 2. Archive Run Log
    try:
        run_archive_ws = get_worksheet(spreadsheet, "Run Log Archive")
        all_logs = sheets_retry(run_log_ws.get_all_values)
        if len(all_logs) > 1:
            headers = all_logs[0]
            keep_logs = []
            move_logs = []
            for row in all_logs[1:]:
                ts_str = str(row[0]).strip() if len(row) > 0 else ""
                date_part = ts_str[:10]
                try:
                    row_date = datetime.strptime(date_part, "%Y-%m-%d").date()
                    if row_date < cutoff:
                        padded = row + [""] * max(0, 4 - len(row))
                        move_logs.append(padded[:4])
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
                logger.info("Archived %d old Run Log row(s).", len(move_logs))
    except Exception:
        logger.exception("Failed archiving old Run Log rows.")


def discover_new_tasks(websites_ws, run_log_ws=None):
    values = sheets_retry(websites_ws.get_all_values)
    if not values:
        logger.warning("Websites tab is empty; no header row found.")
        if run_log_ws:
            write_run_log(run_log_ws, "WARNING", "DISCOVERY_STOP", "Websites tab is empty")
        return 0

    headers = [str(value).strip() for value in values[0]]
    header_map = {name: idx for idx, name in enumerate(headers)}
    required = ["Website", "Email", "Status", "Task ID", "Created At"]
    missing = [name for name in required if name not in header_map]
    if missing:
        message = "Missing Websites columns: " + ", ".join(missing)
        logger.error(message)
        if run_log_ws:
            write_run_log(run_log_ws, "ERROR", "DISCOVERY_HEADER_ERROR", message)
        return 0

    def cell(row, column):
        idx = header_map[column]
        return str(row[idx]).strip() if idx < len(row) and row[idx] is not None else ""

    created_count = 0
    skipped = []
    for sheet_row, row in enumerate(values[1:], start=2):
        website = cell(row, "Website")
        raw_email = cell(row, "Email")
        recipient = normalize_email(raw_email)
        status = cell(row, "Status")

        if not website and not raw_email and not status:
            continue
        if not website:
            skipped.append(f"row {sheet_row}: Website is blank")
            continue
        if not recipient:
            skipped.append(f"row {sheet_row}: Email is blank")
            continue
        if not valid_email(recipient):
            skipped.append(f"row {sheet_row}: Email format is invalid")
            continue
        if status:
            skipped.append(f"row {sheet_row}: Status already '{status}'")
            continue

        task_id = hashlib.sha256(
            f"{website}|{recipient}".encode("utf-8")
        ).hexdigest()[:16]

        sheets_retry(
            websites_ws.update,
            range_name=f"C{sheet_row}:E{sheet_row}",
            values=[["PENDING", task_id, timestamp()]]
        )
        created_count += 1
        logger.info("Created PENDING task for Websites row %d (%s).", sheet_row, recipient)

    summary = f"Created PENDING={created_count}; " + (" | ".join(skipped[:12]) if skipped else "No skipped populated rows")
    logger.info("Task discovery: %s", summary)
    if run_log_ws:
        write_run_log(run_log_ws, "INFO", "TASK_DISCOVERY", summary)
    return created_count


# =========================================================
# TEMPLATES, SUPPRESSION & BOUNCE GUARD
# =========================================================

def get_active_templates(templates_ws):
    rows = get_rows(templates_ws)
    active = []
    for row in rows:
        template_id = str(row.get("Template ID", "")).strip()
        subject = str(row.get("Subject", "")).strip()
        body = str(row.get("Body", "")).strip()
        active_value = str(row.get("Active", "")).strip().lower()

        if active_value not in ("true", "yes", "1", "active"):
            continue
        if not template_id or not subject or not body:
            continue

        active.append({
            "id": template_id,
            "subject": subject,
            "body": body
        })
    return active


def select_template(templates, task_number):
    if not templates:
        return None
    return templates[task_number % len(templates)]


def get_suppressed_emails(suppression_ws):
    rows = get_rows(suppression_ws)
    return {
        normalize_email(row.get("Email", ""))
        for row in rows
        if row.get("Email")
    }


def get_contacted_emails(history_rows):
    contacted = set()
    for row in history_rows:
        if str(row.get("Status", "")).strip().upper() == "SENT":
            recipient = normalize_email(row.get("Email", ""))
            if recipient:
                contacted.add(recipient)
    return contacted


def add_to_suppression(suppression_ws, address, reason):
    address = normalize_email(address)
    if not address:
        return
    existing = get_suppressed_emails(suppression_ws)
    if address not in existing:
        sheets_retry(
            suppression_ws.append_row,
            [address, reason, timestamp()]
        )


def is_bounce_rate_too_high(history_rows, suppression_ws):
    """Halt sending if recent bounce rate exceeds threshold."""
    recent_sends = [
        r for r in history_rows
        if str(r.get("Status", "")).strip().upper() == "SENT"
    ][-50:]

    if len(recent_sends) < 10:
        return False, 0.0

    suppression_rows = get_rows(suppression_ws)
    bounced_emails = {
        normalize_email(r.get("Email", ""))
        for r in suppression_rows
        if str(r.get("Reason", "")).strip().upper() == "BOUNCE"
    }

    recent_bounces = sum(
        1 for r in recent_sends
        if normalize_email(r.get("Email", "")) in bounced_emails
    )
    bounce_rate = recent_bounces / len(recent_sends)
    return bounce_rate > 0.15, bounce_rate


def append_history(
    history_ws,
    recipient,
    website,
    mailbox,
    status,
    message_id="",
    details="",
    template_id=""
):
    sheets_retry(
        history_ws.append_row,
        [
            timestamp(),
            now_local().date().isoformat(),
            recipient,
            website,
            mailbox,
            status,
            message_id,
            details,
            template_id
        ]
    )


# =========================================================
# SMTP SENDING (NO UNSUBSCRIBE HEADER)
# =========================================================

def get_signature_html():
    if not HAMZA_SIGNATURE:
        raise RuntimeError(
            "HAMZA_SIGNATURE GitHub Secret is missing. Email was not sent."
        )

    try:
        parsed = json.loads(HAMZA_SIGNATURE)
        if isinstance(parsed, dict):
            signature = str(parsed.get("html", "")).strip()
        elif isinstance(parsed, str):
            signature = parsed.strip()
        else:
            signature = HAMZA_SIGNATURE
    except json.JSONDecodeError:
        signature = HAMZA_SIGNATURE

    if not signature:
        raise RuntimeError("HAMZA_SIGNATURE is empty. Email was not sent.")

    return signature


def remove_body_signoff(body):
    normalized = (body or "").replace("\r\n", "\n").replace("\r", "\n")
    lines = normalized.split("\n")
    for index, line in enumerate(lines):
        if re.fullmatch(r"\s*best regards\s*,?\s*", line, flags=re.IGNORECASE):
            return "\n".join(lines[:index]).rstrip()
    return normalized.strip()


def html_signature_to_text(signature_html):
    text_value = re.sub(r"(?i)<\s*br\s*/?\s*>", "\n", signature_html)
    text_value = re.sub(r"(?i)</\s*(p|div|tr|li|h[1-6])\s*>", "\n", text_value)
    text_value = re.sub(r"<[^>]+>", "", text_value)
    text_value = html.unescape(text_value)
    text_value = re.sub(r"[ \t]+", " ", text_value)
    text_value = re.sub(r"\n[ \t]*\n[ \t]*\n+", "\n\n", text_value)
    return text_value.strip()


def send_email(mailbox, recipient, subject, body):
    signature_html = get_signature_html()
    clean_body = remove_body_signoff(body)
    signature_text = html_signature_to_text(signature_html)

    message = EmailMessage()
    message["From"] = formataddr((mailbox.get("name", "Western Arsal"), mailbox["email"]))
    message["To"] = recipient
    message["Subject"] = subject
    message["Message-ID"] = make_msgid()

    plain_content = clean_body.rstrip()
    if signature_text:
        plain_content += "\n\n" + signature_text
    message.set_content(plain_content)

    normalized_body = clean_body.replace("\r\n", "\n").replace("\r", "\n").strip()
    paragraphs = re.split(r"\n[ \t]*\n+", normalized_body) if normalized_body else []
    body_blocks = [
        '<p style="margin:0 0 16px 0;">' + html.escape(p).replace("\n", "<br>\n") + "</p>"
        for p in paragraphs
    ]

    body_html = "".join(body_blocks)
    signature_block = (
        '<div style="margin-top:24px; padding-top:0;">'
        + signature_html
        + "</div>"
    )
    html_content = (
        "<!doctype html><html><body>"
        '<div style="font-family:Arial,Helvetica,sans-serif;'
        'font-size:14px;line-height:1.6;color:#333333;'
        'text-align:left;margin:0;padding:0;">'
        f"{body_html}{signature_block}"
        "</div></body></html>"
    )
    message.add_alternative(html_content, subtype="html")

    context = ssl.create_default_context()
    smtp_host = mailbox["smtp_host"]
    smtp_port = int(mailbox["smtp_port"])

    if smtp_port == 465:
        with smtplib.SMTP_SSL(smtp_host, smtp_port, context=context, timeout=30) as server:
            server.login(mailbox["email"], mailbox["password"])
            server.send_message(message)
    else:
        with smtplib.SMTP(smtp_host, smtp_port, timeout=30) as server:
            server.ehlo()
            server.starttls(context=context)
            server.ehlo()
            server.login(mailbox["email"], mailbox["password"])
            server.send_message(message)

    return message["Message-ID"]


# =========================================================
# PROCESS OUTGOING EMAILS
# =========================================================

def is_send_time():
    current = now_local()
    if current.weekday() >= 5:
        return False
    if not SEND_START_HOUR <= current.hour < SEND_END_HOUR:
        return False
    if PAUSE_START_HOUR <= current.hour < PAUSE_END_HOUR:
        return False
    return True


def process_sends(
    spreadsheet,
    mailboxes,
    websites_ws,
    templates_ws,
    history_ws,
    suppression_ws,
    run_log_ws
):
    history_rows = get_rows(history_ws)

    # 1. Bounce guard
    too_many_bounces, bounce_rate = is_bounce_rate_too_high(history_rows, suppression_ws)
    if too_many_bounces:
        msg = f"Sending paused: recent bounce rate is {bounce_rate:.1%} (>15%)."
        logger.warning(msg)
        write_run_log(run_log_ws, "WARNING", "BOUNCE_GUARD_STOP", msg)
        return

    # 2. Light randomness: skip ~12% of runs naturally
    if random.random() < 0.12:
        logger.info("Random jitter: skipping send on this run for natural cadence.")
        write_run_log(run_log_ws, "INFO", "JITTER_SKIP", "Skipped send cycle randomly")
        return

    templates = get_active_templates(templates_ws)
    if not templates:
        logger.warning("No active templates found.")
        return

    rows = get_rows(websites_ws)
    suppressed = get_suppressed_emails(suppression_ws)
    contacted = get_contacted_emails(history_rows)

    # 3. Filter eligible mailboxes (respecting warmup limits and 3-failure cooldown)
    ordered_mailboxes = get_ordered_candidate_mailboxes(mailboxes, history_rows)
    eligible_mailboxes = []
    for mb in ordered_mailboxes:
        mb_email = mb["email"]
        cooled_down, fail_count, _ = is_mailbox_cooled_down(history_rows, mb_email, cooldown_hours=4)
        if not cooled_down:
            logger.info("Mailbox %s is in cooldown (%d consecutive failures).", mb_email, fail_count)
            continue

        limit = get_mailbox_daily_limit(mb, history_rows)
        sent_today = daily_sent_count(history_rows, mb_email)
        if sent_today >= limit:
            logger.info("Mailbox %s reached its daily limit (%d/%d).", mb_email, sent_today, limit)
            continue

        eligible_mailboxes.append(mb)

    if not eligible_mailboxes:
        logger.info("No eligible mailboxes available right now (daily limits or cooldown active).")
        write_run_log(run_log_ws, "INFO", "NO_ELIGIBLE_MAILBOX", "All mailboxes at daily limit or in cooldown")
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

        if recipient in contacted:
            sheets_retry(websites_ws.update_cell, sheet_row, 3, "ALREADY_CONTACTED")
            continue

        # 4. Website name confidence check
        name_info = ai_website_name(website)
        if not name_info["sure"]:
            existing_notes = str(row.get("Notes", "")).strip()
            note_addition = f"Suggested: {name_info['name']}"
            new_notes = f"{existing_notes} | {note_addition}".strip(" |") if existing_notes else note_addition
            sheets_retry(
                websites_ws.update,
                range_name=f"C{sheet_row}:I{sheet_row}",
                values=[[
                    "NAME_UNSURE",
                    row.get("Task ID", ""),
                    row.get("Created At", ""),
                    row.get("Last Sent At", ""),
                    row.get("Mailbox", ""),
                    row.get("Reply Status", ""),
                    new_notes
                ]]
            )
            logger.warning("Marked row %d (%s) as NAME_UNSURE (candidate='%s').", sheet_row, website, name_info["name"])
            write_run_log(run_log_ws, "WARNING", "NAME_UNSURE", f"Row {sheet_row}: {website} -> {name_info['name']}")
            continue

        website_name = name_info["name"]
        raw_template = select_template(templates, index)
        template = ai_format_template(raw_template)
        template_id = template["id"]

        subject = template["subject"].replace("{website_name}", website_name)
        body = template["body"].replace("{website_name}", website_name)

        # 5. Try up to 3 mailboxes in the same run on failure
        attempts_to_try = eligible_mailboxes[:3]
        last_exc = None
        last_mailbox_email = ""

        for attempt_num, mailbox in enumerate(attempts_to_try, start=1):
            last_mailbox_email = mailbox["email"]

            append_history(
                history_ws,
                recipient,
                website,
                mailbox["email"],
                "SENDING",
                details=f"Attempt {attempt_num}/{len(attempts_to_try)}",
                template_id=template_id
            )

            sheets_retry(
                websites_ws.update,
                range_name=f"C{sheet_row}:G{sheet_row}",
                values=[[
                    "SENDING",
                    row.get("Task ID", ""),
                    row.get("Created At", ""),
                    "",
                    mailbox["email"]
                ]]
            )

            try:
                message_id = send_email(mailbox, recipient, subject, body)

                append_history(
                    history_ws,
                    recipient,
                    website,
                    mailbox["email"],
                    "SENT",
                    message_id,
                    "Email sent successfully",
                    template_id=template_id
                )

                sheets_retry(
                    websites_ws.update,
                    range_name=f"C{sheet_row}:G{sheet_row}",
                    values=[[
                        "SENT",
                        row.get("Task ID", ""),
                        row.get("Created At", ""),
                        timestamp(),
                        mailbox["email"]
                    ]]
                )

                logger.info("Email sent to %s via %s", recipient, mailbox["email"])
                return

            except Exception as exc:
                last_exc = exc
                logger.warning(
                    "Attempt %d/%d failed for %s using mailbox %s: %s",
                    attempt_num, len(attempts_to_try), recipient, mailbox["email"], type(exc).__name__
                )
                append_history(
                    history_ws,
                    recipient,
                    website,
                    mailbox["email"],
                    "FAILED",
                    details=f"Attempt {attempt_num}: {type(exc).__name__}",
                    template_id=template_id
                )
                if attempt_num < len(attempts_to_try):
                    time.sleep(2)

        # All attempts (up to 3 mailboxes) failed for this recipient
        logger.error("All %d send attempt(s) failed for %s.", len(attempts_to_try), recipient)
        sheets_retry(
            websites_ws.update,
            range_name=f"C{sheet_row}:G{sheet_row}",
            values=[[
                "FAILED",
                row.get("Task ID", ""),
                row.get("Created At", ""),
                row.get("Last Sent At", ""),
                last_mailbox_email
            ]]
        )
        return


# =========================================================
# IMAP REPLY CHECKING
# =========================================================

def classify_reply(subject, body):
    text = f"{subject} {body}".lower()

    if any(term in text for term in [
        "unsubscribe",
        "remove me",
        "do not contact",
        "stop emailing",
        "opt out"
    ]):
        return "UNSUBSCRIBE"

    if any(term in text for term in [
        "mailbox not found",
        "user unknown",
        "address rejected",
        "delivery failed",
        "undelivered mail"
    ]):
        return "BOUNCE"

    return "REPLY"


def extract_body(message):
    parts = []

    if message.is_multipart():
        for part in message.walk():
            if part.get_content_type() != "text/plain":
                continue
            if part.get_content_disposition() == "attachment":
                continue
            payload = part.get_payload(decode=True)
            if payload:
                charset = part.get_content_charset() or "utf-8"
                parts.append(payload.decode(charset, errors="replace"))
    else:
        payload = message.get_payload(decode=True)
        if payload:
            charset = message.get_content_charset() or "utf-8"
            parts.append(payload.decode(charset, errors="replace"))

    return "\n".join(parts).strip()


def read_replies(mailboxes, replies_ws, suppression_ws, websites_ws):
    existing_ids = {
        str(row.get("Message ID", "")).strip()
        for row in get_rows(replies_ws)
        if row.get("Message ID")
    }

    for mailbox in mailboxes:
        if not mailbox.get("imap_host"):
            continue

        try:
            connection = imaplib.IMAP4_SSL(
                mailbox["imap_host"],
                int(mailbox.get("imap_port", 993)),
                ssl_context=ssl.create_default_context()
            )
            connection.login(mailbox["email"], mailbox["password"])
            connection.select("INBOX")

            status, data = connection.search(None, "UNSEEN")
            if status != "OK":
                connection.logout()
                continue

            for number in data[0].split():
                status, message_data = connection.fetch(number, "(RFC822)")
                if status != "OK":
                    continue

                raw_email = message_data[0][1]
                message = email.message_from_bytes(raw_email)
                message_id = str(message.get("Message-ID", "")).strip()

                if message_id and message_id in existing_ids:
                    connection.store(number, "+FLAGS", "\\Seen")
                    continue

                from_address = normalize_email(parseaddr(message.get("From", ""))[1])
                subject = decode_text(message.get("Subject", ""))
                body = extract_body(message)
                classification = classify_reply(subject, body)
                snippet = body[:300].replace("\n", " ")

                sheets_retry(
                    replies_ws.append_row,
                    [
                        timestamp(),
                        from_address,
                        subject,
                        message_id,
                        classification,
                        snippet
                    ]
                )

                if message_id:
                    existing_ids.add(message_id)

                if classification in ("UNSUBSCRIBE", "BOUNCE"):
                    add_to_suppression(suppression_ws, from_address, classification)

                website_rows = get_rows(websites_ws)
                for index, website_row in enumerate(website_rows, start=2):
                    if normalize_email(website_row.get("Email", "")) == from_address:
                        sheets_retry(websites_ws.update_cell, index, 8, classification)

                connection.store(number, "+FLAGS", "\\Seen")

            connection.logout()

        except Exception:
            logger.exception("IMAP reply check failed for %s", mailbox["email"])


# =========================================================
# DASHBOARD
# =========================================================

def update_dashboard(dashboard_ws, history_ws, websites_ws, mailboxes):
    history = get_rows(history_ws)
    websites = get_rows(websites_ws)

    counts = {
        "SENT": 0,
        "FAILED": 0,
        "SENDING": 0
    }

    today = now_local().date().isoformat()

    for row in history:
        if str(row.get("Date", "")) != today:
            continue
        status = str(row.get("Status", "")).upper()
        if status in counts:
            counts[status] += 1

    pending = sum(
        1 for row in websites
        if str(row.get("Status", "")).strip().upper() == "PENDING"
    )
    name_unsure = sum(
        1 for row in websites
        if str(row.get("Status", "")).strip().upper() == "NAME_UNSURE"
    )

    dashboard_data = [
        ["Metric", "Value"],
        ["Eastern Date", today],
        ["Active Mailboxes", len(mailboxes)],
        ["Max Daily Limit Per Mailbox", DAILY_LIMIT_PER_MAILBOX],
        ["Today's Sent", counts["SENT"]],
        ["Today's Failed", counts["FAILED"]],
        ["Currently Sending", counts["SENDING"]],
        ["Pending Websites", pending],
        ["Name Unsure Review", name_unsure]
    ]

    sheets_retry(dashboard_ws.clear)
    sheets_retry(
        dashboard_ws.update,
        range_name="A1:B9",
        values=dashboard_data
    )


# =========================================================
# MAIN
# =========================================================

def main():
    logger.info("============================================================")
    logger.info("Email automation started.")
    logger.info("Automation timezone: %s", now_local().strftime("%Y-%m-%d %H:%M:%S %Z"))

    required = {
        "GOOGLE_SHEET_ID": GOOGLE_SHEET_ID,
        "GOOGLE_SERVICE_ACCOUNT_JSON": GOOGLE_SERVICE_ACCOUNT_JSON,
        "MAILBOXES_JSON": MAILBOXES_JSON,
    }
    missing = [key for key, value in required.items() if not value]
    if missing:
        raise RuntimeError("Missing required environment variables: " + ", ".join(missing))

    if not any((GEMINI_API_KEY, OPENROUTER_API_KEY, OLLAMA_API_KEY)):
        raise RuntimeError("At least one AI API key (GEMINI, OPENROUTER, or OLLAMA) must be set.")

    spreadsheet = connect_google_sheets()
    setup_tabs(spreadsheet)

    websites_ws = get_worksheet(spreadsheet, "Websites")
    templates_ws = get_worksheet(spreadsheet, "Templates")
    history_ws = get_worksheet(spreadsheet, "History")
    replies_ws = get_worksheet(spreadsheet, "Replies")
    dashboard_ws = get_worksheet(spreadsheet, "Dashboard")
    suppression_ws = get_worksheet(spreadsheet, "Suppression")
    run_log_ws = get_worksheet(spreadsheet, "Run Log")

    write_run_log(run_log_ws, "INFO", "RUN_STARTED", f"Spreadsheet: {spreadsheet.title}")

    mailboxes = load_mailboxes()
    write_run_log(run_log_ws, "INFO", "MAILBOXES_LOADED", f"Active mailboxes: {len(mailboxes)}")

    # 1. Repair rows stuck in SENDING from any interrupted run
    repair_stuck_sending_rows(websites_ws, history_ws, run_log_ws)

    # 2. Archive History & Run Log rows older than 30 days
    archive_old_logs(spreadsheet, history_ws, run_log_ws, retention_days=30)

    # 3. Discover new tasks
    created_tasks = discover_new_tasks(websites_ws, run_log_ws)
    write_run_log(run_log_ws, "INFO", "TASKS_CREATED", f"Automatically marked PENDING: {created_tasks}")

    website_rows = get_rows(websites_ws)
    pending_before = pending_task_count(website_rows)
    write_run_log(run_log_ws, "INFO", "WEBSITE_ROWS_SCANNED", f"Rows={len(website_rows)}; pending={pending_before}")

    # 4. Read replies & classify bounces/unsubscribes
    read_replies(mailboxes, replies_ws, suppression_ws, websites_ws)

    # 5. Process outgoing sends
    process_sends(
        spreadsheet,
        mailboxes,
        websites_ws,
        templates_ws,
        history_ws,
        suppression_ws,
        run_log_ws
    )

    update_dashboard(dashboard_ws, history_ws, websites_ws, mailboxes)
    pending_after = pending_task_count(get_rows(websites_ws))
    write_run_log(run_log_ws, "INFO", "RUN_FINISHED", f"Pending after run={pending_after}")
    logger.info("Email automation finished. Pending tasks remaining: %d", pending_after)


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        logger.exception("FATAL: email automation stopped: %s", exc)
        raise
