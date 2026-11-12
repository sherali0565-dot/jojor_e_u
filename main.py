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
import requests

from datetime import datetime
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
    ]
}


def column_letter(number):
    letters = ""
    while number > 0:
        number, remainder = divmod(number - 1, 26)
        letters = chr(65 + remainder) + letters
    return letters


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
            ws.update(
                range_name=f"A1:{end_column}1",
                values=[headers]
            )
            logger.info("Created worksheet: %s", name)
            continue

        ws = existing[name]
        current = ws.row_values(1)

        if not current:
            ws.update(
                range_name=f"A1:{end_column}1",
                values=[headers]
            )
            continue

        # Old History tab without the Template column:
        # add the new header automatically.
        if (
            name == "History"
            and current[:8] == headers[:8]
            and len(current) == 8
        ):
            ws.update_cell(1, 9, "Template")
            continue

        if current[:len(headers)] != headers:
            raise RuntimeError(
                f"Header mismatch in worksheet: {name}. "
                f"Expected: {headers}. Found: {current}"
            )


def get_worksheet(spreadsheet, name):
    return spreadsheet.worksheet(name)


def get_rows(worksheet):
    return worksheet.get_all_records()


def timestamp():
    return datetime.now(EASTERN_TZ).strftime(
        "%Y-%m-%d %H:%M:%S %z"
    )


def now_local():
    return datetime.now(EASTERN_TZ)


def write_run_log(ws, level, event, details=""):
    ws.append_row([
        timestamp(),
        level,
        event,
        str(details)[:4000]
    ])


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
