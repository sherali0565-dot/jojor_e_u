
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

from datetime import datetime, date
from email.message import EmailMessage
from email.header import decode_header
from email.utils import parseaddr

import gspread
from google.oauth2.service_account import Credentials


# =========================================================
# CONFIGURATION
# =========================================================

GOOGLE_SHEET_ID = os.environ["GOOGLE_SHEET_ID"]
GOOGLE_SERVICE_ACCOUNT_JSON = os.environ["GOOGLE_SERVICE_ACCOUNT_JSON"]
MAILBOXES_JSON = os.environ["MAILBOXES_JSON"]
HAMZA_SIGNATURE = os.getenv("HAMZA_SIGNATURE", "").strip()

SEND_INTERVAL_SECONDS = 12 * 60
DAILY_LIMIT_PER_MAILBOX = 10

SEND_START_HOUR = 8
SEND_END_HOUR = 16

PAUSE_START_HOUR = 13
PAUSE_END_HOUR = 14

DRY_RUN = os.getenv("DRY_RUN", "false").lower() == "true"

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s"
)

logger = logging.getLogger(__name__)


# =========================================================
# GOOGLE SHEETS
# =========================================================

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets"
]


def connect_google_sheets():
    service_info = json.loads(GOOGLE_SERVICE_ACCOUNT_JSON)

    credentials = Credentials.from_service_account_info(
        service_info,
        scopes=SCOPES
    )

    client = gspread.authorize(credentials)

    return client.open_by_key(GOOGLE_SHEET_ID)


HEADERS = {
    "Websites": [
        "Website",
        "Email",
        "Status",
        "Task ID",
        "Created At",
        "Last Sent At",
        "Mailbox",
        "Reply Status",
        "Notes"
    ],

    "Templates": [
        "Template ID",
        "Subject",
        "Body",
        "Active"
    ],

    "History": [
        "Timestamp",
        "Date",
        "Email",
        "Website",
        "Mailbox",
        "Status",
        "Message ID",
        "Details"
    ],

    "Replies": [
        "Received At",
        "From",
        "Subject",
        "Message ID",
        "Classification",
        "Snippet"
    ],

    "Dashboard": [
        "Metric",
        "Value"
    ],

    "Suppression": [
        "Email",
        "Reason",
        "Added At"
    ]
}


def setup_tabs(spreadsheet):
    """
    Create missing tabs and headers.
    Existing user data is not intentionally erased.
    """

    existing_sheets = {
        ws.title: ws
        for ws in spreadsheet.worksheets()
    }

    for tab_name, headers in HEADERS.items():

        if tab_name not in existing_sheets:
            worksheet = spreadsheet.add_worksheet(
                title=tab_name,
                rows=1000,
                cols=max(len(headers), 10)
            )

            worksheet.append_row(headers)
            continue

        worksheet = existing_sheets[tab_name]

        current_headers = worksheet.row_values(1)

        if not current_headers:
            worksheet.append_row(headers)

        elif current_headers != headers:
            raise RuntimeError(
                f"Header mismatch in '{tab_name}'. "
                "Please review the first row before running."
            )


def get_worksheet(spreadsheet, name):
    return spreadsheet.worksheet(name)


# =========================================================
# MAILBOX CONFIGURATION
# =========================================================

def load_mailboxes():
    """
    Expected MAILBOXES_JSON format:

    [
      {
        "name": "Mailbox 1",
        "email": "sender@example.com",
        "password": "APP_PASSWORD",
        "smtp_host": "mail.example.com",
        "smtp_port": 587,
        "imap_host": "mail.example.com",
        "imap_port": 993,
        "active": true
      }
    ]
    """

    raw_mailboxes = json.loads(MAILBOXES_JSON)

    if not isinstance(raw_mailboxes, list):
        raise ValueError("MAILBOXES_JSON must be a JSON list.")

    active_mailboxes = []

    for mailbox in raw_mailboxes:

        if not isinstance(mailbox, dict):
            raise ValueError("Each mailbox must be a JSON object.")

        if mailbox.get("active", True) is not True:
            continue

        required_fields = [
            "email",
            "password",
            "smtp_host",
            "smtp_port"
        ]

        for field in required_fields:
            if not mailbox.get(field):
                raise ValueError(
                    f"Missing '{field}' in an active mailbox."
                )

        mailbox["email"] = mailbox["email"].strip().lower()

        if not valid_email(mailbox["email"]):
            raise ValueError(
                f"Invalid mailbox email: {mailbox['email']}"
            )

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


# =========================================================
# GENERAL HELPERS
# =========================================================

def now_local():
    """Return the computer's local date and time."""
    return datetime.now()


def timestamp():
    return now_local().strftime("%Y-%m-%d %H:%M:%S")


def valid_email(address):
    if not address:
        return False

    address = address.strip()

    pattern = r"^[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}$"

    return bool(re.fullmatch(pattern, address))


def normalize_email(address):
    return (address or "").strip().lower()


def is_send_time():
    current = now_local()

    if current.weekday() >= 5:
        return False

    if not SEND_START_HOUR <= current.hour < SEND_END_HOUR:
        return False

    if PAUSE_START_HOUR <= current.hour < PAUSE_END_HOUR:
        return False

    return True


def decode_text(value):
    if not value:
        return ""

    decoded_parts = []

    for part, encoding in decode_header(value):
        if isinstance(part, bytes):
            decoded_parts.append(
                part.decode(encoding or "utf-8", errors="replace")
            )
        else:
            decoded_parts.append(part)

    return "".join(decoded_parts)


def get_rows(worksheet):
    return worksheet.get_all_records()


def append_history(
    history_ws,
    recipient,
    website,
    mailbox,
    status,
    message_id="",
    details=""
):
    history_ws.append_row([
        timestamp(),
        now_local().date().isoformat(),
        recipient,
        website,
        mailbox,
        status,
        message_id,
        details
    ])


# =========================================================
# WEBSITE TASKS
# =========================================================

def discover_new_tasks(websites_ws):
    """
    Website and Email are the only required manual fields.
    New rows are automatically marked PENDING.
    """

    rows = get_rows(websites_ws)

    for index, row in enumerate(rows, start=2):

        website = str(row.get("Website", "")).strip()
        recipient = normalize_email(row.get("Email", ""))
        status = str(row.get("Status", "")).strip()

        if not website or not recipient or status:
            continue

        task_id = hashlib.sha256(
            f"{website}|{recipient}".encode()
        ).hexdigest()[:16]

        websites_ws.update(
            range_name=f"C{index}:E{index}",
            values=[[
                "PENDING",
                task_id,
                timestamp()
            ]]
        )


# =========================================================
# TEMPLATES
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


# =========================================================
# SUPPRESSION AND DUPLICATE PROTECTION
# =========================================================

def get_suppressed_emails(suppression_ws):
    rows = get_rows(suppression_ws)

    return {
        normalize_email(row.get("Email", ""))
        for row in rows
        if row.get("Email")
    }


def get_contacted_emails(history_ws):
    rows = get_rows(history_ws)

    contacted = set()

    for row in rows:

        status = str(row.get("Status", "")).upper()

        if status in ("SENT", "SENDING"):
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
        suppression_ws.append_row([
            address,
            reason,
            timestamp()
        ])


# =========================================================
# MAILBOX ROTATION AND DAILY LIMITS
# =========================================================

def get_next_mailbox(mailboxes, history_ws):
    rows = get_rows(history_ws)

    last_mailbox = None

    for row in reversed(rows):
        if str(row.get("Status", "")).upper() == "SENT":
            last_mailbox = normalize_email(
                row.get("Mailbox", "")
            )
            break

    if not last_mailbox:
        return mailboxes[0]

    addresses = [
        normalize_email(m["email"])
        for m in mailboxes
    ]

    if last_mailbox not in addresses:
        return mailboxes[0]

    last_index = addresses.index(last_mailbox)

    return mailboxes[(last_index + 1) % len(mailboxes)]


def daily_sent_count(history_ws, mailbox_email):
    today = now_local().date().isoformat()
    mailbox_email = normalize_email(mailbox_email)

    rows = get_rows(history_ws)

    count = 0

    for row in rows:

        if str(row.get("Date", "")) != today:
            continue

        if normalize_email(row.get("Mailbox", "")) != mailbox_email:
            continue

        if str(row.get("Status", "")).upper() == "SENT":
            count += 1

    return count


def last_successful_send(history_ws):
    rows = get_rows(history_ws)

    for row in reversed(rows):

        if str(row.get("Status", "")).upper() != "SENT":
            continue

        try:
            return datetime.strptime(
                row["Timestamp"],
                "%Y-%m-%d %H:%M:%S"
            )

        except (KeyError, ValueError):
            continue

    return None


# =========================================================
# SMTP SENDING AND HTML SIGNATURE
# =========================================================

def get_signature_html():
    """
    Read the signature from the HAMZA_SIGNATURE GitHub Secret.

    Supports:
    1. Raw HTML stored directly in the Secret.
    2. JSON containing an "html" property.
    """

    if not HAMZA_SIGNATURE:
        logger.warning("HAMZA_SIGNATURE is empty.")
        return ""

    try:
        parsed = json.loads(HAMZA_SIGNATURE)

        if isinstance(parsed, dict):
            return str(parsed.get("html", ""))

        if isinstance(parsed, str):
            return parsed

    except json.JSONDecodeError:
        pass

    return HAMZA_SIGNATURE


def send_email(mailbox, recipient, subject, body):

    message = EmailMessage()

    message["From"] = mailbox["email"]

    # Receiver display name
    message["To"] = email.utils.formataddr(
        ("Western Arsal", recipient)
    )

    message["Subject"] = subject
    message["Message-ID"] = email.utils.make_msgid()

    # Plain text email version
    message.set_content(body)

    # Add HTML signature
    signature_html = get_signature_html()

    if signature_html:

        body_html = html.escape(body).replace(
            "\n",
            "<br>\n"
        )

        html_content = (
            "<!doctype html>"
            "<html>"
            "<body>"
            f"<div>{body_html}</div>"
            f"<div>{signature_html}</div>"
            "</body>"
            "</html>"
        )

        message.add_alternative(
            html_content,
            subtype="html"
        )

    if DRY_RUN:
        logger.info(
            "DRY RUN: Would send to %s using %s",
            recipient,
            mailbox["email"]
        )

        return message["Message-ID"]

    context = ssl.create_default_context()

    smtp_host = mailbox["smtp_host"]
    smtp_port = int(mailbox["smtp_port"])

    if smtp_port == 465:

        with smtplib.SMTP_SSL(
            smtp_host,
            smtp_port,
            context=context,
            timeout=30
        ) as server:

            server.login(
                mailbox["email"],
                mailbox["password"]
            )

            server.send_message(message)

    else:

        with smtplib.SMTP(
            smtp_host,
            smtp_port,
            timeout=30
        ) as server:

            server.ehlo()
            server.starttls(context=context)
            server.ehlo()

            server.login(
                mailbox["email"],
                mailbox["password"]
            )

            server.send_message(message)

    return message["Message-ID"]


# =========================================================
# PROCESS OUTGOING EMAILS
# =========================================================

def process_sends(
    spreadsheet,
    mailboxes,
    websites_ws,
    templates_ws,
    history_ws,
    suppression_ws
):

    templates = get_active_templates(templates_ws)

    if not templates:
        logger.info("No active templates. Nothing will be sent.")
        return

    rows = get_rows(websites_ws)

    suppressed = get_suppressed_emails(suppression_ws)
    contacted = get_contacted_emails(history_ws)

    # First email sends immediately if there is no previous successful send.
    # Following emails wait at least 12 minutes.
    last_send = last_successful_send(history_ws)

    if last_send:

        elapsed = (
            now_local() - last_send
        ).total_seconds()

        if elapsed < SEND_INTERVAL_SECONDS:
            logger.info(
                "Waiting for the next sending interval."
            )
            return

    mailbox = get_next_mailbox(mailboxes, history_ws)

    if daily_sent_count(
        history_ws,
        mailbox["email"]
    ) >= DAILY_LIMIT_PER_MAILBOX:

        logger.info("Selected mailbox reached its daily limit.")

        remaining = [
            m for m in mailboxes
            if daily_sent_count(
                history_ws,
                m["email"]
            ) < DAILY_LIMIT_PER_MAILBOX
        ]

        if not remaining:
            logger.info(
                "All mailboxes reached their daily limits."
            )
            return

        mailbox = remaining[0]

    for index, row in enumerate(rows):

        website = str(row.get("Website", "")).strip()
        recipient = normalize_email(row.get("Email", ""))
        status = str(row.get("Status", "")).strip().upper()

        sheet_row = index + 2

        if status != "PENDING":
            continue

        if not website or not valid_email(recipient):

            websites_ws.update_cell(
                sheet_row,
                3,
                "INVALID"
            )

            continue

        if recipient in suppressed:

            websites_ws.update_cell(
                sheet_row,
                3,
                "SUPPRESSED"
            )

            continue

        if recipient in contacted:

            websites_ws.update_cell(
                sheet_row,
                3,
                "ALREADY_CONTACTED"
            )

            continue

        template = select_template(templates, index)

        subject = template["subject"].replace(
            "{website_name}",
            website
        )

        body = template["body"].replace(
            "{website_name}",
            website
        )

        body += (
            "\n\nIf you do not wish to receive further messages, "
            "please reply with 'unsubscribe'."
        )

        # Record SENDING before contacting SMTP.
        # If a run crashes here, manual review is required.
        append_history(
            history_ws,
            recipient,
            website,
            mailbox["email"],
            "SENDING",
            details="Attempt started"
        )

        websites_ws.update(
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

            message_id = send_email(
                mailbox,
                recipient,
                subject,
                body
            )

            final_status = (
                "DRY_RUN" if DRY_RUN else "SENT"
            )

            append_history(
                history_ws,
                recipient,
                website,
                mailbox["email"],
                final_status,
                message_id,
                "Email processed"
            )

            if not DRY_RUN:

                websites_ws.update(
                    range_name=f"C{sheet_row}:G{sheet_row}",
                    values=[[
                        "SENT",
                        row.get("Task ID", ""),
                        row.get("Created At", ""),
                        timestamp(),
                        mailbox["email"]
                    ]]
                )

            else:

                websites_ws.update_cell(
                    sheet_row,
                    3,
                    "PENDING"
                )

            logger.info(
                "Email processing completed for %s",
                recipient
            )

        except Exception as exc:

            logger.exception("Email sending failed.")

            append_history(
                history_ws,
                recipient,
                website,
                mailbox["email"],
                "FAILED",
                details=type(exc).__name__
            )

            websites_ws.update_cell(
                sheet_row,
                3,
                "FAILED"
            )

        # Only one recipient per workflow run.
        # The next scheduled run handles the next recipient.
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

                charset = (
                    part.get_content_charset() or "utf-8"
                )

                parts.append(
                    payload.decode(
                        charset,
                        errors="replace"
                    )
                )

    else:

        payload = message.get_payload(decode=True)

        if payload:

            charset = (
                message.get_content_charset() or "utf-8"
            )

            parts.append(
                payload.decode(
                    charset,
                    errors="replace"
                )
            )

    return "\n".join(parts).strip()


def read_replies(
    mailboxes,
    replies_ws,
    suppression_ws,
    websites_ws
):

    existing_ids = {
        str(row.get("Message ID", "")).strip()
        for row in get_rows(replies_ws)
        if row.get("Message ID")
    }

    for mailbox in mailboxes:

        if not mailbox.get("imap_host"):

            logger.warning(
                "IMAP host missing for %s",
                mailbox["email"]
            )

            continue

        try:

            connection = imaplib.IMAP4_SSL(
                mailbox["imap_host"],
                int(mailbox.get("imap_port", 993)),
                ssl_context=ssl.create_default_context()
            )

            connection.login(
                mailbox["email"],
                mailbox["password"]
            )

            connection.select("INBOX")

            status, data = connection.search(
                None,
                "UNSEEN"
            )

            if status != "OK":
                connection.logout()
                continue

            message_numbers = data[0].split()

            for number in message_numbers:

                status, message_data = connection.fetch(
                    number,
                    "(RFC822)"
                )

                if status != "OK":
                    continue

                raw_email = message_data[0][1]

                message = email.message_from_bytes(raw_email)

                message_id = str(
                    message.get("Message-ID", "")
                ).strip()

                if message_id and message_id in existing_ids:

                    connection.store(
                        number,
                        "+FLAGS",
                        "\\Seen"
                    )

                    continue

                from_address = normalize_email(
                    parseaddr(
                        message.get("From", "")
                    )[1]
                )

                subject = decode_text(
                    message.get("Subject", "")
                )

                body = extract_body(message)

                classification = classify_reply(
                    subject,
                    body
                )

                snippet = body[:300].replace(
                    "\n",
                    " "
                )

                replies_ws.append_row([
                    timestamp(),
                    from_address,
                    subject,
                    message_id,
                    classification,
                    snippet
                ])

                if message_id:
                    existing_ids.add(message_id)

                if classification in (
                    "UNSUBSCRIBE",
                    "BOUNCE"
                ):

                    add_to_suppression(
                        suppression_ws,
                        from_address,
                        classification
                    )

                # Update matching website reply status.
                website_rows = get_rows(websites_ws)

                for index, website_row in enumerate(
                    website_rows,
                    start=2
                ):

                    if normalize_email(
                        website_row.get("Email", "")
                    ) == from_address:

                        websites_ws.update_cell(
                            index,
                            8,
                            classification
                        )

                connection.store(
                    number,
                    "+FLAGS",
                    "\\Seen"
                )

            connection.logout()

        except Exception:

            logger.exception(
                "IMAP reply check failed for %s",
                mailbox["email"]
            )


# =========================================================
# DASHBOARD
# =========================================================

def update_dashboard(
    dashboard_ws,
    history_ws,
    websites_ws,
    mailboxes
):

    history = get_rows(history_ws)
    websites = get_rows(websites_ws)

    counts = {
        "SENT": 0,
        "FAILED": 0,
        "SENDING": 0,
        "DRY_RUN": 0
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
        if str(row.get("Status", "")).upper() == "PENDING"
    )

    dashboard_data = [
        ["Metric", "Value"],
        ["Local Date", today],
        ["Active Mailboxes", len(mailboxes)],
        ["Daily Limit Per Mailbox", DAILY_LIMIT_PER_MAILBOX],
        ["Today's Sent", counts["SENT"]],
        ["Today's Failed", counts["FAILED"]],
        ["Currently Sending", counts["SENDING"]],
        ["Dry Runs", counts["DRY_RUN"]],
        ["Pending Websites", pending]
    ]

    dashboard_ws.clear()

    dashboard_ws.update(
        range_name="A1:B9",
        values=dashboard_data
    )


# =========================================================
# MAIN
# =========================================================

def main():

    logger.info("Email automation started.")

    mailboxes = load_mailboxes()

    spreadsheet = connect_google_sheets()

    setup_tabs(spreadsheet)

    websites_ws = get_worksheet(
        spreadsheet,
        "Websites"
    )

    templates_ws = get_worksheet(
        spreadsheet,
        "Templates"
    )

    history_ws = get_worksheet(
        spreadsheet,
        "History"
    )

    replies_ws = get_worksheet(
        spreadsheet,
        "Replies"
    )

    dashboard_ws = get_worksheet(
        spreadsheet,
        "Dashboard"
    )

    suppression_ws = get_worksheet(
        spreadsheet,
        "Suppression"
    )

    discover_new_tasks(websites_ws)

    read_replies(
        mailboxes,
        replies_ws,
        suppression_ws,
        websites_ws
    )

    process_sends(
        spreadsheet,
        mailboxes,
        websites_ws,
        templates_ws,
        history_ws,
        suppression_ws
    )

    update_dashboard(
        dashboard_ws,
        history_ws,
        websites_ws,
        mailboxes
    )

    logger.info("Email automation finished.")


if __name__ == "__main__":
    main()
