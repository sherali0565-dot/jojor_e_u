"""
replies.py - reading incoming replies (IMAP) and the one-time "best price" auto-reply.

This module does NOT import main.py (that would be a circular import).
main.py passes itself in as `core`, which gives access to the shared helpers
(sheet access, send_email, logging, ...).
"""

import re
import ssl
import email
import imaplib
from email.utils import parseaddr

EMAIL_FINDER = re.compile(r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}")

AUTO_SUBJECT_TERMS = (
    "out of office", "automatic reply", "auto-reply", "autoreply",
    "auto reply", "away from my desk", "vacation reply",
)
NOREPLY_PREFIXES = ("noreply@", "no-reply@", "donotreply@", "do-not-reply@")

# A stronger status never gets overwritten by a weaker one.
_STATUS_PRIORITY = {"": 0, "AUTOREPLY": 1, "REPLY": 2, "BOUNCE": 3, "UNSUBSCRIBE": 3}


def _norm(address):
    return (address or "").strip().lower()


# =========================================================
# CLASSIFICATION & PARSING
# =========================================================

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
    if any(term in text for term in ["mailbox not found", "user unknown", "address rejected", "delivery failed", "undelivered mail", "delivery status notification"]):
        return "BOUNCE"
    if auto_responder or _looks_automatic(subject, sender):
        return "AUTOREPLY"
    if any(term in text for term in ["unsubscribe", "remove me", "do not contact", "stop emailing", "opt out"]):
        return "UNSUBSCRIBE"
    return "REPLY"


def extract_body(message):
    parts = []
    if message.is_multipart():
        for part in message.walk():
            if part.get_content_type() == "text/plain" and part.get_content_disposition() != "attachment":
                payload = part.get_payload(decode=True)
                if payload:
                    parts.append(payload.decode(part.get_content_charset() or "utf-8", errors="replace"))
    else:
        payload = message.get_payload(decode=True)
        if payload:
            parts.append(payload.decode(message.get_content_charset() or "utf-8", errors="replace"))
    return "\n".join(parts).strip()


def extract_bounced_recipients(message, body, known_recipients):
    """A bounce comes from mailer-daemon, not from the person we emailed.
    Only addresses that exist in the Websites sheet are accepted, so
    signature/support addresses inside the bounce text are ignored."""
    candidates = set()
    header = str(message.get("X-Failed-Recipients", ""))
    for addr in EMAIL_FINDER.findall(header):
        candidates.add(_norm(addr))
    for addr in EMAIL_FINDER.findall(body or ""):
        candidates.add(_norm(addr))
    return {a for a in candidates if a in known_recipients}


# =========================================================
# READ REPLIES (IMAP)
# =========================================================

def read_replies(mailboxes, replies_ws, suppression_ws, websites_ws, run_log_ws, core):
    existing_ids = {str(row.get("Message ID", "")).strip() for row in core.get_rows(replies_ws) if row.get("Message ID")}

    # Load Websites once (not once per incoming message).
    website_rows = core.get_rows(websites_ws)
    rows_by_email = {}
    status_by_row = {}
    for idx, website_row in enumerate(website_rows, start=2):
        em = _norm(website_row.get("Email", ""))
        if em:
            rows_by_email.setdefault(em, []).append(idx)
        status_by_row[idx] = str(website_row.get("Reply Status", "")).strip().upper()
    known_recipients = set(rows_by_email)
    suppressed_now = core.get_suppressed_emails(suppression_ws)

    for mailbox in mailboxes:
        if not mailbox.get("imap_host"):
            continue
        connection = None
        try:
            connection = imaplib.IMAP4_SSL(mailbox["imap_host"], int(mailbox.get("imap_port", 993)), ssl_context=ssl.create_default_context())
            connection.login(mailbox["email"], mailbox["password"])
            connection.select("INBOX")
            status, data = connection.search(None, "UNSEEN")
            if status != "OK":
                continue
            for number in data[0].split():
                status, message_data = connection.fetch(number, "(RFC822)")
                if status != "OK":
                    continue
                message = email.message_from_bytes(message_data[0][1])
                message_id = str(message.get("Message-ID", "")).strip()
                if message_id and message_id in existing_ids:
                    connection.store(number, "+FLAGS", "\\Seen")
                    continue
                from_address = _norm(parseaddr(message.get("From", ""))[1])
                subject = core.decode_text(message.get("Subject", ""))
                body = extract_body(message)
                classification = classify_reply(subject, body, from_address, is_auto_responder(message))
                snippet = body[:300].replace("\n", " ")
                core.sheets_retry(replies_ws.append_row, [core.timestamp(), from_address, subject, message_id, classification, snippet])
                if message_id:
                    existing_ids.add(message_id)

                if classification == "BOUNCE":
                    targets = extract_bounced_recipients(message, body, known_recipients)
                else:
                    targets = {from_address} if from_address else set()

                for target in targets:
                    if classification in ("UNSUBSCRIBE", "BOUNCE") and target not in suppressed_now:
                        core.add_to_suppression(suppression_ws, target, classification)
                        suppressed_now.add(target)
                    for sheet_row in rows_by_email.get(target, []):
                        current = status_by_row.get(sheet_row, "")
                        if _STATUS_PRIORITY.get(classification, 0) >= _STATUS_PRIORITY.get(current, 0):
                            core.sheets_retry(websites_ws.update_cell, sheet_row, 8, classification)
                            status_by_row[sheet_row] = classification

                connection.store(number, "+FLAGS", "\\Seen")
        except Exception as exc:
            core.logger.exception("Reading replies failed for %s", mailbox.get("email"))
            if run_log_ws:
                core.write_run_log(run_log_ws, "ERROR", "IMAP_FAILED", f"{mailbox.get('email')}: {type(exc).__name__}: {exc}")
        finally:
            if connection is not None:
                try:
                    connection.logout()
                except Exception:
                    pass


# =========================================================
# AUTO-REPLY ("best price", sent once per contact)
# =========================================================

def backfill_auto_reply_flag(websites_ws, retry):
    """Runs once, when the 'Auto Reply' column is first created.
    Rows that already have a Reply Status get 'SKIP_OLD' so that old replies
    are NOT auto-answered after this feature is switched on."""
    values = retry(websites_ws.get_all_values)
    if len(values) < 2:
        return 0
    headers = [str(v).strip() for v in values[0]]
    if "Reply Status" not in headers or "Auto Reply" not in headers:
        return 0
    reply_idx = headers.index("Reply Status")
    auto_idx = headers.index("Auto Reply")
    column = []
    marked = 0
    for row in values[1:]:
        has_reply = reply_idx < len(row) and str(row[reply_idx]).strip() != ""
        column.append(["SKIP_OLD" if has_reply else ""])
        marked += 1 if has_reply else 0
    if marked:
        letter = _column_letter(auto_idx + 1)
        retry(websites_ws.update, range_name=f"{letter}2:{letter}{len(values)}", values=column)
    return marked


def _column_letter(number):
    letters = ""
    while number > 0:
        number, remainder = divmod(number - 1, 26)
        letters = chr(65 + remainder) + letters
    return letters


def find_auto_reply_template(templates_ws, core):
    """Row of the Templates sheet whose Template ID equals AUTO_REPLY_TEMPLATE_ID.
    The Active column works as an on/off switch for the auto-reply."""
    wanted = core.AUTO_REPLY_TEMPLATE_ID.strip().lower()
    for row in core.get_rows(templates_ws):
        if str(row.get("Template ID", "")).strip().lower() != wanted:
            continue
        if str(row.get("Active", "")).strip().lower() not in ("true", "yes", "1", "active"):
            return None
        subject = str(row.get("Subject", "")).strip()
        body = str(row.get("Body", "")).strip()
        if subject and body:
            return {"id": str(row.get("Template ID", "")).strip(), "subject": subject, "body": body}
    return None


def _latest_reply_for(replies_rows, recipient):
    for row in reversed(replies_rows):
        if _norm(row.get("From", "")) == recipient and str(row.get("Classification", "")).strip().upper() == "REPLY":
            return str(row.get("Subject", "")).strip(), str(row.get("Message ID", "")).strip()
    return "", ""


def _reply_subject(their_subject, template_subject):
    subject = re.sub(r"\s+", " ", their_subject or "").strip()
    if not subject:
        return template_subject
    return subject if re.match(r"(?i)^re\s*:", subject) else f"Re: {subject}"


def send_auto_replies(mailboxes, websites_ws, replies_ws, templates_ws, history_ws, suppression_ws, run_log_ws, core):
    """Sends the 'best price' template once to every contact whose
    Reply Status is REPLY and whose 'Auto Reply' cell is empty.
    The cell is written BEFORE sending ('SENDING ...') so a crash can never
    cause a second email; afterwards it becomes 'REPLIED <time>'.
    On failure it becomes 'FAILED <time>' (clear the cell to retry)."""
    if not core.AUTO_REPLY_ENABLED:
        return 0

    values = core.sheets_retry(websites_ws.get_all_values)
    if len(values) < 2:
        return 0
    headers = [str(v).strip() for v in values[0]]
    hmap = {name: idx for idx, name in enumerate(headers)}
    needed = ("Website", "Email", "Status", "Mailbox", "Reply Status", "Auto Reply")
    missing = [name for name in needed if name not in hmap]
    if missing:
        core.write_run_log(run_log_ws, "WARNING", "AUTO_REPLY_COLUMN_MISSING", f"Missing columns: {missing}")
        return 0
    auto_col = hmap["Auto Reply"] + 1  # 1-based for update_cell

    def cell(row, name):
        idx = hmap[name]
        return str(row[idx]).strip() if idx < len(row) and row[idx] is not None else ""

    candidates = []
    for sheet_row, row in enumerate(values[1:], start=2):
        if cell(row, "Reply Status").upper() != "REPLY":
            continue
        if cell(row, "Auto Reply"):
            continue
        if cell(row, "Status").upper() != "SENT":
            continue
        recipient = _norm(cell(row, "Email"))
        if recipient:
            candidates.append((sheet_row, cell(row, "Website"), recipient, _norm(cell(row, "Mailbox"))))
    if not candidates:
        return 0

    if not core.within_send_window():
        return 0  # nothing is marked, so it is retried on a later run

    template = find_auto_reply_template(templates_ws, core)
    if template is None:
        core.write_run_log(
            run_log_ws, "WARNING", "AUTO_REPLY_TEMPLATE_MISSING",
            f"No active template with Template ID '{core.AUTO_REPLY_TEMPLATE_ID}'. Waiting replies: {len(candidates)}"
        )
        return 0

    suppressed = core.get_suppressed_emails(suppression_ws)
    replies_rows = core.get_rows(replies_ws)
    mailbox_by_email = {_norm(m["email"]): m for m in mailboxes}

    sent = 0
    for sheet_row, website, recipient, mailbox_email in candidates:
        if sent >= core.AUTO_REPLY_MAX_PER_RUN:
            break

        if recipient in suppressed:
            core.sheets_retry(websites_ws.update_cell, sheet_row, auto_col, "SKIPPED_SUPPRESSED")
            continue

        mailbox = mailbox_by_email.get(mailbox_email) or mailboxes[0]

        try:
            core.get_signature_html(mailbox, mailboxes)  # never reply without a signature
        except Exception as exc:
            core.write_run_log(run_log_ws, "ERROR", "AUTO_REPLY_NO_SIGNATURE", str(exc))
            return sent

        try:
            website_name = core.ai_website_name(website)["name"]
        except Exception:
            website_name = core.fallback_clean_website_name(website)

        their_subject, their_message_id = _latest_reply_for(replies_rows, recipient)
        subject = _reply_subject(their_subject, template["subject"]).replace("{website_name}", website_name)
        body = template["body"].replace("{website_name}", website_name)

        stamp = core.timestamp()
        core.sheets_retry(websites_ws.update_cell, sheet_row, auto_col, f"SENDING {stamp}")
        try:
            message_id = core.send_email(
                mailbox, recipient, subject, body, mailboxes,
                in_reply_to=their_message_id or None,
                references=their_message_id or None,
            )
        except Exception as exc:
            core.logger.warning("Auto-reply failed via %s: %s", mailbox["email"], exc)
            core.sheets_retry(websites_ws.update_cell, sheet_row, auto_col, f"FAILED {core.timestamp()}")
            core.append_history(history_ws, recipient, website, mailbox["email"], "AUTO_REPLY_FAILED",
                                details=f"{type(exc).__name__}", template_id=template["id"])
            core.write_run_log(run_log_ws, "ERROR", "AUTO_REPLY_FAILED", f"{recipient}: {type(exc).__name__}: {exc}")
            continue

        core.sheets_retry(websites_ws.update_cell, sheet_row, auto_col, f"REPLIED {core.timestamp()}")
        core.append_history(history_ws, recipient, website, mailbox["email"], "AUTO_REPLY_SENT",
                            message_id, "Best price auto-reply", template_id=template["id"])
        core.write_run_log(run_log_ws, "INFO", "AUTO_REPLY_SENT", f"{recipient} via {mailbox['email']}")
        sent += 1

    return sent
