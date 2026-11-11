import os
import json
import time
import random
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.utils import formataddr
import gspread
from google.oauth2.service_account import Credentials

# --- Configuration & Constants ---
COMPANY_NAME = "Western Arsal"

def get_google_sheet():
    """Google Sheet se connection banata hai service account JSON ke zariye"""
    scope = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
    service_account_info = json.loads(os.environ["GOOGLE_SERVICE_ACCOUNT_JSON"])
    creds = Credentials.from_service_account_info(service_account_info, scopes=scope)
    client = gspread.authorize(creds)
    sheet_id = os.environ["GOOGLE_SHEET_ID"]
    return client.open_by_key(sheet_id).sheet1

def send_email(mailbox, recipient_email, subject, body):
    """Hostinger SMTP ke zariye email bhejta hai"""
    msg = MIMEMultipart()
    msg["From"] = formataddr((COMPANY_NAME, mailbox["email"]))
    msg["To"] = recipient_email
    msg["Subject"] = subject

    msg.attach(MIMEText(body, "plain"))

    # Dry Run check: Agar DRY_RUN true hai toh sirf print karega, email nahi bhejega (Testing ke liye safe)
    is_dry_run = os.environ.get("DRY_RUN", "true").lower() == "true"
    if is_dry_run:
        print(f"[DRY RUN] Would send email from {mailbox['email']} to {recipient_email}")
        return True

    try:
        # Hostinger SMTP configuration (Port 465 SSL)
        server = smtplib.SMTP_SSL("smtp.hostinger.com", 465)
        server.login(mailbox["email"], mailbox["password"])
        server.sendmail(mailbox["email"], recipient_email, msg.as_string())
        server.quit()
        return True
    except Exception as e:
        print(f"Failed to send email from {mailbox['email']} to {recipient_email}. Error: {e}")
        return False

def main():
    print("Email automation script shuru ho rahi hai...")
    
    # 1. Mailboxes load karein jo GitHub Secret (MAILBOXES_JSON) mein rakhe hain
    mailboxes_raw = os.environ.get("MAILBOXES_JSON")
    if not mailboxes_raw:
        raise ValueError("MAILBOXES_JSON secret missing hai!")
    mailboxes = json.loads(mailboxes_raw)
    
    if not mailboxes:
        raise ValueError("Koi mailbox configuration nahi mili!")

    # 2. Google Sheet load karein
    sheet = get_google_sheet()
    records = sheet.get_all_records()
    
    mailbox_index = 0
    total_mailboxes = len(mailboxes)

    # Email template
    subject_template = "Question regarding guest posts on {website}"
    base_body_template = """Hello {website},

I was browsing your website and wanted to check what your terms and conditions are for publishing guest articles. Do you accept sponsored posts or contributor articles, and what are your rates?

Looking forward to your response.

Best regards,
Western Arsal"""

    for index, row in enumerate(records, start=2): # Row 2 se start kyunki row 1 header hai
        status = row.get("Status", "").strip()
        
        # Agar pehle hi email ja chuki hai toh skip karein
        if status.lower() == "sent":
            continue
            
        website = row.get("Website", "").strip()
        email = row.get("Email", "").strip()
        
        if not email or not website:
            continue

        # Current mailbox select karein (Rotation ke liye)
        current_mailbox = mailboxes[mailbox_index]
        
        print(f"Processing row {index}: Sending to {email} using mailbox {current_mailbox['email']}")

        subject = subject_template.format(website=website)
        body = base_body_template.format(website=website)

        # Email send karein
        success = send_email(current_mailbox, email, subject, body)
        
        if success:
            # Google Sheet mein status update kar dein (Column 3 ko 'Sent' kar rahe hain)
            try:
                sheet.update_cell(index, 3, "Sent")
            except Exception as ex:
                print(f"Sheet update error: {ex}")

            print(f"Email successfully processed for {email}")
            
            # Next mailbox par move karein (Rotation)
            mailbox_index = (mailbox_index + 1) % total_mailboxes
            
            # 10 se 15 minute (600 se 900 seconds) ka random gap taaki spam filter na pakde
            delay = random.randint(600, 900)
            print(f"Agli email bhejne se pehle {delay // 60} minute ka wait ho raha hai...")
            time.sleep(delay)
        else:
            print(f"Error sending to {email}, retrying/skipping...")
            time.sleep(10)

    print("Sabhi emails process ho chuki hain!")

if __name__ == "__main__":
    main()
