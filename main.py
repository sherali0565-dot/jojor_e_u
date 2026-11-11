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
from google import genai

# --- Configuration & Constants ---
COMPANY_NAME = "Western Arsal"  # Company name wapas Western Arsal kar diya hai
SHEET_NAME = "Outreach_List"  # Apni Google Sheet ka naam yahan likhein

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

def personalize_email_with_ai(website_name, raw_body):
    """Gemini AI ke zariye email ko website ke mutabiq personalize karta hai"""
    api_key = os.environ.get("GEMINI_API_KEY")
    if not api_key:
        return raw_body  # Agar key na ho toh default body return karega
    
    try:
        client = genai.Client(api_key=api_key)
        prompt = f"""
        Aap ek professional outreach expert hain. Neeche di gayi email body ko website '{website_name}' ke liye thoda customize aur natural bana dein. 
        Tone polite aur professional honi chahiye. Signature ya placeholders ko kharab mat karna.
        
        Original Email Body:
        {raw_body}
        """
        response = client.models.generate_content(
            model="gemini-2.5-flash",
            contents=prompt,
        )
        return response.text.strip()
    except Exception as e:
        print(f"Gemini AI Error: {e}. Using default body.")
        return raw_body

def send_email(mailbox, recipient_email, subject, body):
    """Hostinger SMTP ke zariye email bhejta hai"""
    msg = MIMEMultipart()
    msg["From"] = formataddr((COMPANY_NAME, mailbox["email"]))
    msg["To"] = recipient_email
    msg["Subject"] = subject

    msg.attach(MIMEText(body, "plain"))

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

    # GitHub Secret se HAMZA_SIGNATURE load karein
    hamza_signature = os.environ.get("HAMZA_SIGNATURE", "")

    # 2. Google Sheet load karein
    sheet = get_google_sheet()
    records = sheet.get_all_records()
    
    mailbox_index = 0
    total_mailboxes = len(mailboxes)

    # Email template
    subject_template = "Question regarding guest posts on {website}"
    base_body_template = """Hello {website},

I was browsing your website and wanted to check what your terms and conditions are for publishing guest articles. Do you accept sponsored posts or contributor articles, and what are your rates?

Looking forward to your response."""

    for index, row in enumerate(records, start=2): # Row 2 se start kyunki row 1 header hai
        status = row.get("Status", "").strip()
        
        # Agar pehle hi email ja chuki hai toh skip karein
        if status.lower() == "sent":
            continue
            
        website = row.get("Website", "").strip()
        email = row.get("Email", "").strip()
        
        if not email or not website:
            continue

        # Current mailbox select karein (rotation ke liye)
        current_mailbox = mailboxes[mailbox_index]
        
        print(f"Processing row {index}: Sending to {email} using mailbox {current_mailbox['email']}")

        # Gemini AI se message personalize karayein
        personalized_body = personalize_email_with_ai(website, base_body_template.format(website=website))
        
        # Dynamic Signature add karna (jo HAMZA_SIGNATURE secret se aaya hai)
        if hamza_signature:
            full_body = personalized_body + "\n\n" + hamza_signature
        else:
            full_body = personalized_body

        subject = subject_template.format(website=website)

        # Email send karein
        success = send_email(current_mailbox, email, subject, full_body)
        
        if success:
            # Google Sheet mein status update kar dein
            sheet.update_cell(index, 3, "Sent") # Column 3 'Status' ka hai
            print(f"Email successfully sent to {email}")
            
            # Next mailbox par move karein (Rotation)
            mailbox_index = (mailbox_index + 1) % total_mailboxes
            
            # 10 se 15 minute (600 se 900 seconds) ka random gap
            delay = random.randint(600, 900)
            print(f"Agli email bhejne se pehle {delay // 60} minute ka wait ho raha hai...")
            time.sleep(delay)
        else:
            print(f"Error sending to {email}, trying next row...")
            time.sleep(10)

    print("Sabhi emails process ho chuki hain!")

if __name__ == "__main__":
    main()
