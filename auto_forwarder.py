import os
import time
import base64
import pickle
import shutil
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.mime.text import MIMEText

from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

# === CONFIG ===
PDF_FOLDER = 'pdfs'
SENT_FOLDER = 'sent_pdfs'
SCOPES = ['https://www.googleapis.com/auth/gmail.send']
TO_EMAIL = '********@tourlane.com'
BODY_TEXT = '*******@sunnycars.de Sunny Cars (EUR)'
BATCH_SIZE = 30
DELAY_SECONDS = 600  # 10 minutes

# === SETUP ===
os.makedirs(SENT_FOLDER, exist_ok=True)

# === AUTH ===
def authenticate_gmail():
    creds = None
    if os.path.exists('gmail_token.pkl'):
        with open('gmail_token.pkl', 'rb') as token:
            creds = pickle.load(token)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('client_secret.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('gmail_token.pkl', 'wb') as token:
            pickle.dump(creds, token)
    return build('gmail', 'v1', credentials=creds)

# === SEND EMAIL WITH ATTACHMENTS ===
def send_email(service, to_email, subject, body_text, attachments):
    message = MIMEMultipart()
    message['to'] = to_email
    message['subject'] = subject

    message.attach(MIMEText(body_text, 'plain'))

    for file_path in attachments:
        filename = os.path.basename(file_path)
        with open(file_path, 'rb') as f:
            part = MIMEApplication(f.read(), _subtype='pdf')
            part.add_header('Content-Disposition', 'attachment', filename=filename)
            message.attach(part)

    raw_message = base64.urlsafe_b64encode(message.as_bytes()).decode('utf-8')
    service.users().messages().send(userId='me', body={'raw': raw_message}).execute()
    print(f"✅ Sent email with {len(attachments)} attachment(s).")

# === MOVE SENT FILES ===s
def move_sent_files(batch):
    for file_path in batch:
        filename = os.path.basename(file_path)
        target_path = os.path.join(SENT_FOLDER, filename)
        shutil.move(file_path, target_path)
        print(f"📂 Moved to {SENT_FOLDER}: {filename}")

# === BATCH SENDING LOOP ===
import traceback

def send_pdfs_in_batches():
    service = authenticate_gmail()
    pdf_files = [os.path.join(PDF_FOLDER, f) for f in os.listdir(PDF_FOLDER) if f.endswith('.pdf')]
    total = len(pdf_files)
    print(f"📋 Total PDF files to send: {total}")

    for i in range(0, total, BATCH_SIZE):
        batch = pdf_files[i:i + BATCH_SIZE]
        batch_number = (i // BATCH_SIZE) + 1

        print(f"\n📦 Sending Batch {batch_number} — {len(batch)} file(s):")
        for file_path in batch:
            size_mb = round(os.path.getsize(file_path) / 1024 / 1024, 2)
            print(f"  - {os.path.basename(file_path)} ({size_mb} MB)")

        try:
            send_email(service, TO_EMAIL, "Sunny Cars Documents", BODY_TEXT, batch)
            move_sent_files(batch)
        except Exception as e:
            print(f"❌ Failed to send batch {batch_number}: {e}")
            traceback.print_exc()

            # Retry once after short wait
            print("🔁 Retrying batch in 30 seconds...")
            time.sleep(30)
            try:
                send_email(service, TO_EMAIL, "Sunny Cars Documents (Retry)", BODY_TEXT, batch)
                move_sent_files(batch)
                print(f"✅ Batch {batch_number} sent on retry.")
            except Exception as retry_e:
                print(f"❌ Retry for batch {batch_number} also failed: {retry_e}")
                traceback.print_exc()

        if i + BATCH_SIZE < total:
            print("🔌 Disconnecting Gmail service...")
            service = None
            print("⏳ Waiting 10 minutes before reconnecting...")
            time.sleep(DELAY_SECONDS)
            print("🔄 Re-authenticating Gmail...")
            service = authenticate_gmail()


# === MAIN ===
if __name__ == '__main__':
    send_pdfs_in_batches()
