import os
import base64
import json
from google.oauth2 import service_account
from googleapiclient.discovery import build

# CONFIGURATION
SCOPES = ['https://www.googleapis.com/auth/gmail.modify']  # Needed for read and label modifications
SERVICE_ACCOUNT_FILE = 'YOUR_JSON_FILE'
USER_EMAIL = 'YOUR_EMAIL_ACCOUNT'  # The Gmail account to access
SENDER_EMAIL = '"Sunny Cars" via Finance <*******@tourlane.com>'
ATTACHMENTS_DIR = 'attachments'  # Where to save files

# Authenticate and build service
credentials = service_account.Credentials.from_service_account_file(
    SERVICE_ACCOUNT_FILE, scopes=SCOPES)
delegated_credentials = credentials.with_subject(USER_EMAIL)
service = build('gmail', 'v1', credentials=delegated_credentials)

# Create attachments folder
os.makedirs(ATTACHMENTS_DIR, exist_ok=True)

# Search for emails from specific sender
query = f'from:{SENDER_EMAIL} has:attachment'
results = service.users().messages().list(userId='me', q=query).execute()
messages = results.get('messages', [])

print(f"Found {len(messages)} messages with attachments from {SENDER_EMAIL}.")

# Process each message
for msg in messages:
    msg_id = msg['id']
    msg_detail = service.users().messages().get(userId='me', id=msg_id).execute()
    parts = msg_detail['payload'].get('parts', [])

    for part in parts:
        filename = part.get('filename')
        body = part.get('body', {})
        if 'attachmentId' in body:
            attachment_id = body['attachmentId']
            attachment = service.users().messages().attachments().get(
                userId='me', messageId=msg_id, id=attachment_id).execute()

            data = attachment['data']
            file_data = base64.urlsafe_b64decode(data.encode('UTF-8'))

            file_path = os.path.join(ATTACHMENTS_DIR, filename)
            with open(file_path, 'wb') as f:
                f.write(file_data)
            print(f"Saved attachment: {file_path}")
