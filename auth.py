import os
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from logger import logger


def get_authenticated_services():
    SCOPES = ['https://www.googleapis.com/auth/gmail.readonly',
              'https://www.googleapis.com/auth/spreadsheets',
              'https://www.googleapis.com/auth/documents',
              "https://www.googleapis.com/auth/drive"]

    creds = None
    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json", SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                "cred.json", SCOPES)
            creds = flow.run_local_server(
                port=0, access_type='offline', prompt='consent')
        # Save the credentials for the next run
        with open("token.json", "w") as token:
            token.write(creds.to_json())
    logger.info("OAuth Successfull")
    logger.info("Starting services")
    gmail_service = build('gmail', 'v1', credentials=creds)
    sheets_service = build('sheets', 'v4', credentials=creds)
    drive_service = build("drive", "v3", credentials=creds)
    docs_service = build('docs', 'v1', credentials=creds)

    return gmail_service, sheets_service, drive_service, docs_service
