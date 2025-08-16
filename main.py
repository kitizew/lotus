from __future__ import print_function
import os.path
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
import json;
from privatData import *

SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets.readonly',
    'https://www.googleapis.com/auth/calendar'
          ]

def main():
    creds = None
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    service = build('sheets', 'v4', credentials=creds)

    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=SPREADSHEET_ID,
                                range=RANGE_NAME).execute()

    rows = result.get('values', [])
    keywords = ['TL  Leadgen&CX', 'TL oGT Sales', 'TL B2B MKT']

    extracted_rows = []

    for row in rows:
        for i, cell in enumerate(row):
            if cell in keywords:
                # беремо попередній рядок як текст, сам TL, та наступну клітинку як дату
                text = row[i - 1] if i - 1 >= 0 else ""
                tl = cell
                date = row[i + 1] if i + 1 < len(row) else ""
                extracted_rows.append([text, tl, date])

    # Записуємо у JSON
    with open('data.json', 'w', encoding='utf-8') as f:
        json.dump(extracted_rows, f, ensure_ascii=False, indent=2)



if __name__ == '__main__':
    main()
