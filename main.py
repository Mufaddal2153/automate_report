import pandas as pd
import os
import google.auth
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from googleapiclient import errors
import base64
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from google_auth_oauthlib.flow import InstalledAppFlow

# Authenticate to Google Cloud APIs
creds = None
if os.path.exists('token.json'):
    creds = Credentials.from_authorized_user_file('token.json', ['https://www.googleapis.com/auth/drive', 'https://www.googleapis.com/auth/gmail.compose'])
    print("Credentials loaded from tokens.json", creds)
if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
    else:
        flow = InstalledAppFlow.from_client_secrets_file('credentials.json', ['https://www.googleapis.com/auth/drive', 'https://www.googleapis.com/auth/gmail.compose'])
        creds = flow.run_local_server(port=0)
    with open('token.json', 'w') as token:
        token.write(creds.to_json())

# Fetch file names and email IDs from email.xlsx using Pandas
# df = pd.read_excel('email.xlsx')

# Search for files on Google Drive and create shareable links
drive_service = build('drive', 'v3', credentials=creds)
sheets_service = build('sheets', 'v4', credentials=creds)
gmail_service = build('gmail', 'v1', credentials=creds)
# for index, row in df.iterrows():
try:
    # query = "name='" + row['filename'] + "'"
    results = drive_service.files().list(q = "mimeType = 'application/vnd.google-apps.folder' and name = 'Reports - test'", fields="nextPageToken, files(id, name, webContentLink, webViewLink)").execute()
    # results = drive_service.files().list(q="name='ATMOSPHERE SEO-SMM REPORT 2023- BOHRADEVELOPERS'" ,fields="nextPageToken, files(id, name, webViewLink)").execute()
    items = results.get('files', [])
    if not items:
        print('File not found')
    else:
        print(items)
        folder_id = items[0]['id']
        res = drive_service.files().list(q = "'" + folder_id + "' in parents", fields="nextPageToken, files(id, name, webContentLink, webViewLink)").execute()
        res_items = res.get('files', [])
        print(res_items)
        item = res_items[0]
        file_id = item['id']
        file_name = item['name']
        print(f'File found: {file_id} {file_name}')
        range_name = 'Traffic Ranking!A1'
        res_sheet = sheets_service.spreadsheets().values().get(spreadsheetId=file_id, range=range_name).execute()
        column_data = res_sheet.get('values', [])
        print('Column data:', column_data[0][0])
        clear_values_request_body = {
            'range': range_name,
        }
        sheets_service.spreadsheets().values().clear(spreadsheetId=file_id, body=clear_values_request_body, range=range_name).execute()
        file_url = item['webViewLink'] if 'webViewLink' in item else None
        # for item in items:
        #     file_id = item['id']
        #     file_name = item['name']
        #     file_download_url = item['webContentLink'] if 'webContentLink' in item else None
        #     file_url = item['webViewLink'] if 'webViewLink' in item else None
        #     print(f'File found: {file_id} {file_name} {file_url}')
        #     email_id = row['email_id']

        #     # Send email with the link to the corresponding file
        # service = build('gmail', 'v1', credentials=creds)
        message = MIMEMultipart()
        email_id = "mufaddalhatim53@gmail.com, bohradevelopers@gmail.com, zainabsaleem.cash@gmail.com"
        cc_email_id = "zahabiyamodi@gmail.com, mariahatimbd@gmail.com"
        message['to'] = email_id
        message['cc'] = cc_email_id
        message['subject'] = 'Your report'
        text = "Salaams, please check your SEO report here: " + file_url + f"\n\n {column_data[0][0]}"
        part = MIMEText(text)
        message.attach(part)
        create_message = {'raw': base64.urlsafe_b64encode(message.as_bytes()).decode()}
        send_message = (gmail_service.users().messages().send(userId="me", body=create_message).execute())
        print(F'sent message to {email_id}, {cc_email_id} Message Id: {send_message["id"]}')
except HttpError as error:
    print(F'An error occurred: {error}')
    send_message = None
    # return send_message
