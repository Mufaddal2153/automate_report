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
import pandas as pd, time

# Authenticate to Google Cloud APIs
# DELETE TOKEN JSON FILE BEFORE RUNNING
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



def create_message(cc, to, subject, html_body):
    message = MIMEMultipart('alternative')
    message['to'] = to
    message['cc'] = cc
    message['subject'] = subject

    # Create a MIMEText object for the HTML part
    html_part = MIMEText(html_body, 'html')

    # Attach the plain text and HTML parts to the message
    message.attach(html_part)

    return {'raw': base64.urlsafe_b64encode(message.as_bytes()).decode()}



try:
    folder_name = drive_service.files().list(q = "mimeType = 'application/vnd.google-apps.folder' and name = 'Reports 2024'", fields="nextPageToken, files(id, name, webContentLink, webViewLink)").execute()
    email_file_name = drive_service.files().list(q = "name = 'SEO Client Emails 2024'", fields="nextPageToken, files(id, name, webContentLink, webViewLink)").execute()
    print(email_file_name)
    folder_reports = folder_name.get('files', [])
    emails = email_file_name.get('files', [])
    if not emails:
        print('File not found')
    else:
        # print(items)
        folder_reports_id = folder_reports[0]['id']
        res = drive_service.files().list(q = "'" + folder_reports_id + "' in parents", fields="nextPageToken, files(id, name, webContentLink, webViewLink)").execute()
        report_items = res.get('files', [])
        res_sheet = sheets_service.spreadsheets().values().get(spreadsheetId=emails[0]['id'], range='Sheet1').execute()
        # to pd dataframe
        df = pd.DataFrame(res_sheet.get('values', [])).iloc[1:, :]
        data = df.to_dict('records')
        # print(data)
        send_data = []
        # range_name = 'Traffic Ranking!A1'
        # clear_values_request_body = {
        #     'range': range_name,
        # }
        count = 0
        for item in data:
            if item[1]:
                for i, j in enumerate(report_items):
                    if item[0].strip() == j['name'].strip():
                        # res_sheet = sheets_service.spreadsheets().values().get(spreadsheetId = j['id'], range=range_name).execute()
                        # column_data = res_sheet.get('values', [])[0][0] if res_sheet.get('values', []) else ''
                        # if column_data:
                        temp1 = item[1]
                        temp1 = temp1.replace("\r", " ")
                        temp1 = temp1.replace("\n", " ")
                        temp2 = item[2] if item[2] else ''
                        temp2 = temp2.replace("\r", " ")
                        temp2 = temp2.replace("\n", " ")
                        send_data.append({'name': item[0], 'to': temp1, 'cc': temp2, 'link': j['webViewLink'], 'id': j['id'], 'text': ''}) #column_data})
                        if count != 0 and count % 30 == 0:
                            print('sleeping')
                            count = -1
                            time.sleep(60)
                        print("sheet count", count)
                        count += 1
        count = 0
        print("send_data", send_data)
        for item in send_data:
            item['text'] = item['text'].replace('\n', '<br>')
            text = "Salaams,<br><br>Please review the SEO report for the month of May 2024.<br><br>" + item['link'] + f"<br><br> {item['text']}"
            client_name = ""
            try:
                client_name = item['name']
                msg = create_message(item['cc'], item['to'], item['name'], text)
                # sheets_service.spreadsheets().values().clear(spreadsheetId=item['id'], body=clear_values_request_body, range=range_name).execute()
                send_message = (gmail_service.users().messages().send(userId="me", body=msg).execute())
                if count != 0 and count % 40 == 0:
                    print('sleeping')
                    count = -1
                    time.sleep(60)
                print("email count", count, item['name'])
                count += 1
            except Exception as error:
                print(error, client_name)

except HttpError as error:
    print(error)