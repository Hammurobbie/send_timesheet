import sys
import smtplib, ssl
import base64
import datetime
import requests
import timeit

import io
import pickle
import os.path
from googleapiclient.http import MediaIoBaseDownload
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

from datetime import date
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

def sendTime():

    startTimer = timeit.default_timer()

    times = sys.argv
    start = times[1]
    stop = times[2]
    startHour = start[0:len(start)//2]
    startMinute = start[len(start)//2:]
    stopHour = stop[0:len(stop)//2]
    stopMinute = stop[len(stop)//2:]
    hourMath = int(stopHour) - int(startHour)
    minuteMath = int(stopMinute) - int(startMinute)
    if minuteMath == -15:
        hourMath -=1
        minuteMath = 45
    elif minuteMath == -30:
        hourMath -=1
        minuteMath = 30
    elif minuteMath == -45:
        hourMath -=1
        minuteMath = 15
    minuteMath -= 30
    if minuteMath == -15:
        hourMath -=1
        minuteMath = 45
    elif minuteMath == -30:
        hourMath -=1
        minuteMath = 30
    elif minuteMath == -45:
        hourMath -=1
        minuteMath = 15
        
    if minuteMath == 15:
        minuteMath = 25
    elif minuteMath == 30:
        minuteMath = 50
    elif minuteMath == 45:
        minuteMath = 75
    hours = str(hourMath) + '.' + str(minuteMath)


    today = date.today()
    dateToday = today.strftime("%m/%d/%y")


    def handleSheets():
        
        SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

        # The ID and ranges of spreadsheets.
        SPREADSHEET_ID = '[id_of_google_sheet_you_want_to_modify]'
        DATE_RANGE = 'Sheet1!A2:N2'
        HOURS_RANGE = 'Sheet1!A26:N26'

        creds = None
        # The file token.pickle stores the user's access and refresh tokens, and is
        # created automatically when the authorization flow completes for the first
        # time.
        if os.path.exists('sheets_token.pickle'):
            with open('sheets_token.pickle', 'rb') as token:
                creds = pickle.load(token)
        # If there are no (valid) credentials available, let the user log in.
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(
                'credentials_sheets.json', SCOPES)
                creds = flow.run_local_server(port=0)
            # Save the credentials for the next run
            with open('sheets_token.pickle', 'wb') as token:
                pickle.dump(creds, token)

        service = build('sheets', 'v4', credentials=creds)

        # Change the spreadsheet's title vv

        requests = []

        requests.append({
            'updateSpreadsheetProperties': {
                'properties': {
                    'title': f"{dateToday} Timesheet"
                },
                'fields': 'title'
            }
        })

        body = {
            'requests': requests
        }
        service.spreadsheets().batchUpdate(
            spreadsheetId=SPREADSHEET_ID,
            body=body).execute()

        # Change the spreadsheet's title ^^

        # Change the spreadsheet's date and hours vv

        values = [
            [
                '', '', '', 'Date:', dateToday
            ],
        ]

        values2 = [
            [
                '[job_number]', '', '[job_name]', '', '', '', '', start, '', stop, '0.5', '', hours
            ],
        ]
        data = [
            {
                'range': DATE_RANGE,
                'values': values
            },
            {
                'range': HOURS_RANGE,
                'values': values2
            },
        ]
        body = {
            'valueInputOption': 'USER_ENTERED',
            'data': data
        }
        result = service.spreadsheets().values().batchUpdate(
            spreadsheetId=SPREADSHEET_ID, body=body).execute()
        print('{0} cells updated.'.format(result.get('totalUpdatedCells')))

        # Change the spreadsheet's date and hours ^^

    handleSheets()

    def handleDownload():
        
        SCOPES = ['https://www.googleapis.com/auth/drive']

        creds = None

        if os.path.exists('drive_token.pickle'):
            with open('drive_token.pickle', 'rb') as token:
                creds = pickle.load(token)
        # If there are no (valid) credentials available, let the user log in.
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(
                    'credentials_drive.json', SCOPES)
                creds = flow.run_local_server(port=0)
            # Save the credentials for the next run
            with open('drive_token.pickle', 'wb') as token:
                pickle.dump(creds, token)

        service = build('drive', 'v3', credentials=creds)
        
        file_id = '[id_of_sheet_saved_in_drive]'
        request = service.files().export_media(fileId=file_id,
                                             mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        fh = io.BytesIO()
        downloader = MediaIoBaseDownload(fh, request)
        done = False
        while done is False:
            status, done = downloader.next_chunk()
            destFilename='//path/and/filename/to/save/to.xlsx'
            print('Download %d%%.' % int(status.progress() * 100))
            open(destFilename, 'wb').write(fh.getvalue())
            
    handleDownload()

    
    def handleEmail():
        
        #access gmail api vv

        SCOPES = ['https://www.googleapis.com/auth/gmail.send']

        creds = None
        # The file token.pickle stores the user's access and refresh tokens, and is
        # created automatically when the authorization flow completes for the first
        # time.
        if os.path.exists('token.pickle'):
            with open('token.pickle', 'rb') as token:
                creds = pickle.load(token)
        # If there are no (valid) credentials available, let the user log in.
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(
                    'credentials.json', SCOPES)
                creds = flow.run_local_server(port=0)
            # Save the credentials for the next run
            with open('token.pickle', 'wb') as token:
                pickle.dump(creds, token)

        service = build('gmail', 'v1', credentials=creds)

        subject = f"{dateToday} timesheet"
        sender_email = '[your_email@whatever.com]'
        receiver_emails = "recipient1@whatever.com, recipient2@whatever.com, recipient3@whatever.com"
        message = MIMEMultipart()
        message["From"] = sender_email
        message["To"] = receiver_emails
        message["Subject"] = subject

        fp = open('same/path/to/file/as/above', 'rb')
        msg = MIMEBase('application', 'vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        msg.set_payload(fp.read())
        fp.close()
        msg.add_header('Content-Disposition', 'attachment', filename=f'{dateToday} Timesheet.xlsx')
        encoders.encode_base64(msg)

        message.attach(msg)

        wrappedMess = {'raw': base64.urlsafe_b64encode(message.as_string().encode()).decode()}


        try:
            messageSend = (service.users().messages().send(userId=sender_email, body=wrappedMess)
               .execute())
            print('Message sent, Id: %s' % messageSend['id'])
            return messageSend
        except Exception as e:
            print(e)

    handleEmail()

    stopTimer = timeit.default_timer()
    print(f'Code ran in {stopTimer-startTimer} seconds')

sendTime()
 
