from __future__ import print_function
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

import smtplib

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']

# The ID and range of a sample spreadsheet.
SAMPLE_SPREADSHEET_ID = '1BgsuLlAtDi0oAt62C8ZhzDL45t82xMZLsr54ZShGTUA'
SAMPLE_RANGE_NAME = 'Sheet1!A2:B8'

def getData():
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

    service = build('sheets', 'v4', credentials=creds)

    # Call the Sheets API
    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                                range=SAMPLE_RANGE_NAME).execute()
    values = result.get('values', [])
    
    dataFile = open('data.txt', 'w')

    if not values:
        print('No data found.')
    else:
        for row in values:
            # Write to a file for later processing
            dataFile.write('{0},{1}\n'.format(row[0], row[1]))
    
    dataFile.close()
            
def emailSending():
    # Garbage Var
    x = 0
    subject = []
    text = []
    
    # Gmail Sign In
    sender = 'aniket.rai@mytuition.co.nz'
    passwd = 'devdeep71'
    recipient = 'aniket@assemblyltd.com'
    
    # Server setup
    server=smtplib.SMTP('smtp.gmail.com', 587)
    server.ehlo()
    server.starttls()
    server.login(sender, passwd)
    
    # Subject and Body Setup
    dataFile = open('data.txt', 'r')
    
    for line in dataFile:
        data = line.split(',')
        subject.append(data[0])
        text.append(data[1])
        
    for teamlead in subject:
        body = '\r\n'.join(['To: %s' % recipient,
                            'From: %s' % sender,
                            'Subject: %s' % teamlead,
                            '', text[x]])
        try:
            server.sendmail(sender, [recipient], body)
            print('email sent')
            x += 1
        except:
            print('error occurred')
            
    
    server.quit()
    
                            
def main():
    getData()
    emailSending()

if __name__ == '__main__':
    main()
