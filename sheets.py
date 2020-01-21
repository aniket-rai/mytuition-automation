from __future__ import print_function
import pickle
import os.path
from os.path import basename
from googleapiclient.discovery import build
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow

import smtplib, ssl, email
from email import encoders
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication


# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']

# The ID and range of the spreadsheet.
SPREADSHEET_ID = '1BgsuLlAtDi0oAt62C8ZhzDL45t82xMZLsr54ZShGTUA'
RANGE_NAME = 'Sheet1!A2:C8'

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
    result = sheet.values().get(spreadsheetId=SPREADSHEET_ID,
                                range=RANGE_NAME).execute()
    values = result.get('values', [])
    
    dataFile = open('teamleads.txt', 'w')

    if not values:
        print('No data found.')
    else:
        for row in values:
            # Write to a file for later processing
            dataFile.write('{0},{1},{2}\n'.format(row[0], row[1], row[2]))
    
    dataFile.close()
            
def teamleadProcessing():
    # Var
    teamleads = []
    passwords = []
    pMessage  = []
    
    # Open teamlead information file
    dataFile = open('teamleads.txt', 'r')
    
    # Sort through file and add to respective arrays
    for line in dataFile:
        data = line.split(',')
        teamleads.append(data[0])
        passwords.append(data[1])
        pMessage.append(data[2])
    
    # Process tutors for each teamlead
    for counter in range(len(teamleads)):
        tutorProcessing(teamleads[counter], passwords[counter], pMessage[counter])
    
def tutorProcessing(teamlead, password, pMessage):
    # Var
    tutors = []
    personalisedMSG = []
    
    
    # Get all tutors and sort them correctly 
    teamleadName = teamlead.split(".")
    teamleadLast = teamleadName[1].capitalize().strip()
    teamleadLast = teamleadLast[:-10]
    teamleadName = teamleadName[0].capitalize().strip()
    tutorFile = open(teamleadName + ".txt", 'r')
    tutors = tutorFile.readlines()
        
    for tutor1 in tutors:
        tutor = tutor1.split(",")
        tutorName = tutor[0]
        tutorEmail = tutor[1].strip()
        
        message = MIMEMultipart()
        message["Subject"] = "MyTuition Monthly Feedback Report"
        message["From"] = teamleadName + " " + teamleadLast
        message["To"] = tutorEmail
        
        # Create the plain-text and HTML version of your message
        greeting = "Hey " + tutorName + ",\n\n"
        body = """Please find your monthly feedback report attached below. Please let us know if you have any questions or want to discuss this report further! :)\n\n"""
        conclusion = "Cheers,\n" + teamleadName + "\n"
        text = greeting + pMessage + "\n" + body + conclusion
        
        # Add plain-text parts to MIMEMultipart message
        message.attach(MIMEText(text, "plain"))
        filename = "blank.pdf"

        with open(filename, "rb") as fil:
            part = MIMEApplication(
                fil.read(),
                Name=basename(filename)
            )
        
        # After the file is closed
        part['Content-Disposition'] = 'attachment; filename="%s"' % basename(filename)
        message.attach(part)

        sendEmail(teamlead, password, tutorEmail, message.as_string())
        
def sendEmail(sender, password, receiver, message):
    # Create secure connection with server and send email
    context = ssl.create_default_context()
    with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
        server.login(sender, password)
        server.sendmail(sender, receiver, message)
                            
def main():
    getData()
    teamleadProcessing()

if __name__ == '__main__':
    main()
