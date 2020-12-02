from __future__ import print_function
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from datetime import datetime

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

# The ID and range of a sample spreadsheet.
SPREADSHEET_ID = '13FKqzuFu-A9_JSsTugiw17dOjlVf-_Zn52caU65qWnk'

# Get credentials from credentials.json or token.pickle
def getCredentials():
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
    return creds

# Update google spreadsheet with the values
def updateSheet(inputvalue, range_):
    service = build('sheets', 'v4', credentials=getCredentials())

    # Call the Sheets API
    print(inputvalue)
    sheet = service.spreadsheets()
    request = sheet.values().update(spreadsheetId=SPREADSHEET_ID, range=range_, valueInputOption='RAW', body=inputvalue)
    response = request.execute()
    print(response)

def colnum_string(n):
    string = ""
    while n > 0:
        n, remainder = divmod(n - 1, 26)
        string = chr(65 + remainder) + string
    return string

# Get column of the date
def getDateColumn():
    service = build('sheets', 'v4', credentials=getCredentials())
    range_ = 'G2:AK2'
    sheet = service.spreadsheets()
    dates = sheet.values().get(spreadsheetId=SPREADSHEET_ID, range=range_).execute()
    values = dates.get('values', [])
    datestr = datetime.now().strftime("%d/%m")
    idx = values[0].index(datestr)
    return colnum_string(idx + 7)
    
# Update value of player currently scanning
def updateValue(player_id):
    row = player_id + 2
    range_ = getDateColumn() + str(row)
    inputvalue = {}
    if datetime.time(datetime.now()) < datetime.time(datetime.strptime("19:30", "%H:%M")):
        inputvalue['values']= [['P']]
    else:
        inputvalue['values']= [['L']]
    inputvalue['majorDimension']="ROWS"
    inputvalue['range']= range_
    updateSheet(inputvalue, range_)

if __name__ == '__main__':
    updateValue(7)
