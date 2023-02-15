from googleapiclient.discovery import build
from google.oauth2 import service_account
import pandas as pd


link_spreadsheat = ''
range = "hoja1!A1:C5"
creds = "credentials.json"

def main():
    spreadsheet_id = getIDfromLink(link_spreadsheat)
    gsheet = getSpreadsheet(creds,spreadsheet_id)

    data = pd.DataFrame(gsheet[1:], columns= gsheet[0])

    print(data)


def getSpreadsheet(service_account_file, spreadsheet_id):
    SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']
    creds = None
    creds = service_account.Credentials.from_service_account_file(
            service_account_file, scopes=SCOPES)

    service = build('sheets', 'v4', credentials=creds)

    # Call the Sheets API
    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=spreadsheet_id,
                                range=range).execute()

    values = result.get('values', [])

    return values

def getIDfromLink(link):
    start = link.find("/d/")+3

    if(start == -1):
        exit("Invalid Link")

    id = link[start: ]
    end = id.find("/")

    if(end == -1):
        exit("Invalid Link")

    id = id[:end]

    return id

if __name__ == '__main__':
    main()