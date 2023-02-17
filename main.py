from googleapiclient.discovery import build
from google.oauth2 import service_account
from email.message import EmailMessage
import pandas as pd
import ssl
import smtplib




link_spreadsheat = ''
range = "hoja1!A1:C6"
message = "Hola {Nombre} tu país es {País}"


email_sender = '@gmail.com'
email_password = ''
creds_spreadsheet = "spreadsheetCreds.json"


def main():
    spreadsheet_id = getIDfromLink(link_spreadsheat)
    gsheet = getSpreadsheet(creds_spreadsheet,spreadsheet_id)
    data = pd.DataFrame(gsheet[1:], columns= gsheet[0])

    for i in data.index:
        dictionary = data.loc[i].to_dict()
        replaced_message = replaceMessage(message, dictionary)
        sendMail(replaced_message, dictionary['Mail'], email_sender,'prueba', email_password)

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
    start = link.find("/d/")

    if(start == -1):
        exit("Invalid Link")

    id = link[start+3: ]
    end = id.find("/")

    if(end == -1):
        exit("Invalid Link")

    id = id[:end]

    return id

def replaceMessage(message: str, dictionary: dict):
    for column in dictionary.keys():
        message = message.replace("{"+column+"}", dictionary[column])
    return message

def sendMail(str_message: str, recipient: str, sender: str, subject: str, password: str):
    message = EmailMessage()
    message.set_content(str_message)
    message['To'] = recipient
    message['From'] = sender
    message['Subject'] = subject
   
    context = ssl.create_default_context()
    with smtplib.SMTP_SSL('smtp.gmail.com', 465, context=context) as smtp:
        smtp.login(sender,password)
        smtp.sendmail(sender,recipient,message.as_string())


if __name__ == '__main__':
    main()