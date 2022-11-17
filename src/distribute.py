from googleapiclient.discovery import build
from google.oauth2 import service_account
from pprint import pprint

#!TODO automate entry of the sheet ID and no of ppl 
#!TODO try to get data by using filters 

SA_FILE="keys.json"
SCOPES=["https://www.googleapis.com/auth/spreadsheets"]

creds= None
creds=service_account.Credentials.from_service_account_file(SA_FILE,scopes=SCOPES)

SAMPLE_SHEET_ID="1uM1QB4R3q41fE9okszs7UJzfHWtpZFLTwPVeCd4pgas"

service=build('sheets','v4',credentials=creds)

sheet=service.spreadsheets()
result= sheet.values().get(spreadsheetId=SAMPLE_SHEET_ID,range="copy of sheet1!D1:D").execute()

values=result.get('values',[])
values.pop(0)
# pprint(values)
count=len(values)-values.count([])

body = {
  "requests": [
    {
      "updateCells": {
        "range": {
          "sheetId": 0,
          "startRowIndex": 0,
          "endRowIndex": 1,
          "startColumnIndex": 0,
          "endColumnIndex": 1
        },
        "rows": [
          {
            "values": [
              {
                "userEnteredFormat": {
                  "backgroundColor": {
                    "red": 1
                  }
                }
              }
            ]
          }
        ],
        "fields": "userEnteredFormat.backgroundColor"
      }
    }
  ]
}
# res = service.spreadsheets().batchUpdate(spreadsheetId=SAMPLE_SHEET_ID, body=body).execute()
print("Done")