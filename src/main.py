from googleapiclient.discovery import build
from google.oauth2 import service_account
from pprint import pprint

#!TODO automate entry of the sheet ID and no of ppl 
#!TODO try to get data by using filters 

SA_FILE="keys.json"
SCOPES=["https://www.googleapis.com/auth/spreadsheets"]

creds= None
creds=service_account.Credentials.from_service_account_file(SA_FILE,scopes=SCOPES)

SAMPLE_SHEET_ID="1mL3ExNCHbxHDX0uCEPBzkehPaZjLV-ZxVI9W7AWj2fI"

service=build('sheets','v4',credentials=creds)

sheet=service.spreadsheets()
result= sheet.values().get(spreadsheetId=SAMPLE_SHEET_ID,range="sheet1").execute()

values=result.get('values',[])
# pprint(values)

# writing to a sheet (it will over write he needeed space so part of the previous one may still exist)
# lis=[["rithvik",45],["brad",50],["Shasha",100]]
# request=sheet.values().update(spreadsheetId=SAMPLE_SHEET_ID,range="sheet2",valueInputOption="USER_ENTERED",body={"values":lis}).execute()

# appending to a sheet
# lis_append=[[],[],["Vanilla",100],["chocolate",80],["kesar",10]]
# result=sheet.values().append(spreadsheetId=SAMPLE_SHEET_ID,range="sheet2",valueInputOption="USER_ENTERED",insertDataOption="INSERT_ROWS",body={"values":lis_append}).execute()


#cerating a new spread sheet using code

# spreadsheet_body = {
#             'properties': {
#                 'title': "first"
#             }
#         }
# result=sheet.create(body=spreadsheet_body).execute()  
# print(result) 


# to delete use get Id method to get the sheet you want to delete and then proceed
# req=[
#     {'deleteSheet':{
#             "sheetId":192161416
#         }
#     },
#     {
#     'addSheet': {
#         'properties':{
#             'title':"test sheet"
#         }
#         }
#     }
# ]
# body={
#     'requests': req
# } 
# response = sheet.batchUpdate(
#             spreadsheetId=SAMPLE_SHEET_ID,
#             body=body).execute()


# A brute force way
lis1=[]
lis2=[]
lis3=[]
for item in values:
    year=item[2][2:4]
    if year=="19":
        lis1.append(item)
    if year=="20":
        lis2.append(item)
    if year=="21":
        lis3.append(item)
NO_OF_FOUTH=50
NO_OF_THIRD=40
NO_OF_SECOND=30
NO_OF_FIRST=25

if len(lis1)>NO_OF_FOUTH:
    lis1=lis1[0:NO_OF_FOUTH]
else:
    NO_OF_THIRD+=(NO_OF_FOUTH-len(lis1))
if len(lis2)>NO_OF_THIRD:
    lis2=lis2[0:NO_OF_THIRD]
else:
    NO_OF_SECOND+=(NO_OF_THIRD-len(lis2))
if len(lis3)>NO_OF_SECOND:
    lis3=lis3[0:NO_OF_SECOND]

fin_lis=[["Email ID","Name","Roll No"]]+lis1+[[]]+lis2+[[]]+lis3

req_body={
    'requests':[{
        'addSheet':{
            'properties':{
                'title':"Selected"
            }
        }
    }]
}
response = sheet.batchUpdate(
            spreadsheetId=SAMPLE_SHEET_ID,
            body=req_body).execute()

result=sheet.values().append(spreadsheetId=SAMPLE_SHEET_ID,range="Selected",valueInputOption="USER_ENTERED",insertDataOption="OVERWRITE",body={"values":fin_lis}).execute()


print("Done")