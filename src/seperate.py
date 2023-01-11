import gspread
import math
from pprint import pprint

def create_sheet(basic_info_list,event_list,name):
    ws_new=gs.add_worksheet(title=name,rows=1000,cols=26)
    ws_new.batch_update([
    {
        'range':'A:C',
        'values':basic_info_list
    },
    {
        'range':'D:E',
        'values':event_list
    },
    {
        'range':'F:F',
        'values':[['Hours(3-5)']]
    }])
    return ws_new

def get_index_list(event_list,N):
    count=len(event_list)-event_list.count([])-1
    size_list=[]
    size= math.floor(count/N)
    rem=count-size*N
    i=0
    while i<N:
        size_list.append(size)
        i=i+1
    i=0
    while rem>0:
        size_list[i]+=1
        i+=1
        rem-=1
    index_list=[1]
    idx=0
    for i in size_list:
        count=0
        while i>0:
            idx+=1
            if event_list[idx]!=[]:
                count+=1
                i-=1
            
        index_list.append(idx+1)
    return index_list

def format_sheet(ws,index_list,N):
    # colours for 5 evaluators
    r=[0.0,1.0,0.9,0.6,0.9]
    g=[0.9,0.6,0.9,0.6,0.0]
    b=[0.9,0.0,0.0,0.9,0.9]
    i=0
    while i<N:

        range="A"+str(index_list[i])+":F"+str(index_list[i+1])
        print(range)
        format={
            "backgroundColor": 
            {
            "red": r[i],
            "green": g[i],
            "blue": b[i]
            },
            "wrapStrategy": "CLIP"
        }
        ws.format(range,format )
        i+=1


gc = gspread.service_account('keys.json')

# Open a sheet from a spreadsheet in one go
# rithvik@gsheetaccess6300.iam.gserviceaccount.com
# Gsheet format is 3 cols of basic info and 2 cols for each event
gs = gc.open_by_url('https://docs.google.com/spreadsheets/d/1tp_aT7d59zsApcoNsxpemf8ThFpMJVVEhz30jL4iWTs/edit#gid=174387397')

ws=gs.get_worksheet(1)
basic_info_list= ws.batch_get(['A:C'])[0]
no_of_events=3
no_of_evaluators=4
i=0
c='D'
while i<no_of_events:
    col_name=c+":"+chr(ord(c)+1)
    name="Event-"+str(i+1)
    event_list=ws.batch_get([col_name])[0]
    ws_new=create_sheet(basic_info_list,event_list,name)
    idx_list=get_index_list(event_list,no_of_evaluators)
    format_sheet(ws_new,idx_list,no_of_evaluators)
    c=chr(ord(c)+2)
    i+=1

total_ws=gs.add_worksheet(title="Total Hours",rows=1000,cols=26)
c='D'
range=c+":"+chr(ord(c)+no_of_events+1)
hours_list=[]
i=0
while i<no_of_events:
    hours_list.append("Event-"+str(i+1)+"(3-5)")
    i+=1
hours_list.append("Total Hours")
hours_list.append("Remarks")
total_ws.batch_update([
    {
        'range':'A:C',
        'values':basic_info_list
    },
    {
        'range':range,
        'values':[hours_list]
    }])

print('Done')
