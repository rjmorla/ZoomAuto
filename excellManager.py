import numpy as np
import pandas as pd
import os

path = os.path.join(os.path.dirname(__file__), 'zoomSheet\zoomSheet.xlsx')

def insertEntry(day, time, meetingID, pw):
    if (pw == ''):
        pw = np.nan
    df = pd.DataFrame({
        'Day': [day],
        'Timings': [time],
        'MeetingID': [meetingID],
        'Passcode': [pw]
    })
    old_df = pd.read_excel(path, 'Sheet1')
    new_df = pd.concat([old_df, df])
    new_df.to_excel(path, 'Sheet1', index=False)

def deleteEntry(day, meetingID):
    df = pd.read_excel(path, 'Sheet1')
    df.drop(df[(df.MeetingID == meetingID) & (df.Day == day)].index, inplace=True)
    print(df.Timings)
    df.to_excel(path, 'Sheet1', index=False)

def editEntry(day, time, meetingID):
    print("To Be Implemented")