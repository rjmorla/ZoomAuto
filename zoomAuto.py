import os
import pandas as pd
import pyautogui
import time
import numpy as np
import cv2
from datetime import datetime

def runSearch(target):
    coordinates = None
    count = 0
    while coordinates == None and count < 20:
        print('Searching for button')
        coordinates = pyautogui.locateCenterOnScreen(target, grayscale=True, confidence=.7)
        count += 1
    return coordinates

def signIn(meetingID, pw):
    #open zoom app
    os.startfile(r"C:\Users\Roy\AppData\Roaming\Zoom\bin\Zoom.exe")
    print('File opened successfully')
    time.sleep(1)
    #join a meeting button
    
    target = os.path.join(os.path.dirname(__file__), 'zoomButtons\joinBtn.PNG')
    print('Found Button in file')
    pyautogui.click(runSearch(target))
    print('clicked join btn')

    target = os.path.join(os.path.dirname(__file__), 'zoomButtons\meetingIDTextField.PNG')
    time.sleep(1)
    pyautogui.click(runSearch(target))
    time.sleep(1)
    pyautogui.write(meetingID,interval=.1)
    print('meetingID entered')
    time.sleep(1)
    
    #check if class has pw or not
    if (pw == 'nan'):
        print('No password found for meeting')
        
        target = os.path.join(os.path.dirname(__file__), 'zoomButtons\joinMeetingBtn.PNG')
        pyautogui.click(runSearch(target))
    else:
        print('Password found for meeting')
        
        passcode=pyautogui.locateCenterOnScreen(os.path.join(os.path.dirname(__file__), 'zoomButtons\meetingPasscode.PNG'))
        pyautogui.moveTo(passcode)
        pyautogui.write(pw)
        time.sleep(1)
        
        joinWithPassword=pyautogui.locateCenterOnScreen(os.path.join(os.path.dirname(__file__), 'zoomButtons\joinMeetingBtnPassword.PNG'))
        pyautogui.moveTo(joinWithPassword)
        pyautogui.click()
        time.sleep(2)

def signOut():
    print("To be implemented")

def joinBreakoutRoom():
    print("To be implemented")

def parseExcel(df):
    while True:
        currentTime = datetime.now().strftime("%H:%M")
        currentDay = datetime.today().weekday()
        time.sleep(2)

        if currentDay in df['Day'] and currentTime in str(df['Timings']):
            df["Timings"] = pd.to_datetime(df['Timings'], format='%H:%M')
            mylist=df["Timings"]
            mylist=[i.strftime("%H:%M") for i in mylist]
            c= [i for i in range(len(mylist)) if mylist[i]==currentTime]
            row = df.loc[c] 
            meetingID = str(row.iloc[0,2])  
            pw= str(row.iloc[0,3])  
            time.sleep(2)
            signIn(meetingID, pw)
            time.sleep(2)
            print('signed in')
            break

# excelFile = os.path.join(os.path.dirname(__file__), 'zoomSheet\zoomSheet.xlsx')
# df = pd.read_excel(excelFile,index_col=False)
# parseExcel(df)
# print("Excel file read")
