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
    while coordinates == None or count < 5:
        coordinates = pyautogui.locateCenterOnScreen(target, grayscale=True, confidence=.8)
        count += 1
    return coordinates

def signIn(meetingID, pw):
    #open zoom app
    os.startfile(r"C:\Users\roymo\AppData\Roaming\Zoom\bin\zoom.exe")
    print('File opened successfully')

    #join a meeting button
    target = r"C:\myWorkspace\zoomauto\zoomButtons\joinBtn.PNG"
    pyautogui.click(runSearch(target))
    print('clicked join btn')

    target = r"C:\myWorkspace\zoomauto\zoomButtons\meetingIDTextField.PNG"
    pyautogui.click(runSearch(target))
    pyautogui.write(meetingID)
    print('meetingID entered')
    
    #check if class has pw or not
    if (pw == 'nan'):
        print('No password found for meeting')
        target = r"C:\myWorkspace\zoomauto\zoomButtons\joinMeetingBtn.PNG"
        pyautogui.click(runSearch(target))
    else:
        print('Password found for meeting')
        passcode=pyautogui.locateCenterOnScreen(r"C:\myWorkspace\zoomauto\zoomButtons\meetingPasscode.PNG")
        pyautogui.moveTo(passcode)
        pyautogui.write(pw)
        time.sleep(1)

        joinWithPassword=pyautogui.locateCenterOnScreen(r"C:\myWorkspace\zoomauto\zoomButtons\joinMeetingBtnPassword.PNG")
        pyautogui.moveTo(joinWithPassword)
        pyautogui.click()
        time.sleep(2)

def signOut():
    print("To be implemented")

def joinBreakoutRoom():
    print("To be implemented")

def testPassword():
    signIn('test', 'yeppers')

def testNoPassword():
    signIn('test', 'nan')

def parseExcel(df):
    while True:
        currentTime = datetime.now().strftime("%H:%M")
        currentDay = datetime.today().weekday()
        time.sleep(2)

        if currentDay in df['Day'] and currentTime in str(df['Timings']):
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

excelFile = os.path.join(os.path.dirname(__file__), 'zoomSheet\zoomSheet.xlsx')
df = pd.read_excel(excelFile,index_col=False)
parseExcel(df)
print("Excel file read")
