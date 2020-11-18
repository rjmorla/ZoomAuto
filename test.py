import pyautogui
import os
import datetime
import cv2
import time

def runSearch(target):
    r = None
    while r == None:
        r = pyautogui.locateCenterOnScreen(target, grayscale=True, confidence=.8)
    return r

os.startfile(r"C:\Users\roymo\AppData\Roaming\Zoom\bin\zoom.exe")
print('File opened successfully')

#join a meeting button
target = r"C:\myWorkspace\zoomauto\zoomButtons\joinBtn.PNG"
pyautogui.click(runSearch(target))

print('clicked join btn')

target = r"C:\myWorkspace\zoomauto\zoomButtons\meetingIDTextField.PNG"
pyautogui.click(runSearch(target))
pyautogui.write("851 4564 9814")
print('meetingID entered')


print('no password for meeting')
target = r"C:\myWorkspace\zoomauto\zoomButtons\joinMeetingBtn.PNG"
pyautogui.click(runSearch(target))

