# This Programme will convert your typed text to speech using win32com.client module
#first you need to pip win32com.client packed using pip install pypiwin32

import win32com.client

speaker= win32com.client.Dispatch("SAPI.SpVoice")

while True:
    s=input("Please Enter your text : ")

    speaker.Speak(s)

    repeat=input("Do you want to try again Y/N : ")
    if repeat=="y" or repeat=="Y":
        continue
    else:
        break
