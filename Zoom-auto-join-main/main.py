import pyautogui
import subprocess
import datetime
import time


def join(id, password):
    subprocess.call(r'C:/Users/anura/AppData/Roaming/Zoom/bin/Zoom.exe')
    while True:
        join1 = pyautogui.locateOnScreen('img_4.png')
        if join1 != None:
            pyautogui.click(join1)
            print("Clicked")
            break
        else:
            print("Could not found")
            time.sleep(2)

    while True:
        join2 = pyautogui.locateOnScreen('img.png')
        if join2 != None:
            pyautogui.click(join2)
            print("Made filed active")
            pyautogui.typewrite(id)
            break
        else:
            print("Could not found")
            time.sleep(2)

    while True:
        join3 = pyautogui.locateOnScreen('img_1.png')
        if join3 != None:
            pyautogui.click(join3)
            print("Made filed active")
            pyautogui.typewrite(id)
            pyautogui.click(pyautogui.locateOnScreen('img_1.png'))
            break
        else:
            print("Could not find the input field")
            time.sleep(2)

    while True:
        join4 = pyautogui.locateOnScreen('img_2.png')
        if join4 != None:
            pyautogui.click(join4)
            print("Made filed 2 active")
            pyautogui.typewrite(password)
            pyautogui.click(pyautogui.locateOnScreen('img_3.png'))
            break
        else:
            print("Could not find the input field 2")
            time.sleep(2)


while True:
    now = datetime.datetime.now()
    current_time = (now.strftime("%m-%d-%y %H:%M"))

    if current_time == "18-01-2022 11:16":
        join("797 4034 2566", "SANS181")
        break
    else:
        print("not the time")
        time.sleep(10)