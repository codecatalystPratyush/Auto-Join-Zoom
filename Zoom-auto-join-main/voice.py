import subprocess
import pyttsx3
import pywhatkit
import speech_recognition as sr
import datetime
import wikipedia
import webbrowser
import os
import winshell
import pyjokes
import ctypes
import time
import requests
import shutil
from clint.textui import progress
from bs4 import BeautifulSoup
import win32com.client as wincl
from urllib.request import urlopen


listener = sr.Recognizer()
engine = pyttsx3.init()
voices = engine.getProperty("voices")
engine.setProperty("voice", voices[0].id)


def talk(text):
    engine.say(text)
    engine.runAndWait()

def wishMe():
    hour = int(datetime.datetime.now().hour)
    if hour>= 0 and hour<12:
        talk("Good Morning Sir !")
  
    elif hour>= 12 and hour<18:
        talk("Good Afternoon Sir !")  
  
    else:
        talk("Good Evening Sir !") 
  
    assname = "Amit Assistant 1 point o"
    talk("I am your Assistant")
    talk(assname)
 
def take_command():
    try:
        with sr.Microphone() as source:
            print('listening...')
            listener.pause_threshold = 1
            listener.adjust_for_ambient_noise(source, duration=1)
            voice = listener.listen(source)
            info = listener.recognize_google(voice)
            print(info)
            return info.lower()
    except:
     pass

def usrname():
    talk("What should i call you sir")
    uname = take_command()
    talk("Welcome Mister")
    talk(uname)
    columns = shutil.get_terminal_size().columns
     
    print("###############################################################################################################################".center(columns))
    print("Welcome Mr.")
    print("###############################################################################################################################".center(columns))
     
    talk("How can i Help you, Sir")    

if __name__ == '__main__':
    clear = lambda: os.system('cls')
     
    clear()
    wishMe()
    usrname()    

def run_alexa():
    command = take_command()
    print(command)
    if "play" in command:
        song = command.replace("play", "")
        talk("playing " + song)
        pywhatkit.playonyt(song)
    elif "time" in command:
        time = datetime.datetime.now().strftime("%I:%M %p")
        talk("Current time is " + time)
    elif "who the heck is" in command:
        person = command.replace("who the heck is", "")
        info = wikipedia.summary(person, 1)
        print(info)
        talk(info)
    elif "date" in command:
        talk("sorry, I have a headache")
    elif "are you single" in command:
        talk("I am in a relationship with wifi")
    elif "joke" in command:
        talk(pyjokes.get_joke())
    elif 'open youtube' in command:
        talk("Here you go to Youtube\n")
        webbrowser.open("youtube.com")

    elif "can we be friends" in command:
        talk("Of course, you are like the Earth to my moon. We make a great team")      
 
    elif 'open google' in command:
        talk("Here you go to Google\n")
        webbrowser.open("google.com")
 
    elif 'open stackoverflow' in command:
        talk("Here you go to Stack Over flow.Happy coding")
        webbrowser.open("stackoverflow.com")  
 
    elif 'the time' in command:
        strTime = datetime.datetime.now().strftime("% H:% M:% S")   
        talk(f"Sir, the time is {strTime}")
 
    elif 'open viber' in command:
        codePath = r"C:\Users\Sujit\AppData\Local\Viber\Viber.exe"
        os.startfile(codePath)

    elif 'open teamviewer' in command:
        codePath = r"C:\Program Files (x86)\TeamViewer\TeamViewer.exe"
        os.startfile(codePath)

    elif 'open pycharm' in command:
        codePath = r"C:\Program Files\JetBrains\PyCharm Community Edition 2021.1.3\bin\pycharm64.exe"
        os.startfile(codePath)

    elif 'open chrome' in command:
        codePath = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
        os.startfile(codePath)

    elif 'open word' in command:
        codePath = r"C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE"
        os.startfile(codePath)

    elif 'open excel' in command:
        codePath = r"C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE"
        os.startfile(codePath)

    elif 'open 1 note' in command:
        codePath = r"C:\Program Files\Microsoft Office\root\Office16\ONENOTE.EXE"
        os.startfile(codePath)

    elif 'open edge' in command:
        codePath = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
        os.startfile(codePath)

    elif 'open powerpoint' in command:
        codePath = r"C:\Program Files\Microsoft Office\root\Office16\POWERPNT.EXE"
        os.startfile(codePath)

    elif 'open kaspersky' in command:
        codePath = r"C:\Program Files (x86)\Kaspersky Lab\Kaspersky Internet Security 21.3\avpui.exe"
        os.startfile(codePath) 

    elif 'open vscode' in command:
        codePath = r"C:\Users\Sujit\AppData\Local\Programs\Microsoft VS Code\Code.exe"
        os.startfile(codePath)                

    elif 'open virtualdj' in command:
        codePath = r"C:\Program Files\VirtualDJ\virtualdj.exe"
        os.startfile(codePath)

    elif 'how are you' in command:
        talk("I am fine, Thank you")
        talk("How are you, Sir")
 
    elif 'fine' in command or "good" in command:
        talk("It's good to know that your fine")
 
    elif "change my name to" in command:
        command = command.replace("change my name to", "")
        Pratyush = command
 
    elif "change name" in command:
        talk("What would you like to call me, Sir ")
        assname = take_command()
        talk("Thanks for naming me")
 
    elif "what's your name" in command or "What is your name" in command:
        talk("My friends call me the Amit Assistant")
 
    elif 'exit' in command:
        talk("Thanks for giving me your time")
        exit()
 
    elif "who made you" in command or "who created you" in command:
        talk("I have been created by Amit.")
             
    elif 'joke' in command:
        talk(pyjokes.get_joke())
             
    elif 'search' in command or 'play' in command:
             
        command = command.replace("search", "")
        command = command.replace("play", "")         
        webbrowser.open(command)
 
    elif "who i am" in command:
        talk("If you talk then definitely your human.")

    elif "why you came to world" in command:
        talk("Thanks to Amit. further It's a secret")
 
    elif 'power point presentation' in command:
        talk("opening Power Point presentation")
        power = r"C:\Users\Sujit\OneDrive\Documents\Microsoft Word\Presentation1"
        os.startfile(power)

    elif 'is love' in command:
        talk("It is 7th sense that destroy all other senses")
 
    elif "who are you" in command:
        talk("I am your virtual assistant created by Amit")
 
    elif 'reason for you' in command:
        talk("I was created as a Minor project by Mister Amit ")
 
    elif 'change background' in command:
        ctypes.windll.user32.SystemParametersInfoW(20,
                                                   0,
                                                   "Location of wallpaper",
                                                   0)
        talk("Background changed successfully")
 
    elif 'open chrome' in command:
        appli = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
        os.startfile(appli)
 
         
    elif 'lock window' in command:
            talk("locking the device")
            ctypes.windll.user32.LockWorkStation()
 
    elif 'shutdown system' in command:
            talk("Hold On a Sec ! Your system is on its way to shut down")
            subprocess.call('shutdown / p /f')
                 
    elif 'empty recycle bin' in command:
        winshell.recycle_bin().empty(confirm = False, show_progress = False, sound = True)
        talk("Recycle Bin Recycled")
 
    elif "don't listen" in command or "stop listening" in command:
        talk("for how much time you want to stop jarvis from listening commands")
        a = int(take_command())
        datetime.sleep(a)
        print(a)
 
    elif "where is" in command:
        command = command.replace("where is", "")
        location = command
        talk("User asked to Locate")
        talk(location)
        webbrowser.open("https://www.google.nl / maps / place/" + location + "")
 
    elif "restart" in command:
        subprocess.call(["shutdown", "/r"])
             
    elif "hibernate" in command or "sleep" in command:
        talk("Hibernating")
        subprocess.call("shutdown / h")
 
    elif "log off" in command or "sign out" in command:
        talk("Make sure all the application are closed before sign-out")
        datetime.sleep(5)
        subprocess.call(["shutdown", "/l"])
 
    elif "write a note" in command:
        talk("What should i write, sir")
        note = take_command()
        file = open('amit kumar datta.txt', 'w')
        talk("Sir, Should i include date and time")
        snfm = take_command()
        if 'yes' in snfm or 'sure' in snfm:
            strTime = datetime.datetime.now().strftime("% H:% M:% S")
            file.write(strTime)
            file.write(" :- ")
            file.write(note)
        else:
            file.write(note)
         
    elif "show note" in command:
        talk("Showing Notes")
        file = open("amit kumar datta.txt", "r")
        print(file.read())
        talk(file.read(6))
 
    elif "update assistant" in command:
        talk("After downloading file please replace this file with the downloaded one")
        url = '# url after uploading file'
        r = requests.get(url, stream = True)
             
        with open("Voice.py", "wb") as Pypdf:
                 
            total_length = int(r.headers.get('content-length'))
                 
            for ch in progress.bar(r.iter_content(chunk_size = 2391975),
                                   expected_size =(total_length / 1024) + 1):
                if ch:
                  Pypdf.write(ch)
                     
    elif "jarvis" in command:
             
        wishMe()
        talk("Jarvis 1 point o in your service Mister")
        talk("Pratyush")
 
    elif "wikipedia" in command:
        webbrowser.open("wikipedia.com")
 
    elif "Good Morning" in command:
        talk("A warm" +command)
        talk("How are you Mister")
        talk('assname')

    elif "will you be my gf" in command or "will you be my bf" in command:  
        talk("I'm not sure about, may be you should give me some time")
 
    elif "how are you" in command:
        talk("I'm fine, glad you me that")
 
    elif "i love you" in command:
        talk("It's hard to understand")

    elif "turn on the Bedroom Light" in command:
        talk("hey lazy man turn on your bedroom light")

    elif "What is the loneliest number" in command:
        talk("I would imagine the number quinnonagintillion is pretty lonely. I mean, how often does it even get used")        

    elif 'search youtube in google' in command:
        talk("searching...\n")
        webbrowser.open("https://www.google.com/search?q=youtube&sxsrf=ALeKk01tCqpcwJDLkYC1CPgzIQSV7ae_FQ%3A1627544445363&ei=fVsCYeGyFZm94-EPrY2i0As&oq=youtube&gs_lcp=Cgdnd3Mtd2l6EAMyBAgjECcyBAgjECcyBAgjECcyCggAELEDEIMBEEMyBQgAEJECMgUIABCRAjIFCAAQkQIyBAgAEEMyCggAELEDEIMBEEMyBAgAEEM6BwgjEOoCECc6DQguEMcBEK8BEOoCECc6BQguEJECOgsIABCABBCxAxCDAToFCAAQgAQ6CAgAEIAEELEDOggILhCABBCxA0oECEEYAFCdsANYjtoDYPLdA2gBcAJ4AIABhAKIAbcMkgEFMC40LjSYAQCgAQGwAQrAAQE&sclient=gws-wiz&ved=0ahUKEwih5ZXE44fyAhWZ3jgGHa2GCLoQ4dUDCA8&uact=5")

while True:
    try:
        run_alexa()
    except:
        print("Sorry . Unavailable to execute the command now")
        time.sleep(10)