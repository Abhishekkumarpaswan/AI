from email.mime import image
import subprocess
import winshell
import cv2
#import wolframclient
import pyttsx3
#import tkinter
import json
import random
import operator
import speech_recognition as sr
import datetime
import wikipedia
import webbrowser
import os
import ctypes
import time
import requests
#import shutil
from bs4 import BeautifulSoup
import win32com.client as wincl
from urllib.request import urlopen

chrome_path = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
  
# First registers the new browser
webbrowser.register('chrome', None, 
                    webbrowser.BackgroundBrowser(chrome_path))
  
assname="Alveno"

def speak(audio):
    engine = pyttsx3.init('sapi5')
    voices = engine.getProperty('voices')
    engine.setProperty('voice', voices[1].id)
    engine.say(audio)
    engine.runAndWait()

def wishme():
    hour = int(datetime.datetime.now().hour)
    if hour >= 0 and hour < 12:
        speak("good morning!")

    elif hour >= 12 and hour < 18:
        speak("good afternoon!")

    else:
        speak("good evening!")

    speak("Tell me how may i help you")
    
    speak("I am your assistant")
    speak(assname)


def takecommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:

        print("listening..")
        r.pause_threshold = 0.8
        audio = r.listen(source)

    try:
        print("Recognizing...")
        query = r.recognize_google(audio, language="en-in")
        print(f"user said:{query}\n")

    except Exception as e:
        print(e)
        speak("Omae wa mou shindeiru. Nani?")
        return "None"
    return query

if __name__ == '__main__':
    clear = lambda: os.system('cls')
    speak("Hello Buddy! how are you?")
    clear()
    
    wishme()
    while True:
        query = takecommand().lower()

        if 'wikipedia' in query:
            speak('searching wikipedia...')
            query = query.replace("wikipedia", "")
            results = wikipedia.summary(query, sentences=2)
            speak("according to wikipedia")
            print(results)
            speak(results)

        elif "open youtube" in query:
            speak("Here you go to youtube\n")
            webbrowser.get('chrome').open("youtube.com")

        elif 'open google' in query:
            webbrowser.get('chrome').open("google.com")

        elif 'open stackoverflow' in query:
            webbrowser.get('chrome').open("stackoverflow.com")

        elif 'play movie' in query:
            movie_dir = 'A:\\'
            movies=os.listdir(movie_dir)
            print(movies)
            choice=int(input("Enter the no. of movie"))
            os.startfile(os.path.join(movie_dir,movies[choice]))

        elif 'the time' in query:
            strTime=datetime.datetime.now().strftime("%H:%M:%S")
            speak(f"Sir, the time is {strTime}")

        elif 'open vs code' in query:
            codepath ="C:\\Users\\ABHISHEK\\AppData\\Local\\Programs\\Microsoft VS Code"
            os.startfile(codepath)

        elif 'opera' in query:
            oppath="C:\\Users\\ABHISHEK\\AppData\\Local\\Programs\\Opera\\launcher.exe"
            os.startfile(oppath)

        elif 'change my name to' in query:
            query = query.replace("change my name to", "")
            assname = query

        elif "change name" in query:
            speak("What would you like to call me, Sir ")
            assname = takecommand()
            speak("Thanks for naming me")

        elif "what's your name" in query or "What is your name" in query:
            speak("My friends call me")
            speak(assname)
            print("My friends call me", assname)

        elif 'exit' in query:
            speak("byee byee I hope we will talk later")
            exit()

        elif "who made you" in query or "who created you" in query:
            speak("I have been created by ABHISHEK.")

        elif 'search' in query or 'play' in query:
            query = query.replace("search", "")
            query = query.replace("play", "")
            webbrowser.get('chrome').open(query)

        elif "who i am" in query:
            speak("If you talk then definately you are  a human.")

        elif "why you came to world" in query:
            speak("Thanks to ABHISHEK. further It's a secret")

        elif 'is love' in query:
            speak("It is 7th sense that destroy all other senses")

        elif "who are you" in query:
            speak("I am your virtual assistant created by ABHISHEK")

        elif 'reason for you' in query:
            speak("I was created as a Minor project by Mister ABHISHEK ")

        elif 'change background' in query:
            ctypes.windll.user32.SystemParametersInfoW(20,0,"Location of wallpaper",0)
            speak("Background changed succesfully")

        elif 'news' in query:

            try:
                jsonObj = urlopen('https://newsapi.org / v1 / articles?source = the-times-of-india&sortBy = top&apiKey =\\times of India Api key\\')
                data = json.load(jsonObj)
                i = 1

                speak('here are some top news from the times of india')
                print('''=============== TIMES OF INDIA ============'''+ '\n')

                for item in data['articles']:
                    print(str(i) + '. ' + item['title'] + '\n')
                    print(item['description'] + '\n')
                    speak(str(i) + '. ' + item['title'] + '\n')
                    i += 1

            except Exception as e:
                print(str(e))

        elif 'lock window' in query:
            speak("locking the device")
            ctypes.windll.user32.LockWorkStation()

        elif 'shutdown system' in query:
            speak("Hold On a Sec ! Your system is on its way to shut down")
            subprocess.call('shutdown /s')

        elif 'empty recycle bin' in query:
            winshell.recycle_bin().empty(confirm = False, show_progress = False, sound = True)
            speak("Recycle Bin Recycled")

        elif "don't listen" in query or "stop listening" in query:
            speak("for how much time you want to stop jarvis from listening commands")
            a = int(takecommand())
            time.sleep(a)
            print(a)

        elif "where is" in query:
            query = query.replace("where is", "")
            location = query
            speak("User asked to Locate")
            speak(location)
            webbrowser.get('chrome').open("https://www.google.nl / maps / place/" + location + "")

        elif "camera" in query or "take a photo" in query:
            speak("Tell how many photos you want to take")
            n=int(input("Enter a number"))
            print(n)
            cam = cv2.VideoCapture(0)
            resuult,image = cam.read()
            for i in range(n-1):
                if image:
                    cv2.imshow("IMAGE BY ALVENO",image)
                    cv2.imwrite("IMAGES BY ALVENO.png",image)
                    cv2.waitKey(0)
                    cv2.destroyWindow("IMAGES BY ALVENO")
                else:
                    print("NO IMAGE IS DETECTED. PLEASE TRY AGAIN !")

        elif "restart" in query:
            subprocess.call(["shutdown", "/r"])

        elif "hibernate" in query or "sleep" in query:
            speak("Hibernating")
            subprocess.call("shutdown / h")

        elif "log off" in query or "sign out" in query:
            speak("Make sure all the application are closed before sign-out")
            time.sleep(5)
            subprocess.call(["shutdown", "/l"])

        elif "write a note" in query:
            speak("What should i write, sir")
            note = takecommand()
            file = open('jarvis.txt', 'w')
            speak("Sir, Should i include date and time")
            snfm = takecommand()
            if 'yes' in snfm or 'sure' in snfm:
                strTime = datetime.datetime.now().strftime("% H:% M:% S")
                file.write(strTime)
                file.write(" :- ")
                file.write(note)
            else:
                file.write(note)

        elif "show note" in query or "so note" in query:
            speak("Showing Notes")
            file = open("jarvis.txt", "r")
            print(file.read())
            speak(file.read(6))
