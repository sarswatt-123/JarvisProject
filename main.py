'''import speech_recognition
import win32com.client

speaker= win32com .client.Dispatch("SAPI.SpVoice")

while 1:
    print("Enter the word u want to speak it out by computer ")
    s=input()
    speaker.speak(s)'''

import pyttsx3
import datetime
import speech_recognition as sr
import wikipedia
import webbrowser
import os

engine= pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
print(voices[1].id)
engine.setProperty('voice',voices[0].id)

def speak(audio):
    engine.say(audio)
    engine.runAndWait()

def wishMe():
    hour = int(datetime.datetime.now().hour)
    if hour>=0 and hour <12:
        speak('Good Morning')

    elif hour>=12 and hour<18:
        speak('Goof afternoon')

    else:
        speak("Good evening")

    speak("I am Jarvis Sir Please tell me  How may i help you ")

def takecommand():
    # it takes microphone  input  from yhe user and return strings output
    r=sr.Recognizer()
    with sr.Microphone() as source:
        print("listening...")
        r.pause_threshold = 1
        audio = r.listen(source )

    try:
        print('Recognizing...')
        query=r.recognize_google(audio,language='en-in')
        print(f"user said:{ query}\n")

    except Exception  as e:
        print("say that again please ...")
        return "none"
    return query


if __name__== "__main__":
    wishMe()
    while True:
        query = takecommand().lower()
        if 'wikipedia' in query :
            speak('searching wikipedia....')
            query = query.replace('wikipedia',"")
            results= wikipedia.summary(query,sentence=2)
            speak('according to wikipedia')
            print(results)
            speak(results)

        elif 'open youtube' in query :
            webbrowser.open("Youtube.com")

        elif 'open google' in query:
            webbrowser.open("google.com")

        elif 'open stackoverflow' in query:
            webbrowser.open("stackoverflow.com")

        elif 'the time ' in query :
            strTime= datetime.datetime.now().strftime("%H:%M:%S")
            speak(f"sir,the time is {strTime}")

        elif ' open code' in query :
            codePath="C:\\Program Files\\JetBrains\\PyCharm Community Edition 2023.1\\New folder\\bin\\PyCharm Community Edition 2023.1\\bin\\pycharm64.exe"
            os.startfile(codePath)
