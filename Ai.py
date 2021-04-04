from win32com.client import Dispatch
import datetime
import speech_recognition as sr
import wikipedia
import webbrowser
import os
import smtplib
import random

def speak(str):
    speak = Dispatch(("SAPI.SpVoice"))
    speak.Speak(str)

def WishMe():
    hour = int(datetime.datetime.now().hour)
    if hour>=0 and hour<12:
        print ("Good Morning Sir\n")
        speak("Good Morning Sir")

    elif hour>=12 and hour<18:
        print("Good Afternoon Sir\n")
        speak("Good Afternoon Sir")

    else:
        print("Good Evening Sir\n")
        speak("Good Evening Sir")

    print(" I Am Turbo How May I help You\n")
    speak("I Am Turbo How May I help You")

def takeCommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print('Listening....')
        r.pause_threshold = 1
        audio = r.listen(source)
 
    try:
        print("Recognising....")
        query = r.recognize_google(audio, language='en-in')
        print(f"User Said:  {query}\n")

    except Exception as e:
        print("Say that again... ")
        return("none")
    return query

def sendEmail(to, content):
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.ehlo()     
    server.starttls()
    server.login("Your_Email", "Your_Password")   
    server.sendmail("Your_Email", to, content)
    server.close()                                           



if __name__ == "__main__":
    WishMe()
    while True:
        query = takeCommand().lower()

        #Logic
        if 'wikipedia' in query:
            print("searching...")
            speak('searching')
            query = query.replace("wikipedia", "")
            result = wikipedia.summary(query, sentences= 3)
            print(f"According to Wikipedia... {result}")
            
            speak("According to Wikipedia")
            speak(result)

        elif 'open youtube' in query:
            webbrowser.open("youtube.com")

        elif 'open google' in query:
            webbrowser.open("google.com")

        elif 'open twitter' in query:
            webbrowser.open("twitter.com")

        elif 'open github' in query:
            webbrowser.open("github.com")

        elif 'open email' in query:
            webbrowser.open("gmail.com")

        elif 'open pixabay' in query:
            webbrowser.open("pixabay.com")

        elif 'the time' in query:
            strTime = datetime.datetime.now().strftime('%H:%M:%S')
            print(strTime)
            speak(strTime)
        
        elif 'open code' in query:
            codePath = "E:\\Microsoft VS Code\\Code.exe"
            os.startfile(codePath)

        elif 'open opera' in query:
            operaPath ="C:\\Program Files (x86)\\Opera\\launcher.exe"
            os.startfile(operaPath)

        elif 'open chrome' in query:
            chromePath = "C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe"
            os.startfile(chromePath)

        elif 'play music' in query:
            music_dir = "D:\\Music\\"
            songs = os.listdir(music_dir)
            ran = random.randint(1, 10)
            os.startfile(os.path.join(music_dir, songs[ran]))


        elif 'email to friend'in query:
            try:
                print("What Should I Say")
                speak("What Should I Say")
                content = takeCommand()
                to = 'friend@email.com'
                sendEmail(to, content)
                print("Email Has Been Sent")
                speak("Email Has Been Sent")
            except Exception as k:
                print("I am Sorry Sir, I Failed To Sent Email\n")
                speak("I AM SORRY SIR, I FAILED TO SENT EMAIL")

        elif 'email to daddy' in query:
            try:
                print("What Should I Say")
                speak("What Should I Say")
                content = takeCommand()
                to = 'daddy@email.com'
                sendEmail(to, content)
                print("Email Has Been Sent")
                speak("Email Has Been Sent")
            except Exception as j:
                print("I am Sorry Sir, I Failed To Sent Email\n")
                speak("I AM SORRY SIR, I FAILED TO SENT EMAIL")

        elif 'email to mom'in query:
            try:
                print("What Should I Say")
                speak("What Should I Say")
                content = takeCommand()
                to = 'mom@Email.com'
                sendEmail(to, content)
                print("Email Has Been Sent")
                speak("Email Has Been Sent")
            except Exception as k:
                print("I am Sorry Sir, I Failed To Sent Email\n")
                speak("I AM SORRY SIR, I FAILED TO SENT EMAIL")

        elif 'who are you' in query:
            print("I am Turbo, Aryan created me. How Can I Help You\n")
            speak("I am Tarbo, Aaryan created me. How Can I Help You")

        elif 'who is aryan' in query:
            print("Aryan Is My Creator. He Is The Programmer\n")
            speak("Aaryan Is My Creator. He Is The Programmer")
        
        elif 'who is my crush' in query:
            print("crush_name\n")
            speak("Crush_name")

        elif 'Who is turbo' in query:
            print("I Am Turbo Sir")
            speak("I AM TURBO SIR")

        elif 'open notepad' in query:
            notePath = "C:\\Windows\system32\\notepad.exe"
            os.startfile(notePath)

        elif 'open audacity' in query:
            audaPath = "E:\\Audacity\\audacity.exe"
            os.startfile(audaPath)

            