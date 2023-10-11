import pyttsx3
import speech_recognition as sr
import win32com.client
import os
import webbrowser
import pywhatkit
import wikipedia
import pyautogui
import keyboard
from keyboard import press
from keyboard import press_and_release
from keyboard import write
import pyjokes
from PyDictionary import PyDictionary as Diction


speaker = win32com.client.Dispatch("SAPI.SpVoice")

Assistant = pyttsx3.init('sapi5')
voices = Assistant.getProperty('voices')
print(voices)
Assistant.setProperty('voices', voices[0].id)
Assistant.setProperty('rate', 200)


def Speak(audio):
    print("   ")
    Assistant.say(audio)
    print(f":   {audio} ")
    Assistant.runAndWait()


def takeCommand():
    command = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening.....")
        command.pause_threshold = 1
        audio = command.listen(source)

        try:
            print("Recognizing....")
            query = command.recognize_google(audio, language='en-in')
            print(f"you said: {query}")

        except Exception as error:
            return "none"

        return query.lower()


def TaskExe():

    def Music():
        Speak("Tell me the name of music")
        musicName = takeCommand()
        pywhatkit.playonyt(musicName)
        Speak("Enjoy your song sir")

    def Song():
        Speak("Wait playing the song sir")
        songname = takeCommand()
        pywhatkit.playonyt(songname)
        Speak("Enjoy your song sir")

    def OpenApps():
        Speak("ok sir , Wait a second!")

        if 'open instagram' in query:
            webbrowser.open('https://www.instagram.com/')

        elif 'open facebook' in query:
            webbrowser.open('https://www.facebook.com/')

        elif 'open whatsapp' in query:
            webbrowser.open('https://web.whatsapp.com/')

    def Dict():
        Speak("tell me the problem")
        probl = takeCommand()

        if 'meaning' in probl:
            probl = probl.replace("what is the meaning", "")
            probl = probl.replace("jarvis", "")
            probl = probl.replace("of", "")
            result = Diction.meaning(probl)
            Speak(f"the meaning of {probl} is {result}")

        elif 'synonym' in probl:
            probl = probl.replace("what is the synonym", "")
            probl = probl.replace("jarvis", "")
            probl = probl.replace("of", "")
            result = Diction.meaning(probl)
            Speak(f"the synonym of {probl} is {result}")

    def YoutubeAuto(command):

        query=str(command)

        if 'pause video' in query:
            keyboard.press('Space bar')
            Speak("Your video is paused sir")

        elif 'restart video' in query:
            keyboard.press('0')
            Speak("your video is restarted sir!")

        elif 'mute video' in query:
            keyboard.press('m')
            Speak("your video is muted sir!")

        elif 'skip video' in query:
            keyboard.press('1')
            Speak("video is skipped sir")

        elif 'full screen' in query:
            keyboard.press('f')
            Speak("done sir")

    def ChromeAuto(command):

            query = str(command)

            if 'new tab' in query:

                press_and_release('ctrl+t')

            elif 'close tab' in query:

                press_and_release('ctrl+w')

            elif 'history' in query:

                press_and_release('ctrl+h')

            elif 'download' in query:

                press_and_release('ctrl+j')

            elif 'incognito' in query:

                press_and_release('ctrl+shift+n')

    while True:

        query = takeCommand()

        if 'hello' in query:
            Speak("hello sir,Im jarvis, Roni personal AI assistant")
            Speak("how may I help you?")

        elif 'how are you' in query:
            Speak("Iam fine sir ")
            Speak("what about you")

        elif 'creator' in query:
            Speak("Roni he is a god and he is my creator")
            Speak("But I can help you in many ways because my master said to help the poor")

        elif 'bye' in query:
            Speak("Bye sir")
            Speak("See you again,hope i made you happy")

        elif 'youtube search' in query:
            Speak("This is what I found for you sir")
            query = query.replace("jarvis", "")
            query = query.replace("youtube search", "")
            web = "https://www.youtube.com/results?search_query=" +query
            webbrowser.open(web)
            Speak("Done sir")

        elif 'google search' in query:
            Speak("This is what I found for you sir")
            query = query.replace("jarvis", "")
            query = query.replace("google search", "")
            pywhatkit.search(query)
            Speak("Done sir")

        elif 'website' in query:
            Speak("ok sir launching....")
            query = query.replace("jarvis", "")
            query = query.replace("website ", "")
            web1 = query.replace("open", "")
            web2 = 'https://wwww.' + web1 + '.com'
            webbrowser.open(web2)
            Speak("Launched")

        elif 'launch' in query:
            Speak("Tell me the name of the website sir")
            name = takeCommand()
            web = 'https://www.' + name + '.com'
            webbrowser.open(web)
            Speak("Launched sir")

        elif 'music' in query:
            Speak("wait Playing the song sir!")
            Music()

        elif 'song' in query:
            Speak("Tell me the name of music")
            Song()

        elif 'wikipedia' in query:
            Speak("Searching in wikipedia....")
            query = query.replace("jarvis", "")
            query = query.replace("wikipedia search", "")
            wiki = wikipedia.summary(query,5)
            Speak(f"According to wikipedia : {wiki}")

        elif 'open instagram' in query:
            Speak("Opening instagram sir")
            OpenApps()

        elif 'open facebook' in query:
            Speak("Opening facebook sir")
            OpenApps()

        elif 'open whatsapp' in query:
            Speak("Opening whatsapp sir")
            OpenApps()

        elif 'pause video' in query:
            keyboard.press('Space bar')
            Speak("Your video is paused sir")

        elif 'restart video' in query:
            keyboard.press('0')
            Speak("your video is restarted sir!")

        elif 'mute video' in query:
            keyboard.press('m')
            Speak("your video is muted sir!")

        elif 'skip video' in query:
            keyboard.press('1')
            Speak("video is skipped sir")

        elif 'full screen' in query:
            keyboard.press('f')
            Speak("done sir")

        elif 'youtube tool' in query:
            Speak("youtube has been automated sir")
            YoutubeAuto()

        elif 'new tab' in query:
            Speak("done sir")
            press_and_release('ctrl+t')

        elif 'close tab' in query:
            Speak("done sir")
            press_and_release('ctrl+w')

        elif 'history' in query:
            Speak("done sir")
            press_and_release('ctrl+h')

        elif 'download' in query:
            Speak("done sir")
            press_and_release('ctrl+j')

        elif 'incognito' in query:
            Speak("done sir")
            press_and_release('ctrl+shift+n')

        elif 'chrome tool' in query:
            Speak("chrome has been automated sir")
            ChromeAuto()

        elif 'jokes' in query:
            get = pyjokes.get_joke()
            Speak(get)

        elif 'dictionary' in query:
            Speak("hope i  1give you the answer sir")
            Dict()

        elif 'Rida eat' in query:
            Speak("she eats tatti/potty")
            print("she eats tatti/potty")


TaskExe()
takeCommand()






