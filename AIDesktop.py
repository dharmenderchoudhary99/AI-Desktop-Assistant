import speech_recognition as sr
import os
import win32com.client
import webbrowser
#
# speaker = win32com.client.Dispatch("SAPI.SpVoice")
#
# while 1:
#     print("Enter the word you want to say")
#     s = input()
#     speaker.Speak(s)


def say(text):
    os.system(f"say {text}")


def takeCommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        r.pause_threshold = 1
        audio = r.listen(source)
        try:
            query = r.recognize_google(audio, language="en-in")
            print(f"User said: {query}")
            return query
        except Exception as e:
            return "Some error Occured"


if __name__ == '__main__':
    print("Hello")
    say("Hello Dharm")
    while True:
        print("Listening....")
        query = takeCommand()
        # query.Speak(query)
        sites=[
            ["youtube","https://www.youtube.com"],
            ["wikipedia","https://www.wikipedia.com"],
            ["google","https://www.google.com"],
        ]
        for site in sites:
            if f"Open {site[0]}".lower() in query.lower():
                say(f"Opening {site[0]} sir....")
                webbrowser.open(site[1])
        # say(query)
