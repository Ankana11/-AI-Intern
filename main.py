import win32com.client
import speech_recognition as sr
import webbrowser
import openai
import datetime

chatStr = ""
def chat(query):
    global chatStr
    print(chatStr)
    openai.api_key = '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'
    chatStr += f"You: {query}\n A.I: "
    response = openai.Completion.create(
        model="text-davinci-003",
        prompt= chatStr,
        temperature=0.7,
        max_tokens=256,
        top_p=1,
        frequency_penalty=0,
        presence_penalty=0

    )
    speak(response["choices"][0]["text"])
    chatStr += f"{response['choices'][0]['text']}\n"
    return response["choices"][0]["text"]

def speak(str):
    from win32com.client import Dispatch
    speak = Dispatch('SAPI.SpVoice')
    speak.Speak(str)
def takeCommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        audio = r.listen(source)
        try:
            print("Recognizing...")
            query = r.recognize_google(audio, language="en-in")
            print(f"User said: {query}")
            return query
        except Exception as e:
            return "Some Error Occurred. Sorry from A.I"

if __name__ == '__main__':
    speak("Hello I am A.I assistent. how can i help you")
    while True:
        print("Listening...")
        query = takeCommand()
        '''Add more sites its upto users choice'''
        sites = [["youtube", "https://www.youtube.com"], ["wikipedia", "https://www.wikipedia.com"], ["google", "https://www.google.com"],]
        for site in sites:
            if f"Open {site[0]}".lower() in query.lower():
                speak(f"Opening {site[0]} mam...")
                webbrowser.open(site[1])

        '''using datetime module'''
        if "the time" in query:
            musicPath = "/Users/harry/Downloads/downfall-21371.mp3"
            hour = datetime.datetime.now().strftime("%H")
            min = datetime.datetime.now().strftime("%M")
            speak(f"Mam time is {hour} and {min} minutes")

        elif "Quit".lower() in query.lower():
            exit()

        elif "reset chat".lower() in query.lower():
            chatStr = ""

        else:
            print("Chatting...")
            chat(query)
