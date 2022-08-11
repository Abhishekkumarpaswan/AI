import pyttsx3
import wikipedia
engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[0].id)

def speak(audio):
    engine.say(audio)
    engine.runAndWait()
if __name__ == '__main__':
    n=int(input("Enter any no."))
    length=int(input("Tell me the length of result\n"))
    while(n>0):
        inp = input("What do you want to search..\n")
        speak("According to Wikipedia..")
        results = wikipedia.summary(inp, sentences=length)
        print(results+"\n")
        advice = input("You want to listen the results...\n")
        if (advice == "yes"):
            speak(results)
        elif (advice == "no"):
            print("Okay")
        else:
            exit()






