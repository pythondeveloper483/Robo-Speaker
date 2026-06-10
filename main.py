import win32com.client as wincl

if __name__ == '__main__':
    print("Welcome to RoboSpeaker 2.0. Created by Shashank Patel")
    speaker = wincl.Dispatch("SAPI.SpVoice")
    while True:
        x = input("Enter what you want me to speak: ")
        if x.lower() == "exit":
            speaker.Speak("Bye Bye Friend, Nice to Meet You!")
            break
        speaker.Speak(x)
