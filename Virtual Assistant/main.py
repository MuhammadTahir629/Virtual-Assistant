import speech_recognition as sr
import pyttsx3
import datetime
import webbrowser
import wikipedia
import os
import platform
from googlesearch import search

# Initialize text-to-speech engine
engine = pyttsx3.init()
engine.setProperty('rate', 150)  # Set speech speed

# Speak out text
def speak(text):
    engine.say(text)
    engine.runAndWait()

# Listen and recognize speech
def recognize_speech():
    recognizer = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening...")
        recognizer.adjust_for_ambient_noise(source)
        try:
            audio = recognizer.listen(source, timeout=5)
            command = recognizer.recognize_google(audio).lower()
            print(f"You said: {command}")
            return command
        except sr.UnknownValueError:
            print("Sorry, I couldn't understand.")
            return None
        except sr.RequestError:
            print("Network error.")
            return None
        except sr.WaitTimeoutError:
            print("You didn’t say anything.")
            return None

# Open Microsoft Office apps (Windows only)
def open_ms_office_tool(tool_name):
    system_platform = platform.system()
    office_tools = {
        "word": "winword",
        "excel": "excel",
        "powerpoint": "powerpnt"
    }

    if system_platform == "Windows" and tool_name.lower() in office_tools:
        os.system(f"start {office_tools[tool_name.lower()]}")
    else:
        speak("Sorry, this tool is only supported on Windows.")

# Search and play top YouTube video
def play_youtube_video(query):
    speak(f"Searching YouTube for {query}")
    try:
        search_results = search(f"{query} site:youtube.com", num_results=1)
        video_url = next(search_results, None)
        if video_url:
            speak("Playing top result from YouTube.")
            webbrowser.open(video_url)
        else:
            speak("No video found.")
    except Exception:
        speak("Sorry, I couldn't fetch the video.")

# Main assistant actions
def execute_command(command):
    if "marvin" in command:
        speak("Hello, I am Marvin! How can I help you?")

        while True:
            command = recognize_speech()
            if command is None:
                continue

            # Time and date
            elif "time" in command or "date" in command:
                now = datetime.datetime.now().strftime("%I:%M %p, %A, %B %d, %Y")
                speak(f"The current time and date is {now}")

            # Wikipedia summary
            elif any(phrase in command for phrase in ["who is", "what is", "who was", "what was"]):
                query = command
                for phrase in ["who is", "what is", "who was", "what was"]:
                    query = query.replace(phrase, "")
                try:
                    summary = wikipedia.summary(query.strip(), sentences=2)
                    speak(summary)
                except wikipedia.exceptions.DisambiguationError:
                    speak("Too many results. Please be more specific.")
                except wikipedia.exceptions.PageError:
                    speak("Topic not found.")

            # Google Search
            elif "search google for" in command:
                query = command.replace("search google for", "").strip()
                speak(f"Searching Google for {query}")
                webbrowser.open(next(search(query, num_results=1)))

            # Open websites
            elif "open facebook" in command:
                speak("Opening Facebook")
                webbrowser.open("https://www.facebook.com")
            elif "open instagram" in command:
                speak("Opening Instagram")
                webbrowser.open("https://www.instagram.com")
            elif "open youtube" in command:
                speak("Opening YouTube")
                webbrowser.open("https://www.youtube.com")
            elif "open mail" in command:
                speak("Opening Gmail")
                webbrowser.open("https://mail.google.com")

            # WhatsApp (Windows only)
            elif "open whatsapp" in command:
                speak("Opening WhatsApp")
                if platform.system() == "Windows":
                    os.system("start WhatsApp")
                else:
                    speak("WhatsApp opening is only supported on Windows.")

            # Open Office tools
            elif "open word" in command:
                speak("Opening Microsoft Word")
                open_ms_office_tool("word")
            elif "open excel" in command:
                speak("Opening Microsoft Excel")
                open_ms_office_tool("excel")
            elif "open powerpoint" in command:
                speak("Opening Microsoft PowerPoint")
                open_ms_office_tool("powerpoint")

            # Play YouTube video
            elif "play youtube video on" in command:
                query = command.replace("play youtube video on", "").strip()
                play_youtube_video(query)

            # Exit
            elif "exit" in command or "stop" in command:
                speak("Goodbye! Have a nice day.")
                break

            # Unknown command
            else:
                speak("Sorry, I didn't understand that.")

# Entry point
if __name__ == "__main__":
    speak("Virtual Assistant is running. Say 'Marvin' to activate.")
    while True:
        wake_command = recognize_speech()
        if wake_command and "marvin" in wake_command:
            execute_command("marvin")
