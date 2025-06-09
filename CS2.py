import os
import subprocess
import webbrowser
import time
import pyttsx3
import pyautogui
import screen_brightness_control as sbc
from fuzzywuzzy import process, fuzz
import spacy
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
import speech_recognition as sr
import requests
import json
from PyQt5.QtWidgets import QApplication, QMainWindow, QTextEdit, QVBoxLayout, QWidget, QLabel
from PyQt5.QtCore import QThread, pyqtSignal
import sys

# Load spaCy English model
nlp = spacy.load("en_core_web_sm")

# Initialize text-to-speech engine
engine = pyttsx3.init()

# Global state variables
active_app = None
chat_history = []
last_command = ""
last_user_query = ""
last_written_text = ""
last_app_opened_time = None
wake_word = "c2"

class VoiceThread(QThread):
    signal_text = pyqtSignal(str)

    def run(self):
        self.signal_text.emit("Hi, I am CS2P. Say 'C2' to wake me up.")
        speak("Hi, I am CS2P. Say 'C2' to wake me up.")
        while True:
            if listen_command(wake_mode=True):
                self.signal_text.emit("Speak...")
                speak("Speak")
                command = listen_command()
                if command:
                    self.signal_text.emit("You: " + command)
                    response = process_command(command)
                    self.signal_text.emit("CS2P: " + response)
                    speak(response)
            time.sleep(0.5)

class CS2PApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("CS2P Assistant")
        self.setGeometry(100, 100, 600, 400)

        self.layout = QVBoxLayout()

        self.label = QLabel("CS2P Voice Assistant")
        self.output_area = QTextEdit()
        self.output_area.setReadOnly(True)

        self.layout.addWidget(self.label)
        self.layout.addWidget(self.output_area)

        container = QWidget()
        container.setLayout(self.layout)
        self.setCentralWidget(container)

        self.voice_thread = VoiceThread()
        self.voice_thread.signal_text.connect(self.update_output)
        self.voice_thread.start()

    def update_output(self, text):
        self.output_area.append(text)

# Core functionalities

def speak(text):
    engine.say(text)
    engine.runAndWait()
    return text

def listen_command(wake_mode=False, timeout=5, phrase_time_limit=7):
    recognizer = sr.Recognizer()
    mic = sr.Microphone()
    with mic as source:
        recognizer.adjust_for_ambient_noise(source)
        try:
            audio_data = recognizer.listen(source, timeout=timeout, phrase_time_limit=phrase_time_limit)
        except sr.WaitTimeoutError:
            return ""

    try:
        text = recognizer.recognize_google(audio_data).lower()
        if wake_mode:
            if any(word in text for word in ["kill", "die"]):
                speak("Goodbye!")
                sys.exit(0)
            return wake_word in text
        else:
            return text
    except:
        return ""

def set_volume(level):
    devices = AudioUtilities.GetSpeakers()
    interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
    volume = cast(interface, POINTER(IAudioEndpointVolume))
    volume.SetMasterVolumeLevelScalar(level / 100.0, None)

def adjust_brightness(level):
    sbc.set_brightness(level)

def open_application(app_name):
    global active_app
    try:
        subprocess.Popen([app_name])
        active_app = app_name.lower()
    except:
        speak(f"Could not open {app_name}.")

def open_website(website_name):
    global active_app
    urls = {
        "youtube": "https://www.youtube.com",
        "google": "https://www.google.com"
    }
    url = urls.get(website_name.lower(), f"https://www.{website_name}.com")
    webbrowser.open(url, new=0)
    active_app = website_name.lower()

def search_youtube(query):
    global active_app
    search_url = f"https://www.youtube.com/results?search_query={query.replace(' ', '+')}"
    webbrowser.open(search_url, new=0 if active_app == "youtube" else 2)
    active_app = "youtube"

def open_notepad_and_write(text="Hello, this is CS2P speaking."):
    global active_app
    subprocess.Popen(["notepad.exe"])
    time.sleep(2)
    active_app = "notepad"
    pyautogui.write(text, interval=0.1)

def write_in_notepad(text):
    if active_app == "notepad":
        pyautogui.write(text, interval=0.1)
    else:
        speak("Notepad is not active.")

def create_folder(folder_name):
    try:
        os.mkdir(folder_name)
        return f"Folder {folder_name} created."
    except:
        return "Could not create folder."

def perform_action_in_active_app():
    if active_app == "notepad":
        pyautogui.write("Default action in Notepad.", interval=0.1)
    elif active_app == "youtube":
        query = listen_command()
        if query:
            search_youtube(query)
    else:
        speak(f"No default action defined for {active_app}.")

def query_ollama(prompt, model="llama2-uncensored"):
    url = "http://127.0.0.1:12345/api/generate"
    payload = {"model": model, "prompt": prompt, "options": {"max_tokens": 100}}
    try:
        response = requests.post(url, json=payload)
        if response.status_code == 200:
            result = response.json()
            return result.get("response", "")
        else:
            return "Ollama error"
    except:
        return "Error connecting to Ollama"

def interpret_command_local(command):
    doc = nlp(command)
    if any(token.lemma_ in ["open", "start", "launch"] for token in doc):
        if "notepad" in command:
            return "open notepad"
        elif "folder" in command:
            return "create folder"
        elif "youtube" in command:
            return "open youtube"
        words = command.split()
        idx = words.index("open") if "open" in words else 0
        app_candidate = " ".join(words[idx+1:]).strip()
        known_apps = ["chrome", "calculator", "word", "excel", "whatsapp", "notepad", "spotify"]
        best_match, score = process.extractOne(app_candidate, known_apps)
        return f"open {best_match}" if score > 60 else f"open {app_candidate}"
    elif "search" in command and "youtube" in command:
        return "search youtube"
    elif active_app == "youtube" and command.startswith("search"):
        return "search youtube"
    elif "increase" in command and "volume" in command:
        return "increase volume"
    elif "decrease" in command and "volume" in command:
        return "decrease volume"
    elif "increase" in command and "brightness" in command:
        return "increase brightness"
    elif "decrease" in command and "brightness" in command:
        return "decrease brightness"
    elif "exit" in command or "stop" in command:
        return "exit"
    elif command.lower().strip() in ["do this", "do that", "execute this", "run this"]:
        return "active action"
    elif command.startswith("write") or command.startswith("type"):
        return "write"
    else:
        return "unknown command"

def process_command(command):
    possible = [
        "open youtube", "search youtube", "open notepad", "create folder",
        "exit", "stop", "increase volume", "decrease volume",
        "increase brightness", "decrease brightness", "write", "active action"
    ]
    best_match, score = process.extractOne(command, possible)
    if score > 70:
        command = best_match
    intent = interpret_command_local(command)

    if intent == "open youtube":
        open_website("youtube")
        return "Opening YouTube."
    elif intent == "search youtube":
        if "for" in command:
            query = command.split("for", 1)[-1].strip()
        elif "search" in command:
            query = command.split("search", 1)[-1].strip()
        else:
            query = ""
        if query:
            search_youtube(query)
            return f"Searching YouTube for {query}."
        else:
            return "No search query provided."
    elif intent == "open notepad":
        open_notepad_and_write("CS2P is now smarter!")
        return "Opened Notepad."
    elif intent == "create folder":
        folder_name = command.split("create folder")[-1].strip()
        return create_folder(folder_name)
    elif intent == "increase volume":
        set_volume(80)
        return "Volume set to 80%."
    elif intent == "decrease volume":
        set_volume(30)
        return "Volume set to 30%."
    elif intent == "increase brightness":
        adjust_brightness(90)
        return "Brightness set to 90%."
    elif intent == "decrease brightness":
        adjust_brightness(40)
        return "Brightness set to 40%."
    elif intent == "exit":
        speak("Goodbye!")
        sys.exit(0)
    elif intent == "write":
        text = command.replace("write", "").replace("type", "").strip()
        write_in_notepad(text)
        return "Typed in Notepad."
    elif intent == "active action":
        perform_action_in_active_app()
        return "Performed action in active app."
    elif intent.startswith("open "):
        app = intent.replace("open ", "")
        open_application(app)
        return f"Opened {app}."
    else:
        response = query_ollama(command)
        return f"Ollama says: {response}"

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = CS2PApp()
    window.show()
    sys.exit(app.exec_())