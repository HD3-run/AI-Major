import os
import sys
import time
import subprocess
import webbrowser
import requests
import json
import pyttsx3
import pyautogui
import pygetwindow as gw
import screen_brightness_control as sbc
from fuzzywuzzy import process
import spacy
import speech_recognition as sr
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QTextEdit, QVBoxLayout, QWidget,
    QLabel, QPushButton, QLineEdit
)
from PyQt5.QtCore import QThread, pyqtSignal, QTimer
import threading
import pythoncom
import re
from bs4 import BeautifulSoup

# ------------------------------
# Global Setup
# ------------------------------
nlp = spacy.load("en_core_web_sm")
wake_word = "c2"
last_math_result = None
last_google_query = None
last_youtube_query = None

class SessionManager:
    def __init__(self, max_history=10, timeout=300):
        self.max_history = max_history
        self.timeout = timeout  # seconds
        self.history = []
        self.last_active = time.time()

    def add(self, text):
        self.history.append(text)
        if len(self.history) > self.max_history:
            self.history.pop(0)
        self.last_active = time.time()

    def get(self):
        return self.history

    def clear(self):
        self.history.clear()
        self.last_active = time.time()

    def is_expired(self):
        return (time.time() - self.last_active) > self.timeout

session = SessionManager()

def listen_command(wake_mode=False, timeout=5, phrase_time_limit=7):
    recognizer = sr.Recognizer()
    try:
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
                    os._exit(0)
                return wake_word in text
            else:
                return text
        except Exception:
            return ""
    except OSError as e:
        print(f"Microphone error: {e}")
        return "[Microphone error: unable to access mic]"

def set_volume(level):
    try:
        devices = AudioUtilities.GetSpeakers()
        interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
        volume = cast(interface, POINTER(IAudioEndpointVolume))
        volume.SetMasterVolumeLevelScalar(level / 100.0, None)
    except Exception as e:
        return f"Volume control error: {e}"

def adjust_brightness(level):
    try:
        sbc.set_brightness(level)
    except Exception as e:
        return f"Brightness control error: {e}"

def open_application(app_name):
    try:
        subprocess.Popen([app_name])
    except Exception:
        pass

def open_website(name):
    urls = {
        "youtube": "https://www.youtube.com",
        "google": "https://www.google.com"
    }
    webbrowser.open(urls.get(name.lower(), f"https://www.{name}.com"))

def search_google(query):
    global last_google_query
    url = "https://www.google.com/search?q=" + query.replace(" ", "+")
    last_google_query = url
    webbrowser.open(url, new=1)

def open_first_google_result():
    if not last_google_query:
        return "No recent Google search to open."
    headers = {'User-Agent': 'Mozilla/5.0'}
    try:
        res = requests.get(last_google_query, headers=headers)
        soup = BeautifulSoup(res.text, 'html.parser')
        for a in soup.find_all('a', href=True):
            href = a['href']
            if href.startswith('/url?q='):
                first_url = href.split('/url?q=')[1].split('&')[0]
                webbrowser.open(first_url, new=1)
                return f"Opened first Google result: {first_url}"
        return "Couldn't find the first Google result."
    except Exception as e:
        return f"Error opening first Google result: {e}"

def search_youtube(query):
    global last_youtube_query
    url = f"https://www.youtube.com/results?search_query={query.replace(' ', '+')}"
    last_youtube_query = url
    webbrowser.open(url, new=1)

def open_first_youtube_video():
    if not last_youtube_query:
        return "No recent YouTube search to open."
    headers = {'User-Agent': 'Mozilla/5.0'}
    try:
        res = requests.get(last_youtube_query, headers=headers)
        soup = BeautifulSoup(res.text, 'html.parser')
        for link in soup.find_all('a', href=True):
            href = link['href']
            if href.startswith('/watch'):
                video_url = "https://www.youtube.com" + href
                webbrowser.open(video_url, new=1)
                return f"Opened first YouTube video: {video_url}"
        return "Couldn't find the first YouTube video."
    except Exception as e:
        return f"Error opening first YouTube video: {e}"

def create_folder(name):
    try:
        os.mkdir(name)
        return f"Folder {name} created."
    except Exception:
        return "Could not create folder."

def query_ollama(prompt):
    try:
        url = "http://127.0.0.1:11434/api/generate"
        payload = {
            "model": "tinyllama",
            "prompt": prompt,
            "options": {"max_tokens": 100}
        }
        with requests.post(url, json=payload, stream=True, timeout=20) as res:
            if res.status_code != 200:
                return f"Ollama error: {res.status_code}"

            full_text = ""
            for line in res.iter_lines(decode_unicode=True):
                if line:
                    try:
                        chunk = json.loads(line)
                        full_text += chunk.get("response", "")
                        if chunk.get("done"):
                            break
                    except Exception:
                        continue
            return full_text.strip()
    except Exception as e:
        return f"Failed to connect to Ollama: {e}"

def open_and_write_notepad(text):
    try:
        subprocess.Popen(["notepad.exe"])
        time.sleep(1.5)  # Wait for Notepad to open
        notepad_windows = gw.getWindowsWithTitle("Notepad")
        if notepad_windows:
            notepad = notepad_windows[0]
            notepad.activate()
            time.sleep(0.5)
            if text:
                pyautogui.write(text, interval=0.1)
                return f"Opened Notepad and wrote: {text}"
            else:
                return "Opened Notepad."
        else:
            return "Could not find Notepad window."
    except Exception as e:
        return f"Could not write in Notepad: {e}"

def type_in_notepad(text):
    try:
        notepad_windows = gw.getWindowsWithTitle("Notepad")
        if notepad_windows:
            notepad = notepad_windows[0]
            notepad.activate()
            time.sleep(0.2)
            pyautogui.write(text, interval=0.1)
    except Exception as e:
        pass

def try_math_answer(command):
    # Try to extract and evaluate simple math expressions
    pattern = r'^\s*([\d\.\+\-\*/\(\)\s]+)=$'
    match = re.match(pattern, command.strip())
    if match:
        expr = match.group(1)
        try:
            # Only allow safe characters
            if not re.match(r'^[\d\.\+\-\*/\(\)\s]+$', expr):
                return None
            result = eval(expr)
            return str(result)
        except Exception:
            return None
    # Also handle "what is 1+2", "calculate 1+2", etc.
    if any(word in command.lower() for word in ["what is", "calculate", "answer to"]):
        expr = re.sub(r'[^0-9\.\+\-\*/\(\)\s]', '', command)
        try:
            if expr:
                result = eval(expr)
                return str(result)
        except Exception:
            return None
    return None

def interpret_command(command):
    command = command.lower()
    # Google search
    if ("search google" in command or "google search" in command or "open google and search" in command):
        if "search google for" in command:
            query = command.split("search google for", 1)[-1].strip()
        elif "open google and search" in command:
            query = command.split("open google and search", 1)[-1].strip()
        elif "google search" in command:
            query = command.split("google search", 1)[-1].strip()
        else:
            query = command.split("search google", 1)[-1].strip()
        return [("search_google", query)]
    if command.strip() == "open the first site":
        return [("open_first_google_result", "")]
    # YouTube search
    if ("search youtube" in command or "youtube search" in command):
        if "search youtube for" in command:
            query = command.split("search youtube for", 1)[-1].strip()
        elif "youtube search" in command:
            query = command.split("youtube search", 1)[-1].strip()
        else:
            query = command.split("search youtube", 1)[-1].strip()
        return [("search_youtube", query)]
    if command.strip() == "open the first video":
        return [("open_first_youtube_video", "")]
    # Compound: open notepad and type/write ...
    if "open notepad" in command and ("type " in command or "write " in command):
        text = ""
        if "type " in command:
            text = command.split("type", 1)[-1].strip().strip('"')
        elif "write " in command:
            text = command.split("write", 1)[-1].strip().strip('"')
        return [("open_and_write_notepad", text)]
    # Compound: open youtube and search ...
    if "open youtube" in command and "search" in command:
        search_part = command.split("search", 1)[-1].strip()
        return [("search_youtube", search_part)]
    # Dictation mode: open notepad
    if "open notepad" in command:
        return [("open_notepad_dictation", "")]
    # Single intents
    doc = nlp(command)
    if any(token.lemma_ in ["open", "launch"] for token in doc):
        if "notepad" in command:
            return [("open_application", "notepad.exe")]
        elif "youtube" in command:
            return [("open_website", "youtube")]
        elif "folder" in command:
            return [("create_folder", command.split("folder")[-1].strip())]
        words = command.split()
        idx = words.index("open") if "open" in words else 0
        candidate = " ".join(words[idx+1:]).strip()
        known = ["notepad", "calculator", "chrome", "spotify"]
        match, score = process.extractOne(candidate, known)
        return [("open_application", match + ".exe" if score > 60 else candidate)]
    elif "search" in command and "youtube" in command:
        return [("search_youtube", command.split("search", 1)[-1].strip())]
    elif "increase volume" in command:
        return [("set_volume", 80)]
    elif "decrease volume" in command:
        return [("set_volume", 30)]
    elif "increase brightness" in command:
        return [("adjust_brightness", 90)]
    elif "decrease brightness" in command:
        return [("adjust_brightness", 40)]
    elif command.strip() in ["exit", "stop"]:
        return [("exit", None)]
    elif command.startswith("write"):
        return [("open_and_write_notepad", command.replace("write", "").strip())]
    return [("query_ollama", command)]

def process_command(command):
    global last_math_result
    # Try math first
    math_result = try_math_answer(command)
    if math_result is not None:
        last_math_result = math_result
        return f"The answer is {math_result}."
    # Handle follow-up
    if command.lower().strip() in ["what is the answer", "what's the answer", "the answer"]:
        if last_math_result is not None:
            return f"The answer is {last_math_result}."
        else:
            return "I don't know the answer yet."
    actions = interpret_command(command)
    result = ""
    for action, arg in actions:
        if action == "search_google":
            search_google(arg)
            result += f"Searched Google for {arg}. "
        elif action == "open_first_google_result":
            result += open_first_google_result() + " "
        elif action == "search_youtube":
            search_youtube(arg)
            result += f"Searched YouTube for {arg}. "
        elif action == "open_first_youtube_video":
            result += open_first_youtube_video() + " "
        elif action == "open_website":
            open_website(arg)
            result += f"Opened {arg.capitalize()}. "
        elif action == "open_and_write_notepad":
            result += open_and_write_notepad(arg) + " "
        elif action == "open_application":
            open_application(arg)
            result += f"Opened {arg}. "
        elif action == "create_folder":
            result += create_folder(arg) + " "
        elif action == "set_volume":
            set_volume(arg)
            result += f"Volume set to {arg}. "
        elif action == "adjust_brightness":
            adjust_brightness(arg)
            result += f"Brightness set to {arg}. "
        elif action == "exit":
            sys.exit(0)
        elif action == "open_notepad_dictation":
            result += "Opened Notepad for dictation. "
        elif action == "query_ollama":
            result += query_ollama(arg) + " "
    return result.strip()

class DictationThread(QThread):
    signal_text = pyqtSignal(str)
    signal_status = pyqtSignal(str)
    signal_exit = pyqtSignal()

    def run(self):
        self.signal_status.emit("Dictation Mode: Listening for text to type. Say 'exit notepad' or 'no' to finish.")
        while True:
            text = listen_command()
            if not text:
                continue
            if "exit notepad" in text.lower() or "no" in text.lower():
                self.signal_text.emit("CS2P: Exiting Notepad dictation mode.")
                self.signal_exit.emit()
                self.signal_status.emit("Idle")
                break
            type_in_notepad(text)
            self.signal_text.emit(f"CS2P: Typed: {text}")
            self.signal_status.emit("Do you need me to type more? Say yes or no.")
            app = QApplication.instance()
            if app:
                for widget in app.topLevelWidgets():
                    if isinstance(widget, CS2PApp):
                        widget.speak("Do you need me to type more?")
                        break
            answer = listen_command()
            if "no" in answer.lower() or "exit notepad" in answer.lower():
                self.signal_text.emit("CS2P: Exiting Notepad dictation mode.")
                self.signal_exit.emit()
                self.signal_status.emit("Idle")
                break
            self.signal_status.emit("Dictation Mode: Listening for next line.")

class ManualWakeThread(QThread):
    signal_text = pyqtSignal(str)
    signal_status = pyqtSignal(str)

    def run(self):
        self.signal_status.emit("Listening...")
        command = listen_command()
        if command:
            self.signal_text.emit(f"You: {command}")
        else:
            self.signal_text.emit("CS2P: I couldn't hear anything.")
        self.signal_status.emit("Idle")

class VoiceThread(QThread):
    signal_text = pyqtSignal(str)
    signal_status = pyqtSignal(str)

    def run(self):
        app = QApplication.instance()
        if app:
            for widget in app.topLevelWidgets():
                if isinstance(widget, CS2PApp):
                    widget.speak("Hi, I'm C2. Say C2 to wake me up.")
                    break
        while True:
            if listen_command(wake_mode=True):
                self.signal_status.emit("Listening...")
                command = listen_command()
                if command:
                    self.signal_text.emit(f"You: {command}")
                    session.add(f"You: {command}")
                    response = self.process_command_with_dictation(command)
                    self.signal_text.emit(f"CS2P: {response}")
                    session.add(f"CS2P: {response}")
                    app = QApplication.instance()
                    if app:
                        for widget in app.topLevelWidgets():
                            if isinstance(widget, CS2PApp):
                                widget.speak(response)
                                break
                else:
                    self.signal_text.emit("CS2P: I didn't catch that.")
                    session.add("CS2P: I didn't catch that.")
                self.signal_status.emit("Idle")
            if session.is_expired():
                self.signal_text.emit("Session expired. Conversation cleared.")
                session.clear()
            time.sleep(0.5)

    def process_command_with_dictation(self, command):
        actions = interpret_command(command)
        # Situation 1: open notepad and type/write ...
        if actions and actions[0][0] == "open_and_write_notepad":
            return open_and_write_notepad(actions[0][1])
        # Situation 2: open notepad (dictation mode)
        elif actions and actions[0][0] == "open_notepad_dictation":
            open_and_write_notepad("")
            app = QApplication.instance()
            if app:
                for widget in app.topLevelWidgets():
                    if isinstance(widget, CS2PApp):
                        widget.start_dictation_mode()
                        break
            return "Opened Notepad. Dictation mode started. Please dictate what you want me to type. Say 'exit notepad' or 'no' when done."
        # Normal command processing
        else:
            return process_command(command)

class CS2PApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("CS2P Voice Assistant")
        self.setGeometry(100, 100, 600, 450)

        self.layout = QVBoxLayout()
        self.label = QLabel("CS2P Assistant")
        self.status = QLabel("Idle")

        self.output_area = QTextEdit()
        self.output_area.setReadOnly(True)

        self.input_box = QLineEdit()
        self.input_box.setPlaceholderText("Type your request here...")
        self.input_box.returnPressed.connect(self.process_text_input)

        self.wake_button = QPushButton("🔊 Wake Up")
        self.wake_button.clicked.connect(self.manual_wake)

        self.clear_button = QPushButton("🧹 Clear Session")
        self.clear_button.clicked.connect(self.clear_session)

        self.stop_speaking_button = QPushButton("🛑 Stop Speaking")
        self.stop_speaking_button.clicked.connect(self.stop_speaking)
        self.stop_speaking_button.setEnabled(False)

        self.layout.addWidget(self.label)
        self.layout.addWidget(self.status)
        self.layout.addWidget(self.output_area)
        self.layout.addWidget(self.input_box)
        self.layout.addWidget(self.wake_button)
        self.layout.addWidget(self.clear_button)
        self.layout.addWidget(self.stop_speaking_button)

        container = QWidget()
        container.setLayout(self.layout)
        self.setCentralWidget(container)

        self.voice_thread = VoiceThread()
        self.voice_thread.signal_text.connect(self.update_output)
        self.voice_thread.signal_status.connect(self.update_status)
        self.voice_thread.start()

        self.session_timer = QTimer()
        self.session_timer.timeout.connect(self.check_session_expiry)
        self.session_timer.start(10000)

        self.manual_wake_thread = None
        self.dictation_thread = None

        self.speaking = False

    def speak(self, text):
        if self.speaking:
            return  # Don't start new speech if already speaking
        self.speaking = True
        self.wake_button.setEnabled(False)
        self.input_box.setEnabled(False)
        self.stop_speaking_button.setEnabled(True)
        def _speak():
            pythoncom.CoInitialize()
            local_engine = pyttsx3.init()
            local_engine.say(text)
            local_engine.runAndWait()
            local_engine.stop()
            self.speaking = False
            self.wake_button.setEnabled(True)
            self.input_box.setEnabled(True)
            self.stop_speaking_button.setEnabled(False)
        threading.Thread(target=_speak).start()

    def stop_speaking(self):
        self.speaking = False
        self.wake_button.setEnabled(True)
        self.input_box.setEnabled(True)
        self.stop_speaking_button.setEnabled(False)
        try:
            engine = pyttsx3.init()
            engine.stop()
        except Exception:
            pass

    def update_output(self, text):
        self.output_area.clear()
        for msg in session.get():
            self.output_area.append(msg)
        if text not in session.get():
            self.output_area.append(text)

    def update_status(self, text):
        self.status.setText(text)

    def manual_wake(self):
        if self.speaking:
            return  # Don't listen while speaking
        if self.manual_wake_thread and self.manual_wake_thread.isRunning():
            return  # Prevent multiple threads
        self.manual_wake_thread = ManualWakeThread()
        self.manual_wake_thread.signal_text.connect(self.handle_voice_input)
        self.manual_wake_thread.signal_status.connect(self.update_status)
        self.manual_wake_thread.start()

    def handle_voice_input(self, text):
        if self.speaking:
            return
        self.update_output(text)
        session.add(text)
        response = self.process_command_with_dictation(text.replace("you:", "").strip())
        self.update_output(f"CS2P: {response}")
        session.add(f"CS2P: {response}")
        self.speak(response)

    def process_text_input(self):
        if self.speaking:
            return
        command = self.input_box.text().strip()
        if command:
            self.update_output(f"You: {command}")
            session.add(f"You: {command}")
            response = self.process_command_with_dictation(command)
            self.update_output(f"CS2P: {response}")
            session.add(f"CS2P: {response}")
            self.speak(response)
            self.input_box.clear()

    def clear_session(self):
        session.clear()
        self.output_area.clear()
        self.update_output("Session cleared.")

    def check_session_expiry(self):
        if session.is_expired():
            session.clear()
            self.output_area.clear()
            self.update_output("Session expired. Conversation cleared.")

    def process_command_with_dictation(self, command):
        actions = interpret_command(command)
        # Situation 1: open notepad and type/write ...
        if actions and actions[0][0] == "open_and_write_notepad":
            return open_and_write_notepad(actions[0][1])
        # Situation 2: open notepad (dictation mode)
        elif actions and actions[0][0] == "open_notepad_dictation":
            open_and_write_notepad("")
            self.start_dictation_mode()
            return "Opened Notepad. Dictation mode started. Please dictate what you want me to type. Say 'exit notepad' or 'no' when done."
        # Normal command processing
        else:
            return process_command(command)

    def start_dictation_mode(self):
        if self.dictation_thread and self.dictation_thread.isRunning():
            return
        self.dictation_thread = DictationThread()
        self.dictation_thread.signal_text.connect(self.update_output)
        self.dictation_thread.signal_status.connect(self.update_status)
        self.dictation_thread.signal_exit.connect(self.exit_dictation_mode)
        self.dictation_thread.start()

    def exit_dictation_mode(self):
        self.update_output("CS2P: Dictation mode ended.")
        self.speak("Dictation mode ended.")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = CS2PApp()
    window.show()
    sys.exit(app.exec_())
import sys
import time
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout,
    QPushButton, QLabel, QTextEdit
)
from PyQt5.QtCore import QThread, pyqtSignal

# --- Dummy implementations for demonstration ---
def listen_command():
    # Simulate listening for a command (replace with real implementation)
    time.sleep(1)
    return input("Simulate speech (type your command): ")

def type_in_notepad(text):
    # Simulate typing in Notepad (replace with real implementation)
    print(f"Typing in Notepad: {text}")

# --- Dictation Thread ---
class DictationThread(QThread):
    signal_text = pyqtSignal(str)
    signal_status = pyqtSignal(str)
    signal_exit = pyqtSignal()

    def __init__(self):
        super().__init__()
        self._stop_flag = False

    def run(self):
        self.signal_status.emit("Dictation Mode: Listening for text to type. Say 'exit notepad' or 'no' to finish.")
        while not self._stop_flag:
            text = listen_command()
            if self._stop_flag:
                break
            if not text:
                continue
            if "exit notepad" in text.lower() or "no" in text.lower():
                self.signal_text.emit("CS2P: Exiting Notepad dictation mode.")
                self.signal_exit.emit()
                self.signal_status.emit("Idle")
                break
            type_in_notepad(text)
            self.signal_text.emit(f"CS2P: Typed: {text}")
            self.signal_status.emit("Do you need me to type more? Say yes or no.")
            answer = listen_command()
            if self._stop_flag:
                break
            if "no" in answer.lower() or "exit notepad" in answer.lower():
                self.signal_text.emit("CS2P: Exiting Notepad dictation mode.")
                self.signal_exit.emit()
                self.signal_status.emit("Idle")
                break
            self.signal_status.emit("Dictation Mode: Listening for next line.")

    def stop(self):
        self._stop_flag = True

# --- Main App ---
class CS2PApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("CS2P Dictation Example")
        self.resize(500, 300)
        self.dictation_thread = None

        # Central widget and layout
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        self.layout = QVBoxLayout()
        central_widget.setLayout(self.layout)

        # Status label
        self.status_label = QLabel("Idle")
        self.layout.addWidget(self.status_label)

        # Output text area
        self.text_output = QTextEdit()
        self.text_output.setReadOnly(True)
        self.layout.addWidget(self.text_output)

        # Start Dictation button
        self.start_dictation_btn = QPushButton("Start Dictation")
        self.start_dictation_btn.clicked.connect(self.start_dictation)
        self.layout.addWidget(self.start_dictation_btn)

        # Stop Dictation button
        self.stop_dictation_btn = QPushButton("Stop Dictation")
        self.stop_dictation_btn.clicked.connect(self.stop_dictation)
        self.stop_dictation_btn.setEnabled(False)
        self.layout.addWidget(self.stop_dictation_btn)

    def start_dictation(self):
        if self.dictation_thread is None or not self.dictation_thread.isRunning():
            self.dictation_thread = DictationThread()
            self.dictation_thread.signal_text.connect(self.display_text)
            self.dictation_thread.signal_status.connect(self.update_status)
            self.dictation_thread.signal_exit.connect(self.on_dictation_exit)
            self.dictation_thread.start()
            self.stop_dictation_btn.setEnabled(True)
            self.start_dictation_btn.setEnabled(False)

    def stop_dictation(self):
        if self.dictation_thread and self.dictation_thread.isRunning():
            self.dictation_thread.stop()
            self.dictation_thread.wait()
            self.stop_dictation_btn.setEnabled(False)
            self.start_dictation_btn.setEnabled(True)
            self.update_status("Dictation stopped by user.")

    def on_dictation_exit(self):
        self.stop_dictation_btn.setEnabled(False)
        self.start_dictation_btn.setEnabled(True)
        self.update_status("Dictation stopped.")

    def display_text(self, text):
        self.text_output.append(text)

    def update_status(self, status):
        self.status_label.setText(status)

# --- Run the app ---
if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = CS2PApp()
    window.show()
    sys.exit(app.exec_())
