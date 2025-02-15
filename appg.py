import speech_recognition as sr 
import pyttsx3
import os
import subprocess
from ctypes import cast, POINTER #Low-level Windows API access ,Used for audio volume control through pycaw
from comtypes import CLSCTX_ALL #same
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume #Windows Core Audio API wrapper
import webbrowser #Basic web browser control
import re #Used for command parsing
from urllib.parse import quote #URL encoding/decoding Used for creating search queries
import sys
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                            QPushButton, QTextEdit, QLabel)
from PyQt5.QtCore import Qt, pyqtSignal, QObject, QThread
from PyQt5.QtGui import QFont, QIcon, QColor, QPalette

import requests
from bs4 import BeautifulSoup
import psutil
import win32gui
import win32con
import win32process
import win32com.client

class SystemController:
    def __init__(self):
        # Initialize audio controller
        devices = AudioUtilities.GetSpeakers()
        interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
        self.volume = cast(interface, POINTER(IAudioEndpointVolume))
        
        # Initialize shell for window control
        self.shell = win32com.client.Dispatch("WScript.Shell")
        
        # Common folders
        self.common_folders = {
            'documents': os.path.expanduser('~/Documents'),
            'downloads': os.path.expanduser('~/Downloads'),
            'desktop': os.path.expanduser('~/Desktop'),
            'music': os.path.expanduser('~/Music'),
            'videos': os.path.expanduser('~/Videos'),
            'pictures': os.path.expanduser('~/Pictures')
        }
        
        # Common apps and their possible locations
        self.common_apps = {
            'chrome': ['C:/Program Files/Google/Chrome/Application/chrome.exe',
                      'C:/Program Files (x86)/Google/Chrome/Application/chrome.exe'],
            'firefox': ['C:/Program Files/Mozilla Firefox/firefox.exe'],
            'edge': ['C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe'],
            'whatsapp': [os.path.expanduser('~/AppData/Local/WhatsApp/WhatsApp.exe')],
            'vlc': ['C:/Program Files/VideoLAN/VLC/vlc.exe'],
            'spotify': [os.path.expanduser('~/AppData/Roaming/Spotify/Spotify.exe')],
            'word': ['C:/Program Files/Microsoft Office/root/Office16/WINWORD.EXE'],
            'excel': ['C:/Program Files/Microsoft Office/root/Office16/EXCEL.EXE'],
            'notepad': ['C:/Windows/notepad.exe'],
            'calculator': ['C:/Windows/System32/calc.exe']
        }
        
        # Browser identifiers
        self.browser_identifiers = {
            'chrome': 'Google Chrome',
            'firefox': 'Mozilla Firefox',
            'edge': 'Microsoft Edge'
        }

    def change_volume(self, command):
        """Change system volume by specific amount"""
        try:
            match = re.search(r'(\d+)', command) #check if command contains a number
            if match:
                target_volume = min(100, max(0, int(match.group(1))))
                self.volume.SetMasterVolumeLevelScalar(target_volume / 100, None)
                return target_volume
            return None
        except Exception as e:
            print(f"Volume error: {str(e)}")
            return None

    def open_file_or_folder(self, name):
        """Open common folder or application"""
        name = name.lower()
        
        # Check common folders
        if name in self.common_folders:
            try:
                os.startfile(self.common_folders[name])
                return True, f"Opening {name} folder"
            except Exception:
                pass

        # Check common apps
        for app_name, paths in self.common_apps.items():
            if name in app_name:
                for path in paths:
                    if os.path.exists(path):
                        try:
                            subprocess.Popen(path)
                            return True, f"Opening {app_name}"
                        except Exception:
                            continue

        # Check Program Files directories for apps
        program_files = ['C:/Program Files', 'C:/Program Files (x86)']
        for directory in program_files:
            if os.path.exists(directory):
                for root, dirs, files in os.walk(directory):
                    for file in files:
                        if file.lower().startswith(name) and file.endswith('.exe'):
                            try:
                                subprocess.Popen(os.path.join(root, file))
                                return True, f"Opening {file}"
                            except Exception:
                                continue

        return False, f"Couldn't find {name}"

    def play_youtube_video(self, query=None):
        """Handle YouTube commands"""
        try:
            if not query:
                # Just open YouTube homepage
                webbrowser.open('https://www.youtube.com')
                return True, "Opening YouTube"
            
            # Search and play first video
            search_url = f"https://www.youtube.com/results?search_query={quote(query)}"
            response = requests.get(search_url)
            soup = BeautifulSoup(response.text, 'html.parser')
            
            # Extract video ID from search results
            results = re.findall(r"watch\?v=(\S{11})", response.text)
            if results:
                first_video = f"https://www.youtube.com/watch?v={results[0]}"
                webbrowser.open(first_video)
                return True, f"Playing {query} on YouTube"
            
            return False, "Couldn't find the video"
        except Exception as e:
            print(f"YouTube error: {str(e)}")
            return False, "Error accessing YouTube"
    def web_search(self, query, browser=None):
        """Perform a web search in specified browser or default"""
        search_url = f"https://www.google.com/search?q={quote(query)}"
        return self.open_website(search_url, browser)

    def open_website(self, url, browser=None):
        """Open specific website in specified browser or default"""
        try:
            # Ensure URL has proper protocol
            if not url.startswith(('http://', 'https://')):
                url = f'https://{url}'

            if browser:
                browser = browser.lower()
                # Find browser executable
                for app_name, paths in self.common_apps.items():
                    if app_name in browser:
                        for path in paths:
                            if os.path.exists(path):
                                subprocess.Popen([path, url])
                                return True, f"Opening {url} in {app_name}"
                        return False, f"{browser} not found"
            
            # Fallback to default browser
            webbrowser.open(url)
            return True, f"Opening {url}"
        except Exception as e:
            print(f"Website error: {str(e)}")
            return False, "Failed to open website"

    def get_window_by_title(self, title):
        """Find window handle by partial title"""
        def callback(hwnd, windows):
            if win32gui.IsWindowVisible(hwnd):
                window_title = win32gui.GetWindowText(hwnd)
                if title.lower() in window_title.lower():
                    windows.append(hwnd)
            return True
        
        windows = []
        win32gui.EnumWindows(callback, windows)
        return windows

    def close_window(self, title):
        """Close window by title"""
        windows = self.get_window_by_title(title)
        if windows:
            for hwnd in windows:
                win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
            return True, f"Closing {title}"
        return False, f"Couldn't find window with title {title}"

    def close_application(self, app_name):
        """Close application by name"""
        app_name = app_name.lower()
        closed = False
        
        for proc in psutil.process_iter(['name', 'pid']):
            try:
                if app_name in proc.info['name'].lower():
                    proc.terminate()
                    closed = True
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                continue
        
        if closed:
            return True, f"Closed {app_name}"
        return False, f"Couldn't find {app_name}"

    def close_browser_tab(self, title=None):
        """Close browser tab by title or active tab"""
        try:
            # Send Alt+F4 to close active tab
            self.shell.SendKeys('%{F4}')
            return True, "Closed browser tab"
        except Exception as e:
            return False, "Couldn't close browser tab"

class VoiceAssistant(QObject):
    textUpdated = pyqtSignal(str)
    
    def __init__(self):
        super().__init__()
        self.recognizer = sr.Recognizer()
        self.engine = pyttsx3.init()
        self.system = SystemController()
        self.is_listening = True
        
    def speak(self, text):
        """Convert text to speech"""
        self.textUpdated.emit(f"Assistant: {text}")
        self.engine.say(text)
        self.engine.runAndWait()
    
    def process_command(self, command):
        """Process and execute voice commands"""
        command = command.lower()
        
        # Volume control
        if "volume" in command or "sound" in command:
            # Handle volume increase cases
            if "increase" in command or "up" in command:
                new_volume = self.system.change_volume(command)
                if new_volume is not None:
                    self.speak(f"Volume set to {new_volume}%")
                else:
                    # Increase by 5% if no specific value
                    current_vol = self.system.volume.GetMasterVolumeLevelScalar() * 100
                    new_vol = min(100, round(current_vol) + 5)
                    self.system.volume.SetMasterVolumeLevelScalar(new_vol / 100, None)
                    self.speak(f"Increased volume to {new_vol}%")
                return
            
            # Handle volume decrease cases
            elif "decrease" in command or "down" in command:
                new_volume = self.system.change_volume(command)
                if new_volume is not None:
                    self.speak(f"Volume set to {new_volume}%")
                else:
                    # Decrease by 5% if no specific value
                    current_vol = self.system.volume.GetMasterVolumeLevelScalar() * 100
                    new_vol = max(0, round(current_vol) - 5)
                    self.system.volume.SetMasterVolumeLevelScalar(new_vol / 100, None)
                    self.speak(f"Decreased volume to {new_vol}%")
                return
            
            # Handle set volume without specific value
            elif "set" in command:
                new_volume = self.system.change_volume(command)
                if new_volume is not None:
                    self.speak(f"Volume set to {new_volume}%")
                else:
                    current_vol = self.system.volume.GetMasterVolumeLevelScalar() * 100
                    self.speak(f"Current volume is {round(current_vol)}%")
                return
        
        # Close commands
        if "close" in command:
            # Close browser tab
            if "tab" in command:
                success, message = self.system.close_browser_tab()
                self.speak(message)
                return
            
            # Close specific application
            if "app" in command or "application" in command:
                app_name = command.replace("close", "").replace("app", "").replace("application", "").strip()
                success, message = self.system.close_application(app_name)
                self.speak(message)
                return
            
            # Close window/folder
            name = command.replace("close", "").replace("window", "").replace("folder", "").strip()
            success, message = self.system.close_window(name)
            self.speak(message)
            return
        
        # YouTube commands
        if "youtube" in command:
            if command.strip() == "play youtube":
                success, message = self.system.play_youtube_video()
            else:
                query = command.replace("play", "").replace("youtube", "").strip()
                success, message = self.system.play_youtube_video(query)
            self.speak(message)
            return
        
        # Play video/song (YouTube)
        if "play" in command:
            query = command.replace("play", "").strip()
            success, message = self.system.play_youtube_video(query)
            self.speak(message)
            return
        
        # Open file/folder/app
        if "open" in command:
            name = command.replace("open", "").replace("app", "").strip()
            success, message = self.system.open_file_or_folder(name)
            self.speak(message)
            return
        
        # Stop command
        if "stop" in command or "exit" in command:
            self.is_listening = False
            self.speak("Stopping voice assistant")
            return
        # Web search command
        if "search" in command:
            query = command.replace("search", "").replace("for", "").strip()
            if query:
                # Check if browser is specified
                browser = None
                if " in " in command:
                    parts = command.split(" in ")
                    query = parts[0].replace("search", "").strip()
                    browser = parts[1].strip()
                
                # Check if query contains website pattern
                is_website = re.search(r'\b[\w-]+\.(com|org|net|gov|edu|io|co|uk|info|biz|app|dev|xyz)\b', query, re.IGNORECASE)
                
                if is_website:
                    # Handle as website URL
                    if not query.startswith(('http://', 'https://')):
                        query = f'https://{query}'
                    success, message = self.system.open_website(query, browser)
                    self.speak(f"Opening {query}")
                else:
                    # Handle as regular search
                    success, message = self.system.web_search(query, browser)
                    self.speak(f"Searching for {query}")
                return
        
            
        else:
            self.speak("Sorry, I don't understand that command")


    def start_listening(self):
        """Main loop to listen for voice commands"""
        with sr.Microphone() as source:
            self.recognizer.adjust_for_ambient_noise(source, duration=1)
            self.speak("Voice assistant is ready")
            
            while self.is_listening:
                try:
                    self.textUpdated.emit("\nListening...")
                    audio = self.recognizer.listen(source, timeout=5)
                    text = self.recognizer.recognize_google(audio)
                    self.textUpdated.emit(f"You said: {text}")
                    self.process_command(text)
                
                except sr.WaitTimeoutError:
                    continue
                except sr.UnknownValueError:
                    continue
                except sr.RequestError:
                    self.speak("Sorry, there was an error with the speech recognition service")
                except Exception as e:
                    print(f"Error: {str(e)}")
                    continue

class VoiceAssistantGUI(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Voice Command Assistant")
        self.setGeometry(100, 100, 800, 600)

        # Set dark theme
        self.set_dark_theme()

        # Create central widget and layout
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)
        layout.setSpacing(15)
        layout.setContentsMargins(20, 20, 20, 20)

        # Create text display
        self.text_display = QTextEdit()
        self.text_display.setReadOnly(True)
        self.text_display.setStyleSheet("""
            QTextEdit {
                background-color: #2E2E2E;
                color: #FFFFFF;
                border: 1px solid #444444;
                border-radius: 5px;
                padding: 10px;
                font-size: 14px;
            }
        """)
        layout.addWidget(self.text_display)

        # Create control buttons
        button_style = """
            QPushButton {
                background-color: #444444;
                color: #FFFFFF;
                border: none;
                border-radius: 5px;
                padding: 10px;
                font-size: 14px;
            }
            QPushButton:hover {
                background-color: #555555;
            }
            QPushButton:disabled {
                background-color: #333333;
                color: #777777;
            }
        """

        self.start_button = QPushButton("Start Listening")
        self.start_button.setIcon(QIcon("icons/start.png"))  # Add an icon
        self.start_button.setStyleSheet(button_style)
        self.start_button.clicked.connect(self.start_assistant)
        layout.addWidget(self.start_button)

        self.stop_button = QPushButton("Stop Listening")
        self.stop_button.setIcon(QIcon("icons/stop.png"))  # Add an icon
        self.stop_button.setStyleSheet(button_style)
        self.stop_button.clicked.connect(self.stop_assistant)
        self.stop_button.setEnabled(False)
        layout.addWidget(self.stop_button)

        # Help text
        help_text = """
            <h3>Available Commands:</h3>
            <ul>
                <li>"Increase volume to 55%" or "Set volume 75%"</li>
                <li>"Open [filename/foldername]" (e.g., "Open Documents")</li>
                <li>"Open [application] app" (e.g., "Open Chrome app")</li>
                <li>"Play youtube" (opens homepage)</li>
                <li>"Play [songname/video]" (plays first YouTube result)</li>
                <li>"Search for Python tutorials" (in default browser)</li>
                <li>"Search cat pictures in Chrome" (specific browser)</li>
                <li>"Open deepseek.com" (in default browser)</li>
                <li>"Open youtube.com on Firefox"</li>
                <li>"Open github in Edge"</li>
                <li>"Close tab" (closes active browser tab)</li>
                <li>"Close [application] app" (e.g., "Close Chrome app")</li>
                <li>"Close [window/folder name]" (closes specific window)</li>
                <li>"Stop listening" or "Stop assistant"</li>
            </ul>
        """
        help_label = QLabel(help_text)
        help_label.setStyleSheet("""
            QLabel {
                color: #FFFFFF;
                font-size: 14px;
                padding: 10px;
                background-color: #2E2E2E;
                border-radius: 5px;
            }
        """)
        help_label.setWordWrap(True)
        layout.addWidget(help_label)

        # Copyright notice
        copyright_label = QLabel("© Copyright 2025 Sandaru Abey. All rights reserved.")
        copyright_label.setStyleSheet("""
            QLabel {
                color: #777777;
                font-size: 12px;
                text-align: center;
                padding: 10px;
            }
        """)
        layout.addWidget(copyright_label)

        # Initialize worker and thread
        self.worker = None
        self.thread = None

    def set_dark_theme(self):
        """Apply a dark theme to the application."""
        dark_palette = QPalette()
        dark_palette.setColor(QPalette.Window, QColor(53, 53, 53))
        dark_palette.setColor(QPalette.WindowText, Qt.white)
        dark_palette.setColor(QPalette.Base, QColor(35, 35, 35))
        dark_palette.setColor(QPalette.AlternateBase, QColor(53, 53, 53))
        dark_palette.setColor(QPalette.ToolTipBase, Qt.white)
        dark_palette.setColor(QPalette.ToolTipText, Qt.white)
        dark_palette.setColor(QPalette.Text, Qt.white)
        dark_palette.setColor(QPalette.Button, QColor(53, 53, 53))
        dark_palette.setColor(QPalette.ButtonText, Qt.white)
        dark_palette.setColor(QPalette.BrightText, Qt.red)
        dark_palette.setColor(QPalette.Highlight, QColor(42, 130, 218))
        dark_palette.setColor(QPalette.HighlightedText, Qt.black)
        QApplication.setPalette(dark_palette)

    def update_text_display(self, text):
        self.text_display.append(text)
        self.text_display.verticalScrollBar().setValue(
            self.text_display.verticalScrollBar().maximum()
        )

    def start_assistant(self):
        self.worker = VoiceAssistant()
        self.thread = QThread()

        self.worker.moveToThread(self.thread)
        self.worker.textUpdated.connect(self.update_text_display)

        self.thread.started.connect(self.worker.start_listening)
        self.thread.start()

        self.start_button.setEnabled(False)
        self.stop_button.setEnabled(True)

    def stop_assistant(self):
        if self.worker:
            self.worker.is_listening = False
        if self.thread:
            self.thread.quit()
            self.thread.wait()

        self.start_button.setEnabled(True)
        self.stop_button.setEnabled(False)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyle("Fusion")  # Use Fusion style for a modern look
    window = VoiceAssistantGUI()
    window.show()
    sys.exit(app.exec_())