import datetime
import math
import os
import random
import re
import sqlite3
import subprocess
import sys
import time
import urllib.parse
import webbrowser
from calendar import (day_name)
import pyaudio
import pyttsx3
import pywhatkit as kit
import speech_recognition as sr
from google import genai

p = pyaudio.PyAudio()
# --- External Library Check & Setup ---
# For pycaw (Windows volume control)
PYCAW_AVAILABLE = False
if os.name == 'nt':
    try:
        import comtypes
        from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
        from ctypes import cast, POINTER

        # comtypes.CoInitialize() # Optional: Initialize COM for stability
        PYCAW_AVAILABLE = True
        print("STATUS: pycaw modules imported successfully for Windows volume control.")
    except ImportError:
        print("WARNING: pycaw not installed. Volume control disabled on Windows. Use 'pip install pycaw'.")
    except Exception as e:
        print(f"WARNING: pycaw initialization failed: {e}")

# For pyautogui (Media Control & Screenshot)
PYAUTOGUI_AVAILABLE = False
try:
    import pyautogui

    PYAUTOGUI_AVAILABLE = True
    print("STATUS: pyautogui imported successfully for media control and screenshot.")
except ImportError:
    print("WARNING: pyautogui not installed. Media control and screenshot disabled. Use 'pip install pyautogui'.")

# For psutil (System Status)
PSUTIL_AVAILABLE = False
try:
    import psutil

    PYAUTOGUI_AVAILABLE = True
    print("STATUS: psutil imported successfully for system monitoring.")
except ImportError:
    print("WARNING: psutil not installed. System monitoring disabled. Use 'pip install psutil'.")

# For pyperclip (Clipboard Access)
PYPERCLIP_AVAILABLE = True
try:
    import pyperclip

    PYPERCLIP_AVAILABLE = True
    print("STATUS: pyperclip imported successfully for clipboard access.")
except ImportError:
    print("WARNING: pyperclip not installed. Clipboard search disabled. Use 'pip install pyperclip'.")

# For Wikipedia Search
WIKIPEDIA_AVAILABLE = False
try:
    import wikipedia

    WIKIPEDIA_AVAILABLE = True
    print("STATUS: wikipedia library imported successfully.")
except ImportError:
    print("WARNING: wikipedia library not installed. Wikipedia search disabled. Use 'pip install wikipedia'.")

# ==============================================================================
# 1. Configuration & Setup
# ==============================================================================
# --- DATABASE CONFIGURATION ---
DB_FILE = "friday_assistant.db"

# NOTE: Your provided Gemini API Key is placed here.
GEMINI_API_KEY = ("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")

# Global Variables
os.environ["GEMINI_API_KEY"] = GEMINI_API_KEY
LLM_AVAILABLE = False
client = None
WA_CONTACTS = {}
EMAIL_CONTACTS = {}

# --- TTS Engine Setup ---
engine = None
try:
    engine = pyttsx3.init()
    print("STATUS: TTS Engine initialized successfully.")
except Exception as e:
    print(f"CRITICAL ERROR: Failed to initialize any TTS engine. FRIDAY cannot speak. {e}")

# --- Application Dictionary for Voice Commands ---
APP_COMMANDS = {
    # Windows applications
    "notepad": "notepad.exe",
    "calculator": "calc.exe",
    "paint": "mspaint.exe",
    "chrome": "chrome.exe",
    "firefox": "firefox.exe",
    "word": "winword.exe",
    "excel": "excel.exe",
    "powerpoint": "powerpnt.exe",
    "file explorer": "explorer.exe",
    "command prompt": "cmd.exe",
    "task manager": "taskmgr.exe",
    "photo shop": "photoshop.exe",
    "corel draw": "coreldraw.exe",
    "tally": "tally.exe",
    "notion": "notion.exe",
    "Needforspeed": "Needforspeed.exe",

    # Cross-platform applications
    "browser": "chrome.exe",
    "text editor": "notepad.exe",
    "music": "spotify.exe",
    "video": "vlc.exe",

    # Web applications
    "youtube": "https://youtube.com",
    "google": "https://google.com",
    "gmail": "https://gmail.com",
    "github": "https://github.com",
    "facebook": "https://facebook.com",
    "twitter": "https://twitter.com",
    "linkedin": "https://linkedin.com",
    "whatsapp": "https://web.whatsapp.com",
    "netflix": "https://netflix.com",
    "amazon": "https://amazon.com",
}

# --- Close Commands Dictionary ---
CLOSE_COMMANDS = {
    "notepad": "notepad.exe",
    "calculator": "calc.exe",
    "paint": "mspaint.exe",
    "chrome": "chrome.exe",
    "firefox": "firefox.exe",
    "word": "winword.exe",
    "excel": "excel.exe",
    "powerpoint": "powerpnt.exe",
    "file explorer": "explorer.exe",
    "command prompt": "cmd.exe",
    "task manager": "taskmgr.exe",
    "photoshop": "photoshop.exe",
    "corel draw": "coreldraw.exe",
    "tally": "tally.exe",
    "notion": "notion.exe",
    "browser": "chrome.exe",
    "text editor": "notepad.exe",
    "music": "spotify.exe",
    "video": "vlc.exe",
}

# Global state to track microphone mute status
MICROPHONE_MUTED = False


# ==============================================================================
# 2. DATABASE MANAGEMENT FUNCTIONS
# ==============================================================================

def init_db():
    """Initializes the SQLite database and creates all tables."""
    global GEMINI_API_KEY
    conn = None
    try:
        conn = sqlite3.connect(DB_FILE)
        cursor = conn.cursor()

        # Configuration Table
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS configuration (
                key TEXT PRIMARY KEY,
                value TEXT NOT NULL
            )
        """)
        cursor.execute("INSERT OR REPLACE INTO configuration (key, value) VALUES (?, ?)",
                       ('GEMINI_API_KEY', GEMINI_API_KEY))

        # Contacts Table
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS contacts (
                contact_id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT UNIQUE NOT NULL,
                phone_number TEXT,
                email_address TEXT
            )
        """)
        initial_contacts = [
            ("mom", "+919890687695", None),
            ("dad", "+919890753754", None),
            ("mahek", None, "maheksundke@gmail.com"),
            ("heer", None, "heervinita1230@gmail.com")
        ]
        for name, phone, email in initial_contacts:
            cursor.execute("INSERT OR IGNORE INTO contacts (name, phone_number, email_address) VALUES (?, ?, ?)",
                           (name, phone, email))

        # Tasks Table
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS tasks (
                task_id INTEGER PRIMARY KEY AUTOINCREMENT,
                description TEXT NOT NULL,
                created_at TEXT NOT NULL,
                is_completed INTEGER NOT NULL DEFAULT 0
            )
        """)

        # Schedule Table
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS schedule (
                event_id INTEGER PRIMARY KEY AUTOINCREMENT,
                description TEXT NOT NULL,
                event_time TEXT,
                created_at TEXT NOT NULL
            )
        """)

        # Notes Table
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS notes (
                note_id INTEGER PRIMARY KEY AUTOINCREMENT,
                content TEXT NOT NULL,
                created_at TEXT NOT NULL
            )
        """)

        conn.commit()
        print("STATUS: SQLite database initialized and tables created/checked.")

    except sqlite3.Error as e:
        print(f"DATABASE ERROR: Failed to initialize database: {e}")
    finally:
        if conn:
            conn.close()


def load_contacts_from_db():
    """Loads contacts from the database into global WA_CONTACTS and EMAIL_CONTACTS."""
    global WA_CONTACTS, EMAIL_CONTACTS
    conn = None
    try:
        conn = sqlite3.connect(DB_FILE)
        cursor = conn.cursor()

        cursor.execute("SELECT name, phone_number, email_address FROM contacts")
        rows = cursor.fetchall()

        WA_CONTACTS = {}
        EMAIL_CONTACTS = {}

        for name, phone, email in rows:
            if phone:
                WA_CONTACTS[name.lower()] = phone
            if email:
                EMAIL_CONTACTS[name.lower()] = email

        print(f"STATUS: Loaded {len(WA_CONTACTS) + len(EMAIL_CONTACTS)} contacts from database.")
    except sqlite3.Error as e:
        print(f"DATABASE ERROR: Failed to load contacts: {e}")
    finally:
        if conn:
            conn.close()


def load_gemini_key():
    """Loads the Gemini API key from the database for initialization."""
    global GEMINI_API_KEY, LLM_AVAILABLE, client
    conn = None
    try:
        conn = sqlite3.connect(DB_FILE)
        cursor = conn.cursor()

        cursor.execute("SELECT value FROM configuration WHERE key = 'GEMINI_API_KEY'")
        result = cursor.fetchone()
        if result and result[0]:
            GEMINI_API_KEY = result[0]
            os.environ["GEMINI_API_KEY"] = GEMINI_API_KEY
            print("STATUS: Gemini API Key loaded from database.")

        # Initialize the Gemini client
        client = genai.Client(api_key=GEMINI_API_KEY)
        # Test the connection
        client.models.generate_content(model='gemini-2.5-flash', contents=["ping"], max_output_tokens=5)
        LLM_AVAILABLE = True
        print("STATUS: Gemini LLM initialized successfully.")

    except Exception as e:
        print(f"ERROR: Gemini Client initialization failed. {e}")
        print("FRIDAY will only execute built-in commands and won't answer general questions.")
    finally:
        if conn:
            conn.close()


# ==============================================================================
# 3. CORE FUNCTIONS
# ==============================================================================

def say(text):
    """Speaks the text using the initialized pyttsx3 engine."""
    print(f"FRIDAY said: {text}")
    if engine:
        engine.say(text)
        engine.runAndWait()
    else:
        print(f"[No speech output] {text}")


def takecommand():
    """Listens to the user's voice command, respecting the global mute state."""
    global MICROPHONE_MUTED
    if MICROPHONE_MUTED:
        print("Microphone is SOFT-MUTED. Listening is disabled.")
        time.sleep(0.5)
        return ""

    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening...")
        r.pause_threshold = 0.8
        r.adjust_for_ambient_noise(source, duration=1.0)
        audio = r.listen(source)

        try:
            query = r.recognize_google(audio, language='en-in')
            print(f"User said: {query}")
            return query
        except sr.UnknownValueError:
            return "unknown"
        except sr.RequestError:
            return "error"
        except Exception as e:
            print(f"CRITICAL MIC ERROR: Check your microphone connection/settings. Details: {e}")
            return ""


# ==============================================================================
# 3.1 NEW/UPDATED: INTERACTIVE GREETING FUNCTION
# ==============================================================================

def interactive_greeting():
    """Greets the user based on the current time of day with added flair."""
    now = datetime.datetime.now()
    hour = now.hour

    # Determine the time-based greeting
    if 5 <= hour < 12:
        greeting_time = "Good morning"
    elif 12 <= hour < 18:
        greeting_time = "Good afternoon"
    else:
        greeting_time = "Good evening"

    # Add a personal touch
    greetings = [
        f"{greeting_time}, sir. How may I be of assistance today?",
        f"{greeting_time}. I'm FRIDAY, ready for your command.",
        f"Hello, sir. {greeting_time}. I'm here and ready.",
        f"{greeting_time}! Systems are operational. What's the plan?"
    ]

    say(random.choice(greetings))


# ==============================================================================
# 4. VOICE COMMAND APP LAUNCHER (NEW FEATURE)
# ==============================================================================

def open_application_via_voice():
    """Enhanced application launcher that listens for voice commands"""
    say("Which application would you like me to open?")
    print("Listening for application name...")

    app_name = takecommand().lower().strip()

    if not app_name or app_name in ["unknown", "error"]:
        say("I didn't hear the application name. Please try again.")
        return

    # Remove common prefixes
    for prefix in ["open", "launch", "start"]:
        if app_name.startswith(prefix):
            app_name = app_name[len(prefix):].strip()
            break

    if not app_name:
        say("Please specify which application to open.")
        return

    say(f"Opening {app_name}")
    result = open_application(app_name)
    say(result)


def open_application(app_name):
    """Enhanced version: Attempts to open a program, file, or website."""
    app_name = app_name.strip().lower()

    # Check if it's a web URL first
    if app_name in APP_COMMANDS:
        target = APP_COMMANDS[app_name]

        # If it's a URL, open in browser
        if target.startswith("http"):
            webbrowser.open(target)
            return f"Opened {app_name} in your browser."

        # If it's an executable, try to run it
        try:
            if os.name == 'nt':  # Windows
                subprocess.Popen(target, shell=True)
            elif os.uname().sysname == 'Darwin':  # macOS
                subprocess.Popen(['open', target])
            else:  # Linux
                subprocess.Popen([target])
            return f"Successfully opened {app_name}."
        except Exception as e:
            return f"Failed to open {app_name}: {str(e)}"

    # If not in predefined commands, try to open directly
    try:
        if os.name == 'nt':  # Windows
            subprocess.Popen(app_name, shell=True)
        elif os.uname().sysname == 'Darwin':  # macOS
            subprocess.Popen(['open', app_name])
        else:  # Linux
            subprocess.Popen([app_name])
        return f"Attempting to open {app_name}."
    except Exception as e:
        return f"Failed to open {app_name}. Error: {str(e)}"


def list_available_apps():
    """Lists all available applications that can be opened by voice command"""
    say("Here are the applications I can open:")
    apps_list = list(APP_COMMANDS.keys())
    apps_text = ", ".join(apps_list)
    print(f"Available apps: {apps_text}")
    say(f"I can open: {apps_text}")


def add_custom_app_command(app_name, app_path):
    """Add a custom application to the command dictionary"""
    global APP_COMMANDS
    APP_COMMANDS[app_name.lower()] = app_path
    return f"Added {app_name} to available applications."


# ==============================================================================
# 5. APPLICATION CLOSING FUNCTIONS
# ==============================================================================

def close_application_via_voice():
    """Closes an application via voice command"""
    say("Which application would you like me to close?")
    print("Listening for application name...")

    app_name = takecommand().lower().strip()

    if not app_name or app_name in ["unknown", "error"]:
        say("I didn't hear the application name. Please try again.")
        return

    result = close_application(app_name)
    say(result)


def close_application(app_name):
    """Closes a specific application by name"""
    app_name = app_name.strip().lower()

    if app_name in CLOSE_COMMANDS:
        process_name = CLOSE_COMMANDS[app_name]

        if process_name.startswith("http"):
            return f"Cannot close web application {app_name}. Please close the browser tab manually."

        try:
            if os.name == 'nt':  # Windows
                # Use taskkill to terminate the process
                subprocess.run(f'taskkill /f /im "{process_name}"', shell=True, check=True)
                return f"Successfully closed {app_name}."
            else:
                # For macOS and Linux, use pkill
                subprocess.run(['pkill', '-f', process_name], check=True)
                return f"Successfully closed {app_name}."
        except subprocess.CalledProcessError:
            return f"Failed to close {app_name}. It may not be running or requires manual closure."
        except Exception as e:
            return f"Error closing {app_name}: {str(e)}"
    else:
        return f"I don't know how to close {app_name}. It's not in my list of closable applications."


def close_all_applications():
    """Closes all common applications"""
    say("Closing all common applications...")
    closed_apps = []
    failed_apps = []

    for app_name, process_name in CLOSE_COMMANDS.items():
        if not process_name.startswith("http"):  # Don't try to close web URLs
            result = close_application(app_name)
            if "Successfully" in result:
                closed_apps.append(app_name)
            else:
                failed_apps.append(app_name)
            time.sleep(0.5)  # Small delay between closing apps

    if closed_apps:
        say(f"Successfully closed: {', '.join(closed_apps)}")
    if failed_apps:
        say(f"Failed to close: {', '.join(failed_apps)}")

    return f"Closed {len(closed_apps)} applications."


def list_running_apps():
    """Lists currently running applications that can be closed"""
    if not PSUTIL_AVAILABLE:
        return "I cannot check running applications without the psutil library."

    try:
        running_apps = []
        for proc in psutil.process_iter(['name']):
            try:
                process_name = proc.info['name'].lower()
                # Check if this process matches any in our close commands
                for app_name, target_name in CLOSE_COMMANDS.items():
                    if target_name.lower() in process_name or app_name in process_name:
                        if app_name not in running_apps:
                            running_apps.append(app_name)
            except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                pass

        if running_apps:
            apps_text = ", ".join(running_apps)
            say(f"Currently running applications: {apps_text}")
            return f"Running apps: {apps_text}"
        else:
            say("No common applications are currently running.")
            return "No common applications running."
    except Exception as e:
        print(f"Error checking running apps: {e}")
        return "Failed to check running applications."


# ==============================================================================
# 6. WIKIPEDIA FUNCTIONS
# ==============================================================================

def wikipedia_voice_search():
    """Comprehensive Wikipedia search function"""
    if not WIKIPEDIA_AVAILABLE:
        say("I cannot search Wikipedia without the wikipedia library. Please install it using 'pip install wikipedia'.")
        return

    say("What would you like me to search on Wikipedia?")
    print("Listening for Wikipedia search query...")

    search_query = takecommand()

    if not search_query or search_query.lower() in ["unknown", "error"]:
        say("I didn't hear your search query. Please try again.")
        return

    search_query = search_query.strip()
    if not search_query:
        say("The search query appears to be empty. Please try again.")
        return

    say(f"Searching Wikipedia for {search_query}")
    print(f"\n{'=' * 60}")
    print(f"WIKIPEDIA SEARCH: {search_query}")
    print(f"{'=' * 60}")

    try:
        wikipedia.set_lang("en")
        summary = wikipedia.summary(search_query, sentences=5, auto_suggest=True)

        print(f"\nSEARCH RESULTS:")
        print(f"{'-' * 60}")
        print(summary)
        print(f"{'-' * 60}")

        # Read it aloud
        say(f"According to Wikipedia, about {search_query}:")

        # Offer to open the full page
        say("Would you like me to open the full Wikipedia page in your browser? Say yes or no.")
        open_choice = takecommand().lower()

        if "yes" in open_choice or "open" in open_choice or "please" in open_choice:
            try:
                page = wikipedia.page(search_query)
                webbrowser.open(page.url)
                say("Opening the full Wikipedia page in your browser.")
            except wikipedia.exceptions.DisambiguationError as e:
                if e.options:
                    try:
                        page = wikipedia.page(e.options[0])
                        webbrowser.open(page.url)
                        say(f"Opening the page for {e.options[0]} in your browser.")
                    except:
                        safe_query = urllib.parse.quote_plus(search_query)
                        webbrowser.open(f"https://en.wikipedia.org/wiki/{safe_query}")
            except:
                safe_query = urllib.parse.quote_plus(search_query)
                webbrowser.open(f"https://en.wikipedia.org/wiki/{safe_query}")
                say("Opening Wikipedia search results in your browser.")

    except wikipedia.exceptions.PageError:
        error_msg = f"I'm sorry, I couldn't find anything on Wikipedia about {search_query}."
        print(f"ERROR: {error_msg}")
        say(error_msg)
    except wikipedia.exceptions.DisambiguationError as e:
        error_msg = f"Your search for {search_query} is ambiguous. There are multiple options: {', '.join(e.options[:5])}"
        if len(e.options) > 5:
            error_msg += f" and {len(e.options) - 5} more."
        print(f"DISAMBIGUATION: {error_msg}")
        say(f"Your search is ambiguous. I found multiple options. {error_msg}")

        say("Would you like me to search for the first option? Say yes or no.")
        choice = takecommand().lower()
        if "yes" in choice:
            wikipedia_voice_search_specific(e.options[0])
    except Exception as e:
        error_msg = f"An unexpected error occurred during the Wikipedia search: {e}"
        print(f"ERROR: {error_msg}")
        say("I apologize, I encountered an issue while searching Wikipedia.")


def wikipedia_voice_search_specific(specific_query):
    """Helper function to search for a specific Wikipedia query"""
    try:
        say(f"Searching for {specific_query} on Wikipedia.")
        summary = wikipedia.summary(specific_query, sentences=5, auto_suggest=True)

        print(f"\nSPECIFIC SEARCH: {specific_query}")
        print(f"{'-' * 60}")
        print(summary)
        print(f"{'-' * 60}")

        say(f"According to Wikipedia: {summary}")

    except Exception as e:
        error_msg = f"Failed to search for {specific_query}: {e}"
        print(f"ERROR: {error_msg}")
        say(f"I couldn't find information about {specific_query}.")


def wikipedia_text_search(search_query):
    """Wikipedia search function for text input (command line)"""
    if not WIKIPEDIA_AVAILABLE:
        return "Wikipedia search requires the wikipedia library. Install with 'pip install wikipedia'."

    try:
        wikipedia.set_lang("en")
        summary = wikipedia.summary(search_query, sentences=5, auto_suggest=True)

        print(f"\n{'=' * 60}")
        print(f"WIKIPEDIA SEARCH: {search_query}")
        print(f"{'=' * 60}")
        print(summary)
        print(f"{'=' * 60}")

        if engine:
            say(f"According to Wikipedia: {summary}")

        return f"Wikipedia search completed for: {search_query}"

    except wikipedia.exceptions.PageError:
        return f"No Wikipedia page found for: {search_query}"
    except wikipedia.exceptions.DisambiguationError as e:
        return f"Ambiguous search. Options: {', '.join(e.options[:5])}"
    except Exception as e:
        return f"Search failed: {e}"


def wikipedia_text_to_speech(query):
    """Searches Wikipedia for the query, prints the summary, and speaks it aloud."""
    if not WIKIPEDIA_AVAILABLE:
        say("I cannot search Wikipedia without the wikipedia library. Please install it using 'pip install wikipedia'.")
        return True

    print(f"DEBUG: Wikipedia search initiated for query: '{query}'")

    if not query or len(query.split()) < 1:
        say("Please specify a topic to search on Wikipedia.")
        return True

    say(f"Searching Wikipedia for {query}")
    print(f"Searching Wikipedia for: {query}")
    try:
        result = wikipedia.summary(query, sentences=3, auto_suggest=True)
        print("-" * 50)
        print(f"WIKIPEDIA SUMMARY for '{query}':\n{result}")
        print("-" * 50)

        say(f"According to Wikipedia: {result}")
        return True
    except wikipedia.exceptions.PageError:
        error_message = f"I'm sorry, I couldn't find anything on Wikipedia about {query}."
        print(f"ERROR: {error_message}")
        say(error_message)
        return True
    except wikipedia.exceptions.DisambiguationError as e:
        first_option = e.options[0]
        say(f"Your search for {query} is ambiguous. I will try the first result: {first_option}.")
        return wikipedia_text_to_speech(first_option)
    except Exception as e:
        error_message = f"An unexpected error occurred during the Wikipedia search: {e}"
        print(f"ERROR: {error_message}")
        say("I apologize, I encountered an issue while searching Wikipedia.")
        return True


# ==============================================================================
# 7. MAIN COMMAND PROCESSING FUNCTION (UPDATED)
# ==============================================================================

def process_query(query):
    """Central function to process a command (either from voice or text)."""
    if not query:
        return False

    query_lower = query.lower()
    command_processed = False

    # --- Interactive Greeting (NEW/UPDATED) ---
    if any(g in query_lower for g in
           ["hello", "hi", "hey", "good morning", "good afternoon", "good evening", "friday", "jarvis"]):
        interactive_greeting()
        command_processed = True

    # --- Exit Command (Always first) ---
    elif "exit" in query_lower or "stop" in query_lower or "quit" in query_lower or "shut down" in query_lower:
        # Check if user means 'shut down computer'
        if "computer" in query_lower or "system" in query_lower:
            shutdown_system()
        else:
            say("Goodbye. Shutting down now.")
            if engine:
                engine.stop()
            sys.exit(0)

    # === SHORTENED APP LAUNCHER COMMANDS ===
    elif query_lower.startswith("open"):
        app_name = query_lower.replace("open", "").strip()
        if app_name:
            say(f"Opening {app_name}")
            open_application(app_name)
        else:
            open_application_via_voice()
        command_processed = True

    elif "list apps" in query_lower or "apps list" in query_lower:
        list_available_apps()
        command_processed = True

    # === APPLICATION CLOSING COMMANDS (SHORTENED) ===
    elif query_lower.startswith("close"):
        target = query_lower.replace("close", "").strip()
        if target:
            if target == "all":
                close_all_applications()
            else:
                say(close_application(target))
        else:
            close_application_via_voice()
        command_processed = True

    elif "running apps" in query_lower:
        list_running_apps()
        command_processed = True

    # --- MICROPHONE CONTROL COMMANDS (SHORTENED) ---
    elif "mute mic" in query_lower or "mute microphone" in query_lower:
        toggle_microphone_mute("mute")
        command_processed = True

    elif "unmute mic" in query_lower or "unmute microphone" in query_lower:
        toggle_microphone_mute("unmute")
        command_processed = True

    # --- WIKIPEDIA VOICE SEARCH COMMANDS (SHORTENED) ---
    elif query_lower.startswith("wiki"):
        topic = query_lower.replace("wiki", "").strip()
        if topic:
            wikipedia_voice_search_specific(topic)
        else:
            wikipedia_voice_search()
        command_processed = True

    elif "tell me about" in query_lower:
        topic = query_lower.replace("tell me about", "").strip()
        if topic:
            wikipedia_voice_search_specific(topic)
        else:
            wikipedia_voice_search()
        command_processed = True

    # --- UNIT CONVERSION (SHORTENED) ---
    elif "convert" in query_lower and re.search(r'\d', query_lower):
        say(convert_units(query_lower))
        command_processed = True

    # --- TO-DO LIST MANAGEMENT (SHORTENED) ---
    elif query_lower.startswith("add task"):
        item = query_lower.replace("add task", "").strip()
        if item:
            say(add_todo_item(item))
        else:
            say("What item would you like to add to your list?")
            item = takecommand()
            if item and item.lower() not in ["unknown", "error"]:
                say(add_todo_item(item))
            else:
                say("No item specified. Task cancelled.")
        command_processed = True

    elif "show tasks" in query_lower or "read tasks" in query_lower:
        say(read_todo_list())
        command_processed = True

    elif "clear tasks" in query_lower or "delete tasks" in query_lower:
        say(clear_todo_list())
        command_processed = True

    # --- SYSTEM CONTROL COMMANDS (SHORTENED) ---
    elif "lock system" in query_lower or "lock computer" in query_lower:
        lock_system()
        command_processed = True

    # --- TIME/DATE COMMANDS (SHORTENED) ---
    elif "what's the time" in query_lower or query_lower == "time":
        tell_time()
        command_processed = True

    elif "what day is it" in query_lower or query_lower == "day":
        tell_day_of_week()
        command_processed = True

    elif "date in a week" in query_lower:
        get_date_in_a_week()
        command_processed = True

    # --- FILE MANAGEMENT COMMANDS (SHORTENED) ---
    elif query_lower.startswith("create file"):
        filename_query = query_lower.replace("create file", "").strip()
        if filename_query:
            filename = filename_query.replace(" dot ", ".").replace(" ", "_").strip()
            say(create_new_file(filename))
        else:
            if not len(sys.argv) > 1:
                say("What would you like to name the new file? Please include the extension, like 'document dot txt'.")
                filename_query = takecommand()

            if filename_query and filename_query.lower() not in ["unknown", "error"]:
                filename = filename_query.replace(" dot ", ".").replace(" ", "_").strip()
                say(create_new_file(filename))
            else:
                say("I did not hear a valid filename. File creation cancelled.")
        command_processed = True

    # --- UTILITY COMMANDS (SHORTENED) ---
    elif query_lower.startswith("google"):
        search_term = query_lower.replace("google", "").strip()
        if search_term:
            google_search(search_term)
        else:
            if not len(sys.argv) > 1:
                say("What would you like me to search for on Google?")
                search_query = takecommand()
                if search_query and search_query.lower() not in ["unknown", "error"]:
                    google_search(search_query)
                else:
                    say("I did not hear a search query.")
            else:
                say("Please provide a search term after the command.")
        command_processed = True

    elif query_lower.startswith("weather"):
        city_name = query_lower.replace("weather", "").strip()
        if city_name:
            say(get_weather_forecast(city_name))
        else:
            if not len(sys.argv) > 1:
                say("Which city would you like the weather forecast for?")
                city_name = takecommand()

            if city_name and city_name.lower() not in ["unknown", "error"]:
                say(get_weather_forecast(city_name))
            else:
                say("I did not catch the city name. Please try again.")
        command_processed = True

    elif query_lower.startswith("take note"):
        note_to_save = query_lower.replace("take note", "").strip()

        if note_to_save:
            say(take_quick_note(note_to_save))
        else:
            if not len(sys.argv) > 1:
                say("What would you like me to remember?")
                note_to_save = takecommand()

            if note_to_save and note_to_save.lower() not in ["unknown", "error"]:
                say(take_quick_note(note_to_save))
            else:
                say("I didn't hear the note. Note cancelled.")
        command_processed = True

    # --- MEDIA AND SYSTEM COMMANDS (SHORTENED) ---
    elif any(cmd in query_lower for cmd in ["start", "pause", "next", "previous", "stop music"]):
        action = None
        if "start" in query_lower or "resume" in query_lower:
            action = "start"
        elif "pause" in query_lower:
            action = "pause"
        elif "next" in query_lower:
            action = "next"
        elif "previous" in query_lower:
            action = "previous"
        elif "stop" in query_lower:
            action = "stop"

        if action:
            media_control(action)
            command_processed = True

    elif query_lower.startswith("play song"):
        song_name = query_lower.replace("play song", "").strip()
        if song_name:
            play_song_on_youtube(song_name)
        else:
            if not len(sys.argv) > 1:
                say("What is the name of the song or artist you want to play?")
                song_query = takecommand()
                if song_query and song_query.lower() not in ["unknown", "error"]:
                    play_song_on_youtube(song_query)
                else:
                    say("I couldn't identify the song name.")
            else:
                say("Please include the song name after the command.")
        command_processed = True

    elif "screenshot" in query_lower or "capture" in query_lower:
        take_screenshot()
        command_processed = True

    elif "system status" in query_lower or "check performance" in query_lower:
        get_system_status()
        command_processed = True

    # --- CALCULATOR COMMANDS (SHORTENED) ---
    elif "calculate" in query_lower or (
            "what is" in query_lower and re.search(r'\d', query_lower)):
        math_query = query_lower.split("calculate", 1)[-1].strip() if "calculate" in query_lower else query_lower
        result_text = perform_calculation(math_query)
        say(result_text)
        command_processed = True

    # --- MESSAGING COMMANDS ---
    elif "whatsapp" in query_lower:
        if not len(sys.argv) > 1:
            say("To whom should I send the WhatsApp message? Say their name.")
            recipient_name = takecommand().lower()

            if recipient_name in WA_CONTACTS:
                recipient_number = WA_CONTACTS[recipient_name]
                say(f"What is the message you want to send to {recipient_name}?")
                message_body = takecommand()
                send_whatsapp_message(recipient_number, message_body)
            else:
                say(f"I don't have a WhatsApp number for {recipient_name}. Cancelling message.")
        else:
            say("For sending a WhatsApp message via text, I recommend composing it manually.")
        command_processed = True

    return command_processed


# ==============================================================================
# 8. LLM INTERACTION FUNCTION
# ==============================================================================

def get_llm_response(prompt):
    """Sends the user's non-command query to the Gemini API for a conversational response."""
    if not LLM_AVAILABLE:
        return random.choice([
            "My external intelligence module is offline. I can still process commands like open and calculate.",
            "I can't answer general questions right now."
        ])

    print("FRIDAY said: Thinking...")
    try:
        response = client.models.generate_content(
            model='gemini-2.5-flash',
            contents=[
                {"role": "user", "parts": [
                    {
                        "text": f"You are a helpful, smart personal assistant named FRIDAY. Be concise and friendly. Answer the following request: {prompt}"}
                ]}
            ]
        )
        return response.text
    except Exception as e:
        print(f"LLM Error during generation: {e}")
        return "I apologize, but I failed to get an answer from the intelligence module."


# ==============================================================================
# 9. SYSTEM CONTROL FUNCTIONS
# ==============================================================================

def toggle_microphone_mute(action):
    """Toggles the soft-mute state of the listening function."""
    global MICROPHONE_MUTED
    action = action.lower()

    if action == "mute" or action == "disable":
        target_state = True
        message = "Microphone muted. I will stop listening until you unmute me."
    elif action == "unmute" or action == "enable":
        target_state = False
        message = "Microphone unmuted. I am now listening again."
    else:
        say("I did not recognize the microphone control command.")
        return

    if MICROPHONE_MUTED != target_state:
        MICROPHONE_MUTED = target_state
        say(message)
    else:
        if target_state:
            say("The microphone is already muted.")
        else:
            say("The microphone is already unmuted.")


def lock_system():
    """Immediately locks the user session."""
    say("Locking the system now. Please ensure your session is password-protected.")
    try:
        if os.name == 'nt':  # Windows
            subprocess.run('rundll32.exe user32.dll,LockWorkStation', shell=True)
        elif os.uname().sysname == 'Darwin':  # macOS
            subprocess.run(['pmset', 'sleepnow'], check=True)
        elif os.name == 'posix':  # General Linux
            subprocess.run(['xdg-screensaver', 'lock'], check=False)
        else:
            say("System lock is not supported on this operating system.")
            return False
        return True
    except Exception as e:
        print(f"System lock error: {e}")
        say("I failed to execute the lock command.")
        return False


def shutdown_system():
    """Immediately initiates system shutdown."""
    say("WARNING: Initiating system shutdown. Say 'cancel' immediately if this is incorrect.")
    time.sleep(3)

    if len(sys.argv) > 1 and "cancel" in " ".join(sys.argv[1:]).lower():
        say("Shutdown cancelled.")
        return

    say("Starting shutdown sequence now. Goodbye.")
    try:
        if os.name == 'nt':  # Windows
            subprocess.run(['shutdown', '/s', '/t', '0'], check=True)
        elif os.uname().sysname == 'Darwin' or os.name == 'posix':
            subprocess.run(['shutdown', 'now'], check=True)
        else:
            say("Shutdown is not supported on this operating system. Exiting Python script instead.")
            sys.exit(0)
    except Exception as e:
        print(f"Shutdown error: {e}")
        say("I failed to execute the shutdown command. Exiting Python script instead.")
        sys.exit(0)

    sys.exit(0)


# ==============================================================================
# 10. DATE AND TIME FUNCTIONS
# ==============================================================================

def tell_time():
    """Tells the current time."""
    now = datetime.datetime.now()
    current_time = now.strftime("%I:%M %p")
    say(f"The current time is {current_time}")


def tell_day_of_week():
    """Tells the current day of the week."""
    current_day_index = datetime.datetime.today().weekday()
    day = day_name[current_day_index]
    say(f"Today is {day}")


def get_date_in_a_week():
    """Calculates and announces the date one week from today."""
    future_date = datetime.datetime.now() + datetime.timedelta(weeks=1)
    date_str = future_date.strftime("%B %d, %Y")
    say(f"The date one week from today will be {date_str}.")


# ==============================================================================
# 11. MEDIA AND SYSTEM CONTROL FUNCTIONS
# ==============================================================================

def media_control(action):
    """Controls media playback using simulated key presses."""
    if not PYAUTOGUI_AVAILABLE:
        say("I cannot control media playback without the pyautogui library.")
        return
    action = action.lower()
    say(f"Executing media action: {action}.")
    try:
        if action == "start" or action == "pause" or action == "resume":
            pyautogui.press('playpause')
        elif action == "next":
            pyautogui.press('nexttrack')
        elif action == "previous":
            pyautogui.press('prevtrack')
        elif action == "stop":
            pyautogui.press('stop')
        else:
            say("I did not recognize that media command.")
    except Exception as e:
        print(f"Media control error: {e}")
        say("I encountered an error trying to send the media command.")


def take_screenshot():
    """Takes a screenshot and saves it to the current directory."""
    if not PYAUTOGUI_AVAILABLE:
        say("I cannot take a screenshot without the pyautogui library.")
        return
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"screenshot_{timestamp}.png"
    try:
        screenshot = pyautogui.screenshot()
        screenshot.save(filename)
        say(f"Screenshot successfully captured and saved as {filename} in the current folder.")
    except Exception as e:
        print(f"Screenshot error: {e}")
        say("I encountered an error while trying to take the screenshot.")


def get_system_status():
    """Provides a report on CPU, RAM, and Battery usage."""
    if not PSUTIL_AVAILABLE:
        say("I cannot check the system status without the psutil library.")
        return
    try:
        cpu_percent = psutil.cpu_percent(interval=1)
        memory = psutil.virtual_memory()
        ram_percent = memory.percent
        battery = psutil.sensors_battery()
        battery_report = ""
        if battery:
            battery_percent = battery.percent
            plugged = "and charging" if battery.power_plugged else "and not charging"
            battery_report = f"Your battery is at {battery_percent} percent, {plugged}. "
        else:
            battery_report = "Battery status is unavailable. "
        report = (
            f"Here is the system status report. "
            f"CPU usage is currently at {cpu_percent} percent. "
            f"System memory usage is {ram_percent} percent. "
            f"{battery_report}"
        )
        say(report)
    except Exception as e:
        print(f"System status error: {e}")
        say("I encountered an error while trying to check the system status.")


# ==============================================================================
# 12. WEB AND SEARCH FUNCTIONS
# ==============================================================================
# ==============================================================================
# 12. WEB AND SEARCH FUNCTIONS (UPDATED FOR STABILITY)
# ==============================================================================

def play_song_on_youtube(song_name):
    """
    Uses the web browser to open a YouTube search for the song,
    providing a more stable alternative to pywhatkit.
    """
    song_name = song_name.strip()
    if not song_name:
        say("Please specify a song name to play.")
        return

    say(f"Searching YouTube for {song_name}. Opening the video search results now.")
    try:
        # 1. Safely encode the search query (e.g., 'A Sky Full of Stars' -> 'A+Sky+Full+of+Stars')
        safe_query = urllib.parse.quote_plus(f"{song_name} song official video")

        # 2. Construct the direct YouTube search URL
        url = f"https://www.youtube.com/results?search_query={safe_query}"

        # --- DEBUG LINE ADDED ---
        print(f"DEBUG: Attempting to open URL: {url}")
        # ------------------------

        # 3. Open the URL in the default web browser
        webbrowser.open(url)

    except Exception as e:
        print(f"YouTube open error: {e}")
        say("I encountered an error trying to open YouTube in your browser.")

def get_weather_forecast(city):
    """Performs a Google search for the weather forecast in a given city."""
    city = city.strip()
    if not city:
        return "Please provide a city name."

    query = f"weather forecast in {city}"
    google_search(query)
    return f"Displaying the weather forecast for {city}."


def google_search(query):
    """Opens a general Google search for a user-provided query."""
    query = query.strip()
    if not query:
        say("Please provide a search term.")
        return
    safe_query = urllib.parse.quote_plus(query)
    url = f"https://www.google.com/search?q={safe_query}"
    say(f"Searching Google for {query}. Opening the search results now.")
    webbrowser.open(url)


def send_whatsapp_message(phone_number, message):
    """Uses pywhatkit to open WhatsApp Web and automatically send the message."""
    say("Opening WhatsApp Web now. The message will be sent automatically after a short delay.")
    try:
        kit.sendwhatmsg_instantly(phone_number, message, wait_time=20, tab_close=True)
        say("The WhatsApp message has been automatically sent.")
    except Exception as e:
        say("I encountered an error while sending the WhatsApp message. Ensure WhatsApp Web is ready.")
        print(f"WhatsApp send error: {e}")


# ==============================================================================
# 13. DATABASE PERSISTENCE FUNCTIONS
# ==============================================================================

def take_quick_note(note_content):
    """Saves a quick note to the SQLite database's notes table."""
    conn = None
    try:
        note_content = note_content.strip()
        if not note_content:
            return "The note cannot be empty."

        conn = sqlite3.connect(DB_FILE)
        cursor = conn.cursor()
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        cursor.execute("INSERT INTO notes (content, created_at) VALUES (?, ?)", (note_content, timestamp))
        conn.commit()
        return f"I have saved the note: '{note_content}'."
    except sqlite3.Error as e:
        print(f"Note taking error: {e}")
        return f"I failed to save the note due to a database error."
    finally:
        if conn:
            conn.close()


def add_todo_item(item):
    """Adds a new item to the global to-do list in the database."""
    conn = None
    try:
        item = item.strip()
        if not item:
            return "The task item cannot be empty."

        conn = sqlite3.connect(DB_FILE)
        cursor = conn.cursor()
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        cursor.execute("INSERT INTO tasks (description, created_at, is_completed) VALUES (?, ?, ?)",
                       (item, timestamp, 0))
        conn.commit()
        return f"I've added '{item}' to your to-do list."
    except sqlite3.Error as e:
        print(f"To-do add error: {e}")
        return f"I failed to add the task due to a database error."
    finally:
        if conn:
            conn.close()


def read_todo_list():
    """Reads and speaks the current items on the to-do list from the database."""
    conn = None
    try:
        conn = sqlite3.connect(DB_FILE)
        cursor = conn.cursor()

        cursor.execute("SELECT description FROM tasks WHERE is_completed = 0 ORDER BY task_id")
        tasks = cursor.fetchall()

        if not tasks:
            return "Your to-do list is empty. You have no pending tasks."
        else:
            task_list = [t[0] for t in tasks]
            output = "Here are your pending tasks: " + ", ".join(task_list)
            return output
    except sqlite3.Error as e:
        print(f"To-do read error: {e}")
        return "I failed to retrieve your to-do list due to a database error."
    finally:
        if conn:
            conn.close()


def clear_todo_list():
    """Clears all tasks from the to-do list in the database."""
    conn = None
    try:
        conn = sqlite3.connect(DB_FILE)
        cursor = conn.cursor()

        cursor.execute("DELETE FROM tasks")
        conn.commit()
        return "Your to-do list has been completely cleared. All tasks deleted."
    except sqlite3.Error as e:
        print(f"To-do clear error: {e}")
        return "I failed to clear the to-do list due to a database error."
    finally:
        if conn:
            conn.close()


# ==============================================================================
# 14. FILE MANAGEMENT FUNCTIONS
# ==============================================================================

def create_new_file(filename):
    """Creates an empty file with the specified name."""
    try:
        filename = filename.strip(' "')
        if not filename:
            return "File creation failed. The filename cannot be empty."

        if os.path.exists(filename):
            return f"A file named {filename} already exists. Please choose a different name."

        with open(filename, 'w') as f:
            pass
        return f"Successfully created a new file named {filename}."
    except Exception as e:
        print(f"File creation error: {e}")
        return f"I failed to create the file. Error: {str(e)}"


# ==============================================================================
# 15. MATH & CALCULATION FUNCTIONS
# ==============================================================================

def convert_units(query):
    """Handles basic unit conversions."""
    match = re.search(r'convert\s+(\d+(\.\d+)?)\s+([\w\s]+?)\s+to\s+([\w\s]+)', query)
    if not match:
        return "I need the command in the format: convert <number> <unit1> to <unit2>."
    try:
        value = float(match.group(1))
        unit1 = match.group(3).strip().lower()
        unit2 = match.group(4).strip().lower()

        conversions = {
            'miles': 1609.34, 'kilometers': 1000, 'meters': 1, 'feet': 0.3048, 'cm': 0.01,
            'kilograms': 1, 'pounds': 0.453592, 'ounces': 0.0283495,
        }

        if ('celsius' in unit1 and 'fahrenheit' in unit2) or ('fahrenheit' in unit1 and 'celsius' in unit2):
            if 'celsius' in unit1:
                result = (value * 9 / 5) + 32
                return f"{value} Celsius is {round(result, 2)} Fahrenheit."
            else:
                result = (value - 32) * 5 / 9
                return f"{value} Fahrenheit is {round(result, 2)} Celsius."

        if unit1 not in conversions or unit2 not in conversions:
            return f"I do not recognize one or both units: {unit1} or {unit2}. I currently support length and mass conversions."

        base_value = value * conversions[unit1]
        final_value = base_value / conversions[unit2]

        return f"{value} {unit1} is approximately {round(final_value, 4)} {unit2}."

    except Exception as e:
        print(f"Conversion error: {e}")
        return "I encountered an error during the unit conversion."


def perform_calculation(math_query):
    """Parses natural language math query and calculates the result."""
    query = math_query.lower().strip()

    # Trigonometry
    if any(op in query for op in ["sin", "cos", "tan"]):
        try:
            op_match = re.search(r'(sin|cos|tan)\s*of\s*(\d+(\.\d+)?)', query)
            if op_match:
                op, value_str = op_match.groups()[0], op_match.groups()[1]
                value = float(value_str)
                rad_value = math.radians(value)

                if op == "sin":
                    result = math.sin(rad_value)
                elif op == "cos":
                    result = math.cos(rad_value)
                elif op == "tan":
                    if value % 180 == 90:
                        return f"Tangent of {value} degrees is undefined."
                    result = math.tan(rad_value)
                return f"The {op} of {value} degrees is approximately {round(result, 4)}."
            else:
                return "Please state the trigonometric function and the number (e.g., 'sin of 30')."
        except Exception as e:
            print(f"Trigonometry error: {e}")
            return "An error occurred during the trigonometric calculation."

    # General Arithmetic
    query = query.replace(" plus ", "+").replace(" minus ", "-").replace(" times ", "*").replace(" multiplied by ", "*")
    query = query.replace(" divided by ", "/").replace(" power ", "**").replace(" raised to ", "**")

    safe_chars = '0123456789.+-*/()** '
    sanitized_query = "".join(c for c in query if c in safe_chars)
    if sanitized_query.strip() != query.strip():
        return "I can only perform basic arithmetic and common math functions. Please use clear numbers and operators."

    try:
        result = eval(sanitized_query)
        return f"The result is {round(result, 4)}."
    except ZeroDivisionError:
        return "Not defined."
    except SyntaxError:
        return "That does not look like a valid mathematical expression."
    except Exception as e:
        print(f"General Calculation error: {e}")
        return "An unknown error occurred during the calculation."


# ==============================================================================
# 16. MAIN EXECUTION
# ==============================================================================

if __name__ == "__main__":
    # Initialize database and services
    init_db()
    load_contacts_from_db()
    load_gemini_key()

    # Text command mode
    if len(sys.argv) > 1:
        text_query = " ".join(sys.argv[1:])
        print(f"Processing text command: {text_query}")

        if text_query.lower().startswith(("wikipedia", "wiki", "search wikipedia")):
            search_term = text_query.lower()
            for prefix in ["wikipedia", "wiki", "search wikipedia"]:
                if search_term.startswith(prefix):
                    search_term = search_term[len(prefix):].strip()
                    break

            if search_term:
                result = wikipedia_text_search(search_term)
                say(result)
            else:
                say("Please specify what to search on Wikipedia.")
        else:
            command_processed = process_query(text_query)
            if not command_processed:
                response = get_llm_response(text_query)
                say(response)

        sys.exit(0)

    # Voice loop mode
    interactive_greeting()  # Using the new interactive greeting

    try:
        while True:
            query = takecommand()

            if not query:
                continue

            query_lower = query.lower()

            if query_lower in ["unknown", "error"]:
                say("I did not catch that. Could you please repeat your request?")
                time.sleep(1)
                continue

            command_processed = process_query(query)

            if not command_processed:
                response = get_llm_response(query)
                say(response)

    except Exception as e:
        print(f"\nCRASH DETECTED: An unexpected error occurred in the main loop: {e}")
        say("I've encountered a critical error and need to stop.")
        sys.exit(1)