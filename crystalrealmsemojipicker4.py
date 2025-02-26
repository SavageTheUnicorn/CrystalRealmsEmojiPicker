# Created by SavageTheUnicorn with a bit of AI love
# Date: 02/24/2025
# File Hash and source code can be found at: "https://github.com/SavageTheUnicorn/CrystalRealmsEmojiPicker"

import tkinter as tk
from tkinter import Entry, Listbox, Scrollbar, Button, Label, Frame, Menu, StringVar, Checkbutton, BooleanVar
from tkinter import colorchooser, Toplevel, ttk, IntVar
import requests
import pyperclip
from fontTools.ttLib import TTFont
import os
import sys
import json
import pystray
from PIL import Image, ImageDraw, ImageTk
from threading import Thread, Lock
import win32gui
import win32process
import win32api
import win32con
import queue
import psutil
import ctypes
import time
import traceback
import threading
import win32clipboard

# Additional import for unique window class
import uuid

# Define hotkey constants
MOD_ALT = 0x0001
MOD_CONTROL = 0x0002
MOD_SHIFT = 0x0004
MOD_WIN = 0x0008
WM_HOTKEY = 0x0312

# Virtual key codes for common keys
VK_CODES = {
    'a': 0x41, 'b': 0x42, 'c': 0x43, 'd': 0x44, 'e': 0x45, 'f': 0x46, 'g': 0x47, 'h': 0x48,
    'i': 0x49, 'j': 0x4A, 'k': 0x4B, 'l': 0x4C, 'm': 0x4D, 'n': 0x4E, 'o': 0x4F, 'p': 0x50,
    'q': 0x51, 'r': 0x52, 's': 0x53, 't': 0x54, 'u': 0x55, 'v': 0x56, 'w': 0x57, 'x': 0x58,
    'y': 0x59, 'z': 0x5A, '0': 0x30, '1': 0x31, '2': 0x32, '3': 0x33, '4': 0x34, '5': 0x35,
    '6': 0x36, '7': 0x37, '8': 0x38, '9': 0x39, 'BackSpace': 0x08, 'Tab': 0x09, 'Return': 0x0D,
    'Escape': 0x1B, 'space': 0x20, 'Delete': 0x2E, 'Prior': 0x21, 'Next': 0x22, 'End': 0x23,
    'Home': 0x24, 'Left': 0x25, 'Up': 0x26, 'Right': 0x27, 'Down': 0x28, 'F1': 0x70, 'F2': 0x71,
    'F3': 0x72, 'F4': 0x73, 'F5': 0x74, 'F6': 0x75, 'F7': 0x76, 'F8': 0x77, 'F9': 0x78,
    'F10': 0x79, 'F11': 0x7A, 'F12': 0x7B, 'slash': 0xBF, 'backslash': 0xDC, 'bracketleft': 0xDB,
    'bracketright': 0xDD, 'minus': 0xBD, 'plus': 0xBB, 'equal': 0xBB, 'comma': 0xBC, 'period': 0xBE,
    'semicolon': 0xBA, 'apostrophe': 0xDE, 'grave': 0xC0, 'numbersign': 0xBF, 'colon': 0xBA
}

EMOJI_API = "https://emojihub.yurace.pro/api/all"
PAGE_SIZE = 20

# Define the theme presets
THEMES = {
    "Default": {
        "name": "Default",
        "background": None,  # Uses system default
        "primary": None,
        "accent": None,
        "text": None,
        "listbox_bg": None,
        "listbox_fg": None,
        "entry_bg": None,
        "entry_fg": None,
        "button_bg": None,
        "button_fg": None,
        "statusbar_bg": None,
        "statusbar_fg": None
    },
    "Classic Light": {
        "name": "Classic Light",
        "background": "#F8F9FA",
        "primary": "#007BFF",
        "accent": "#17A2B8",
        "text": "#212529",
        "listbox_bg": "#FFFFFF",
        "listbox_fg": "#212529",
        "entry_bg": "#FFFFFF",
        "entry_fg": "#212529",
        "button_bg": "#007BFF",
        "button_fg": "#FFFFFF",
        "statusbar_bg": "#E9ECEF",
        "statusbar_fg": "#212529"
    },
    "Dark Mode": {
        "name": "Dark Mode",
        "background": "#121212",
        "primary": "#BB86FC",
        "accent": "#03DAC6",
        "text": "#E0E0E0",
        "listbox_bg": "#1E1E1E",
        "listbox_fg": "#E0E0E0",
        "entry_bg": "#2D2D2D",
        "entry_fg": "#E0E0E0",
        "button_bg": "#BB86FC",
        "button_fg": "#121212",
        "statusbar_bg": "#1E1E1E",
        "statusbar_fg": "#E0E0E0"
    },
    "Cyberpunk Neon": {
        "name": "Cyberpunk Neon",
        "background": "#080808",
        "primary": "#FF00FF",
        "accent": "#00FFFF",
        "text": "#FFFFFF",
        "listbox_bg": "#101010",
        "listbox_fg": "#FFFFFF",
        "entry_bg": "#181818",
        "entry_fg": "#FF00FF",
        "button_bg": "#FF00FF",
        "button_fg": "#000000",
        "statusbar_bg": "#101010",
        "statusbar_fg": "#00FFFF"
    },
    "Solarized Warm": {
        "name": "Solarized Warm",
        "background": "#FDF6E3",
        "primary": "#268BD2",
        "accent": "#DC322F",
        "text": "#073642",
        "listbox_bg": "#EEE8D5",
        "listbox_fg": "#073642",
        "entry_bg": "#FDF6E3",
        "entry_fg": "#073642",
        "button_bg": "#268BD2",
        "button_fg": "#FDF6E3",
        "statusbar_bg": "#EEE8D5",
        "statusbar_fg": "#073642"
    },
    "Forest Green": {
        "name": "Forest Green",
        "background": "#1B2B1A",
        "primary": "#3D9970",
        "accent": "#FFD700",
        "text": "#EAEAEA",
        "listbox_bg": "#243623",
        "listbox_fg": "#EAEAEA",
        "entry_bg": "#2D4129",
        "entry_fg": "#EAEAEA",
        "button_bg": "#3D9970",
        "button_fg": "#EAEAEA",
        "statusbar_bg": "#243623",
        "statusbar_fg": "#FFD700"
    },
    "Ocean Blue": {
        "name": "Ocean Blue",
        "background": "#1B2F40",
        "primary": "#0074D9",
        "accent": "#FFDC00",
        "text": "#EAEAEA",
        "listbox_bg": "#243A4F",
        "listbox_fg": "#EAEAEA",
        "entry_bg": "#2D465D",
        "entry_fg": "#EAEAEA",
        "button_bg": "#0074D9",
        "button_fg": "#EAEAEA",
        "statusbar_bg": "#243A4F",
        "statusbar_fg": "#FFDC00"
    }
}

def get_data_path():
    """Get path to data file in user's AppData directory"""
    app_data = os.path.join(os.getenv('APPDATA'), 'EmojiPicker')
    if not os.path.exists(app_data):
        os.makedirs(app_data)
    return os.path.join(app_data, 'emoji_data.json')

def get_settings_path():
    """Get path to settings file in user's AppData directory"""
    app_data = os.path.join(os.getenv('APPDATA'), 'EmojiPicker')
    if not os.path.exists(app_data):
        os.makedirs(app_data)
    return os.path.join(app_data, 'settings.json')

def resource_path(relative_path):
    """Get absolute path to resource, works for dev and for PyInstaller"""
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    
    return os.path.join(base_path, relative_path)

NOTO_EMOJI_PATH = resource_path("NotoEmoji-Regular.ttf")

class EmojiData:
    def __init__(self, name, unicode_key, count=0, pinned=False, manual_order=float('inf'), keybind=None):
        self.name = name
        self.unicode_key = unicode_key
        self.count = count
        self.pinned = pinned
        self.manual_order = manual_order
        self.keybind = keybind

class Settings:
    def __init__(self):
        self.theme = "Default"
        self.keybinds_enabled = True
        self.direct_paste = True
        self.return_focus = True
        self.always_on_top = False
        
    @classmethod
    def load(cls):
        settings = cls()
        try:
            settings_path = get_settings_path()
            if os.path.exists(settings_path):
                with open(settings_path, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    settings.theme = data.get('theme', 'Default')
                    settings.keybinds_enabled = data.get('keybinds_enabled', True)
                    settings.direct_paste = data.get('direct_paste', True)
                    settings.return_focus = data.get('return_focus', True)
                    settings.always_on_top = data.get('always_on_top', False)
        except Exception:
            pass
        return settings
    
    def save(self):
        try:
            settings_path = get_settings_path()
            data = {
                'theme': self.theme,
                'keybinds_enabled': self.keybinds_enabled,
                'direct_paste': self.direct_paste,
                'return_focus': self.return_focus,
                'always_on_top': self.always_on_top
            }
            with open(settings_path, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
        except Exception:
            pass

def get_supported_characters():
    """Get all characters supported by Noto Emoji font."""
    if not os.path.exists(NOTO_EMOJI_PATH):
        print(f"Warning: Noto Emoji font not found at {NOTO_EMOJI_PATH}")
        return set()
    
    try:
        font = TTFont(NOTO_EMOJI_PATH)
        chars = set()
        
        for table in font['cmap'].tables:
            for char_code in table.cmap.keys():
                try:
                    char = chr(char_code)
                    chars.add(char)
                except ValueError:
                    continue
        
        return chars
    except Exception as e:
        print(f"Error reading font file: {e}")
        return set()

def fetch_emojis():
    try:
        response = requests.get(EMOJI_API, timeout=10)
        if response.status_code == 200:
            return response.json()
    except:
        pass
    return []

def load_data():
    """Load emoji data from AppData file"""
    try:
        data_path = get_data_path()
        if not os.path.exists(data_path):
            return {}
        
        with open(data_path, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception:
        return {}

def save_data(emoji_dict):
    """Save emoji data to AppData file"""
    try:
        data_path = get_data_path()
        with open(data_path, 'w', encoding='utf-8') as f:
            json.dump(emoji_dict, f, ensure_ascii=False, indent=2)
    except Exception:
        pass

def find_crystal_realms_window():
    """Find Crystal Realms window by its title and process name"""
    def enum_window_callback(hwnd, result):
        if win32gui.IsWindowVisible(hwnd):
            try:
                window_title = win32gui.GetWindowText(hwnd)
                if "crystal realms" in window_title.lower():
                    try:
                        _, process_id = win32process.GetWindowThreadProcessId(hwnd)
                        process = psutil.Process(process_id)
                        if "crystal_realms" in process.name().lower():
                            result.append(hwnd)
                    except:
                        pass
            except:
                pass
        return True
    
    result = []
    win32gui.EnumWindows(enum_window_callback, result)
    
    # Also try finding by executable name directly
    if not result:
        for proc in psutil.process_iter(['pid', 'name']):
            if "crystal_realms" in proc.info['name'].lower():
                try:
                    def callback(hwnd, pid):
                        if win32gui.IsWindowVisible(hwnd) and win32gui.GetWindowText(hwnd):
                            _, window_pid = win32process.GetWindowThreadProcessId(hwnd)
                            if window_pid == pid:
                                result.append(hwnd)
                        return True
                    
                    win32gui.EnumWindows(callback, proc.info['pid'])
                except:
                    continue
    
    return result[0] if result else None

def send_emoji_to_game(emoji, game_hwnd=None):
    """Send emoji directly to Crystal Realms game window"""
    if not game_hwnd:
        game_hwnd = find_crystal_realms_window()
    
    if not game_hwnd:
        return False
    
    # First copy to clipboard to ensure the emoji is available
    pyperclip.copy(emoji)
    
    # Store current foreground window
    current_hwnd = win32gui.GetForegroundWindow()
    
    try:
        # Bring game window to foreground
        win32gui.SetForegroundWindow(game_hwnd)
        # Reduce delay for faster response
        win32api.Sleep(20)  # Further reduced delay
        
        # Send Ctrl+V keystroke with optimized timing
        # Press Ctrl down
        ctypes.windll.user32.keybd_event(0x11, 0, 0, 0)  # CTRL key
        win32api.Sleep(10)  # Further reduced delay
        
        # Press V down
        ctypes.windll.user32.keybd_event(0x56, 0, 0, 0)  # V key
        win32api.Sleep(10)  # Further reduced delay
        
        # Release V first
        ctypes.windll.user32.keybd_event(0x56, 0, win32con.KEYEVENTF_KEYUP, 0)
        win32api.Sleep(10)  # Further reduced delay
        
        # Release Ctrl second
        ctypes.windll.user32.keybd_event(0x11, 0, win32con.KEYEVENTF_KEYUP, 0)
        
        # Return focus to emoji picker if option is enabled
        if hasattr(current_hwnd, 'return_focus') and current_hwnd.return_focus:
            win32api.Sleep(20)  # Further reduced delay
            win32gui.SetForegroundWindow(current_hwnd)
        return True
    except Exception:
        # Try to restore focus
        try:
            win32gui.SetForegroundWindow(current_hwnd)
        except:
            pass
        return False

class KeyBindDialog:
    def __init__(self, parent, emoji_data):
        self.parent = parent
        self.emoji_data = emoji_data
        self.result = None
        self.create_dialog()
        
    def create_dialog(self):
        self.dialog = Toplevel(self.parent)
        self.dialog.title(f"Set Keybind for {self.emoji_data.unicode_key}")
        self.dialog.geometry("400x180")
        self.dialog.transient(self.parent)
        self.dialog.grab_set()
        
        # Center dialog
        self.dialog.update_idletasks()
        width = self.dialog.winfo_width()
        height = self.dialog.winfo_height()
        x = (self.dialog.winfo_screenwidth() // 2) - (width // 2)
        y = (self.dialog.winfo_screenheight() // 2) - (height // 2)
        self.dialog.geometry(f"{width}x{height}+{x}+{y}")
        
        # Apply theme
        if hasattr(self.parent, 'current_theme'):
            theme = self.parent.current_theme
            if theme != "Default":
                self.dialog.configure(bg=THEMES[theme]["background"])
        
        # Instruction label
        msg = f"Press the key you want to bind to the emoji: {self.emoji_data.unicode_key}"
        Label(
            self.dialog, 
            text=msg,
            font=("Arial", 12),
            wraplength=380,
            pady=10
        ).pack(pady=10)
        
        # Display current keybind if any
        current_text = "Current: None" if not self.emoji_data.keybind else f"Current: {self.emoji_data.keybind}"
        self.current_keybind_label = Label(
            self.dialog,
            text=current_text,
            font=("Arial", 10)
        )
        self.current_keybind_label.pack(pady=5)
        
        # Key input field (read-only)
        self.key_entry = Entry(
            self.dialog,
            font=("Arial", 14),
            justify="center",
            state="readonly"
        )
        self.key_entry.pack(pady=10, padx=20, fill="x")
        
        # Buttons frame
        btn_frame = Frame(self.dialog)
        btn_frame.pack(pady=10, fill="x")
        
        # Apply button (disabled until a key is pressed)
        self.apply_btn = Button(
            btn_frame,
            text="Apply",
            command=self.apply_keybind,
            state="disabled",
            width=8
        )
        self.apply_btn.pack(side="left", padx=10)
        
        # Clear button
        Button(
            btn_frame,
            text="Clear",
            command=self.clear_keybind,
            width=8
        ).pack(side="left", padx=10)
        
        # Cancel button
        Button(
            btn_frame,
            text="Cancel",
            command=self.cancel,
            width=8
        ).pack(side="right", padx=10)
        
        # Bind key press event
        self.dialog.bind("<KeyPress>", self.on_key_press)
        
        # Wait for dialog to close
        self.parent.wait_window(self.dialog)
    
    def on_key_press(self, event):
        # Get the key name
        key_name = event.keysym
        
        # Skip modifier keys
        if key_name in ('Shift_L', 'Shift_R', 'Control_L', 'Control_R', 'Alt_L', 'Alt_R'):
            return
        
        # Skip some keys that shouldn't be bound
        if key_name in ('Escape', 'Tab', 'Return', 'BackSpace'):
            return
            
        # Display the key
        self.key_entry.config(state="normal")
        self.key_entry.delete(0, "end")
        self.key_entry.insert(0, key_name)
        self.key_entry.config(state="readonly")
        
        # Enable the apply button
        self.apply_btn.config(state="normal")
        
        # Store the key
        self.result = key_name
    
    def apply_keybind(self):
        self.dialog.destroy()
    
    def clear_keybind(self):
        self.result = ""  # Empty string means clear the keybind
        self.dialog.destroy()
    
    def cancel(self):
        self.result = None  # None means no change
        self.dialog.destroy()

class EmojiPicker:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Crystal Realms EmojiPicker - Made By SavageTheUnicorn")
        self.root.geometry("600x450")
        self.root.minsize(300, 350)
        
        # Add these for the queue-based hotkey handling
        self.global_lock = threading.RLock()  # Reentrant lock for thread safety
        self.hotkey_queue = queue.Queue()
        self.hotkey_processor_running = False
        
        # Hotkey Handling (keep these for backward compatibility)
        self.last_hotkey_time = 0
        self.last_hotkey_id = None
        self.processing_hotkey = False
        
        # Load settings
        self.settings = Settings.load()
        
        # Initialize threading lock for UI updates
        self.ui_lock = Lock()
        self.save_lock = threading.Lock()
        
        # Set current theme from settings
        self.current_theme = self.settings.theme
        
        self.selected_index = None  # Track currently selected emoji index
        self.last_click_time = 0    # Time of last click
        self.click_throttle = 0.05  # Minimum seconds between spam clicks (50ms)
        
        # Set window icon
        icon_path = resource_path("app.ico")
        if os.path.exists(icon_path):
            self.root.iconbitmap(icon_path)
            
        self.current_page = 0
        self.search_var = tk.StringVar()
        self.emoji_data = {}
        self.filtered_emojis = []
        self.supported_chars = get_supported_characters()
        
        # Set values from settings
        self.always_on_top = self.settings.always_on_top
        self.direct_paste = self.settings.direct_paste
        self.return_focus = self.settings.return_focus
        self.keybinds_enabled = self.settings.keybinds_enabled
        
        self.game_hwnd = None    # Will store game window handle
        self.tray_icon = None    # Will store the tray icon instance
        
        # Hot key registration
        self.hotkey_ids = {}     # Store registered hotkey IDs
        self.hotkey_atoms = {}   # Store atom references to prevent garbage collection
        
        # New variables for improved clicking behavior
        self.selected_index = None  # Track currently selected emoji index
        self.last_click_time = 0    # Time of last click
        self.click_throttle = 0.05  # Minimum seconds between spam clicks (50ms)
        
        # Initialize key bindings maps
        self.keybind_to_emoji = {}  # Maps key -> emoji unicode
        
        # Initialize system tray icon
        self.setup_tray()
        
        # Bind window close event to hide window instead of closing app
        self.root.protocol('WM_DELETE_WINDOW', self.on_closing)
        
        if not self.supported_chars:
            self.show_font_warning()
        
        self.setup_gui()
        self.apply_theme(self.current_theme)
        self.load_emojis()
        
        # Start the hotkey processor before setting up hotkeys
        self.start_hotkey_processor()
        
        # Set up message handler for hotkeys
        self.setup_win32_hotkeys()
        
        # Try to find the game window on startup
        self.find_game_window()

    # FIXED METHODS START HERE - Added/Modified to fix GIL issues
    def start_hotkey_processor(self):
        """Start a background thread to process hotkey events from the queue"""
        if self.hotkey_processor_running:
            return
            
        self.hotkey_processor_running = True
        
        def process_hotkeys():
            while self.hotkey_processor_running:
                try:
                    # Get an event from the queue, with a short timeout to allow checking the running flag
                    try:
                        hotkey_id = self.hotkey_queue.get(timeout=0.1)
                    except queue.Empty:
                        continue
                    
                    # Process the hotkey
                    try:
                        # Much lighter throttling - only 75ms between keypresses
                        # This is quick enough for chat but still prevents total spam
                        current_time = time.time()
                        if current_time - self.last_hotkey_time < 0.075:  # 75ms throttle - reduced for MMO chat
                            continue
                            
                        self.last_hotkey_time = current_time
                        
                        # Schedule emoji handling on main thread with after_idle
                        if hotkey_id in self.hotkey_ids:
                            emoji_unicode = self.hotkey_ids[hotkey_id]
                            self.root.after_idle(lambda u=emoji_unicode: self.process_hotkey_on_main_thread(u))
                    finally:
                        # Always mark the task as done to keep the queue clean
                        self.hotkey_queue.task_done()
                        
                except Exception as e:
                    print(f"Error in hotkey processor: {e}")
                    traceback.print_exc()
                    # Sleep a bit to avoid tight loops in case of errors
                    time.sleep(0.1)  # Reduced sleep on error
        
        # Start the processor thread
        t = threading.Thread(target=process_hotkeys, daemon=True)
        t.name = "HotkeyProcessor"  # Name the thread for easier debugging
        t.start()

    def process_hotkey_on_main_thread(self, emoji_unicode):
        """Process a hotkey event safely on the main thread"""
        try:
            # Only process if we're not already handling one
            if self.processing_hotkey:
                return
                
            with self.global_lock:
                self.processing_hotkey = True
                
            # Get the emoji data 
            emoji_data = self.emoji_data.get(emoji_unicode)
            if not emoji_data:
                return
                
            # Copy to clipboard
            pyperclip.copy(emoji_data.unicode_key)
            
            # Get active window
            try:
                active_window = win32gui.GetForegroundWindow()
                window_name = win32gui.GetWindowText(active_window) if active_window else "Unknown"
            except:
                active_window = None
                window_name = "Unknown"
            
            # Faster keystrokes with shorter delays
            try:
                # Press Ctrl down
                ctypes.windll.user32.keybd_event(0x11, 0, 0, 0)
                time.sleep(0.015)  # 15ms instead of 50ms
                
                # Press V down
                ctypes.windll.user32.keybd_event(0x56, 0, 0, 0)
                time.sleep(0.015)  # 15ms
                
                # Release V
                ctypes.windll.user32.keybd_event(0x56, 0, win32con.KEYEVENTF_KEYUP, 0)
                time.sleep(0.015)  # 15ms
                
                # Release Ctrl
                ctypes.windll.user32.keybd_event(0x11, 0, win32con.KEYEVENTF_KEYUP, 0)
            except Exception as e:
                print(f"Error sending keystrokes: {e}")
            
            # Update statistics
            emoji_data.count += 1
            
            # Update status if window is visible
            if self.root.state() == 'normal':
                self.update_status(f"Sent {emoji_data.unicode_key} to {window_name}")
            
            # Schedule saving for later to not block the main thread
            self.root.after(500, self.thread_safe_save_emoji_data)
        except Exception as e:
            print(f"Error processing hotkey on main thread: {e}")
            traceback.print_exc()
        finally:
            # Always release the processing flag
            with self.global_lock:
                self.processing_hotkey = False
                
    def setup_win32_hotkeys(self):
        """Set up a reliable Windows hotkey system using a custom window procedure"""
        # Create a unique window class name to prevent conflicts
        unique_class_name = f"EmojiPickerHotkeyWindow_{uuid.uuid4().hex}"

        # Unregister any existing window class and hotkeys first
        try:
            self.unregister_all_hotkeys()
        except Exception as e:
            print(f"Error unregistering hotkeys: {e}")

        # Window message handler - Optimized for faster response
        def _wnd_proc(hwnd, msg, wparam, lparam):
            if msg == WM_HOTKEY:
                try:
                    # Increased queue limit to allow faster processing
                    if self.hotkey_queue.qsize() < 10:  # Allow more events to be queued (from 5)
                        self.hotkey_queue.put_nowait(wparam)
                except Exception:
                    pass  # Silently ignore queue errors
                # Return 1 to indicate we've handled this message
                return 1
            # Default processing for other messages
            return ctypes.windll.user32.DefWindowProcW(hwnd, msg, wparam, lparam)

        # Create a window class
        wc = win32gui.WNDCLASS()
        wc.lpfnWndProc = _wnd_proc
        wc.lpszClassName = unique_class_name

        try:
            # Register the window class
            win32gui.RegisterClass(wc)

            # Create a message-only window (no UI, just for message handling)
            self.hotkey_hwnd = win32gui.CreateWindowEx(
                0,
                wc.lpszClassName,
                "Emoji Picker Hotkey Window",
                0, 0, 0, 0, 0,
                win32con.HWND_MESSAGE,  # Message-only window
                0, wc.hInstance, None
            )

            # Register hotkeys if enabled
            if self.keybinds_enabled:
                self.register_system_wide_hotkeys()

        except Exception as e:
            print(f"Error setting up hotkey window: {e}")
            traceback.print_exc()

    def register_system_wide_hotkeys(self):
        """Register hotkeys for all emojis with keybindings"""
        if not self.keybinds_enabled or not hasattr(self, 'hotkey_hwnd'):
            return

        # Unregister any existing hotkeys first
        self.unregister_all_hotkeys()

        # Track hotkey IDs to avoid conflicts
        hotkey_id = 1
        bound_vk_codes = set()  # Track which keys are already bound

        # Register each emoji hotkey with error handling for each key
        for key, emoji_unicode in self.keybind_to_emoji.items():
            vk_code = self.get_vk_code(key)
            if not vk_code or vk_code in bound_vk_codes:
                continue
                
            try:
                # Register with no modifier for maximum speed
                result = ctypes.windll.user32.RegisterHotKey(
                    self.hotkey_hwnd,  # Window handle
                    hotkey_id,         # Unique ID for this hotkey
                    0,                 # No modifier keys
                    vk_code            # Virtual key code
                )
                
                if result:
                    self.hotkey_ids[hotkey_id] = emoji_unicode
                    bound_vk_codes.add(vk_code)
                    print(f"Registered hotkey for {key} (VK: {vk_code})")
                    hotkey_id += 1
                else:
                    error_code = ctypes.windll.kernel32.GetLastError()
                    print(f"Failed to register hotkey for {key} (Error: {error_code})")
            except Exception as e:
                print(f"Error registering hotkey {key}: {e}")

    def thread_safe_save_emoji_data(self):
        """Thread-safe version of save_emoji_data to prevent concurrent file access"""
        # Use a non-blocking acquire to prevent deadlocks
        if hasattr(self, 'save_lock') and self.save_lock.acquire(blocking=False):
            try:
                data = {}
                # Create a copy of the data to avoid long lock times
                with self.global_lock:
                    data = {
                        unicode_key: {
                            'count': emoji.count,
                            'pinned': emoji.pinned,
                            'manual_order': emoji.manual_order,
                            'keybind': emoji.keybind
                        }
                        for unicode_key, emoji in self.emoji_data.items()
                    }
                    
                # Save data outside the global lock
                try:
                    data_path = get_data_path()
                    temp_path = data_path + ".tmp"
                    
                    # Write to temp file first to avoid corruption
                    with open(temp_path, 'w', encoding='utf-8') as f:
                        json.dump(data, f, ensure_ascii=False, indent=2)
                        f.flush()
                        os.fsync(f.fileno())  # Ensure data is written to disk
                    
                    # Now rename the temp file to the actual file (atomic operation)
                    if os.path.exists(data_path):
                        os.replace(temp_path, data_path)
                    else:
                        os.rename(temp_path, data_path)
                        
                except Exception as e:
                    print(f"Error saving emoji data: {e}")
                    traceback.print_exc()
            finally:
                self.save_lock.release()
                
    def update_emoji_stats(self, emoji_data, window_hwnd=None):
        """Update emoji statistics and UI from the main thread"""
        try:
            # Increment usage count
            emoji_data.count += 1
            
            # Update UI status
            if window_hwnd:
                window_name = win32gui.GetWindowText(window_hwnd)
                self.update_status(f"Sent {emoji_data.unicode_key} to {window_name}")
            else:
                self.update_status(f"Sent {emoji_data.unicode_key}")
                
            # Save emoji data after a delay to batch updates
            self.root.after(500, self.safe_save_emoji_data)
        except Exception as e:
            print(f"Error updating emoji stats: {e}")
            traceback.print_exc()
    
    def safe_save_emoji_data(self):
        """Thread-safe emoji data saving method"""
        with self.global_lock:
            data = {
                unicode_key: {
                    'count': emoji.count,
                    'pinned': emoji.pinned,
                    'manual_order': emoji.manual_order,
                    'keybind': emoji.keybind
                }
                for unicode_key, emoji in self.emoji_data.items()
            }
            save_data(data)

    def get_vk_code(self, key):
        """Convert key name to virtual key code for Windows hotkey registration"""
        # Handle special cases first
        special_keys = {
            'space': 0x20,
            'BackSpace': 0x08,
            'Tab': 0x09,
            'Return': 0x0D,
            'Escape': 0x1B,
            'Delete': 0x2E,
            'Prior': 0x21,  # Page Up
            'Next': 0x22,   # Page Down
            'End': 0x23,
            'Home': 0x24,
            'Left': 0x25,
            'Up': 0x26,
            'Right': 0x27,
            'Down': 0x28,
            'F1': 0x70,
            'F2': 0x71,
            'F3': 0x72,
            'F4': 0x73,
            'F5': 0x74,
            'F6': 0x75,
            'F7': 0x76,
            'F8': 0x77,
            'F9': 0x78,
            'F10': 0x79,
            'F11': 0x7A,
            'F12': 0x7B,
        }
        
        # If it's a special key, return its code
        if key in special_keys:
            return special_keys[key]
        
        # Convert to lowercase for standard keys
        key = key.lower()
        
        # Use VK_CODES dictionary from earlier in the script
        return VK_CODES.get(key)

    def toggle_keybinds_enabled(self):
        """Toggle whether keybinds are enabled"""
        self.keybinds_enabled = not self.keybinds_enabled
        self.keybinds_var.set(self.keybinds_enabled)
        status = "enabled" if self.keybinds_enabled else "disabled"
        self.update_status(f"Keybinds {status}")
        
        # Update settings
        self.settings.keybinds_enabled = self.keybinds_enabled
        self.settings.save()
        
        # Update hotkeys registration if already set up
        try:
            if hasattr(self, 'hwnd'):
                if self.keybinds_enabled:
                    self.register_system_wide_hotkeys()
                else:
                    self.unregister_all_hotkeys()
        except Exception as e:
            print(f"Error toggling hotkeys: {e}")
            traceback.print_exc()

    def on_closing(self):
        """Handle window closing properly"""
        # Just hide the window instead of closing the app
        self.root.withdraw()

    def find_game_window(self):
        """Find the game window and store its handle"""
        self.game_hwnd = find_crystal_realms_window()
        game_status = "Game window found" if self.game_hwnd else "Game window not found"
        self.update_status(game_status)
        return self.game_hwnd is not None

    def update_status(self, message):
        """Update status label with message"""
        if hasattr(self, 'status_label'):
            self.status_label.config(text=message)

    def create_tray_icon(self):
        """Create a simple emoji icon for the tray (fallback)"""
        icon_size = 32
        image = Image.new('RGB', (icon_size, icon_size), color='#7289DA')
        draw = ImageDraw.Draw(image)
        draw.text((icon_size//4, icon_size//4), "🦄", fill='white', size=32)
        return image

    def setup_tray(self):
        """Setup the system tray icon and menu"""
        try:
            icon_path = resource_path("app.ico")
            if os.path.exists(icon_path):
                icon_image = Image.open(icon_path)
            else:
                # Fallback to creating a simple emoji icon
                icon_image = self.create_tray_icon()
                
            self.tray_icon = pystray.Icon(
                "emoji_picker",
                icon_image,
                "Crystal Realms EmojiPicker",
                menu=pystray.Menu(
                    pystray.MenuItem("Show/Hide", self.toggle_window),
                    pystray.MenuItem("Find Game Window", self.find_game_window),
                    pystray.MenuItem("Exit", self.quit_app)
                )
            )
            # Start the tray icon in a separate thread
            Thread(target=self.tray_icon.run, daemon=True).start()
        except Exception:
            pass

    def toggle_window(self, _=None):
        """Toggle the main window visibility"""
        try:
            if self.root.state() == 'withdrawn':
                self.root.deiconify()
                self.root.lift()
                if self.always_on_top:
                    self.root.attributes('-topmost', True)
            else:
                self.root.withdraw()
        except Exception:
            pass

    def quit_app(self, _=None):
        """Properly close the application"""
        try:
            # Stop the hotkey processor
            self.hotkey_processor_running = False
            
            # Save settings first
            self.settings.save()
            
            # Stop the tray icon thread
            if self.tray_icon:
                self.tray_icon.stop()
            
            # Destroy the root window
            self.root.destroy()
        except Exception:
            # Force exit as last resort
            try:
                self.root.destroy()
            except:
                pass

    def send_enter_key_to_game(self):
        """Send Enter key to Crystal Realms game window"""
        if not self.game_hwnd:
            self.game_hwnd = find_crystal_realms_window()
        
        if not self.game_hwnd:
            self.update_status("Game window not found. Cannot send Enter key.")
            return False
        
        # Store current foreground window
        current_hwnd = win32gui.GetForegroundWindow()
        
        try:
            # Bring game window to foreground
            win32gui.SetForegroundWindow(self.game_hwnd)
            win32api.Sleep(20)  # Small delay
            
            # Send Enter keystroke
            ctypes.windll.user32.keybd_event(0x0D, 0, 0, 0)  # VK_RETURN (Enter key) press
            win32api.Sleep(10)  # Small delay
            ctypes.windll.user32.keybd_event(0x0D, 0, win32con.KEYEVENTF_KEYUP, 0)  # Enter key release
            
            # Return focus to emoji picker if option is enabled
            if self.return_focus:
                win32api.Sleep(20)  # Small delay
                win32gui.SetForegroundWindow(current_hwnd)
            
            self.update_status("Enter key sent to game")
            return True
        except Exception:
            # Try to restore focus
            try:
                win32gui.SetForegroundWindow(current_hwnd)
            except:
                pass
            self.update_status("Failed to send Enter key to game")
            return False

    def show_font_warning(self):
        """Show warning if font file is not found."""
        warning_window = tk.Toplevel(self.root)
        warning_window.title("Font Not Found")
        warning_window.geometry("400x150")
        
        Label(
            warning_window,
            text="Noto Emoji font file not found!\n\n"
                 "Please download the font from:\n"
                 "https://fonts.google.com/noto/specimen/Noto+Emoji\n\n"
                 "Place the TTF file in the same directory as this script\n"
                 "and name it 'NotoEmoji-Regular.ttf'",
            justify=tk.LEFT,
            padx=20,
            pady=20
        ).pack()

    def is_emoji_supported(self, emoji_chars):
        """Check if all characters in an emoji are supported by the font."""
        if not self.supported_chars:
            return True  # If we couldn't load the font, show all emojis
            
        for char in emoji_chars:
            if char not in self.supported_chars:
                return False
        return True

    def load_emojis(self):
        raw_emojis = fetch_emojis()
        saved_data = load_data()
        
        # Clear existing keybind maps
        self.keybind_to_emoji = {}
        
        for emoji in raw_emojis:
            if 'unicode' not in emoji or not emoji['unicode']:
                continue
            
            try:
                # Handle single unicode value
                if isinstance(emoji['unicode'], str):
                    unicode_val = emoji['unicode']
                    if unicode_val.startswith('U+'):
                        hex_val = unicode_val[2:]
                    else:
                        hex_val = unicode_val
                    unicode_char = chr(int(hex_val, 16))
                # Handle multiple unicode values (like flags)
                else:
                    unicode_chars = []
                    for unicode_val in emoji['unicode']:
                        if unicode_val.startswith('U+'):
                            hex_val = unicode_val[2:]
                        else:
                            hex_val = unicode_val
                        unicode_chars.append(chr(int(hex_val, 16)))
                    unicode_char = ''.join(unicode_chars)
                
                # Only add emoji if it's supported by the font
                if self.is_emoji_supported(unicode_char):
                    name = emoji['name']
                    saved = saved_data.get(unicode_char, {})
                    count = saved.get('count', 0)
                    pinned = saved.get('pinned', False)
                    manual_order = saved.get('manual_order', float('inf'))
                    keybind = saved.get('keybind', None)
                    
                    # Create emoji data
                    self.emoji_data[unicode_char] = EmojiData(
                        name, unicode_char, count, pinned, manual_order, keybind
                    )
                    
                    # Register keybind if it exists
                    if keybind:
                        self.keybind_to_emoji[keybind] = unicode_char
            except:
                continue
        
        # Skip hotkey registration during initialization
        # We'll register them after the UI is fully set up
        
        self.filter_and_sort_emojis()

    def apply_theme(self, theme_name):
        """Apply the selected theme to all UI elements"""
        try:
            if theme_name == "Default":
                # Reset to system defaults
                self.root.configure(bg=self.root.cget("bg"))
                
                # Reset specific widgets to system defaults
                default_bg = self.root.cget("bg")
                default_fg = "black"
                
                widgets_to_reset = [
                    self.container, self.controls_label, self.search_frame, 
                    self.search_entry, self.search_icon, self.list_frame, 
                    self.listbox, self.status_label, self.options_frame, 
                    self.nav_frame, self.center_frame, self.center_content
                ]
                
                for widget in widgets_to_reset:
                    try:
                        widget.configure(bg=default_bg, fg=default_fg)
                    except:
                        pass
                
                # Reset specific widget colors
                try:
                    self.listbox.configure(bg="white", fg="black")
                except:
                    pass
            else:
                # Apply the selected theme
                theme = THEMES[theme_name]
                
                # Update root window background
                self.root.configure(bg=theme["background"])
                
                # Configure main container and frames
                widgets_to_theme = [
                    (self.container, "background"),
                    (self.controls_label, ["background", "foreground"]),
                    (self.search_frame, "background"),
                    (self.search_entry, ["bg", "fg"]),
                    (self.search_icon, ["background", "foreground"]),
                    (self.list_frame, "background"),
                    (self.listbox, ["bg", "fg"]),
                    (self.status_label, ["background", "foreground"]),
                    (self.options_frame, "background"),
                    (self.nav_frame, "background"),
                    (self.center_frame, "background"),
                    (self.center_content, "background"),
                    # Fix the white background
                    (self.left_buttons, "background"),
                    (self.right_buttons, "background")
                ]
                
                for widget, attrs in widgets_to_theme:
                    try:
                        if isinstance(attrs, list):
                            for attr in attrs:
                                if attr == "background" or attr == "bg":
                                    widget.configure(bg=theme["background"])
                                elif attr == "foreground" or attr == "fg":
                                    widget.configure(fg=theme["text"])
                        else:
                            widget.configure(bg=theme["background"])
                    except Exception as e:
                        print(f"Error theming {widget}: {e}")
                
                # Theme buttons and checkboxes
                button_widgets = [
                    self.page_label, 
                    self.always_on_top_btn
                ]
                
                for widget in button_widgets:
                    try:
                        widget.configure(
                            bg=theme.get("button_bg", theme["background"]),
                            fg=theme.get("button_fg", theme["text"])
                        )
                    except Exception as e:
                        print(f"Error theming button {widget}: {e}")
                
                # Theme checkboxes in options frame
                for widget in self.options_frame.winfo_children():
                    if isinstance(widget, Checkbutton):
                        try:
                            # Specific color handling for problematic themes
                            if theme_name in ["Dark Mode", "Cyberpunk Neon", "Forest Green", "Ocean Blue"]:
                                # Use a light color for checkmark visibility
                                widget.configure(
                                    bg=theme["background"],
                                    fg=theme["text"],
                                    selectcolor="#EEE8D5"  # Solarized light color
                                )
                            else:
                                widget.configure(
                                    bg=theme["background"],
                                    fg=theme["text"]
                                )
                        except Exception as e:
                            print(f"Error theming checkbox: {e}")
                
                # Theme navigation buttons
                nav_buttons = []
                
                # Collect buttons from different frames
                for frame in [self.left_buttons, self.right_buttons, self.center_content]:
                    for widget in frame.winfo_children():
                        if isinstance(widget, Button):
                            nav_buttons.append(widget)
                
                # Additional specific buttons
                try:
                    # Send Enter button
                    send_enter_btn = [btn for btn in self.options_frame.winfo_children() if isinstance(btn, Button)][0]
                    nav_buttons.append(send_enter_btn)
                except:
                    pass
                
                # Theme all collected buttons
                for btn in nav_buttons:
                    try:
                        btn.configure(
                            bg=theme.get("button_bg", theme["background"]),
                            fg=theme.get("button_fg", theme["text"]),
                            activebackground=theme.get("primary", theme["background"]),
                            activeforeground=theme.get("button_fg", theme["text"])
                        )
                    except Exception as e:
                        print(f"Error theming button {btn}: {e}")
            
            # Save the current theme
            self.current_theme = theme_name
            self.settings.theme = theme_name
            self.settings.save()
        
        except Exception as e:
            print(f"Error applying theme {theme_name}: {e}")
            import traceback
            traceback.print_exc()

    def open_theme_selector(self):
        """Open a dialog to select a theme"""
        theme_dialog = Toplevel(self.root)
        theme_dialog.title("Select Theme")
        theme_dialog.geometry("400x450")  # Made taller to fit all content
        theme_dialog.resizable(False, False)  # Fixed size
        theme_dialog.transient(self.root)
        theme_dialog.grab_set()
        
        # Set icon
        icon_path = resource_path("app.ico")
        if os.path.exists(icon_path):
            theme_dialog.iconbitmap(icon_path)
            
        # Center dialog
        theme_dialog.update_idletasks()
        width = theme_dialog.winfo_width()
        height = theme_dialog.winfo_height()
        x = (theme_dialog.winfo_screenwidth() // 2) - (width // 2)
        y = (theme_dialog.winfo_screenheight() // 2) - (height // 2)
        theme_dialog.geometry(f"{width}x{height}+{x}+{y}")
        
        # Apply current theme to dialog
        if self.current_theme != "Default":
            theme = THEMES[self.current_theme]
            theme_dialog.configure(bg=theme["background"])
        
        # Main content frame with fixed margins
        main_frame = Frame(theme_dialog, padx=20, pady=20)
        main_frame.pack(fill="both", expand=True)
        
        if self.current_theme != "Default":
            main_frame.configure(bg=theme["background"])
        
        # Heading
        heading = Label(
            main_frame,
            text="Select a Theme",
            font=("Arial", 16, "bold"),
            pady=10
        )
        heading.pack(fill="x")
        
        if self.current_theme != "Default":
            heading.configure(bg=theme["background"], fg=theme["text"])
        
        # Theme selection variable
        theme_var = StringVar(value=self.current_theme)
        
        # Create a scrollable frame for the theme options
        scroll_frame = Frame(main_frame)
        scroll_frame.pack(fill="both", expand=True, pady=10)
        
        if self.current_theme != "Default":
            scroll_frame.configure(bg=theme["background"])
        
        # Create radio buttons for each theme
        for i, theme_name in enumerate(THEMES.keys()):
            rb = tk.Radiobutton(
                scroll_frame,
                text=theme_name,
                variable=theme_var,
                value=theme_name,
                font=("Arial", 12),
                command=lambda tn=theme_name: self.preview_theme(tn, preview_label)
            )
            rb.pack(anchor="w", pady=5)
            
            # Apply theme to radio button
            if self.current_theme != "Default":
                rb.configure(
                    bg=theme["background"],
                    fg=theme["text"],
                    selectcolor=theme["background"],
                    activebackground=theme["primary"],
                    activeforeground=theme["text"]
                )
        
        # Create a preview label
        preview_frame = Frame(main_frame, padx=10, pady=10)
        preview_frame.pack(fill="x", pady=10)
        
        if self.current_theme != "Default":
            preview_frame.configure(bg=theme["background"])
        
        preview_label = Label(
            preview_frame,
            text="Theme Preview",
            width=40,
            height=2,
            relief="ridge"
        )
        preview_label.pack(fill="x")
        
        # Initially show preview of current theme
        self.preview_theme(self.current_theme, preview_label)
        
        # Buttons frame
        btn_frame = Frame(main_frame)
        btn_frame.pack(fill="x", pady=20)
        
        if self.current_theme != "Default":
            btn_frame.configure(bg=theme["background"])
        
        # Apply button - much larger and more visible
        apply_btn = Button(
            btn_frame,
            text="Apply Theme",
            command=lambda: [self.apply_theme(theme_var.get()), theme_dialog.destroy()],
            width=15,
            height=2,
            font=("Arial", 11, "bold")
        )
        apply_btn.pack(side="left", padx=5)
        
        # Cancel button - larger as well
        cancel_btn = Button(
            btn_frame,
            text="Cancel",
            command=theme_dialog.destroy,
            width=15,
            height=2,
            font=("Arial", 11)
        )
        cancel_btn.pack(side="right", padx=5)
        
        # If using a custom theme, style the buttons
        if self.current_theme != "Default":
            apply_btn.configure(
                bg=theme["button_bg"],
                fg=theme["button_fg"],
                activebackground=theme["primary"],
                activeforeground=theme["button_fg"]
            )
            cancel_btn.configure(
                bg=theme["button_bg"],
                fg=theme["button_fg"],
                activebackground=theme["primary"],
                activeforeground=theme["button_fg"]
            )
        
        # Wait for dialog to close
        self.root.wait_window(theme_dialog)

    def preview_theme(self, theme_name, preview_label):
            """Show a preview of the selected theme"""
            if theme_name == "Default":
                preview_label.configure(
                    bg=self.root.cget("bg"),
                    fg="black",
                    text="Default Theme (System Colors)"
                )
                # Immediately apply the theme
                self.apply_theme("Default")
            else:
                theme = THEMES[theme_name]
                preview_label.configure(
                    bg=theme["background"],
                    fg=theme["text"],
                    text=f"{theme_name} Theme"
                )
                # Immediately apply the theme
                self.apply_theme(theme_name)

    def filter_and_sort_emojis(self):
        """Filter and sort emojis based on search query and other criteria"""
        query = self.search_var.get().lower()
        emojis = list(self.emoji_data.values())
    
        # Store the unicode key of the selected emoji if there is one
        selected_unicode = None
        if self.selected_index is not None and 0 <= self.selected_index < len(self.filtered_emojis):
            selected_unicode = self.filtered_emojis[self.selected_index].unicode_key
    
        if query:
            emojis = [e for e in emojis if query in e.name.lower()]
            emojis.sort(key=lambda x: (not x.name.lower().startswith(query), x.name.lower()))
        else:
            # Sort by: pinned first, then pin order, then usage count, then name
            emojis.sort(key=lambda x: (
                not x.pinned,
                x.manual_order if x.pinned else float('inf'),
                -x.count,
                x.name.lower()
            ))
    
        self.filtered_emojis = emojis
    
        # Attempt to preserve selection
        if selected_unicode:
            try:
                # Find the new index of the previously selected emoji
                new_idx = next(i for i, e in enumerate(self.filtered_emojis) 
                             if e.unicode_key == selected_unicode)
                self.selected_index = new_idx
            except StopIteration:
                # If not found, reset selection
                self.selected_index = None
    
        # Only reset page if necessary
        if self.selected_index is not None:
            # Calculate which page the selected emoji is on
            selected_page = self.selected_index // PAGE_SIZE
            # Only reset current_page if selection is no longer visible
            if selected_page != self.current_page:
                self.current_page = selected_page
        else:
            # If no selection, reset to first page
            self.current_page = 0
        
        self.display_page()

    def handle_global_keypress(self, event):
        """Handle global key presses for keybinds"""
        if not self.keybinds_enabled:
            return
            
        # Skip non-visible window state
        if self.root.state() != 'normal':
            return
            
        # Get the key name
        key_name = event.keysym
        
        # Skip modifier keys
        if key_name in ('Shift_L', 'Shift_R', 'Control_L', 'Control_R', 'Alt_L', 'Alt_R'):
            return
            
        # Check if this key is bound to an emoji
        if key_name in self.keybind_to_emoji:
            unicode_char = self.keybind_to_emoji[key_name]
            emoji_data = self.emoji_data.get(unicode_char)
            
            if emoji_data:
                # Trigger the emoji sending
                self.handle_emoji_action(emoji_data=emoji_data, update_ui=True)
                self.update_status(f"Keybind: {key_name} → {emoji_data.unicode_key}")
                return "break"  # Prevent further processing
            
    def unregister_all_hotkeys(self):
        """Unregister all previously registered hotkeys"""
        if not hasattr(self, 'hotkey_hwnd'):
            return
    
        # Unregister hotkeys from the dedicated window
        for hotkey_id in list(self.hotkey_ids.keys()):
            try:
                ctypes.windll.user32.UnregisterHotKey(self.hotkey_hwnd, hotkey_id)
                print(f"Unregistered hotkey ID {hotkey_id}")
            except Exception as e:
                print(f"Error unregistering hotkey {hotkey_id}: {e}")
    
        # Clear the hotkey map
        self.hotkey_ids.clear()
        
    def handle_hotkey(self, hotkey_id):
        """Handle a hotkey event"""
        try:
            # Get the emoji data for this hotkey
            if hotkey_id not in self.hotkey_ids:
                return
                
            emoji_unicode = self.hotkey_ids[hotkey_id]
            emoji_data = self.emoji_data.get(emoji_unicode)
            
            if not emoji_data:
                return
                
            # Copy the emoji to clipboard
            pyperclip.copy(emoji_data.unicode_key)
            
            # Send paste command to active window
            active_window = win32gui.GetForegroundWindow()
            
            # Send Ctrl+V keystroke
            win32api.keybd_event(0x11, 0, 0, 0)  # CTRL down
            win32api.Sleep(10)  # Small delay
            win32api.keybd_event(0x56, 0, 0, 0)  # V down
            win32api.Sleep(10)  # Small delay
            win32api.keybd_event(0x56, 0, win32con.KEYEVENTF_KEYUP, 0)  # V up
            win32api.Sleep(10)  # Small delay
            win32api.keybd_event(0x11, 0, win32con.KEYEVENTF_KEYUP, 0)  # CTRL up
            
            # Update emoji statistics
            emoji_data.count += 1
            self.save_emoji_data()
            
            # Update status if window is visible
            if self.root.state() == 'normal':
                window_name = win32gui.GetWindowText(active_window)
                self.update_status(f"Sent {emoji_data.unicode_key} to {window_name}")
        except Exception as e:
            print(f"Error handling hotkey: {e}")
            traceback.print_exc()
        

    def handle_click(self, event=None):
        """
        Improved click handler:
        - First click selects emoji
        - Subsequent clicks on same emoji send it quickly
        - Clicking different emoji selects that one instead
        """
        # Get current time for throttling
        current_time = time.time()
    
        # Get clicked index from listbox
        clicked_idx = self.listbox.nearest(event.y)
        if clicked_idx < 0 or clicked_idx >= len(self.listbox.get(0, tk.END)):
            return
        
        # Calculate absolute index in filtered emojis
        abs_idx = self.current_page * PAGE_SIZE + clicked_idx
    
        # Ensure index is valid
        if abs_idx >= len(self.filtered_emojis):
            return
        
        # If this is a different emoji than previously selected
        if self.selected_index != abs_idx:
            # Select new emoji
            self.listbox.selection_clear(0, tk.END)
            self.listbox.selection_set(clicked_idx)
            self.selected_index = abs_idx
            self.update_status(f"Selected: {self.filtered_emojis[abs_idx].name}")
            return
    
        # If it's the same emoji and enough time has passed (throttle for performance)
        if current_time - self.last_click_time >= self.click_throttle:
            # Send the emoji
            emoji_data = self.filtered_emojis[abs_idx]
            # Update click time BEFORE sending to prevent double execution
            self.last_click_time = current_time
        
            # Now send the emoji
            self.handle_emoji_action(emoji_data=emoji_data, update_ui=True)
        
            # Always maintain selection
            self.listbox.selection_clear(0, tk.END)
            self.listbox.selection_set(clicked_idx)
        
            # Update status with a counter to show rapid sending
            status_suffix = "..."
            self.update_status(f"Sending: {emoji_data.name}{status_suffix}")

    def handle_emoji_action(self, event=None, emoji_data=None, update_ui=True):
        """Handle action when an emoji is selected (copy or direct paste)"""
        if event and not emoji_data:
            selection = self.listbox.curselection()
            if not selection:
                return
            emoji_data = self.filtered_emojis[self.current_page * PAGE_SIZE + selection[0]]
    
        if not emoji_data:
            return
        
        # Always copy to clipboard
        pyperclip.copy(emoji_data.unicode_key)
    
        # Update count
        emoji_data.count += 1
    
        # Always save data for accurate counting 
        self.save_emoji_data()
    
        # If direct paste is enabled, attempt to paste to game
        if self.direct_paste:
            # Try to find game window if we don't have it
            if not self.game_hwnd:
                self.game_hwnd = find_crystal_realms_window()
        
            if self.game_hwnd:
                success = send_emoji_to_game(emoji_data.unicode_key, self.game_hwnd)
                if success and update_ui:
                    self.update_status(f"Emoji sent to game: {emoji_data.unicode_key}")
                elif not success and update_ui:
                    self.update_status("Failed to send emoji to game")
            elif update_ui:
                self.update_status("Game window not found. Emoji copied to clipboard.")
        elif update_ui:
            self.update_status(f"Emoji copied to clipboard: {emoji_data.unicode_key}")
    
        # Update display to reflect new counts and sort order
        # We need to resort and redisplay the current page
        if update_ui:
            # First preserve the current page
            current_page = self.current_page 
            # Resort emojis
            self.filter_and_sort_emojis()
            # Restore the page
            self.navigate_to(current_page)
            # Update the specific item on current page
            self.update_specific_item(emoji_data)

    def update_specific_item(self, emoji_data):
        """Update a specific item in the listbox without redrawing everything"""
        # Check if the emoji is on the current page
        start_idx = self.current_page * PAGE_SIZE
        end_idx = start_idx + PAGE_SIZE
    
        # Find the emoji's new position in filtered_emojis
        try:
            abs_idx = next(i for i, e in enumerate(self.filtered_emojis) 
                          if e.unicode_key == emoji_data.unicode_key)
        except StopIteration:
            return  # Emoji not found in filtered list
    
        # Check if it's on the current page
        if start_idx <= abs_idx < end_idx:
            page_idx = abs_idx - start_idx
            pin_indicator = "📌 " if emoji_data.pinned else ""
            count_indicator = f" ({emoji_data.count})" if emoji_data.count > 0 else ""
            keybind_indicator = f" [{emoji_data.keybind}]" if emoji_data.keybind else ""
            display_text = f"{emoji_data.unicode_key} {pin_indicator}{emoji_data.name}{count_indicator}{keybind_indicator}"
            
            # Update the item in the listbox
            self.listbox.delete(page_idx)
            self.listbox.insert(page_idx, display_text)
        
            # If this was the selected item, keep it selected
            if self.selected_index == abs_idx:
                self.listbox.selection_clear(0, tk.END)
                self.listbox.selection_set(page_idx)

    def reset_count(self, emoji_data):
        emoji_data.count = 0
        self.save_emoji_data()
        self.filter_and_sort_emojis()

    def toggle_pin(self, emoji_data):
        new_state = not emoji_data.pinned
        
        emoji_data.pinned = new_state
        if emoji_data.pinned:
            # Find lowest manual_order among pinned items
            min_order = min((e.manual_order for e in self.emoji_data.values() if e.pinned), 
                          default=-1)
            emoji_data.manual_order = min_order - 1 if min_order != float('inf') else 0
        else:
            emoji_data.manual_order = float('inf')
        
        self.save_emoji_data()
        self.filter_and_sort_emojis()

    def move_pinned(self, emoji_data, direction):
        if not emoji_data.pinned:
            return
            
        pinned = sorted(
            [e for e in self.emoji_data.values() if e.pinned],
            key=lambda x: x.manual_order
        )
        
        idx = next(i for i, e in enumerate(pinned) if e.unicode_key == emoji_data.unicode_key)
        
        if direction == 'up' and idx > 0:
            pinned[idx].manual_order, pinned[idx-1].manual_order = \
                pinned[idx-1].manual_order, pinned[idx].manual_order
        elif direction == 'down' and idx < len(pinned) - 1:
            pinned[idx].manual_order, pinned[idx+1].manual_order = \
                pinned[idx+1].manual_order, pinned[idx].manual_order
        
        self.save_emoji_data()
        self.filter_and_sort_emojis()

    def set_keybind(self, emoji_data):
        """Open dialog to set a keybind for the emoji"""
        try:
            # Create the dialog and store it as an instance variable
            self.current_keybind_dialog = KeyBindDialog(self.root, emoji_data)
            
            # Schedule the keybind update after the dialog closes
            def safe_update_keybind():
                try:
                    # Check if a result exists and wasn't canceled
                    if hasattr(self.current_keybind_dialog, 'result') and self.current_keybind_dialog.result is not None:
                        new_keybind = self.current_keybind_dialog.result
                        old_keybind = emoji_data.keybind
                        
                        # First unregister ALL hotkeys to prevent any issues
                        self.unregister_all_hotkeys()
                        
                        # Remove old keybind if it exists
                        if old_keybind and old_keybind in self.keybind_to_emoji:
                            del self.keybind_to_emoji[old_keybind]
                        
                        # Clear the new keybind or set it
                        if new_keybind:
                            # Check if the new keybind is already in use
                            if new_keybind in self.keybind_to_emoji:
                                # Find the emoji currently using this keybind
                                conflicting_unicode = self.keybind_to_emoji[new_keybind]
                                conflicting_emoji = self.emoji_data.get(conflicting_unicode)
                                
                                if conflicting_emoji:
                                    # Clear the existing emoji's keybind
                                    conflicting_emoji.keybind = None
                                    # Remove the conflicting binding
                                    del self.keybind_to_emoji[new_keybind]
                                
                            # Now set the new keybind (after resolving conflicts)
                            emoji_data.keybind = new_keybind
                            self.keybind_to_emoji[new_keybind] = emoji_data.unicode_key
                            self.update_status(f"Keybind set: {new_keybind} → {emoji_data.unicode_key}")
                        else:
                            # Clear the keybind
                            emoji_data.keybind = None
                            self.update_status(f"Keybind cleared for {emoji_data.unicode_key}")
                        
                        # Save updated data
                        self.save_emoji_data()
                        
                        # Re-register hotkeys
                        self.register_system_wide_hotkeys()
                        
                        # Update UI
                        self.filter_and_sort_emojis()
                
                except Exception as e:
                    print(f"Error updating keybind: {e}")
                    traceback.print_exc()
                
                # Clear the dialog reference
                if hasattr(self, 'current_keybind_dialog'):
                    del self.current_keybind_dialog
            
            # Schedule the update to run after dialog closes
            self.root.after(50, safe_update_keybind)
        
        except Exception as e:
            print(f"Error in set_keybind: {e}")
            traceback.print_exc()

    def show_context_menu(self, event):
        try:
            idx = self.listbox.nearest(event.y)
            if idx < 0 or idx >= len(self.listbox.get(0, tk.END)):
                return
                
            self.listbox.selection_clear(0, tk.END)
            self.listbox.selection_set(idx)
            
            # Update selected index to match right-clicked item
            self.selected_index = self.current_page * PAGE_SIZE + idx
            
            emoji_data = self.filtered_emojis[self.selected_index]
            
            menu = Menu(self.root, tearoff=0)
            menu.add_command(
                label="Unpin" if emoji_data.pinned else "Pin to Top",
                command=lambda: self.toggle_pin(emoji_data)
            )
            
            if emoji_data.pinned:
                menu.add_command(
                    label="Move Up",
                    command=lambda: self.move_pinned(emoji_data, 'up')
                )
                menu.add_command(
                    label="Move Down",
                    command=lambda: self.move_pinned(emoji_data, 'down')
                )
            
            menu.add_separator()
            
            # Add keybind options
            keybind_label = f"Set/Clear Keybind [{emoji_data.keybind}]" if emoji_data.keybind else "Set/Clear Keybind"
            menu.add_command(
                label=keybind_label,
                command=lambda: self.set_keybind(emoji_data)
            )
            
            menu.add_separator()
            menu.add_command(
                label="Send to Game",
                command=lambda: self.handle_emoji_action(emoji_data=emoji_data)
            )
            menu.add_command(
                label="Copy",
                command=lambda: self.copy_only(emoji_data)
            )
            menu.add_command(
                label="Reset",
                command=lambda: self.reset_count(emoji_data)
            )
            
            menu.tk_popup(event.x_root, event.y_root)
            menu.grab_release()
        except Exception:
            pass

    def copy_only(self, emoji_data):
        """Just copy to clipboard without sending to game"""
        pyperclip.copy(emoji_data.unicode_key)
        emoji_data.count += 1
        self.save_emoji_data()
        self.update_status(f"Emoji copied to clipboard: {emoji_data.unicode_key}")
        
        # Update UI to reflect the change immediately
        current_page = self.current_page
        self.filter_and_sort_emojis()
        self.navigate_to(current_page)

    def save_emoji_data(self):
        data = {
            unicode_key: {
                'count': emoji.count,
                'pinned': emoji.pinned,
                'manual_order': emoji.manual_order,
                'keybind': emoji.keybind
            }
            for unicode_key, emoji in self.emoji_data.items()
        }
        save_data(data)

    def display_page(self):
        """Display the current page of emojis in the listbox with improved stability"""
        try:
            # Prepare data first without UI lock to minimize lock time
            start = self.current_page * PAGE_SIZE
            end = start + PAGE_SIZE
            
            # Create safe copy of data to prevent concurrent modification issues
            try:
                page_items = self.filtered_emojis[start:end] if start < len(self.filtered_emojis) else []
                
                # Pre-format all display texts outside the lock
                display_texts = []
                for emoji in page_items:
                    try:
                        pin_indicator = "📌 " if emoji.pinned else ""
                        count_indicator = f" ({emoji.count})" if emoji.count > 0 else ""
                        keybind_indicator = f" [{emoji.keybind}]" if emoji.keybind else ""
                        display_text = f"{emoji.unicode_key} {pin_indicator}{emoji.name}{count_indicator}{keybind_indicator}"
                        display_texts.append(display_text)
                    except Exception:
                        display_texts.append("Error loading emoji")
                
                total_pages = max(1, (len(self.filtered_emojis) + PAGE_SIZE - 1) // PAGE_SIZE)
                page_label_text = f"Page {self.current_page + 1} / {total_pages}"
                
            except Exception:
                display_texts = ["Error loading page"]
                page_label_text = "Error loading page"
            
            # No UI lock for faster operations - the stability fix from earlier
            try:
                # Clear current items
                self.listbox.delete(0, tk.END)
                
                # Insert new items in batches with UI refreshes
                for text in display_texts:
                    self.listbox.insert(tk.END, text)
                    # Process UI events every few items to keep responsiveness
                    if display_texts.index(text) % 5 == 0:
                        self.root.update_idletasks()
            
                # Update selection if needed
                if self.selected_index is not None:
                    selected_page_idx = self.selected_index - (self.current_page * PAGE_SIZE)
                    if 0 <= selected_page_idx < min(PAGE_SIZE, len(display_texts)):
                        self.listbox.selection_clear(0, tk.END)
                        self.listbox.selection_set(selected_page_idx)
                        # Ensure the selected item is visible by scrolling to it
                        self.listbox.see(selected_page_idx)
            
                # Update page indicator
                self.page_label.config(text=page_label_text)
                
                # Force UI update to show changes
                self.root.update_idletasks()
                
            except Exception:
                self.update_status("Display error - try searching or changing pages")
                
        except Exception:
            # Try to recover and show something to the user
            try:
                self.update_status("Error displaying page. Please try again.")
                # Force UI update
                self.root.update_idletasks()
            except:
                pass

    def toggle_always_on_top(self):
        self.always_on_top = not self.always_on_top
        self.root.attributes('-topmost', self.always_on_top)
        self.always_on_top_btn.config(
            text="Always on Top (On)" if self.always_on_top else "Always on Top (Off)"
        )
        
        # Update settings
        self.settings.always_on_top = self.always_on_top
        self.settings.save()

    def toggle_direct_paste(self):
        self.direct_paste = not self.direct_paste
        self.direct_paste_var.set(self.direct_paste)
        mode_text = "Send to game" if self.direct_paste else "Copy only"
        self.update_status(f"Mode changed to: {mode_text}")
        
        # Update settings
        self.settings.direct_paste = self.direct_paste
        self.settings.save()

    def toggle_return_focus(self):
        self.return_focus = not self.return_focus
        self.return_focus_var.set(self.return_focus)
        state_text = "enabled" if self.return_focus else "disabled"
        self.update_status(f"Return focus {state_text}")
        
        # Update settings
        self.settings.return_focus = self.return_focus
        self.settings.save()

    def setup_gui(self):
        # Create a main container frame
        self.container = Frame(self.root)
        self.container.pack(fill=tk.BOTH, expand=True)
        
        # Configure grid weights to ensure proper resizing
        self.container.grid_rowconfigure(2, weight=1)  # Listbox row
        self.container.grid_columnconfigure(0, weight=1)
        
        # Controls reminder - Updated for new click behavior
        self.controls_label = Label(
            self.container,
            text="Click Once to Select • Click Again to Copy/Send • Right Click for Options • Don't Hold Hotkeys",
            font=("Arial", 10),
            pady=5
        )
        self.controls_label.grid(row=0, column=0, sticky="ew")
        
        # Search frame
        self.search_frame = Frame(self.container)
        self.search_frame.grid(row=1, column=0, padx=5, pady=(2, 2), sticky="ew")
        self.search_frame.grid_columnconfigure(0, weight=1)  # Entry gets all extra space
        
        self.search_entry = Entry(
            self.search_frame,
            textvariable=self.search_var,
            font=("Arial", 14)
        )
        self.search_entry.grid(row=0, column=0, sticky="ew")
        
        # Magnifying glass icon/label on the right side
        self.search_icon = Label(
            self.search_frame,
            text="🔍",
            font=("Arial", 14)
        )
        self.search_icon.grid(row=0, column=1, padx=(5, 0))
        
        self.search_var.trace('w', lambda *args: self.filter_and_sort_emojis())
        
        # Listbox frame with fixed minimum height
        self.list_frame = Frame(self.container, height=175)  # Reduced to match screenshot 1 exactly
        self.list_frame.grid(row=2, column=0, padx=5, sticky="nsew")
        self.list_frame.grid_columnconfigure(0, weight=1)
        self.list_frame.grid_rowconfigure(0, weight=1)
        self.list_frame.pack_propagate(False)  # Using pack_propagate as well for more control
        self.list_frame.grid_propagate(False)  # Prevent frame from shrinking

        scrollbar = Scrollbar(self.list_frame)
        scrollbar.grid(row=0, column=1, sticky="ns")
        
        # Increased emoji font size by 2px and decreased name font size by 2px
        self.listbox = Listbox(
            self.list_frame,
            yscrollcommand=scrollbar.set,
            font=("Arial", 20),  # Increased from 18 to 20
            activestyle='none',  # Removes underline from active selection
            height=4  # Reduced to show exactly 4 items as in screenshot 1
        )
        self.listbox.grid(row=0, column=0, sticky="nsew")
        
        scrollbar.config(command=self.listbox.yview)
        
        # Bind to single clicks for the improved click behavior
        self.listbox.bind("<Button-1>", self.handle_click)
        self.listbox.bind("<Button-3>", self.show_context_menu)
        
        # Status label - reduced padding
        self.status_label = Label(
            self.container,
            text="Starting up...",
            font=("Arial", 10),
            bd=1,
            relief=tk.SUNKEN,
            anchor=tk.W
        )
        self.status_label.grid(row=3, column=0, sticky="ew", pady=(2, 0))
        
        # Options frame - reduced padding
        self.options_frame = Frame(self.container)
        self.options_frame.grid(row=4, column=0, padx=5, pady=(2, 2), sticky="ew")
        
        # Direct paste option
        self.direct_paste_var = BooleanVar(value=self.direct_paste)
        direct_paste_cb = Checkbutton(
            self.options_frame,
            text="Send directly to game",
            variable=self.direct_paste_var,
            command=self.toggle_direct_paste
        )
        direct_paste_cb.pack(side=tk.LEFT, padx=5)
        
        # Return focus option
        self.return_focus_var = BooleanVar(value=self.return_focus)
        return_focus_cb = Checkbutton(
            self.options_frame,
            text="Return focus",
            variable=self.return_focus_var,
            command=self.toggle_return_focus
        )
        return_focus_cb.pack(side=tk.LEFT, padx=5)
        
        # Keybinds enabled option
        self.keybinds_var = BooleanVar(value=self.keybinds_enabled)
        keybinds_cb = Checkbutton(
            self.options_frame,
            text="Enable keybinds (restart app lmaooo)",
            variable=self.keybinds_var,
            command=self.toggle_keybinds_enabled
        )
        keybinds_cb.pack(side=tk.LEFT, padx=5)
        
        # Send Enter key button
        send_enter_button = Button(
            self.options_frame,
            text="Send Enter",
            command=self.send_enter_key_to_game
        )
        send_enter_button.pack(side=tk.RIGHT, padx=5)
        
        # Navigation frame - reduced padding
        self.nav_frame = Frame(self.container)
        self.nav_frame.grid(row=5, column=0, padx=5, pady=(0, 2), sticky="ew")
        self.nav_frame.grid_columnconfigure(2, weight=1)  # Center column gets extra space
        
        # Left side buttons (in their own frame)
        self.left_buttons = Frame(self.nav_frame)
        self.left_buttons.grid(row=0, column=0)
        
        Button(
            self.left_buttons,
            text="⟨⟨ First",
            command=lambda: self.navigate_to(0),
            font=("Arial", 14)
        ).pack(side=tk.LEFT, padx=2)
        
        Button(
            self.left_buttons,
            text="⟨ Prev",
            command=lambda: self.navigate_to(self.current_page - 1),
            font=("Arial", 14)
        ).pack(side=tk.LEFT, padx=2)
        
        # Center elements (in their own frame with proper centering)
        self.center_frame = Frame(self.nav_frame)
        self.center_frame.grid(row=0, column=2, sticky="nsew")
        self.center_frame.grid_columnconfigure(0, weight=1)
        
        self.center_content = Frame(self.center_frame)
        self.center_content.grid(row=0, column=0)
        
        # Theme selector button
        theme_button = Button(
            self.center_content,
            text="Select Theme",
            command=self.open_theme_selector,
            font=("Arial", 12)
        )
        theme_button.pack(pady=(0, 2))
        
        # Exit application button centered above page indicator
        exit_button = Button(
            self.center_content,
            text="Exit Application",
            command=self.quit_app,
            font=("Arial", 12)
        )
        exit_button.pack(pady=(0, 2))
        
        self.page_label = Label(self.center_content, font=("Arial", 12))
        self.page_label.pack()
        
        self.always_on_top_btn = Button(
            self.center_content,
            text="Always on Top (On)" if self.always_on_top else "Always on Top (Off)",
            command=self.toggle_always_on_top,
            font=("Arial", 10)
        )
        self.always_on_top_btn.pack(pady=(2, 0))
        
        # Right side buttons (in their own frame)
        self.right_buttons = Frame(self.nav_frame)
        self.right_buttons.grid(row=0, column=3)
        
        Button(
            self.right_buttons,
            text="Next ⟩",
            command=lambda: self.navigate_to(self.current_page + 1),
            font=("Arial", 14)
        ).pack(side=tk.LEFT, padx=2)
        
        Button(
            self.right_buttons,
            text="Last ⟩⟩",
            command=lambda: self.navigate_to(float('inf')),
            font=("Arial", 14)
        ).pack(side=tk.LEFT, padx=2)
        
        # Set always on top state from settings
        if self.always_on_top:
            self.root.attributes('-topmost', True)

    def navigate_to(self, page):
        max_page = max(0, (len(self.filtered_emojis) - 1) // PAGE_SIZE)
        new_page = max(0, min(page, max_page))
        self.current_page = new_page
        self.display_page()

    def run(self):
        try:
            # Set up exception handler for the mainloop
            def handle_exception(exc_type, exc_value, exc_traceback):
                print(f"Uncaught exception: {exc_type.__name__}: {exc_value}")
                
                # Try to recover
                try:
                    self.update_status(f"Error: {exc_type.__name__}: {exc_value}")
                except:
                    pass
                
                # Don't exit the app on error
                return True
                
            # Install exception handler for Tkinter
            self.root.report_callback_exception = handle_exception
            
            # Update status to ready
            self.update_status("Application ready")
            
            # Now that the UI is fully set up, register hotkeys
            try:
                # Set up the Windows hook for hotkeys
                self.setup_win32_hotkeys()
            except Exception as e:
                print(f"Error setting up hotkeys: {e}")
                traceback.print_exc()
            
            # Start the main loop
            self.root.mainloop()
            
        except Exception as e:
            print(f"Fatal error in mainloop: {e}")
            traceback.print_exc()
            self.quit_app()

if __name__ == "__main__":
    try:
        app = EmojiPicker()
        app.run()
    except Exception as e:
        print(f"Fatal error: {e}")
        traceback.print_exc()
        # Try to clean up if possible
        try:
            if hasattr(app, 'tray_icon') and app.tray_icon:
                app.tray_icon.stop()
        except:
            pass
