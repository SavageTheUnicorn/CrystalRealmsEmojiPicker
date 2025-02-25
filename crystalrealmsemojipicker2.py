# Created by SavageTheUnicorn
# Date: 02/24/2025
# File Hash and source code can be found at: "https://github.com/SavageTheUnicorn/CrystalRealmsEmojiPicker"

import tkinter as tk
from tkinter import Entry, Listbox, Scrollbar, Button, Label, Frame, Menu, StringVar, Checkbutton
import requests
import pyperclip
from fontTools.ttLib import TTFont
import os
import sys
import json
import pystray
from PIL import Image, ImageDraw
from threading import Thread
import win32gui
import win32process
import win32api
import win32con
import psutil
import ctypes
import time

EMOJI_API = "https://emojihub.yurace.pro/api/all"
PAGE_SIZE = 20

def get_data_path():
    """Get path to data file in user's AppData directory"""
    app_data = os.path.join(os.getenv('APPDATA'), 'EmojiPicker')
    if not os.path.exists(app_data):
        os.makedirs(app_data)
    return os.path.join(app_data, 'emoji_data.json')

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
    def __init__(self, name, unicode_key, count=0, pinned=False, manual_order=float('inf')):
        self.name = name
        self.unicode_key = unicode_key
        self.count = count
        self.pinned = pinned
        self.manual_order = manual_order

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
    response = requests.get(EMOJI_API)
    if response.status_code == 200:
        return response.json()
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
    except Exception as e:
        print(f"Error saving data: {e}")

def find_crystal_realms_window():
    """Find Crystal Realms window by its title and process name"""
    def enum_window_callback(hwnd, result):
        if win32gui.IsWindowVisible(hwnd):
            window_title = win32gui.GetWindowText(hwnd)
            if "crystal realms" in window_title.lower():
                try:
                    _, process_id = win32process.GetWindowThreadProcessId(hwnd)
                    process = psutil.Process(process_id)
                    if "crystal_realms" in process.name().lower():
                        result.append(hwnd)
                except (psutil.NoSuchProcess, psutil.AccessDenied):
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
                except (psutil.NoSuchProcess, psutil.AccessDenied):
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
    except Exception as e:
        print(f"Error sending emoji to game: {e}")
        # Try to restore focus
        try:
            win32gui.SetForegroundWindow(current_hwnd)
        except:
            pass
        return False

class EmojiPicker:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Crystal Realms EmojiPicker - Made By SavageTheUnicorn")
        self.root.geometry("600x450")  # Adjusted height to better match screenshot 1
        self.root.minsize(300, 350)  # Adjusted minimum height to match
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
        self.always_on_top = False
        self.direct_paste = True  # Default to direct paste enabled
        self.return_focus = True  # Default to returning focus to picker
        self.game_hwnd = None    # Will store game window handle
        self.tray_icon = None    # Will store the tray icon instance
        
        # New variables for improved clicking behavior
        self.selected_index = None  # Track currently selected emoji index
        self.last_click_time = 0    # Time of last click
        self.click_throttle = 0.05  # Minimum seconds between spam clicks (50ms)
        
        # Initialize system tray icon
        self.setup_tray()
        
        # Bind window close event to hide window instead of closing app
        self.root.protocol('WM_DELETE_WINDOW', self.on_closing)
        
        if not self.supported_chars:
            self.show_font_warning()
        
        self.setup_gui()
        self.load_emojis()
        
        # Try to find the game window on startup
        self.find_game_window()

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
        except Exception as e:
            print(f"Error setting up tray icon: {e}")

    def toggle_window(self, _=None):
        """Toggle the main window visibility"""
        if self.root.state() == 'withdrawn':
            self.root.deiconify()
            self.root.lift()
            if self.always_on_top:
                self.root.attributes('-topmost', True)
        else:
            self.root.withdraw()

    def quit_app(self, _=None):
        """Properly close the application"""
        try:
            if self.tray_icon:
                self.tray_icon.stop()
        except Exception as e:
            print(f"Error stopping tray icon: {e}")
        finally:
            self.root.destroy()  # Use destroy instead of quit

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
            # Reduce delay for faster response
            win32api.Sleep(20)  # Small delay
            
            # Send Enter keystroke
            ctypes.windll.user32.keybd_event(0x0D, 0, 0, 0)  # VK_RETURN (Enter key) press
            win32api.Sleep(10)  # Small delay
            ctypes.windll.user32.keybd_event(0x0D, 0, win32con.KEYEVENTF_KEYUP, 0)  # Enter key release
            
            # Return focus to emoji picker if option is enabled
            if hasattr(current_hwnd, 'return_focus') and current_hwnd.return_focus:
                win32api.Sleep(20)  # Small delay
                win32gui.SetForegroundWindow(current_hwnd)
            
            self.update_status("Enter key sent to game")
            return True
        except Exception as e:
            print(f"Error sending Enter key to game: {e}")
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
                    
                    self.emoji_data[unicode_char] = EmojiData(name, unicode_char, count, pinned, manual_order)
            except (ValueError, TypeError) as e:
                print(f"Error processing emoji {emoji.get('name', 'unknown')}: {e}")
                continue
        
        self.filter_and_sort_emojis()

    def filter_and_sort_emojis(self):
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
            display_text = f"{emoji_data.unicode_key} {pin_indicator}{emoji_data.name}{count_indicator}"
        
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
        emoji_data.pinned = not emoji_data.pinned
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

    def show_context_menu(self, event):
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
                'manual_order': emoji.manual_order
            }
            for unicode_key, emoji in self.emoji_data.items()
        }
        save_data(data)

    def display_page(self):
        self.listbox.delete(0, tk.END)
        start = self.current_page * PAGE_SIZE
        end = start + PAGE_SIZE
    
        for emoji in self.filtered_emojis[start:end]:
            pin_indicator = "📌 " if emoji.pinned else ""
            count_indicator = f" ({emoji.count})" if emoji.count > 0 else ""
            display_text = f"{emoji.unicode_key} {pin_indicator}{emoji.name}{count_indicator}"
            self.listbox.insert(tk.END, display_text)
    
        # If we have a selected index, make sure it's selected in the UI if visible
        if self.selected_index is not None:
            selected_page_idx = self.selected_index - (self.current_page * PAGE_SIZE)
            if 0 <= selected_page_idx < PAGE_SIZE:
                self.listbox.selection_clear(0, tk.END)
                self.listbox.selection_set(selected_page_idx)
                # Ensure the selected item is visible by scrolling to it
                self.listbox.see(selected_page_idx)
    
        total_pages = max(1, (len(self.filtered_emojis) + PAGE_SIZE - 1) // PAGE_SIZE)
        self.page_label.config(text=f"Page {self.current_page + 1} / {total_pages}")

    def toggle_always_on_top(self):
        self.always_on_top = not self.always_on_top
        self.root.attributes('-topmost', self.always_on_top)
        self.always_on_top_btn.config(
            text="Always on Top (On)" if self.always_on_top else "Always on Top (Off)"
        )

    def toggle_direct_paste(self):
        self.direct_paste = not self.direct_paste
        self.direct_paste_var.set(self.direct_paste)
        mode_text = "Send to game" if self.direct_paste else "Copy only"
        self.update_status(f"Mode changed to: {mode_text}")

    def setup_gui(self):
        # Create a main container frame
        container = Frame(self.root)
        container.pack(fill=tk.BOTH, expand=True)
        
        # Configure grid weights to ensure proper resizing
        container.grid_rowconfigure(2, weight=1)  # Listbox row
        container.grid_columnconfigure(0, weight=1)
        
        # Controls reminder - Updated for new click behavior
        controls_label = Label(
            container,
            text="Click Once to Select • Click Again to Copy/Send • Right Click for Options",
            font=("Arial", 10),
            pady=5
        )
        controls_label.grid(row=0, column=0, sticky="ew")
        
        # Search frame
        search_frame = Frame(container)
        search_frame.grid(row=1, column=0, padx=5, pady=(2, 2), sticky="ew")
        search_frame.grid_columnconfigure(0, weight=1)  # Entry gets all extra space
        
        search_entry = Entry(
            search_frame,
            textvariable=self.search_var,
            font=("Arial", 14)
        )
        search_entry.grid(row=0, column=0, sticky="ew")
        
        # Magnifying glass icon/label on the right side
        search_icon = Label(
            search_frame,
            text="🔍",
            font=("Arial", 14)
        )
        search_icon.grid(row=0, column=1, padx=(5, 0))
        
        self.search_var.trace('w', lambda *args: self.filter_and_sort_emojis())
        
        # Listbox frame with fixed minimum height
        list_frame = Frame(container, height=175)  # Reduced to match screenshot 1 exactly
        list_frame.grid(row=2, column=0, padx=5, sticky="nsew")
        list_frame.grid_columnconfigure(0, weight=1)
        list_frame.grid_rowconfigure(0, weight=1)
        list_frame.pack_propagate(False)  # Using pack_propagate as well for more control
        list_frame.grid_propagate(False)  # Prevent frame from shrinking

        scrollbar = Scrollbar(list_frame)
        scrollbar.grid(row=0, column=1, sticky="ns")
        
        self.listbox = Listbox(
            list_frame,
            yscrollcommand=scrollbar.set,
            font=("Arial", 18),
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
            container,
            text="Starting up...",
            font=("Arial", 10),
            bd=1,
            relief=tk.SUNKEN,
            anchor=tk.W
        )
        self.status_label.grid(row=3, column=0, sticky="ew", pady=(2, 0))
        
        # Options frame - reduced padding
        options_frame = Frame(container)
        options_frame.grid(row=4, column=0, padx=5, pady=(2, 2), sticky="ew")
        
        # Direct paste option
        self.direct_paste_var = StringVar(value=int(self.direct_paste))
        direct_paste_cb = Checkbutton(
            options_frame,
            text="Send directly to game(Make sure chat is open)",
            variable=self.direct_paste_var,
            onvalue="1",
            offvalue="0",
            command=self.toggle_direct_paste
        )
        direct_paste_cb.pack(side=tk.LEFT, padx=5)
        
        # Send Enter key button
        send_enter_button = Button(
            options_frame,
            text="Send Enter",
            command=self.send_enter_key_to_game
        )
        send_enter_button.pack(side=tk.RIGHT, padx=5)
        
        # Navigation frame - reduced padding
        nav_frame = Frame(container)
        nav_frame.grid(row=5, column=0, padx=5, pady=(0, 2), sticky="ew")
        nav_frame.grid_columnconfigure(2, weight=1)  # Center column gets extra space
        
        # Left side buttons (in their own frame)
        left_buttons = Frame(nav_frame)
        left_buttons.grid(row=0, column=0)
        
        Button(
            left_buttons,
            text="⟨⟨ First",
            command=lambda: self.navigate_to(0),
            font=("Arial", 14)
        ).pack(side=tk.LEFT, padx=2)
        
        Button(
            left_buttons,
            text="⟨ Prev",
            command=lambda: self.navigate_to(self.current_page - 1),
            font=("Arial", 14)
        ).pack(side=tk.LEFT, padx=2)
        
        # Center elements (in their own frame with proper centering)
        center_frame = Frame(nav_frame)
        center_frame.grid(row=0, column=2, sticky="nsew")
        center_frame.grid_columnconfigure(0, weight=1)
        
        center_content = Frame(center_frame)
        center_content.grid(row=0, column=0)
        
        # Exit application button centered above page indicator
        exit_button = Button(
            center_content,
            text="Exit Application",
            command=self.quit_app,
            font=("Arial", 12)
        )
        exit_button.pack(pady=(0, 2))
        
        self.page_label = Label(center_content, font=("Arial", 12))
        self.page_label.pack()
        
        self.always_on_top_btn = Button(
            center_content,
            text="Always on Top (Off)",
            command=self.toggle_always_on_top,
            font=("Arial", 10)
        )
        self.always_on_top_btn.pack(pady=(2, 0))
        
        # Right side buttons (in their own frame)
        right_buttons = Frame(nav_frame)
        right_buttons.grid(row=0, column=3)
        
        Button(
            right_buttons,
            text="Next ⟩",
            command=lambda: self.navigate_to(self.current_page + 1),
            font=("Arial", 14)
        ).pack(side=tk.LEFT, padx=2)
        
        Button(
            right_buttons,
            text="Last ⟩⟩",
            command=lambda: self.navigate_to(float('inf')),
            font=("Arial", 14)
        ).pack(side=tk.LEFT, padx=2)

    def navigate_to(self, page):
        max_page = (len(self.filtered_emojis) - 1) // PAGE_SIZE
        self.current_page = max(0, min(page, max_page))
        self.display_page()

    def run(self):
        try:
            self.root.mainloop()
        except Exception as e:
            print(f"Error in main loop: {e}")
            self.quit_app()

if __name__ == "__main__":
    try:
        app = EmojiPicker()
        app.run()
    except Exception as e:
        print(f"Fatal error: {e}")
        # Try to clean up if possible
        try:
            if hasattr(app, 'tray_icon') and app.tray_icon:
                app.tray_icon.stop()
        except:
            pass
