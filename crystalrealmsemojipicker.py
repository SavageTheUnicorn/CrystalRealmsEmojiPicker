# Created by SavageTheUnicorn
# Date: 02/24/2025
# File Hash and source code can be found at: "https://github.com/SavageTheUnicorn/CrystalRealmsEmojiPicker"

import tkinter as tk
from tkinter import Entry, Listbox, Scrollbar, Button, Label, Frame, Menu
import requests
import pyperclip
from fontTools.ttLib import TTFont
import os
import sys
import json
import pystray
from PIL import Image, ImageDraw
from threading import Thread

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

class EmojiPicker:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Crystal Realms EmojiPicker - Double Click to Copy, Right Click for Options")
        self.root.geometry("600x500")
        self.root.minsize(300, 200)
        
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
        
        # Initialize system tray icon
        self.setup_tray()
        
        # Bind window close event
        self.root.protocol('WM_DELETE_WINDOW', self.toggle_window)
        
        if not self.supported_chars:
            self.show_font_warning()
        
        self.setup_gui()
        self.load_emojis()

    def create_tray_icon(self):
        """Create a simple emoji icon for the tray (fallback)"""
        icon_size = 32
        image = Image.new('RGB', (icon_size, icon_size), color='#7289DA')
        draw = ImageDraw.Draw(image)
        draw.text((icon_size//4, icon_size//4), "🦄", fill='white', size=32)
        return image

    def setup_tray(self):
        """Setup the system tray icon and menu"""
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
                pystray.MenuItem("Exit", self.quit_app)
            )
        )
        # Start the tray icon in a separate thread
        Thread(target=self.tray_icon.run, daemon=True).start()

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
        self.tray_icon.stop()
        self.root.quit()

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
        self.current_page = 0
        self.display_page()

    def copy_emoji(self, event=None):
        selection = self.listbox.curselection()
        if not selection:
            return
            
        emoji_data = self.filtered_emojis[self.current_page * PAGE_SIZE + selection[0]]
        pyperclip.copy(emoji_data.unicode_key)
        
        emoji_data.count += 1
        self.save_emoji_data()
        self.filter_and_sort_emojis()

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
        
        emoji_data = self.filtered_emojis[self.current_page * PAGE_SIZE + idx]
        
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
            label="Copy",
            command=lambda: self.copy_emoji()
        )
        menu.add_command(
            label="Reset",
            command=lambda: self.reset_count(emoji_data)
        )
        
        menu.tk_popup(event.x_root, event.y_root)
        menu.grab_release()

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
        
        total_pages = max(1, (len(self.filtered_emojis) + PAGE_SIZE - 1) // PAGE_SIZE)
        self.page_label.config(text=f"Page {self.current_page + 1} / {total_pages}")

    def toggle_always_on_top(self):
        self.always_on_top = not self.always_on_top
        self.root.attributes('-topmost', self.always_on_top)
        self.always_on_top_btn.config(
            text="Always on Top (On)" if self.always_on_top else "Always on Top (Off)"
        )

    def setup_gui(self):
        # Create a main container frame
        container = Frame(self.root)
        container.pack(fill=tk.BOTH, expand=True)
        
        # Configure grid weights to ensure proper resizing
        container.grid_rowconfigure(2, weight=1)  # Listbox row
        container.grid_columnconfigure(0, weight=1)
        
        # Controls reminder
        controls_label = Label(
            container,
            text="Controls: Double Click to Copy • Right Click for Options",
            font=("Arial", 10),
            pady=5
        )
        controls_label.grid(row=0, column=0, sticky="ew")
        
        # Search frame
        search_frame = Frame(container)
        search_frame.grid(row=1, column=0, padx=5, pady=5, sticky="ew")
        search_frame.grid_columnconfigure(0, weight=1)
        
        search_entry = Entry(
            search_frame,
            textvariable=self.search_var,
            font=("Arial", 14)
        )
        search_entry.grid(row=0, column=0, sticky="ew")
        self.search_var.trace('w', lambda *args: self.filter_and_sort_emojis())
        
        # Listbox frame
        list_frame = Frame(container)
        list_frame.grid(row=2, column=0, padx=5, sticky="nsew")
        list_frame.grid_columnconfigure(0, weight=1)
        list_frame.grid_rowconfigure(0, weight=1)
        
        scrollbar = Scrollbar(list_frame)
        scrollbar.grid(row=0, column=1, sticky="ns")
        
        self.listbox = Listbox(
            list_frame,
            yscrollcommand=scrollbar.set,
            font=("Arial", 18)
        )
        self.listbox.grid(row=0, column=0, sticky="nsew")
        
        scrollbar.config(command=self.listbox.yview)
        
        self.listbox.bind("<Double-Button-1>", self.copy_emoji)
        self.listbox.bind("<Button-3>", self.show_context_menu)
        
        # Navigation frame
        nav_frame = Frame(container)
        nav_frame.grid(row=3, column=0, padx=5, pady=5, sticky="ew")
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
        
        self.page_label = Label(center_content, font=("Arial", 14))
        self.page_label.pack()
        
        self.always_on_top_btn = Button(
            center_content,
            text="Always on Top (Off)",
            command=self.toggle_always_on_top,
            font=("Arial", 12)
        )
        self.always_on_top_btn.pack(pady=(5, 0))
        
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
        self.root.mainloop()

if __name__ == "__main__":
    app = EmojiPicker()
    app.run()
