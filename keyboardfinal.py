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
import socket
import platform
import unicodedata
import re
import queue
import psutil
import ctypes
from ctypes import wintypes
import time
import traceback
import threading
import win32clipboard
import tempfile
import atexit

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

# Primary API endpoint
EMOJI_API = "https://emojihub.yurace.pro/api/all"

# Alternative API endpoints as fallbacks
EMOJI_API_FALLBACKS = []

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

# Built-in emoji dataset from noto_emojis_converted.txt
BUILT_IN_EMOJI_DATA = [
    {'name': 'Leftwards Arrow', 'unicode': '2190', 'category': 'Symbols'},
    {'name': 'Upwards Arrow', 'unicode': '2191', 'category': 'Symbols'},
    {'name': 'Rightwards Arrow', 'unicode': '2192', 'category': 'Symbols'},
    {'name': 'Downwards Arrow', 'unicode': '2193', 'category': 'Symbols'},
    {'name': 'Left Right Arrow', 'unicode': '2194', 'category': 'Symbols'},
    {'name': 'Up Down Arrow', 'unicode': '2195', 'category': 'Symbols'},
    {'name': 'North West Arrow', 'unicode': '2196', 'category': 'Symbols'},
    {'name': 'North East Arrow', 'unicode': '2197', 'category': 'Symbols'},
    {'name': 'South East Arrow', 'unicode': '2198', 'category': 'Symbols'},
    {'name': 'South West Arrow', 'unicode': '2199', 'category': 'Symbols'},
    {'name': 'Leftwards Arrow With Stroke', 'unicode': '219A', 'category': 'Symbols'},
    {'name': 'Rightwards Arrow With Stroke', 'unicode': '219B', 'category': 'Symbols'},
    {'name': 'Leftwards Wave Arrow', 'unicode': '219C', 'category': 'Symbols'},
    {'name': 'Rightwards Wave Arrow', 'unicode': '219D', 'category': 'Symbols'},
    {'name': 'Leftwards Two Headed Arrow', 'unicode': '219E', 'category': 'Symbols'},
    {'name': 'Upwards Two Headed Arrow', 'unicode': '219F', 'category': 'Symbols'},
    {'name': 'Rightwards Two Headed Arrow', 'unicode': '21A0', 'category': 'Symbols'},
    {'name': 'Downwards Two Headed Arrow', 'unicode': '21A1', 'category': 'Symbols'},
    {'name': 'Leftwards Arrow With Tail', 'unicode': '21A2', 'category': 'Symbols'},
    {'name': 'Rightwards Arrow With Tail', 'unicode': '21A3', 'category': 'Symbols'},
    {'name': 'Leftwards Arrow From Bar', 'unicode': '21A4', 'category': 'Symbols'},
    {'name': 'Upwards Arrow From Bar', 'unicode': '21A5', 'category': 'Symbols'},
    {'name': 'Rightwards Arrow From Bar', 'unicode': '21A6', 'category': 'Symbols'},
    {'name': 'Downwards Arrow From Bar', 'unicode': '21A7', 'category': 'Symbols'},
    {'name': 'Up Down Arrow With Base', 'unicode': '21A8', 'category': 'Symbols'},
    {'name': 'Leftwards Arrow With Hook', 'unicode': '21A9', 'category': 'Symbols'},
    {'name': 'Rightwards Arrow With Hook', 'unicode': '21AA', 'category': 'Symbols'},
    {'name': 'Leftwards Arrow With Loop', 'unicode': '21AB', 'category': 'Symbols'},
    {'name': 'Rightwards Arrow With Loop', 'unicode': '21AC', 'category': 'Symbols'},
    {'name': 'Left Right Wave Arrow', 'unicode': '21AD', 'category': 'Symbols'},
    {'name': 'Left Right Arrow With Stroke', 'unicode': '21AE', 'category': 'Symbols'},
    {'name': 'Downwards Zigzag Arrow', 'unicode': '21AF', 'category': 'Symbols'},
    {'name': 'Upwards Arrow With Tip Leftwards', 'unicode': '21B0', 'category': 'Symbols'},
    {'name': 'Upwards Arrow With Tip Rightwards', 'unicode': '21B1', 'category': 'Symbols'},
    {'name': 'Downwards Arrow With Tip Leftwards', 'unicode': '21B2', 'category': 'Symbols'},
    {'name': 'Downwards Arrow With Tip Rightwards', 'unicode': '21B3', 'category': 'Symbols'},
    {'name': 'Rightwards Arrow With Corner Downwards', 'unicode': '21B4', 'category': 'Symbols'},
    {'name': 'Downwards Arrow With Corner Leftwards', 'unicode': '21B5', 'category': 'Symbols'},
    {'name': 'Anticlockwise Top Semicircle Arrow', 'unicode': '21B6', 'category': 'Symbols'},
    {'name': 'Clockwise Top Semicircle Arrow', 'unicode': '21B7', 'category': 'Symbols'},
    {'name': 'North West Arrow To Long Bar', 'unicode': '21B8', 'category': 'Symbols'},
    {'name': 'Leftwards Arrow To Bar Over Rightwards Arrow To Bar', 'unicode': '21B9', 'category': 'Symbols'},
    {'name': 'Anticlockwise Open Circle Arrow', 'unicode': '21BA', 'category': 'Symbols'},
    {'name': 'Clockwise Open Circle Arrow', 'unicode': '21BB', 'category': 'Symbols'},
    {'name': 'Leftwards Harpoon With Barb Upwards', 'unicode': '21BC', 'category': 'Symbols'},
    {'name': 'Leftwards Harpoon With Barb Downwards', 'unicode': '21BD', 'category': 'Symbols'},
    {'name': 'Upwards Harpoon With Barb Rightwards', 'unicode': '21BE', 'category': 'Symbols'},
    {'name': 'Upwards Harpoon With Barb Leftwards', 'unicode': '21BF', 'category': 'Symbols'},
    {'name': 'Rightwards Harpoon With Barb Upwards', 'unicode': '21C0', 'category': 'Symbols'},
    {'name': 'Rightwards Harpoon With Barb Downwards', 'unicode': '21C1', 'category': 'Symbols'},
    {'name': 'Downwards Harpoon With Barb Rightwards', 'unicode': '21C2', 'category': 'Symbols'},
    {'name': 'Downwards Harpoon With Barb Leftwards', 'unicode': '21C3', 'category': 'Symbols'},
    {'name': 'Rightwards Arrow Over Leftwards Arrow', 'unicode': '21C4', 'category': 'Symbols'},
    {'name': 'Upwards Arrow Leftwards Of Downwards Arrow', 'unicode': '21C5', 'category': 'Symbols'},
    {'name': 'Leftwards Arrow Over Rightwards Arrow', 'unicode': '21C6', 'category': 'Symbols'},
    {'name': 'Leftwards Paired Arrows', 'unicode': '21C7', 'category': 'Symbols'},
    {'name': 'Upwards Paired Arrows', 'unicode': '21C8', 'category': 'Symbols'},
    {'name': 'Rightwards Paired Arrows', 'unicode': '21C9', 'category': 'Symbols'},
    {'name': 'Downwards Paired Arrows', 'unicode': '21CA', 'category': 'Symbols'},
    {'name': 'Leftwards Harpoon Over Rightwards Harpoon', 'unicode': '21CB', 'category': 'Symbols'},
    {'name': 'Rightwards Harpoon Over Leftwards Harpoon', 'unicode': '21CC', 'category': 'Symbols'},
    {'name': 'Leftwards Double Arrow With Stroke', 'unicode': '21CD', 'category': 'Symbols'},
    {'name': 'Left Right Double Arrow With Stroke', 'unicode': '21CE', 'category': 'Symbols'},
    {'name': 'Rightwards Double Arrow With Stroke', 'unicode': '21CF', 'category': 'Symbols'},
    {'name': 'Leftwards Double Arrow', 'unicode': '21D0', 'category': 'Symbols'},
    {'name': 'Upwards Double Arrow', 'unicode': '21D1', 'category': 'Symbols'},
    {'name': 'Rightwards Double Arrow', 'unicode': '21D2', 'category': 'Symbols'},
    {'name': 'Downwards Double Arrow', 'unicode': '21D3', 'category': 'Symbols'},
    {'name': 'Left Right Double Arrow', 'unicode': '21D4', 'category': 'Symbols'},
    {'name': 'Up Down Double Arrow', 'unicode': '21D5', 'category': 'Symbols'},
    {'name': 'North West Double Arrow', 'unicode': '21D6', 'category': 'Symbols'},
    {'name': 'North East Double Arrow', 'unicode': '21D7', 'category': 'Symbols'},
    {'name': 'South East Double Arrow', 'unicode': '21D8', 'category': 'Symbols'},
    {'name': 'South West Double Arrow', 'unicode': '21D9', 'category': 'Symbols'},
    {'name': 'Leftwards Triple Arrow', 'unicode': '21DA', 'category': 'Symbols'},
    {'name': 'Rightwards Triple Arrow', 'unicode': '21DB', 'category': 'Symbols'},
    {'name': 'Leftwards Squiggle Arrow', 'unicode': '21DC', 'category': 'Symbols'},
    {'name': 'Rightwards Squiggle Arrow', 'unicode': '21DD', 'category': 'Symbols'},
    {'name': 'Upwards Arrow With Double Stroke', 'unicode': '21DE', 'category': 'Symbols'},
    {'name': 'Downwards Arrow With Double Stroke', 'unicode': '21DF', 'category': 'Symbols'},
    {'name': 'Leftwards Dashed Arrow', 'unicode': '21E0', 'category': 'Symbols'},
    {'name': 'Upwards Dashed Arrow', 'unicode': '21E1', 'category': 'Symbols'},
    {'name': 'Rightwards Dashed Arrow', 'unicode': '21E2', 'category': 'Symbols'},
    {'name': 'Downwards Dashed Arrow', 'unicode': '21E3', 'category': 'Symbols'},
    {'name': 'Leftwards Arrow To Bar', 'unicode': '21E4', 'category': 'Symbols'},
    {'name': 'Rightwards Arrow To Bar', 'unicode': '21E5', 'category': 'Symbols'},
    {'name': 'Leftwards White Arrow', 'unicode': '21E6', 'category': 'Symbols'},
    {'name': 'Upwards White Arrow', 'unicode': '21E7', 'category': 'Symbols'},
    {'name': 'Rightwards White Arrow', 'unicode': '21E8', 'category': 'Symbols'},
    {'name': 'Downwards White Arrow', 'unicode': '21E9', 'category': 'Symbols'},
    {'name': 'Upwards White Arrow From Bar', 'unicode': '21EA', 'category': 'Symbols'},
    {'name': 'Upwards White Arrow On Pedestal', 'unicode': '21EB', 'category': 'Symbols'},
    {'name': 'Upwards White Arrow On Pedestal With Horizontal Bar', 'unicode': '21EC', 'category': 'Symbols'},
    {'name': 'Upwards White Arrow On Pedestal With Vertical Bar', 'unicode': '21ED', 'category': 'Symbols'},
    {'name': 'Upwards White Double Arrow', 'unicode': '21EE', 'category': 'Symbols'},
    {'name': 'Upwards White Double Arrow On Pedestal', 'unicode': '21EF', 'category': 'Symbols'},
    {'name': 'Rightwards White Arrow From Wall', 'unicode': '21F0', 'category': 'Symbols'},
    {'name': 'North West Arrow To Corner', 'unicode': '21F1', 'category': 'Symbols'},
    {'name': 'South East Arrow To Corner', 'unicode': '21F2', 'category': 'Symbols'},
    {'name': 'Up Down White Arrow', 'unicode': '21F3', 'category': 'Symbols'},
    {'name': 'Right Arrow With Small Circle', 'unicode': '21F4', 'category': 'Symbols'},
    {'name': 'Downwards Arrow Leftwards Of Upwards Arrow', 'unicode': '21F5', 'category': 'Symbols'},
    {'name': 'Three Rightwards Arrows', 'unicode': '21F6', 'category': 'Symbols'},
    {'name': 'Leftwards Arrow With Vertical Stroke', 'unicode': '21F7', 'category': 'Symbols'},
    {'name': 'Rightwards Arrow With Vertical Stroke', 'unicode': '21F8', 'category': 'Symbols'},
    {'name': 'Left Right Arrow With Vertical Stroke', 'unicode': '21F9', 'category': 'Symbols'},
    {'name': 'Leftwards Arrow With Double Vertical Stroke', 'unicode': '21FA', 'category': 'Symbols'},
    {'name': 'Rightwards Arrow With Double Vertical Stroke', 'unicode': '21FB', 'category': 'Symbols'},
    {'name': 'Left Right Arrow With Double Vertical Stroke', 'unicode': '21FC', 'category': 'Symbols'},
    {'name': 'Leftwards Open-Headed Arrow', 'unicode': '21FD', 'category': 'Symbols'},
    {'name': 'Rightwards Open-Headed Arrow', 'unicode': '21FE', 'category': 'Symbols'},
    {'name': 'Left Right Open-Headed Arrow', 'unicode': '21FF', 'category': 'Symbols'},
    {'name': 'Diameter Sign', 'unicode': '2300', 'category': 'Symbols'},
    {'name': 'Electric Arrow', 'unicode': '2301', 'category': 'Symbols'},
    {'name': 'House', 'unicode': '2302', 'category': 'Symbols'},
    {'name': 'Up Arrowhead', 'unicode': '2303', 'category': 'Symbols'},
    {'name': 'Down Arrowhead', 'unicode': '2304', 'category': 'Symbols'},
    {'name': 'Projective', 'unicode': '2305', 'category': 'Symbols'},
    {'name': 'Perspective', 'unicode': '2306', 'category': 'Symbols'},
    {'name': 'Wavy Line', 'unicode': '2307', 'category': 'Symbols'},
    {'name': 'Left Ceiling', 'unicode': '2308', 'category': 'Symbols'},
    {'name': 'Right Ceiling', 'unicode': '2309', 'category': 'Symbols'},
    {'name': 'Left Floor', 'unicode': '230A', 'category': 'Symbols'},
    {'name': 'Right Floor', 'unicode': '230B', 'category': 'Symbols'},
    {'name': 'Bottom Right Crop', 'unicode': '230C', 'category': 'Symbols'},
    {'name': 'Bottom Left Crop', 'unicode': '230D', 'category': 'Symbols'},
    {'name': 'Top Right Crop', 'unicode': '230E', 'category': 'Symbols'},
    {'name': 'Top Left Crop', 'unicode': '230F', 'category': 'Symbols'},
    {'name': 'Reversed Not Sign', 'unicode': '2310', 'category': 'Symbols'},
    {'name': 'Square Lozenge', 'unicode': '2311', 'category': 'Symbols'},
    {'name': 'Arc', 'unicode': '2312', 'category': 'Symbols'},
    {'name': 'Segment', 'unicode': '2313', 'category': 'Symbols'},
    {'name': 'Sector', 'unicode': '2314', 'category': 'Symbols'},
    {'name': 'Telephone Recorder', 'unicode': '2315', 'category': 'Symbols'},
    {'name': 'Position Indicator', 'unicode': '2316', 'category': 'Symbols'},
    {'name': 'Viewdata Square', 'unicode': '2317', 'category': 'Symbols'},
    {'name': 'Place Of Interest Sign', 'unicode': '2318', 'category': 'Travel & Places'},
    {'name': 'Turned Not Sign', 'unicode': '2319', 'category': 'Symbols'},
    {'name': 'Watch', 'unicode': '231A', 'category': 'Symbols'},
    {'name': 'Hourglass', 'unicode': '231B', 'category': 'Symbols'},
    {'name': 'Top Left Corner', 'unicode': '231C', 'category': 'Symbols'},
    {'name': 'Top Right Corner', 'unicode': '231D', 'category': 'Symbols'},
    {'name': 'Bottom Left Corner', 'unicode': '231E', 'category': 'Symbols'},
    {'name': 'Bottom Right Corner', 'unicode': '231F', 'category': 'Symbols'},
    {'name': 'Top Half Integral', 'unicode': '2320', 'category': 'Symbols'},
    {'name': 'Bottom Half Integral', 'unicode': '2321', 'category': 'Symbols'},
    {'name': 'Frown', 'unicode': '2322', 'category': 'Symbols'},
    {'name': 'Smile', 'unicode': '2323', 'category': 'Smileys & Emotion'},
    {'name': 'Up Arrowhead Between Two Horizontal Bars', 'unicode': '2324', 'category': 'Symbols'},
    {'name': 'Option Key', 'unicode': '2325', 'category': 'Symbols'},
    {'name': 'Erase To The Right', 'unicode': '2326', 'category': 'Symbols'},
    {'name': 'X In A Rectangle Box', 'unicode': '2327', 'category': 'Symbols'},
    {'name': 'Keyboard', 'unicode': '2328', 'category': 'Symbols'},
    {'name': 'Left-Pointing Angle Bracket', 'unicode': '2329', 'category': 'Symbols'},
    {'name': 'Right-Pointing Angle Bracket', 'unicode': '232A', 'category': 'Symbols'},
    {'name': 'Erase To The Left', 'unicode': '232B', 'category': 'Symbols'},
    {'name': 'Benzene Ring', 'unicode': '232C', 'category': 'Symbols'},
    {'name': 'Cylindricity', 'unicode': '232D', 'category': 'Symbols'},
    {'name': 'All Around-Profile', 'unicode': '232E', 'category': 'Symbols'},
    {'name': 'Symmetry', 'unicode': '232F', 'category': 'Symbols'},
    {'name': 'Total Runout', 'unicode': '2330', 'category': 'Symbols'},
    {'name': 'Dimension Origin', 'unicode': '2331', 'category': 'Symbols'},
    {'name': 'Conical Taper', 'unicode': '2332', 'category': 'Symbols'},
    {'name': 'Slope', 'unicode': '2333', 'category': 'Symbols'},
    {'name': 'Counterbore', 'unicode': '2334', 'category': 'Symbols'},
    {'name': 'Countersink', 'unicode': '2335', 'category': 'Symbols'},
    {'name': 'Apl Functional Symbol I-Beam', 'unicode': '2336', 'category': 'Symbols'},
    {'name': 'Apl Functional Symbol Squish Quad', 'unicode': '2337', 'category': 'Symbols'},
    {'name': 'Apl Functional Symbol Quad Equal', 'unicode': '2338', 'category': 'Symbols'},
    {'name': 'Apl Functional Symbol Quad Divide', 'unicode': '2339', 'category': 'Symbols'},
    {'name': 'Apl Functional Symbol Quad Diamond', 'unicode': '233A', 'category': 'Symbols'},
    {'name': 'Apl Functional Symbol Quad Jot', 'unicode': '233B', 'category': 'Symbols'},
    {'name': 'Apl Functional Symbol Quad Circle', 'unicode': '233C', 'category': 'Symbols'},
    {'name': 'Apl Functional Symbol Circle Stile', 'unicode': '233D', 'category': 'Symbols'},
    {'name': 'Apl Functional Symbol Circle Jot', 'unicode': '233E', 'category': 'Symbols'},
    {'name': 'Apl Functional Symbol Slash Bar', 'unicode': '233F', 'category': 'Symbols'},
    {'name': 'Apl Functional Symbol Backslash Bar', 'unicode': '2340', 'category': 'Symbols'},
    {'name': 'Apl Functional Symbol Quad Slash', 'unicode': '2341', 'category': 'Symbols'},
    {'name': 'Apl Functional Symbol Quad Backslash', 'unicode': '2342', 'category': 'Symbols'},
    {'name': 'Apl Functional Symbol Quad Less-Than', 'unicode': '2343', 'category': 'Symbols'},
    {'name': 'Apl Functional Symbol Quad Greater-Than', 'unicode': '2344', 'category': 'Food & Drink'},
    {'name': 'Apl Functional Symbol Leftwards Vane', 'unicode': '2345', 'category': 'Symbols'},
    {'name': 'Apl Functional Symbol Rightwards Vane', 'unicode': '2346', 'category': 'Symbols'},
    {'name': 'Apl Functional Symbol Quad Leftwards Arrow', 'unicode': '2347', 'category': 'Symbols'},
    {'name': 'Apl Functional Symbol Quad Rightwards Arrow', 'unicode': '2348', 'category': 'Symbols'},
    {'name': 'Apl Functional Symbol Circle Backslash', 'unicode': '2349', 'category': 'Symbols'},
    {'name': 'Apl Functional Symbol Down Tack Underbar', 'unicode': '234A', 'category': 'Symbols'},
    {'name': 'Apl Functional Symbol Delta Stile', 'unicode': '234B', 'category': 'Symbols'},
    {'name': 'Apl Functional Symbol Quad Down Caret', 'unicode': '234C', 'category': 'Travel & Places'},
    {'name': 'Apl Functional Symbol Quad Delta', 'unicode': '234D', 'category': 'Symbols'},
    {'name': 'Apl Functional Symbol Down Tack Jot', 'unicode': '234E', 'category': 'Symbols'},
    {'name': 'Apl Functional Symbol Upwards Vane', 'unicode': '234F', 'category': 'Symbols'},
    {'name': 'Apl Functional Symbol Quad Upwards Arrow', 'unicode': '2350', 'category': 'Symbols'},
    {'name': 'Apl Functional Symbol Up Tack Overbar', 'unicode': '2351', 'category': 'Symbols'},
    {'name': 'Apl Functional Symbol Del Stile', 'unicode': '2352', 'category': 'Symbols'},
    {'name': 'Apl Functional Symbol Quad Up Caret', 'unicode': '2353', 'category': 'Travel & Places'},
    {'name': 'Apl Functional Symbol Quad Del', 'unicode': '2354', 'category': 'Symbols'},
    {'name': 'Apl Functional Symbol Up Tack Jot', 'unicode': '2355', 'category': 'Symbols'},
    {'name': 'Apl Functional Symbol Downwards Vane', 'unicode': '2356', 'category': 'Symbols'},
    {'name': 'Apl Functional Symbol Quad Downwards Arrow', 'unicode': '2357', 'category': 'Symbols'},
    {'name': 'Apl Functional Symbol Quote Underbar', 'unicode': '2358', 'category': 'Symbols'},
    {'name': 'Apl Functional Symbol Delta Underbar', 'unicode': '2359', 'category': 'Symbols'},
    {'name': 'Apl Functional Symbol Diamond Underbar', 'unicode': '235A', 'category': 'Symbols'},
    {'name': 'Apl Functional Symbol Jot Underbar', 'unicode': '235B', 'category': 'Symbols'},
    {'name': 'Apl Functional Symbol Circle Underbar', 'unicode': '235C', 'category': 'Symbols'},
    {'name': 'Apl Functional Symbol Up Shoe Jot', 'unicode': '235D', 'category': 'Symbols'},
    {'name': 'Apl Functional Symbol Quote Quad', 'unicode': '235E', 'category': 'Symbols'},
    {'name': 'Apl Functional Symbol Circle Star', 'unicode': '235F', 'category': 'Symbols'},
    {'name': 'Apl Functional Symbol Quad Colon', 'unicode': '2360', 'category': 'Symbols'},
    {'name': 'Apl Functional Symbol Up Tack Diaeresis', 'unicode': '2361', 'category': 'Symbols'},
    {'name': 'Apl Functional Symbol Del Diaeresis', 'unicode': '2362', 'category': 'Symbols'},
    {'name': 'Apl Functional Symbol Star Diaeresis', 'unicode': '2363', 'category': 'Symbols'},
    {'name': 'Apl Functional Symbol Jot Diaeresis', 'unicode': '2364', 'category': 'Symbols'},
    {'name': 'Apl Functional Symbol Circle Diaeresis', 'unicode': '2365', 'category': 'Symbols'},
    {'name': 'Apl Functional Symbol Down Shoe Stile', 'unicode': '2366', 'category': 'Symbols'},
    {'name': 'Apl Functional Symbol Left Shoe Stile', 'unicode': '2367', 'category': 'Symbols'},
    {'name': 'Apl Functional Symbol Tilde Diaeresis', 'unicode': '2368', 'category': 'Symbols'},
    {'name': 'Apl Functional Symbol Greater-Than Diaeresis', 'unicode': '2369', 'category': 'Food & Drink'},
    {'name': 'Apl Functional Symbol Comma Bar', 'unicode': '236A', 'category': 'Symbols'},
    {'name': 'Apl Functional Symbol Del Tilde', 'unicode': '236B', 'category': 'Symbols'},
    {'name': 'Apl Functional Symbol Zilde', 'unicode': '236C', 'category': 'Symbols'},
    {'name': 'Apl Functional Symbol Stile Tilde', 'unicode': '236D', 'category': 'Symbols'},
    {'name': 'Apl Functional Symbol Semicolon Underbar', 'unicode': '236E', 'category': 'Symbols'},
    {'name': 'Apl Functional Symbol Quad Not Equal', 'unicode': '236F', 'category': 'Symbols'},
    {'name': 'Apl Functional Symbol Quad Question', 'unicode': '2370', 'category': 'Symbols'},
    {'name': 'Apl Functional Symbol Down Caret Tilde', 'unicode': '2371', 'category': 'Travel & Places'},
    {'name': 'Apl Functional Symbol Up Caret Tilde', 'unicode': '2372', 'category': 'Travel & Places'},
    {'name': 'Apl Functional Symbol Iota', 'unicode': '2373', 'category': 'Symbols'},
    {'name': 'Apl Functional Symbol Rho', 'unicode': '2374', 'category': 'Symbols'},
    {'name': 'Apl Functional Symbol Omega', 'unicode': '2375', 'category': 'Symbols'},
    {'name': 'Apl Functional Symbol Alpha Underbar', 'unicode': '2376', 'category': 'Symbols'},
    {'name': 'Apl Functional Symbol Epsilon Underbar', 'unicode': '2377', 'category': 'Symbols'},
    {'name': 'Apl Functional Symbol Iota Underbar', 'unicode': '2378', 'category': 'Symbols'},
    {'name': 'Apl Functional Symbol Omega Underbar', 'unicode': '2379', 'category': 'Symbols'},
    {'name': 'Apl Functional Symbol Alpha', 'unicode': '237A', 'category': 'Symbols'},
    {'name': 'Not Check Mark', 'unicode': '237B', 'category': 'Symbols'},
    {'name': 'Right Angle With Downwards Zigzag Arrow', 'unicode': '237C', 'category': 'Symbols'},
    {'name': 'Shouldered Open Box', 'unicode': '237D', 'category': 'Symbols'},
    {'name': 'Bell Symbol', 'unicode': '237E', 'category': 'Symbols'},
    {'name': 'Vertical Line With Middle Dot', 'unicode': '237F', 'category': 'Symbols'},
    {'name': 'Insertion Symbol', 'unicode': '2380', 'category': 'Symbols'},
    {'name': 'Continuous Underline Symbol', 'unicode': '2381', 'category': 'Symbols'},
    {'name': 'Discontinuous Underline Symbol', 'unicode': '2382', 'category': 'Symbols'},
    {'name': 'Emphasis Symbol', 'unicode': '2383', 'category': 'Symbols'},
    {'name': 'Composition Symbol', 'unicode': '2384', 'category': 'Symbols'},
    {'name': 'White Square With Centre Vertical Line', 'unicode': '2385', 'category': 'Symbols'},
    {'name': 'Enter Symbol', 'unicode': '2386', 'category': 'Symbols'},
    {'name': 'Alternative Key Symbol', 'unicode': '2387', 'category': 'Symbols'},
    {'name': 'Helm Symbol', 'unicode': '2388', 'category': 'Symbols'},
    {'name': 'Circled Horizontal Bar With Notch', 'unicode': '2389', 'category': 'Symbols'},
    {'name': 'Circled Triangle Down', 'unicode': '238A', 'category': 'Symbols'},
    {'name': 'Broken Circle With Northwest Arrow', 'unicode': '238B', 'category': 'Symbols'},
    {'name': 'Undo Symbol', 'unicode': '238C', 'category': 'Symbols'},
    {'name': 'Monostable Symbol', 'unicode': '238D', 'category': 'Symbols'},
    {'name': 'Hysteresis Symbol', 'unicode': '238E', 'category': 'Symbols'},
    {'name': 'Open-Circuit-Output H-Type Symbol', 'unicode': '238F', 'category': 'Symbols'},
    {'name': 'Open-Circuit-Output L-Type Symbol', 'unicode': '2390', 'category': 'Symbols'},
    {'name': 'Passive-Pull-Down-Output Symbol', 'unicode': '2391', 'category': 'Symbols'},
    {'name': 'Passive-Pull-Up-Output Symbol', 'unicode': '2392', 'category': 'Symbols'},
    {'name': 'Direct Current Symbol Form Two', 'unicode': '2393', 'category': 'Symbols'},
    {'name': 'Software-Function Symbol', 'unicode': '2394', 'category': 'Symbols'},
    {'name': 'Apl Functional Symbol Quad', 'unicode': '2395', 'category': 'Symbols'},
    {'name': 'Decimal Separator Key Symbol', 'unicode': '2396', 'category': 'Symbols'},
    {'name': 'Previous Page', 'unicode': '2397', 'category': 'Symbols'},
    {'name': 'Next Page', 'unicode': '2398', 'category': 'Symbols'},
    {'name': 'Print Screen Symbol', 'unicode': '2399', 'category': 'Symbols'},
    {'name': 'Clear Screen Symbol', 'unicode': '239A', 'category': 'Symbols'},
    {'name': 'Left Parenthesis Upper Hook', 'unicode': '239B', 'category': 'Symbols'},
    {'name': 'Left Parenthesis Extension', 'unicode': '239C', 'category': 'Symbols'},
    {'name': 'Left Parenthesis Lower Hook', 'unicode': '239D', 'category': 'Symbols'},
    {'name': 'Right Parenthesis Upper Hook', 'unicode': '239E', 'category': 'Symbols'},
    {'name': 'Right Parenthesis Extension', 'unicode': '239F', 'category': 'Symbols'},
    {'name': 'Right Parenthesis Lower Hook', 'unicode': '23A0', 'category': 'Symbols'},
    {'name': 'Left Square Bracket Upper Corner', 'unicode': '23A1', 'category': 'Symbols'},
    {'name': 'Left Square Bracket Extension', 'unicode': '23A2', 'category': 'Symbols'},
    {'name': 'Left Square Bracket Lower Corner', 'unicode': '23A3', 'category': 'Symbols'},
    {'name': 'Right Square Bracket Upper Corner', 'unicode': '23A4', 'category': 'Symbols'},
    {'name': 'Right Square Bracket Extension', 'unicode': '23A5', 'category': 'Symbols'},
    {'name': 'Right Square Bracket Lower Corner', 'unicode': '23A6', 'category': 'Symbols'},
    {'name': 'Left Curly Bracket Upper Hook', 'unicode': '23A7', 'category': 'Symbols'},
    {'name': 'Left Curly Bracket Middle Piece', 'unicode': '23A8', 'category': 'Symbols'},
    {'name': 'Left Curly Bracket Lower Hook', 'unicode': '23A9', 'category': 'Symbols'},
    {'name': 'Curly Bracket Extension', 'unicode': '23AA', 'category': 'Symbols'},
    {'name': 'Right Curly Bracket Upper Hook', 'unicode': '23AB', 'category': 'Symbols'},
    {'name': 'Right Curly Bracket Middle Piece', 'unicode': '23AC', 'category': 'Symbols'},
    {'name': 'Right Curly Bracket Lower Hook', 'unicode': '23AD', 'category': 'Symbols'},
    {'name': 'Integral Extension', 'unicode': '23AE', 'category': 'Symbols'},
    {'name': 'Horizontal Line Extension', 'unicode': '23AF', 'category': 'Symbols'},
    {'name': 'Upper Left Or Lower Right Curly Bracket Section', 'unicode': '23B0', 'category': 'Symbols'},
    {'name': 'Upper Right Or Lower Left Curly Bracket Section', 'unicode': '23B1', 'category': 'Symbols'},
    {'name': 'Summation Top', 'unicode': '23B2', 'category': 'Symbols'},
    {'name': 'Summation Bottom', 'unicode': '23B3', 'category': 'Symbols'},
    {'name': 'Top Square Bracket', 'unicode': '23B4', 'category': 'Symbols'},
    {'name': 'Bottom Square Bracket', 'unicode': '23B5', 'category': 'Symbols'},
    {'name': 'Bottom Square Bracket Over Top Square Bracket', 'unicode': '23B6', 'category': 'Symbols'},
    {'name': 'Radical Symbol Bottom', 'unicode': '23B7', 'category': 'Symbols'},
    {'name': 'Left Vertical Box Line', 'unicode': '23B8', 'category': 'Symbols'},
    {'name': 'Right Vertical Box Line', 'unicode': '23B9', 'category': 'Symbols'},
    {'name': 'Horizontal Scan Line-1', 'unicode': '23BA', 'category': 'Symbols'},
    {'name': 'Horizontal Scan Line-3', 'unicode': '23BB', 'category': 'Symbols'},
    {'name': 'Horizontal Scan Line-7', 'unicode': '23BC', 'category': 'Symbols'},
    {'name': 'Horizontal Scan Line-9', 'unicode': '23BD', 'category': 'Symbols'},
    {'name': 'Dentistry Symbol Light Vertical And Top Right', 'unicode': '23BE', 'category': 'Symbols'},
    {'name': 'Dentistry Symbol Light Vertical And Bottom Right', 'unicode': '23BF', 'category': 'Symbols'},
    {'name': 'Dentistry Symbol Light Vertical With Circle', 'unicode': '23C0', 'category': 'Symbols'},
    {'name': 'Dentistry Symbol Light Down And Horizontal With Circle', 'unicode': '23C1', 'category': 'Symbols'},
    {'name': 'Dentistry Symbol Light Up And Horizontal With Circle', 'unicode': '23C2', 'category': 'Symbols'},
    {'name': 'Dentistry Symbol Light Vertical With Triangle', 'unicode': '23C3', 'category': 'Symbols'},
    {'name': 'Dentistry Symbol Light Down And Horizontal With Triangle', 'unicode': '23C4', 'category': 'Symbols'},
    {'name': 'Dentistry Symbol Light Up And Horizontal With Triangle', 'unicode': '23C5', 'category': 'Symbols'},
    {'name': 'Dentistry Symbol Light Vertical And Wave', 'unicode': '23C6', 'category': 'Symbols'},
    {'name': 'Dentistry Symbol Light Down And Horizontal With Wave', 'unicode': '23C7', 'category': 'Symbols'},
    {'name': 'Dentistry Symbol Light Up And Horizontal With Wave', 'unicode': '23C8', 'category': 'Symbols'},
    {'name': 'Dentistry Symbol Light Down And Horizontal', 'unicode': '23C9', 'category': 'Symbols'},
    {'name': 'Dentistry Symbol Light Up And Horizontal', 'unicode': '23CA', 'category': 'Symbols'},
    {'name': 'Dentistry Symbol Light Vertical And Top Left', 'unicode': '23CB', 'category': 'Symbols'},
    {'name': 'Dentistry Symbol Light Vertical And Bottom Left', 'unicode': '23CC', 'category': 'Symbols'},
    {'name': 'Square Foot', 'unicode': '23CD', 'category': 'Symbols'},
    {'name': 'Return Symbol', 'unicode': '23CE', 'category': 'Symbols'},
    {'name': 'Eject Symbol', 'unicode': '23CF', 'category': 'Symbols'},
    {'name': 'Vertical Line Extension', 'unicode': '23D0', 'category': 'Symbols'},
    {'name': 'Metrical Breve', 'unicode': '23D1', 'category': 'Symbols'},
    {'name': 'Metrical Long Over Short', 'unicode': '23D2', 'category': 'Symbols'},
    {'name': 'Metrical Short Over Long', 'unicode': '23D3', 'category': 'Symbols'},
    {'name': 'Metrical Long Over Two Shorts', 'unicode': '23D4', 'category': 'Symbols'},
    {'name': 'Metrical Two Shorts Over Long', 'unicode': '23D5', 'category': 'Symbols'},
    {'name': 'Metrical Two Shorts Joined', 'unicode': '23D6', 'category': 'Symbols'},
    {'name': 'Metrical Triseme', 'unicode': '23D7', 'category': 'Symbols'},
    {'name': 'Metrical Tetraseme', 'unicode': '23D8', 'category': 'Symbols'},
    {'name': 'Metrical Pentaseme', 'unicode': '23D9', 'category': 'Symbols'},
	{'name': 'Earth Ground', 'unicode': '23DA', 'category': 'Symbols'},
    {'name': 'Fuse', 'unicode': '23DB', 'category': 'Symbols'},
    {'name': 'Top Parenthesis', 'unicode': '23DC', 'category': 'Symbols'},
    {'name': 'Bottom Parenthesis', 'unicode': '23DD', 'category': 'Symbols'},
    {'name': 'Top Curly Bracket', 'unicode': '23DE', 'category': 'Symbols'},
    {'name': 'Bottom Curly Bracket', 'unicode': '23DF', 'category': 'Symbols'},
    {'name': 'Top Tortoise Shell Bracket', 'unicode': '23E0', 'category': 'Symbols'},
    {'name': 'Bottom Tortoise Shell Bracket', 'unicode': '23E1', 'category': 'Symbols'},
    {'name': 'White Trapezium', 'unicode': '23E2', 'category': 'Symbols'},
    {'name': 'Benzene Ring With Circle', 'unicode': '23E3', 'category': 'Symbols'},
    {'name': 'Straightness', 'unicode': '23E4', 'category': 'Symbols'},
    {'name': 'Flatness', 'unicode': '23E5', 'category': 'Symbols'},
    {'name': 'Ac Current', 'unicode': '23E6', 'category': 'Symbols'},
    {'name': 'Electrical Intersection', 'unicode': '23E7', 'category': 'Symbols'},
    {'name': 'Decimal Exponent Symbol', 'unicode': '23E8', 'category': 'Symbols'},
    {'name': 'Black Right-Pointing Double Triangle', 'unicode': '23E9', 'category': 'Symbols'},
    {'name': 'Black Left-Pointing Double Triangle', 'unicode': '23EA', 'category': 'Symbols'},
    {'name': 'Black Up-Pointing Double Triangle', 'unicode': '23EB', 'category': 'Symbols'},
    {'name': 'Black Down-Pointing Double Triangle', 'unicode': '23EC', 'category': 'Symbols'},
    {'name': 'Black Right-Pointing Double Triangle With Vertical Bar', 'unicode': '23ED', 'category': 'Symbols'},
    {'name': 'Black Left-Pointing Double Triangle With Vertical Bar', 'unicode': '23EE', 'category': 'Symbols'},
    {'name': 'Black Right-Pointing Triangle With Double Vertical Bar', 'unicode': '23EF', 'category': 'Symbols'},
    {'name': 'Alarm Clock', 'unicode': '23F0', 'category': 'Symbols'},
    {'name': 'Stopwatch', 'unicode': '23F1', 'category': 'Symbols'},
    {'name': 'Timer Clock', 'unicode': '23F2', 'category': 'Symbols'},
    {'name': 'Hourglass With Flowing Sand', 'unicode': '23F3', 'category': 'Symbols'},
    {'name': 'Black Medium Left-Pointing Triangle', 'unicode': '23F4', 'category': 'Symbols'},
    {'name': 'Black Medium Right-Pointing Triangle', 'unicode': '23F5', 'category': 'Symbols'},
    {'name': 'Black Medium Up-Pointing Triangle', 'unicode': '23F6', 'category': 'Symbols'},
    {'name': 'Black Medium Down-Pointing Triangle', 'unicode': '23F7', 'category': 'Symbols'},
    {'name': 'Double Vertical Bar', 'unicode': '23F8', 'category': 'Symbols'},
    {'name': 'Black Square For Stop', 'unicode': '23F9', 'category': 'Symbols'},
    {'name': 'Black Circle For Record', 'unicode': '23FA', 'category': 'Symbols'},
    {'name': 'Power Symbol', 'unicode': '23FB', 'category': 'Symbols'},
    {'name': 'Power On-Off Symbol', 'unicode': '23FC', 'category': 'Symbols'},
    {'name': 'Power On Symbol', 'unicode': '23FD', 'category': 'Symbols'},
    {'name': 'Power Sleep Symbol', 'unicode': '23FE', 'category': 'Symbols'},
    {'name': 'Observer Eye Symbol', 'unicode': '23FF', 'category': 'Symbols'},
    {'name': 'Black Square', 'unicode': '25A0', 'category': 'Symbols'},
    {'name': 'White Square', 'unicode': '25A1', 'category': 'Symbols'},
    {'name': 'White Square With Rounded Corners', 'unicode': '25A2', 'category': 'Symbols'},
    {'name': 'White Square Containing Black Small Square', 'unicode': '25A3', 'category': 'Symbols'},
    {'name': 'Square With Horizontal Fill', 'unicode': '25A4', 'category': 'Symbols'},
    {'name': 'Square With Vertical Fill', 'unicode': '25A5', 'category': 'Symbols'},
    {'name': 'Square With Orthogonal Crosshatch Fill', 'unicode': '25A6', 'category': 'Symbols'},
    {'name': 'Square With Upper Left To Lower Right Fill', 'unicode': '25A7', 'category': 'Symbols'},
    {'name': 'Square With Upper Right To Lower Left Fill', 'unicode': '25A8', 'category': 'Symbols'},
    {'name': 'Square With Diagonal Crosshatch Fill', 'unicode': '25A9', 'category': 'Symbols'},
    {'name': 'Black Small Square', 'unicode': '25AA', 'category': 'Symbols'},
    {'name': 'White Small Square', 'unicode': '25AB', 'category': 'Symbols'},
    {'name': 'Black Rectangle', 'unicode': '25AC', 'category': 'Symbols'},
    {'name': 'White Rectangle', 'unicode': '25AD', 'category': 'Symbols'},
    {'name': 'Black Vertical Rectangle', 'unicode': '25AE', 'category': 'Symbols'},
    {'name': 'White Vertical Rectangle', 'unicode': '25AF', 'category': 'Symbols'},
    {'name': 'Black Parallelogram', 'unicode': '25B0', 'category': 'Symbols'},
    {'name': 'White Parallelogram', 'unicode': '25B1', 'category': 'Symbols'},
    {'name': 'Black Up-Pointing Triangle', 'unicode': '25B2', 'category': 'Symbols'},
    {'name': 'White Up-Pointing Triangle', 'unicode': '25B3', 'category': 'Symbols'},
    {'name': 'Black Up-Pointing Small Triangle', 'unicode': '25B4', 'category': 'Symbols'},
    {'name': 'White Up-Pointing Small Triangle', 'unicode': '25B5', 'category': 'Symbols'},
    {'name': 'Black Right-Pointing Triangle', 'unicode': '25B6', 'category': 'Symbols'},
    {'name': 'White Right-Pointing Triangle', 'unicode': '25B7', 'category': 'Symbols'},
    {'name': 'Black Right-Pointing Small Triangle', 'unicode': '25B8', 'category': 'Symbols'},
    {'name': 'White Right-Pointing Small Triangle', 'unicode': '25B9', 'category': 'Symbols'},
    {'name': 'Black Right-Pointing Pointer', 'unicode': '25BA', 'category': 'Symbols'},
    {'name': 'White Right-Pointing Pointer', 'unicode': '25BB', 'category': 'Symbols'},
    {'name': 'Black Down-Pointing Triangle', 'unicode': '25BC', 'category': 'Symbols'},
    {'name': 'White Down-Pointing Triangle', 'unicode': '25BD', 'category': 'Symbols'},
    {'name': 'Black Down-Pointing Small Triangle', 'unicode': '25BE', 'category': 'Symbols'},
    {'name': 'White Down-Pointing Small Triangle', 'unicode': '25BF', 'category': 'Symbols'},
    {'name': 'Black Left-Pointing Triangle', 'unicode': '25C0', 'category': 'Symbols'},
    {'name': 'White Left-Pointing Triangle', 'unicode': '25C1', 'category': 'Symbols'},
    {'name': 'Black Left-Pointing Small Triangle', 'unicode': '25C2', 'category': 'Symbols'},
    {'name': 'White Left-Pointing Small Triangle', 'unicode': '25C3', 'category': 'Symbols'},
    {'name': 'Black Left-Pointing Pointer', 'unicode': '25C4', 'category': 'Symbols'},
    {'name': 'White Left-Pointing Pointer', 'unicode': '25C5', 'category': 'Symbols'},
    {'name': 'Black Diamond', 'unicode': '25C6', 'category': 'Symbols'},
    {'name': 'White Diamond', 'unicode': '25C7', 'category': 'Symbols'},
    {'name': 'White Diamond Containing Black Small Diamond', 'unicode': '25C8', 'category': 'Symbols'},
    {'name': 'Fisheye', 'unicode': '25C9', 'category': 'Symbols'},
    {'name': 'Lozenge', 'unicode': '25CA', 'category': 'Symbols'},
    {'name': 'White Circle', 'unicode': '25CB', 'category': 'Symbols'},
    {'name': 'Dotted Circle', 'unicode': '25CC', 'category': 'Symbols'},
    {'name': 'Circle With Vertical Fill', 'unicode': '25CD', 'category': 'Symbols'},
    {'name': 'Bullseye', 'unicode': '25CE', 'category': 'Symbols'},
    {'name': 'Black Circle', 'unicode': '25CF', 'category': 'Symbols'},
    {'name': 'Circle With Left Half Black', 'unicode': '25D0', 'category': 'Symbols'},
    {'name': 'Circle With Right Half Black', 'unicode': '25D1', 'category': 'Symbols'},
    {'name': 'Circle With Lower Half Black', 'unicode': '25D2', 'category': 'Symbols'},
    {'name': 'Circle With Upper Half Black', 'unicode': '25D3', 'category': 'Symbols'},
    {'name': 'Circle With Upper Right Quadrant Black', 'unicode': '25D4', 'category': 'Symbols'},
    {'name': 'Circle With All But Upper Left Quadrant Black', 'unicode': '25D5', 'category': 'Symbols'},
    {'name': 'Left Half Black Circle', 'unicode': '25D6', 'category': 'Symbols'},
    {'name': 'Right Half Black Circle', 'unicode': '25D7', 'category': 'Symbols'},
    {'name': 'Inverse Bullet', 'unicode': '25D8', 'category': 'Symbols'},
    {'name': 'Inverse White Circle', 'unicode': '25D9', 'category': 'Symbols'},
    {'name': 'Upper Half Inverse White Circle', 'unicode': '25DA', 'category': 'Symbols'},
    {'name': 'Lower Half Inverse White Circle', 'unicode': '25DB', 'category': 'Symbols'},
    {'name': 'Upper Left Quadrant Circular Arc', 'unicode': '25DC', 'category': 'Symbols'},
    {'name': 'Upper Right Quadrant Circular Arc', 'unicode': '25DD', 'category': 'Symbols'},
    {'name': 'Lower Right Quadrant Circular Arc', 'unicode': '25DE', 'category': 'Symbols'},
    {'name': 'Lower Left Quadrant Circular Arc', 'unicode': '25DF', 'category': 'Symbols'},
    {'name': 'Upper Half Circle', 'unicode': '25E0', 'category': 'Symbols'},
    {'name': 'Lower Half Circle', 'unicode': '25E1', 'category': 'Symbols'},
    {'name': 'Black Lower Right Triangle', 'unicode': '25E2', 'category': 'Symbols'},
    {'name': 'Black Lower Left Triangle', 'unicode': '25E3', 'category': 'Symbols'},
    {'name': 'Black Upper Left Triangle', 'unicode': '25E4', 'category': 'Symbols'},
    {'name': 'Black Upper Right Triangle', 'unicode': '25E5', 'category': 'Symbols'},
    {'name': 'White Bullet', 'unicode': '25E6', 'category': 'Symbols'},
    {'name': 'Square With Left Half Black', 'unicode': '25E7', 'category': 'Symbols'},
    {'name': 'Square With Right Half Black', 'unicode': '25E8', 'category': 'Symbols'},
    {'name': 'Square With Upper Left Diagonal Half Black', 'unicode': '25E9', 'category': 'Symbols'},
    {'name': 'Square With Lower Right Diagonal Half Black', 'unicode': '25EA', 'category': 'Symbols'},
    {'name': 'White Square With Vertical Bisecting Line', 'unicode': '25EB', 'category': 'Symbols'},
    {'name': 'White Up-Pointing Triangle With Dot', 'unicode': '25EC', 'category': 'Symbols'},
    {'name': 'Up-Pointing Triangle With Left Half Black', 'unicode': '25ED', 'category': 'Symbols'},
    {'name': 'Up-Pointing Triangle With Right Half Black', 'unicode': '25EE', 'category': 'Symbols'},
    {'name': 'Large Circle', 'unicode': '25EF', 'category': 'Symbols'},
    {'name': 'White Square With Upper Left Quadrant', 'unicode': '25F0', 'category': 'Symbols'},
    {'name': 'White Square With Lower Left Quadrant', 'unicode': '25F1', 'category': 'Symbols'},
    {'name': 'White Square With Lower Right Quadrant', 'unicode': '25F2', 'category': 'Symbols'},
    {'name': 'White Square With Upper Right Quadrant', 'unicode': '25F3', 'category': 'Symbols'},
    {'name': 'White Circle With Upper Left Quadrant', 'unicode': '25F4', 'category': 'Symbols'},
    {'name': 'White Circle With Lower Left Quadrant', 'unicode': '25F5', 'category': 'Symbols'},
    {'name': 'White Circle With Lower Right Quadrant', 'unicode': '25F6', 'category': 'Symbols'},
    {'name': 'White Circle With Upper Right Quadrant', 'unicode': '25F7', 'category': 'Symbols'},
    {'name': 'Upper Left Triangle', 'unicode': '25F8', 'category': 'Symbols'},
    {'name': 'Upper Right Triangle', 'unicode': '25F9', 'category': 'Symbols'},
    {'name': 'Lower Left Triangle', 'unicode': '25FA', 'category': 'Symbols'},
    {'name': 'White Medium Square', 'unicode': '25FB', 'category': 'Symbols'},
    {'name': 'Black Medium Square', 'unicode': '25FC', 'category': 'Symbols'},
    {'name': 'White Medium Small Square', 'unicode': '25FD', 'category': 'Symbols'},
    {'name': 'Black Medium Small Square', 'unicode': '25FE', 'category': 'Symbols'},
    {'name': 'Lower Right Triangle', 'unicode': '25FF', 'category': 'Symbols'},
    {'name': 'Black Sun With Rays', 'unicode': '2600', 'category': 'Symbols'},
    {'name': 'Cloud', 'unicode': '2601', 'category': 'Symbols'},
    {'name': 'Umbrella', 'unicode': '2602', 'category': 'Symbols'},
    {'name': 'Snowman', 'unicode': '2603', 'category': 'Symbols'},
    {'name': 'Comet', 'unicode': '2604', 'category': 'Symbols'},
    {'name': 'Black Star', 'unicode': '2605', 'category': 'Symbols'},
    {'name': 'White Star', 'unicode': '2606', 'category': 'Symbols'},
    {'name': 'Lightning', 'unicode': '2607', 'category': 'Symbols'},
    {'name': 'Thunderstorm', 'unicode': '2608', 'category': 'Symbols'},
    {'name': 'Sun', 'unicode': '2609', 'category': 'Symbols'},
    {'name': 'Ascending Node', 'unicode': '260A', 'category': 'Symbols'},
    {'name': 'Descending Node', 'unicode': '260B', 'category': 'Symbols'},
    {'name': 'Conjunction', 'unicode': '260C', 'category': 'Symbols'},
    {'name': 'Opposition', 'unicode': '260D', 'category': 'Symbols'},
    {'name': 'Black Telephone', 'unicode': '260E', 'category': 'Symbols'},
    {'name': 'White Telephone', 'unicode': '260F', 'category': 'Symbols'},
    {'name': 'Ballot Box', 'unicode': '2610', 'category': 'Symbols'},
    {'name': 'Ballot Box With Check', 'unicode': '2611', 'category': 'Symbols'},
    {'name': 'Ballot Box With X', 'unicode': '2612', 'category': 'Symbols'},
    {'name': 'Saltire', 'unicode': '2613', 'category': 'Symbols'},
    {'name': 'Umbrella With Rain Drops', 'unicode': '2614', 'category': 'Symbols'},
    {'name': 'Hot Beverage', 'unicode': '2615', 'category': 'Food & Drink'},
    {'name': 'White Shogi Piece', 'unicode': '2616', 'category': 'Symbols'},
    {'name': 'Black Shogi Piece', 'unicode': '2617', 'category': 'Symbols'},
    {'name': 'Shamrock', 'unicode': '2618', 'category': 'Symbols'},
    {'name': 'Reversed Rotated Floral Heart Bullet', 'unicode': '2619', 'category': 'Symbols'},
    {'name': 'Black Left Pointing Index', 'unicode': '261A', 'category': 'Symbols'},
    {'name': 'Black Right Pointing Index', 'unicode': '261B', 'category': 'Symbols'},
    {'name': 'White Left Pointing Index', 'unicode': '261C', 'category': 'Symbols'},
    {'name': 'White Up Pointing Index', 'unicode': '261D', 'category': 'Symbols'},
    {'name': 'White Right Pointing Index', 'unicode': '261E', 'category': 'Symbols'},
    {'name': 'White Down Pointing Index', 'unicode': '261F', 'category': 'Symbols'},
    {'name': 'Skull And Crossbones', 'unicode': '2620', 'category': 'Symbols'},
    {'name': 'Caution Sign', 'unicode': '2621', 'category': 'Symbols'},
    {'name': 'Radioactive Sign', 'unicode': '2622', 'category': 'Symbols'},
    {'name': 'Biohazard Sign', 'unicode': '2623', 'category': 'Symbols'},
    {'name': 'Caduceus', 'unicode': '2624', 'category': 'Symbols'},
    {'name': 'Ankh', 'unicode': '2625', 'category': 'Symbols'},
    {'name': 'Orthodox Cross', 'unicode': '2626', 'category': 'Symbols'},
    {'name': 'Chi Rho', 'unicode': '2627', 'category': 'Symbols'},
    {'name': 'Cross Of Lorraine', 'unicode': '2628', 'category': 'Symbols'},
    {'name': 'Cross Of Jerusalem', 'unicode': '2629', 'category': 'Symbols'},
    {'name': 'Star And Crescent', 'unicode': '262A', 'category': 'Symbols'},
    {'name': 'Farsi Symbol', 'unicode': '262B', 'category': 'Symbols'},
    {'name': 'Adi Shakti', 'unicode': '262C', 'category': 'Symbols'},
    {'name': 'Hammer And Sickle', 'unicode': '262D', 'category': 'Symbols'},
    {'name': 'Peace Symbol', 'unicode': '262E', 'category': 'Symbols'},
    {'name': 'Yin Yang', 'unicode': '262F', 'category': 'Symbols'},
    {'name': 'Trigram For Heaven', 'unicode': '2630', 'category': 'Smileys & Emotion'},
    {'name': 'Trigram For Lake', 'unicode': '2631', 'category': 'Smileys & Emotion'},
    {'name': 'Trigram For Fire', 'unicode': '2632', 'category': 'Smileys & Emotion'},
    {'name': 'Trigram For Thunder', 'unicode': '2633', 'category': 'Smileys & Emotion'},
    {'name': 'Trigram For Wind', 'unicode': '2634', 'category': 'Smileys & Emotion'},
    {'name': 'Trigram For Water', 'unicode': '2635', 'category': 'Smileys & Emotion'},
    {'name': 'Trigram For Mountain', 'unicode': '2636', 'category': 'Smileys & Emotion'},
    {'name': 'Trigram For Earth', 'unicode': '2637', 'category': 'Smileys & Emotion'},
    {'name': 'Wheel Of Dharma', 'unicode': '2638', 'category': 'Smileys & Emotion'},
    {'name': 'White Frowning Face', 'unicode': '2639', 'category': 'Smileys & Emotion'},
    {'name': 'White Smiling Face', 'unicode': '263A', 'category': 'Smileys & Emotion'},
    {'name': 'Black Smiling Face', 'unicode': '263B', 'category': 'Smileys & Emotion'},
    {'name': 'White Sun With Rays', 'unicode': '263C', 'category': 'Smileys & Emotion'},
    {'name': 'First Quarter Moon', 'unicode': '263D', 'category': 'Smileys & Emotion'},
    {'name': 'Last Quarter Moon', 'unicode': '263E', 'category': 'Smileys & Emotion'},
    {'name': 'Mercury', 'unicode': '263F', 'category': 'Smileys & Emotion'},
    {'name': 'Female Sign', 'unicode': '2640', 'category': 'Symbols'},
    {'name': 'Earth', 'unicode': '2641', 'category': 'Symbols'},
    {'name': 'Male Sign', 'unicode': '2642', 'category': 'Symbols'},
    {'name': 'Jupiter', 'unicode': '2643', 'category': 'Symbols'},
    {'name': 'Saturn', 'unicode': '2644', 'category': 'Symbols'},
    {'name': 'Uranus', 'unicode': '2645', 'category': 'Symbols'},
    {'name': 'Neptune', 'unicode': '2646', 'category': 'Symbols'},
    {'name': 'Pluto', 'unicode': '2647', 'category': 'Symbols'},
    {'name': 'Aries', 'unicode': '2648', 'category': 'Symbols'},
    {'name': 'Taurus', 'unicode': '2649', 'category': 'Symbols'},
    {'name': 'Gemini', 'unicode': '264A', 'category': 'Symbols'},
    {'name': 'Cancer', 'unicode': '264B', 'category': 'Symbols'},
    {'name': 'Leo', 'unicode': '264C', 'category': 'Symbols'},
    {'name': 'Virgo', 'unicode': '264D', 'category': 'Symbols'},
    {'name': 'Libra', 'unicode': '264E', 'category': 'Symbols'},
    {'name': 'Scorpius', 'unicode': '264F', 'category': 'Symbols'},
    {'name': 'Sagittarius', 'unicode': '2650', 'category': 'Symbols'},
    {'name': 'Capricorn', 'unicode': '2651', 'category': 'Symbols'},
    {'name': 'Aquarius', 'unicode': '2652', 'category': 'Symbols'},
    {'name': 'Pisces', 'unicode': '2653', 'category': 'Symbols'},
    {'name': 'White Chess King', 'unicode': '2654', 'category': 'Symbols'},
    {'name': 'White Chess Queen', 'unicode': '2655', 'category': 'Symbols'},
    {'name': 'White Chess Rook', 'unicode': '2656', 'category': 'Symbols'},
    {'name': 'White Chess Bishop', 'unicode': '2657', 'category': 'Symbols'},
    {'name': 'White Chess Knight', 'unicode': '2658', 'category': 'Symbols'},
    {'name': 'White Chess Pawn', 'unicode': '2659', 'category': 'Symbols'},
    {'name': 'Black Chess King', 'unicode': '265A', 'category': 'Symbols'},
    {'name': 'Black Chess Queen', 'unicode': '265B', 'category': 'Symbols'},
    {'name': 'Black Chess Rook', 'unicode': '265C', 'category': 'Symbols'},
    {'name': 'Black Chess Bishop', 'unicode': '265D', 'category': 'Symbols'},
    {'name': 'Black Chess Knight', 'unicode': '265E', 'category': 'Symbols'},
    {'name': 'Black Chess Pawn', 'unicode': '265F', 'category': 'Symbols'},
    {'name': 'Black Spade Suit', 'unicode': '2660', 'category': 'Symbols'},
    {'name': 'White Heart Suit', 'unicode': '2661', 'category': 'Symbols'},
    {'name': 'White Diamond Suit', 'unicode': '2662', 'category': 'Symbols'},
    {'name': 'Black Club Suit', 'unicode': '2663', 'category': 'Symbols'},
    {'name': 'White Spade Suit', 'unicode': '2664', 'category': 'Symbols'},
    {'name': 'Black Heart Suit', 'unicode': '2665', 'category': 'Symbols'},
    {'name': 'Black Diamond Suit', 'unicode': '2666', 'category': 'Symbols'},
    {'name': 'White Club Suit', 'unicode': '2667', 'category': 'Symbols'},
    {'name': 'Hot Springs', 'unicode': '2668', 'category': 'Symbols'},
    {'name': 'Quarter Note', 'unicode': '2669', 'category': 'Symbols'},
    {'name': 'Eighth Note', 'unicode': '266A', 'category': 'Symbols'},
    {'name': 'Beamed Eighth Notes', 'unicode': '266B', 'category': 'Symbols'},
    {'name': 'Beamed Sixteenth Notes', 'unicode': '266C', 'category': 'Symbols'},
    {'name': 'Music Flat Sign', 'unicode': '266D', 'category': 'Symbols'},
    {'name': 'Music Natural Sign', 'unicode': '266E', 'category': 'Symbols'},
    {'name': 'Music Sharp Sign', 'unicode': '266F', 'category': 'Symbols'},
    {'name': 'West Syriac Cross', 'unicode': '2670', 'category': 'Symbols'},
    {'name': 'East Syriac Cross', 'unicode': '2671', 'category': 'Symbols'},
    {'name': 'Universal Recycling Symbol', 'unicode': '2672', 'category': 'Symbols'},
    {'name': 'Recycling Symbol For Type-1 Plastics', 'unicode': '2673', 'category': 'Symbols'},
    {'name': 'Recycling Symbol For Type-2 Plastics', 'unicode': '2674', 'category': 'Symbols'},
    {'name': 'Recycling Symbol For Type-3 Plastics', 'unicode': '2675', 'category': 'Symbols'},
    {'name': 'Recycling Symbol For Type-4 Plastics', 'unicode': '2676', 'category': 'Symbols'},
    {'name': 'Recycling Symbol For Type-5 Plastics', 'unicode': '2677', 'category': 'Symbols'},
    {'name': 'Recycling Symbol For Type-6 Plastics', 'unicode': '2678', 'category': 'Symbols'},
    {'name': 'Recycling Symbol For Type-7 Plastics', 'unicode': '2679', 'category': 'Symbols'},
    {'name': 'Recycling Symbol For Generic Materials', 'unicode': '267A', 'category': 'Symbols'},
    {'name': 'Black Universal Recycling Symbol', 'unicode': '267B', 'category': 'Symbols'},
    {'name': 'Recycled Paper Symbol', 'unicode': '267C', 'category': 'Symbols'},
    {'name': 'Partially-Recycled Paper Symbol', 'unicode': '267D', 'category': 'Symbols'},
    {'name': 'Permanent Paper Sign', 'unicode': '267E', 'category': 'Symbols'},
    {'name': 'Wheelchair Symbol', 'unicode': '267F', 'category': 'Symbols'},
    {'name': 'Die Face-1', 'unicode': '2680', 'category': 'Smileys & Emotion'},
    {'name': 'Die Face-2', 'unicode': '2681', 'category': 'Smileys & Emotion'},
    {'name': 'Die Face-3', 'unicode': '2682', 'category': 'Smileys & Emotion'},
    {'name': 'Die Face-4', 'unicode': '2683', 'category': 'Smileys & Emotion'},
    {'name': 'Die Face-5', 'unicode': '2684', 'category': 'Smileys & Emotion'},
    {'name': 'Die Face-6', 'unicode': '2685', 'category': 'Smileys & Emotion'},
    {'name': 'White Circle With Dot Right', 'unicode': '2686', 'category': 'Symbols'}
]

# Add rest of emojis (continuing from the first part)
BUILT_IN_EMOJI_DATA.extend([
    {'name': 'White Circle With Two Dots', 'unicode': '2687', 'category': 'Symbols'},
    {'name': 'Black Circle With White Dot Right', 'unicode': '2688', 'category': 'Symbols'},
    {'name': 'Black Circle With Two White Dots', 'unicode': '2689', 'category': 'Symbols'},
    {'name': 'Monogram For Yang', 'unicode': '268A', 'category': 'Symbols'},
    {'name': 'Monogram For Yin', 'unicode': '268B', 'category': 'Symbols'},
    {'name': 'Digram For Greater Yang', 'unicode': '268C', 'category': 'Food & Drink'},
    {'name': 'Digram For Lesser Yin', 'unicode': '268D', 'category': 'Symbols'},
    {'name': 'Digram For Lesser Yang', 'unicode': '268E', 'category': 'Symbols'},
    {'name': 'Digram For Greater Yin', 'unicode': '268F', 'category': 'Food & Drink'},
    {'name': 'White Flag', 'unicode': '2690', 'category': 'Symbols'},
    {'name': 'Black Flag', 'unicode': '2691', 'category': 'Symbols'},
    {'name': 'Hammer And Pick', 'unicode': '2692', 'category': 'Symbols'},
    {'name': 'Anchor', 'unicode': '2693', 'category': 'Symbols'},
    {'name': 'Crossed Swords', 'unicode': '2694', 'category': 'Symbols'},
    {'name': 'Staff Of Aesculapius', 'unicode': '2695', 'category': 'Symbols'},
    {'name': 'Scales', 'unicode': '2696', 'category': 'Symbols'},
    {'name': 'Alembic', 'unicode': '2697', 'category': 'Symbols'},
    {'name': 'Flower', 'unicode': '2698', 'category': 'Animals & Nature'},
    {'name': 'Gear', 'unicode': '2699', 'category': 'Symbols'},
    {'name': 'Staff Of Hermes', 'unicode': '269A', 'category': 'Symbols'},
    {'name': 'Atom Symbol', 'unicode': '269B', 'category': 'Symbols'},
    {'name': 'Fleur-De-Lis', 'unicode': '269C', 'category': 'Symbols'},
    {'name': 'Outlined White Star', 'unicode': '269D', 'category': 'Symbols'},
    {'name': 'Three Lines Converging Right', 'unicode': '269E', 'category': 'Symbols'},
    {'name': 'Three Lines Converging Left', 'unicode': '269F', 'category': 'Symbols'},
    {'name': 'Warning Sign', 'unicode': '26A0', 'category': 'Symbols'},
    {'name': 'High Voltage Sign', 'unicode': '26A1', 'category': 'Symbols'},
    {'name': 'Doubled Female Sign', 'unicode': '26A2', 'category': 'Symbols'},
    {'name': 'Doubled Male Sign', 'unicode': '26A3', 'category': 'Symbols'},
    {'name': 'Interlocked Female And Male Sign', 'unicode': '26A4', 'category': 'Symbols'},
    {'name': 'Male And Female Sign', 'unicode': '26A5', 'category': 'Symbols'},
    {'name': 'Male With Stroke Sign', 'unicode': '26A6', 'category': 'Symbols'},
    {'name': 'Male With Stroke And Male And Female Sign', 'unicode': '26A7', 'category': 'Symbols'},
    {'name': 'Vertical Male With Stroke Sign', 'unicode': '26A8', 'category': 'Symbols'},
    {'name': 'Horizontal Male With Stroke Sign', 'unicode': '26A9', 'category': 'Symbols'},
    {'name': 'Medium White Circle', 'unicode': '26AA', 'category': 'Symbols'},
    {'name': 'Medium Black Circle', 'unicode': '26AB', 'category': 'Symbols'},
    {'name': 'Medium Small White Circle', 'unicode': '26AC', 'category': 'Symbols'},
    {'name': 'Marriage Symbol', 'unicode': '26AD', 'category': 'Symbols'},
    {'name': 'Divorce Symbol', 'unicode': '26AE', 'category': 'Symbols'},
    {'name': 'Unmarried Partnership Symbol', 'unicode': '26AF', 'category': 'Symbols'},
    {'name': 'Coffin', 'unicode': '26B0', 'category': 'Symbols'},
    {'name': 'Funeral Urn', 'unicode': '26B1', 'category': 'Symbols'},
    {'name': 'Neuter', 'unicode': '26B2', 'category': 'Symbols'},
    {'name': 'Ceres', 'unicode': '26B3', 'category': 'Symbols'},
    {'name': 'Pallas', 'unicode': '26B4', 'category': 'Symbols'},
    {'name': 'Juno', 'unicode': '26B5', 'category': 'Symbols'},
    {'name': 'Vesta', 'unicode': '26B6', 'category': 'Symbols'},
    {'name': 'Chiron', 'unicode': '26B7', 'category': 'Symbols'},
    {'name': 'Black Moon Lilith', 'unicode': '26B8', 'category': 'Symbols'},
    {'name': 'Sextile', 'unicode': '26B9', 'category': 'Symbols'},
    {'name': 'Semisextile', 'unicode': '26BA', 'category': 'Symbols'},
    {'name': 'Quincunx', 'unicode': '26BB', 'category': 'Symbols'},
    {'name': 'Sesquiquadrate', 'unicode': '26BC', 'category': 'Symbols'},
    {'name': 'Soccer Ball', 'unicode': '26BD', 'category': 'Symbols'},
    {'name': 'Baseball', 'unicode': '26BE', 'category': 'Symbols'},
    {'name': 'Squared Key', 'unicode': '26BF', 'category': 'Symbols'},
    {'name': 'White Draughts Man', 'unicode': '26C0', 'category': 'Symbols'},
    {'name': 'White Draughts King', 'unicode': '26C1', 'category': 'Symbols'},
    {'name': 'Black Draughts Man', 'unicode': '26C2', 'category': 'Symbols'},
    {'name': 'Black Draughts King', 'unicode': '26C3', 'category': 'Symbols'},
    {'name': 'Snowman Without Snow', 'unicode': '26C4', 'category': 'Symbols'},
    {'name': 'Sun Behind Cloud', 'unicode': '26C5', 'category': 'Symbols'},
    {'name': 'Rain', 'unicode': '26C6', 'category': 'Symbols'},
    {'name': 'Black Snowman', 'unicode': '26C7', 'category': 'Symbols'},
    {'name': 'Thunder Cloud And Rain', 'unicode': '26C8', 'category': 'Symbols'},
    {'name': 'Turned White Shogi Piece', 'unicode': '26C9', 'category': 'Symbols'},
    {'name': 'Turned Black Shogi Piece', 'unicode': '26CA', 'category': 'Symbols'},
    {'name': 'White Diamond In Square', 'unicode': '26CB', 'category': 'Symbols'},
    {'name': 'Crossing Lanes', 'unicode': '26CC', 'category': 'Symbols'},
    {'name': 'Disabled Car', 'unicode': '26CD', 'category': 'Travel & Places'},
    {'name': 'Ophiuchus', 'unicode': '26CE', 'category': 'Symbols'},
    {'name': 'Pick', 'unicode': '26CF', 'category': 'Symbols'},
    {'name': 'Car Sliding', 'unicode': '26D0', 'category': 'Travel & Places'},
    {'name': 'Helmet With White Cross', 'unicode': '26D1', 'category': 'Symbols'},
    {'name': 'Circled Crossing Lanes', 'unicode': '26D2', 'category': 'Symbols'},
    {'name': 'Chains', 'unicode': '26D3', 'category': 'Symbols'},
    {'name': 'No Entry', 'unicode': '26D4', 'category': 'Symbols'},
    {'name': 'Alternate One-Way Left Way Traffic', 'unicode': '26D5', 'category': 'Symbols'},
    {'name': 'Black Two-Way Left Way Traffic', 'unicode': '26D6', 'category': 'Symbols'},
    {'name': 'White Two-Way Left Way Traffic', 'unicode': '26D7', 'category': 'Symbols'},
    {'name': 'Black Left Lane Merge', 'unicode': '26D8', 'category': 'Symbols'},
    {'name': 'White Left Lane Merge', 'unicode': '26D9', 'category': 'Symbols'},
    {'name': 'Drive Slow Sign', 'unicode': '26DA', 'category': 'Symbols'},
    {'name': 'Heavy White Down-Pointing Triangle', 'unicode': '26DB', 'category': 'Symbols'},
    {'name': 'Left Closed Entry', 'unicode': '26DC', 'category': 'Symbols'},
    {'name': 'Squared Saltire', 'unicode': '26DD', 'category': 'Symbols'},
    {'name': 'Falling Diagonal In White Circle In Black Square', 'unicode': '26DE', 'category': 'Symbols'},
    {'name': 'Black Truck', 'unicode': '26DF', 'category': 'Symbols'},
    {'name': 'Restricted Left Entry-1', 'unicode': '26E0', 'category': 'Symbols'},
    {'name': 'Restricted Left Entry-2', 'unicode': '26E1', 'category': 'Symbols'},
    {'name': 'Astronomical Symbol For Uranus', 'unicode': '26E2', 'category': 'Symbols'},
    {'name': 'Heavy Circle With Stroke And Two Dots Above', 'unicode': '26E3', 'category': 'Symbols'},
    {'name': 'Pentagram', 'unicode': '26E4', 'category': 'Symbols'},
    {'name': 'Right-Handed Interlaced Pentagram', 'unicode': '26E5', 'category': 'People & Body'},
    {'name': 'Left-Handed Interlaced Pentagram', 'unicode': '26E6', 'category': 'People & Body'},
    {'name': 'Inverted Pentagram', 'unicode': '26E7', 'category': 'Symbols'},
    {'name': 'Black Cross On Shield', 'unicode': '26E8', 'category': 'Symbols'},
    {'name': 'Shinto Shrine', 'unicode': '26E9', 'category': 'Symbols'},
    {'name': 'Church', 'unicode': '26EA', 'category': 'Symbols'},
    {'name': 'Castle', 'unicode': '26EB', 'category': 'Symbols'},
    {'name': 'Historic Site', 'unicode': '26EC', 'category': 'Symbols'},
    {'name': 'Gear Without Hub', 'unicode': '26ED', 'category': 'Symbols'},
    {'name': 'Gear With Handles', 'unicode': '26EE', 'category': 'People & Body'},
    {'name': 'Map Symbol For Lighthouse', 'unicode': '26EF', 'category': 'Symbols'},
    {'name': 'Mountain', 'unicode': '26F0', 'category': 'Travel & Places'},
    {'name': 'Umbrella On Ground', 'unicode': '26F1', 'category': 'Travel & Places'},
    {'name': 'Fountain', 'unicode': '26F2', 'category': 'Travel & Places'},
    {'name': 'Flag In Hole', 'unicode': '26F3', 'category': 'Travel & Places'},
    {'name': 'Ferry', 'unicode': '26F4', 'category': 'Travel & Places'},
    {'name': 'Sailboat', 'unicode': '26F5', 'category': 'Travel & Places'},
    {'name': 'Square Four Corners', 'unicode': '26F6', 'category': 'Travel & Places'},
    {'name': 'Skier', 'unicode': '26F7', 'category': 'Travel & Places'},
    {'name': 'Ice Skate', 'unicode': '26F8', 'category': 'Travel & Places'},
    {'name': 'Person With Ball', 'unicode': '26F9', 'category': 'Travel & Places'},
    {'name': 'Tent', 'unicode': '26FA', 'category': 'Travel & Places'},
    {'name': 'Japanese Bank Symbol', 'unicode': '26FB', 'category': 'Travel & Places'},
    {'name': 'Headstone Graveyard Symbol', 'unicode': '26FC', 'category': 'Travel & Places'},
    {'name': 'Fuel Pump', 'unicode': '26FD', 'category': 'Travel & Places'},
    {'name': 'Cup On Black Square', 'unicode': '26FE', 'category': 'Travel & Places'},
    {'name': 'White Flag With Horizontal Middle Black Stripe', 'unicode': '26FF', 'category': 'Travel & Places'},
    {'name': 'Black Safety Scissors', 'unicode': '2700', 'category': 'Symbols'},
    {'name': 'Upper Blade Scissors', 'unicode': '2701', 'category': 'Symbols'},
    {'name': 'Black Scissors', 'unicode': '2702', 'category': 'Symbols'},
    {'name': 'Lower Blade Scissors', 'unicode': '2703', 'category': 'Symbols'},
    {'name': 'White Scissors', 'unicode': '2704', 'category': 'Symbols'},
    {'name': 'White Heavy Check Mark', 'unicode': '2705', 'category': 'Symbols'},
    {'name': 'Telephone Location Sign', 'unicode': '2706', 'category': 'Symbols'},
    {'name': 'Tape Drive', 'unicode': '2707', 'category': 'Symbols'},
    {'name': 'Airplane', 'unicode': '2708', 'category': 'Symbols'},
    {'name': 'Envelope', 'unicode': '2709', 'category': 'Symbols'},
    {'name': 'Raised Fist', 'unicode': '270A', 'category': 'Symbols'},
    {'name': 'Raised Hand', 'unicode': '270B', 'category': 'People & Body'},
    {'name': 'Victory Hand', 'unicode': '270C', 'category': 'People & Body'},
    {'name': 'Writing Hand', 'unicode': '270D', 'category': 'People & Body'},
    {'name': 'Lower Right Pencil', 'unicode': '270E', 'category': 'Symbols'},
    {'name': 'Pencil', 'unicode': '270F', 'category': 'Symbols'},
    {'name': 'Upper Right Pencil', 'unicode': '2710', 'category': 'Symbols'},
    {'name': 'White Nib', 'unicode': '2711', 'category': 'Symbols'},
    {'name': 'Black Nib', 'unicode': '2712', 'category': 'Symbols'},
    {'name': 'Check Mark', 'unicode': '2713', 'category': 'Symbols'},
    {'name': 'Heavy Check Mark', 'unicode': '2714', 'category': 'Symbols'},
    {'name': 'Multiplication X', 'unicode': '2715', 'category': 'Symbols'},
    {'name': 'Heavy Multiplication X', 'unicode': '2716', 'category': 'Symbols'},
    {'name': 'Ballot X', 'unicode': '2717', 'category': 'Symbols'},
    {'name': 'Heavy Ballot X', 'unicode': '2718', 'category': 'Symbols'}
])

EMOJI_NAME_MAPPING = {
    # Smileys & Faces
    0x1F600: "Grinning Face",
    0x1F601: "Beaming Face with Smiling Eyes",
    0x1F602: "Face with Tears of Joy",
    0x1F603: "Grinning Face with Big Eyes",
    0x1F604: "Grinning Face with Smiling Eyes",
    0x1F605: "Grinning Face with Sweat",
    0x1F606: "Grinning Squinting Face", 
    0x1F607: "Smiling Face with Halo",
    0x1F608: "Smiling Face with Horns",
    0x1F609: "Winking Face",
    0x1F60A: "Smiling Face with Smiling Eyes",
    0x1F60B: "Face Savoring Food",
    0x1F60C: "Relieved Face",
    0x1F60D: "Smiling Face with Heart-Eyes",
    0x1F60E: "Smiling Face with Sunglasses",
    # Animals
    0x1F980: "Crab",
    0x1F981: "Lion",
    0x1F984: "Unicorn",
    0x1F985: "Eagle",
    0x1F98B: "Butterfly",
    0x1F98C: "Deer",
    0x1F98D: "Gorilla",
    0x1F98E: "Lizard",
    0x1F98F: "Rhinoceros",
    0x1F990: "Shrimp",
    0x1F991: "Squid",
    0x1F992: "Giraffe",
    0x1F993: "Zebra",
    0x1F994: "Hedgehog",
    0x1F995: "Sauropod",
    0x1F996: "T-Rex",
    # Food items
    0x1F32D: "Hot Dog",
    0x1F32E: "Taco",
    0x1F32F: "Burrito",
    0x1F330: "Chestnut",
    0x1F354: "Hamburger",
    0x1F355: "Pizza",
    0x1F356: "Meat on Bone",
    0x1F357: "Poultry Leg",
    0x1F35A: "Cooked Rice",
    0x1F35C: "Steaming Bowl",
    0x1F35E: "Bread",
    0x1F35F: "French Fries",
    0x1F363: "Sushi",
    0x1F366: "Soft Ice Cream",
    0x1F369: "Doughnut",
    0x1F36A: "Cookie",
    # Plants & Flowers
    0x1F331: "Seedling",
    0x1F332: "Evergreen Tree",
    0x1F333: "Deciduous Tree",
    0x1F334: "Palm Tree",
    0x1F335: "Cactus",
    0x1F337: "Tulip",
    0x1F338: "Cherry Blossom",
    0x1F339: "Rose",
    0x1F33A: "Hibiscus",
    0x1F33B: "Sunflower",
}

# Global variable to store the lock file object
_lock_file = None

def prevent_multiple_instances():
    """
    Prevents multiple instances of the application from running simultaneously.
    If another instance is already running, activates that window and exits.
    
    Returns:
        bool: True if this is the only instance, False if another instance exists
    """
    global _lock_file
    
    # Create a lock file path in the temp directory
    lock_file_path = os.path.join(tempfile.gettempdir(), "CrystalRealmsEmojiPicker.lock")
    
    try:
        # Try to create the lock file
        if os.path.exists(lock_file_path):
            # Check if the lock file is stale (from a crashed instance)
            try:
                with open(lock_file_path, 'r') as f:
                    pid = int(f.read().strip())
                    # Check if process with this PID exists
                    if pid_exists(pid):
                        # Process exists, find and activate the window
                        activate_existing_window()
                        return False
                    else:
                        # Stale lock file, remove it
                        os.remove(lock_file_path)
            except:
                # If any error occurs, assume the lock file is invalid and remove it
                try:
                    os.remove(lock_file_path)
                except:
                    pass
        
        # Create a new lock file
        _lock_file = open(lock_file_path, 'w')
        _lock_file.write(str(os.getpid()))
        _lock_file.flush()
        
        # Register cleanup function
        atexit.register(cleanup_lock_file, lock_file_path)
        
        # This is the only instance
        return True
        
    except:
        # If any error occurs during lock creation, assume another instance exists
        activate_existing_window()
        return False

def pid_exists(pid):
    """Check if a process with given PID exists"""
    try:
        # Windows implementation
        import ctypes
        kernel32 = ctypes.windll.kernel32
        process = kernel32.OpenProcess(1, 0, pid)
        if process != 0:
            kernel32.CloseHandle(process)
            return True
        return False
    except:
        return False

def activate_existing_window():
    """Find and activate the existing application window"""
    try:
        import win32gui
        import win32con
        
        def enum_windows_callback(hwnd, result):
            if win32gui.IsWindowVisible(hwnd):
                window_title = win32gui.GetWindowText(hwnd)
                if "Crystal Realms EmojiPicker" in window_title:
                    result.append(hwnd)
            return True
        
        window_handles = []
        win32gui.EnumWindows(enum_windows_callback, window_handles)
        
        if window_handles:
            # Restore and bring to foreground
            hwnd = window_handles[0]
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
            win32gui.SetForegroundWindow(hwnd)
            print(f"Another instance is already running. Activating existing window: {hwnd}")
    except Exception as e:
        print(f"Error activating existing window: {e}")

def cleanup_lock_file(lock_file_path):
    """Clean up the lock file when the application exits"""
    global _lock_file
    try:
        if _lock_file:
            _lock_file.close()
        if os.path.exists(lock_file_path):
            os.remove(lock_file_path)
    except:
        pass

def is_already_running():
    """
    Check if another instance of the application is already running
    using a file lock approach that works in both development and production.
    
    Returns True if another instance is running, False otherwise.
    """
    global lock_file
    
    try:
        # Create a lock file path in the temp directory
        lock_file_path = os.path.join(tempfile.gettempdir(), "CrystalRealmsEmojiPicker.lock")
        
        # Try to open the file in exclusive creation mode
        try:
            lock_file = open(lock_file_path, "w+")
            
            # Try to acquire an exclusive lock immediately without waiting
            try:
                # Windows locking with msvcrt
                msvcrt.locking(lock_file.fileno(), msvcrt.LK_NBLCK, 1)
                
                # If we get here, we got the lock
                # Write our PID to the lock file
                lock_file.write(str(os.getpid()))
                lock_file.flush()
                
                # Register a cleanup function to release the lock on exit
                atexit.register(cleanup_lock)
                
                return False
                
            except IOError:
                # Lock failed, another instance has the lock
                lock_file.close()
                lock_file = None
        except IOError:
            # File already exists and couldn't be opened exclusively
            pass
            
        # If we reach here, another instance is running
        # Try to find and activate the existing window
        try:
            def enum_windows_callback(hwnd, result):
                if win32gui.IsWindowVisible(hwnd):
                    window_title = win32gui.GetWindowText(hwnd)
                    if "Crystal Realms EmojiPicker" in window_title:
                        result.append(hwnd)
                return True
            
            window_handles = []
            win32gui.EnumWindows(enum_windows_callback, window_handles)
            
            if window_handles:
                # Restore and bring to foreground
                for hwnd in window_handles:
                    try:
                        win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                        win32gui.SetForegroundWindow(hwnd)
                        print(f"Another instance is already running. Activating existing window: {hwnd}")
                        break
                    except Exception as e:
                        print(f"Error activating window {hwnd}: {e}")
                        continue
            else:
                print("Another instance is running but window not found")
                
                # Wait a moment and try again - the window might be initializing
                time.sleep(1)
                win32gui.EnumWindows(enum_windows_callback, window_handles)
                
                if window_handles:
                    for hwnd in window_handles:
                        try:
                            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                            win32gui.SetForegroundWindow(hwnd)
                            print(f"Found window on second attempt: {hwnd}")
                            break
                        except Exception as e:
                            print(f"Error activating window {hwnd}: {e}")
                            continue
        except Exception as e:
            print(f"Error finding existing window: {e}")
        
        return True
            
    except Exception as e:
        print(f"Error checking for running instance: {e}")
        traceback.print_exc()
        # If we can't check, assume no other instance is running
        return False

def cleanup_lock():
    """Clean up the lock file when the application exits"""
    global lock_file
    try:
        if lock_file is not None:
            # Release the lock
            msvcrt.locking(lock_file.fileno(), msvcrt.LK_UNLCK, 1)
            lock_file.close()
            
            # Try to remove the lock file
            try:
                lock_file_path = os.path.join(tempfile.gettempdir(), "CrystalRealmsEmojiPicker.lock")
                if os.path.exists(lock_file_path):
                    os.remove(lock_file_path)
            except:
                pass
    except:
        pass

def get_proper_emoji_name(unicode_char, default_name=None):
    """Get a proper descriptive name for an emoji character."""
    try:
        # Get code point of first character
        code_point = ord(unicode_char[0])
        
        # Try to get from our mapping
        if code_point in EMOJI_NAME_MAPPING:
            return EMOJI_NAME_MAPPING[code_point]
        
        # Try Unicode name directly
        try:
            name = unicodedata.name(unicode_char[0])
            # Clean up Unicode names to be more readable
            name = re.sub(r'_', ' ', name)
            name = name.title()  # Convert to title case
            return name
        except (ValueError, TypeError):
            pass
        
        # If default name isn't generic, use it
        if default_name and not default_name.startswith("Emoji U+"):
            return default_name
            
        # Check if it's a BUILT_IN_EMOJI_DATA entry
        for emoji in BUILT_IN_EMOJI_DATA:
            try:
                if int(emoji['unicode'], 16) == code_point:
                    return emoji['name']
            except:
                continue
                
        # Fallback to default if provided
        return default_name or f"Emoji U+{code_point:X}"
        
    except (IndexError, TypeError):
        return default_name or "Unknown Emoji"

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

def get_emoji_cache_path():
    """Get path to emoji cache file in user's AppData directory"""
    app_data = os.path.join(os.getenv('APPDATA'), 'EmojiPicker')
    if not os.path.exists(app_data):
        os.makedirs(app_data)
    return os.path.join(app_data, 'emoji_cache.json')

def save_emoji_cache(emoji_list):
    """Save downloaded emoji list to local cache file"""
    try:
        cache_path = get_emoji_cache_path()
        with open(cache_path, 'w', encoding='utf-8') as f:
            json.dump(emoji_list, f, ensure_ascii=False, indent=2)
        print(f"Emoji cache saved to {cache_path}")
        return True
    except Exception as e:
        print(f"Error saving emoji cache: {e}")
        return False

def load_emoji_cache():
    """Load emoji list from local cache file"""
    try:
        cache_path = get_emoji_cache_path()
        if os.path.exists(cache_path):
            with open(cache_path, 'r', encoding='utf-8') as f:
                emoji_list = json.load(f)
            print(f"Loaded {len(emoji_list)} emojis from cache")
            return emoji_list
        else:
            print("No emoji cache found")
            return None
    except Exception as e:
        print(f"Error loading emoji cache: {e}")
        return None

def resource_path(relative_path):
    """Get absolute path to resource, works for dev and for PyInstaller"""
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    
    return os.path.join(base_path, relative_path)

NOTO_EMOJI_PATH = resource_path("NotoEmoji-Regular.ttf")

def get_comprehensive_emoji_dataset():
    """
    Generate a comprehensive emoji dataset that includes standard Unicode emojis.
    This creates a dataset with ~1800 emojis similar to what the API would provide.
    """
    emoji_data = []
    
    # Basic emoji ranges to include (this will get us to ~1800 emojis)
    emoji_ranges = [
        # Emoticons
        (0x1F600, 0x1F64F),  # Emoticons
        (0x1F300, 0x1F5FF),  # Misc Symbols and Pictographs
        (0x1F680, 0x1F6FF),  # Transport and Map
        (0x1F700, 0x1F77F),  # Alchemical Symbols
        (0x1F780, 0x1F7FF),  # Geometric Shapes
        (0x1F800, 0x1F8FF),  # Supplemental Arrows-C
        (0x1F900, 0x1F9FF),  # Supplemental Symbols and Pictographs
        (0x2600, 0x26FF),    # Misc symbols
        (0x2700, 0x27BF),    # Dingbats
        (0x3030, 0x303F),    # CJK Symbols
        (0x1F1E0, 0x1F1FF),  # Flags
        (0x1F200, 0x1F2FF),  # Enclosed Ideographic Supplement
        (0x1FA70, 0x1FAFF),  # Symbols and Pictographs Extended-A
    ]
    
    # First try to load file-based emojis
    file_emojis = load_noto_emojis_from_file()
    emoji_dict = {int(e['unicode'], 16): e for e in file_emojis}
    
    # Then add all emoji ranges
    for start, end in emoji_ranges:
        for code_point in range(start, end + 1):
            # Skip if emoji already exists from file
            if code_point in emoji_dict:
                continue
                
            try:
                char = chr(code_point)
                # Generate a name if not in the emoji_dict
                category = "Symbols"
                if 0x1F600 <= code_point <= 0x1F64F:
                    category = "Smileys & Emotion"
                elif 0x1F300 <= code_point <= 0x1F5FF:
                    category = "Symbols & Pictographs"
                elif 0x1F680 <= code_point <= 0x1F6FF:
                    category = "Transport & Map"
                
                name = f"Emoji U+{code_point:X}"
                
                emoji_data.append({
                    'name': name,
                    'unicode': f"{code_point:X}",
                    'category': category
                })
            except:
                pass
    
    # Add all file-based emojis
    for emoji in file_emojis:
        if emoji not in emoji_data:
            emoji_data.append(emoji)
    
    print(f"Created comprehensive emoji dataset with {len(emoji_data)} emojis")
    return emoji_data

def load_noto_emojis_from_file():
    """
    Load emoji data directly from the noto_emojis.txt or noto_emojis_converted.txt file.
    """
    try:
        # Try to find the file in the script directory
        file_path = resource_path("noto_emojis.txt")
        if not os.path.exists(file_path):
            file_path = resource_path("noto_emojis_converted.txt")
            
        if not os.path.exists(file_path):
            print("Emoji file not found, using built-in dataset")
            return [{'name': emoji['name'], 'unicode': emoji['unicode'], 'category': emoji['category']} 
                    for emoji in BUILT_IN_EMOJI_DATA]
        
        emojis = []
        with open(file_path, 'r', encoding='utf-8') as f:
            lines = f.readlines()
            for line in lines:
                try:
                    # Parse line, expecting format like: {'name': 'Name', 'unicode': 'XXXX', 'category': 'Category'}
                    emoji_data = eval(line.strip())
                    emojis.append(emoji_data)
                except:
                    pass
        
        print(f"Loaded {len(emojis)} emojis from {os.path.basename(file_path)}")
        return emojis
    except Exception as e:
        print(f"Error loading emoji file: {e}")
        return [{'name': emoji['name'], 'unicode': emoji['unicode'], 'category': emoji['category']} 
                for emoji in BUILT_IN_EMOJI_DATA]

def set_window_title_bar_colors(hwnd, bg_color=None, text_color=None):
    """
    Sets the title bar color of a window on Windows 10/11.
    Uses win32gui to get a more reliable window handle.
    
    Args:
        hwnd: Window handle integer
        bg_color: Background color in hex format (#RRGGBB)
        text_color: Text color in hex format (#RRGGBB)
    
    Returns:
        bool: True if successful, False otherwise
    """
    # If colors are None, we don't change anything
    if bg_color is None or text_color is None:
        return False
        
    # Only works on Windows 10/11
    if platform.system() != "Windows" or int(platform.version().split('.')[0]) < 10:
        print("Title bar coloring requires Windows 10 or later")
        return False
        
    try:
        # Convert Tkinter's window ID to an HWND 
        # This is more reliable than using winfo_id directly
        if isinstance(hwnd, int):
            try:
                # Verify the handle is valid
                window_text = win32gui.GetWindowText(hwnd)
                print(f"Window found: {window_text}")
            except Exception:
                print(f"Invalid hwnd: {hwnd}, attempting to find by class")
                # Try to find the window by class name
                hwnd = win32gui.FindWindow("TkTopLevel", None)
                if not hwnd:
                    print("Could not find Tkinter window")
                    return False
        
        print(f"Using window handle: {hwnd}")
        
        # Constants for DwmSetWindowAttribute
        DWMWA_CAPTION_COLOR = 35  # Available from Windows build 18312+
        DWMWA_TEXT_COLOR = 36     # Available from Windows build 18312+
        
        # Convert hex colors to RGB and then to COLORREF (which is BGR for Windows)
        def hex_to_colorref(hex_color):
            # Remove the '#' if present
            hex_color = hex_color.lstrip('#')
            
            # Convert hex to RGB
            r = int(hex_color[0:2], 16)
            g = int(hex_color[2:4], 16)
            b = int(hex_color[4:6], 16)
            
            # Convert to COLORREF (DWORD) which is in BGR format: 0x00BBGGRR
            return b | (g << 8) | (r << 16)
        
        # Load the dwmapi.dll
        dwmapi = ctypes.WinDLL("dwmapi")
        
        # Define the function prototype
        dwmapi.DwmSetWindowAttribute.argtypes = [
            wintypes.HWND,        # hwnd
            wintypes.DWORD,       # dwAttribute
            ctypes.POINTER(ctypes.c_int),  # pvAttribute
            wintypes.DWORD        # cbAttribute
        ]
        dwmapi.DwmSetWindowAttribute.restype = ctypes.HRESULT
        
        # Convert colors to COLORREF
        bg_colorref = hex_to_colorref(bg_color)
        text_colorref = hex_to_colorref(text_color)
        
        # Print for debugging
        print(f"Setting title bar colors - BG: {bg_color} ({bg_colorref}), Text: {text_color} ({text_colorref})")
        
        # Set background color
        bg_color_value = ctypes.c_int(bg_colorref)
        result1 = dwmapi.DwmSetWindowAttribute(
            hwnd, 
            DWMWA_CAPTION_COLOR, 
            ctypes.byref(bg_color_value), 
            ctypes.sizeof(bg_color_value)
        )
        
        # Set text color
        text_color_value = ctypes.c_int(text_colorref)
        result2 = dwmapi.DwmSetWindowAttribute(
            hwnd, 
            DWMWA_TEXT_COLOR, 
            ctypes.byref(text_color_value), 
            ctypes.sizeof(text_color_value)
        )
        
        if result1 == 0 and result2 == 0:
            print("Title bar colors set successfully")
            return True
        else:
            print(f"Failed to set title bar colors: {result1}, {result2}")
            return False
            
    except Exception as e:
        print(f"Error setting title bar colors: {e}")
        traceback.print_exc()
        return False

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
    """
    Enhanced function to get all characters supported by Noto Emoji font.
    Includes better error handling, caching and more comprehensive font checking.
    """
    # Check if we have a cached set of characters
    cache_file = os.path.join(os.getenv('APPDATA'), 'EmojiPicker', 'font_cache.json')
    
    try:
        if os.path.exists(cache_file):
            with open(cache_file, 'r') as f:
                # Load cached character codes
                char_codes = json.load(f)
                chars = set(chr(code) for code in char_codes)
                print(f"Loaded {len(chars)} characters from font cache")
                return chars
    except Exception as e:
        print(f"Error loading font cache: {e}")
    
    # If no cache, load from font file
    if not os.path.exists(NOTO_EMOJI_PATH):
        print(f"Warning: Noto Emoji font not found at {NOTO_EMOJI_PATH}")
        chars = get_fallback_supported_chars()
    else:
        try:
            font = TTFont(NOTO_EMOJI_PATH)
            chars = set()
            
            # Process all cmap tables for maximum coverage
            for table in font['cmap'].tables:
                for char_code in table.cmap.keys():
                    try:
                        # Some character codes might be invalid, so we need to handle exceptions
                        char = chr(char_code)
                        chars.add(char)
                    except ValueError:
                        continue
            
            # Check if we found a reasonable number of characters
            if len(chars) < 100:
                print(f"Warning: Only {len(chars)} characters found in Noto Emoji font. Using fallback.")
                chars = get_fallback_supported_chars()
            else:
                print(f"Loaded {len(chars)} supported characters from Noto Emoji font")
                
                # Save to cache for future use
                try:
                    os.makedirs(os.path.dirname(cache_file), exist_ok=True)
                    with open(cache_file, 'w') as f:
                        # Save character codes as integers to keep file small
                        json.dump([ord(c) for c in chars], f)
                except Exception as e:
                    print(f"Error saving font cache: {e}")
                
        except Exception as e:
            print(f"Error reading Noto Emoji font file: {e}")
            chars = get_fallback_supported_chars()
    
    return chars

def get_fallback_supported_chars():
    """
    Create a fallback set of supported characters when the font can't be read.
    This includes only the most common emoji ranges that are widely supported.
    """
    supported_chars = set()
    
    # Common emoji ranges that are widely supported
    safe_ranges = [
        (0x2600, 0x26FF),    # Miscellaneous Symbols
        (0x2700, 0x27BF),    # Dingbats
        (0x1F300, 0x1F5FF),  # Miscellaneous Symbols and Pictographs (excluding problematic ranges)
        (0x1F600, 0x1F64F),  # Emoticons
        (0x1F680, 0x1F6FF),  # Transport and Map Symbols
        # Exclude problematic ranges like U+1F20X
    ]
    
    # Explicitly exclude problematic ranges
    excluded_ranges = [
        (0x1F200, 0x1F2FF),  # Enclosed Ideographic Supplement (includes U+1F20B to U+1F210)
        # Add other problematic ranges here
    ]
    
    # Add characters from safe ranges
    for start, end in safe_ranges:
        for code_point in range(start, end + 1):
            supported_chars.add(chr(code_point))
    
    # Remove characters from excluded ranges
    for start, end in excluded_ranges:
        for code_point in range(start, end + 1):
            try:
                supported_chars.discard(chr(code_point))
            except:
                pass
    
    # Add all characters from the built-in emoji data as they should be safe
    for emoji in BUILT_IN_EMOJI_DATA:
        try:
            hex_val = emoji['unicode']
            unicode_char = chr(int(hex_val, 16))
            supported_chars.add(unicode_char)
        except:
            continue
    
    print(f"Using fallback with {len(supported_chars)} supported characters")
    return supported_chars

def fetch_emojis():
    """
    Enhanced emoji fetching with robust fallback:
    1. Try cache first (fast local option)
    2. Use built-in dataset if cache fails
    3. Try API last (can be slow)
    """
    # Try cached data first (fastest)
    cached_emojis = load_emoji_cache()
    if cached_emojis and len(cached_emojis) > 100:
        print(f"Using {len(cached_emojis)} emojis from cache")
        # Start API fetch in background for future use
        threading.Thread(target=lambda: background_api_fetch(), daemon=True).start()
        return cached_emojis
    
    # Use comprehensive dataset as second option
    dataset = get_comprehensive_emoji_dataset()
    
    # Try API in background for next time 
    threading.Thread(target=lambda: background_api_fetch(), daemon=True).start()
    
    return dataset

def background_api_fetch():
    """Fetch API data in background and save to cache for next run"""
    try:
        print(f"Fetching emojis from API in background: {EMOJI_API}")
        response = requests.get(EMOJI_API, timeout=5)  # Reduced timeout
        if response.status_code == 200:
            emoji_list = response.json()
            if len(emoji_list) > 100:
                print(f"Successfully fetched {len(emoji_list)} emojis from API")
                save_emoji_cache(emoji_list)
    except Exception as e:
        print(f"Background API fetch error: {str(e)}")

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
        
        # Center the window
        window_width = 600
        window_height = 450
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        x = (screen_width // 2) - (window_width // 2)
        y = (screen_height // 2) - (window_height // 2)
        self.root.geometry(f"{window_width}x{window_height}+{x}+{y}")
        
        # Create splash screen immediately
        self.create_splash_screen()
        
        # Set window icon right away
        icon_path = resource_path("app.ico")
        if os.path.exists(icon_path):
            self.root.iconbitmap(icon_path)
        
        # Add these for the queue-based hotkey handling
        self.global_lock = threading.RLock()  # Reentrant lock for thread safety
        self.hotkey_queue = queue.Queue()
        self.hotkey_processor_running = False
        
        # This tells us if title bar colors have been applied
        self.title_colors_applied = False
        
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
        
        self.current_page = 0
        self.search_var = tk.StringVar()
        self.emoji_data = {}
        self.filtered_emojis = []
        
        # Run font loading in a separate thread
        self.supported_chars = None
        threading.Thread(target=self.async_load_font, daemon=True).start()
        
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
        
        # Setup main UI
        self.setup_gui()
        
        # Apply theme only after UI is created
        self.apply_theme(self.current_theme)
        
        # Start the hotkey processor
        self.start_hotkey_processor()
        
        # Set up message handler for hotkeys (only after UI is ready)
        self.root.after(500, self.setup_win32_hotkeys)
        
        # Load emojis after a short delay to let the UI become responsive first
        self.root.after(100, self.load_emojis)
        
        # Try to find the game window after UI is ready
        self.root.after(1000, self.find_game_window)
    
    def update_emoji_names(self):
        """Update emoji names for existing emojis"""
        renamed_count = 0
        
        for unicode_char, emoji_data in list(self.emoji_data.items()):
            # Check if the name is generic (starts with "Emoji U+")
            if emoji_data.name.startswith("Emoji U+"):
                # Get a proper name
                proper_name = get_proper_emoji_name(unicode_char, emoji_data.name)
                
                # Update the name
                emoji_data.name = proper_name
                renamed_count += 1
        
        if renamed_count > 0:
            # Save the updated data
            print(f"Updated {renamed_count} emoji names")
            self.save_emoji_data()
            # Update the display
            self.filter_and_sort_emojis()
    
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
        """Process a hotkey event safely on the main thread, with immediate UI update"""
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
            
            # Update the specific item immediately on the UI
            if self.root.state() == 'normal':
                # Check if the emoji is on the current page
                start_idx = self.current_page * PAGE_SIZE
                end_idx = start_idx + PAGE_SIZE
                
                # Find the emoji's position in filtered_emojis
                try:
                    abs_idx = next(i for i, e in enumerate(self.filtered_emojis) 
                                  if e.unicode_key == emoji_data.unicode_key)
                                  
                    # If it's on the current page, update it in place
                    if start_idx <= abs_idx < end_idx:
                        page_idx = abs_idx - start_idx
                        pin_indicator = "📌 " if emoji_data.pinned else ""
                        count_indicator = f" ({emoji_data.count})" if emoji_data.count > 0 else ""
                        keybind_indicator = f" [{emoji_data.keybind}]" if emoji_data.keybind else ""
                        display_text = f"{emoji_data.unicode_key} {pin_indicator}{emoji_data.name}{count_indicator}{keybind_indicator}"
                        
                        # Schedule UI update on main thread
                        self.root.after_idle(lambda: self.update_listbox_item(page_idx, display_text, abs_idx))
                except StopIteration:
                    pass  # Emoji not found in filtered list
            
            # Schedule saving for later to not block the main thread
            self.root.after(500, self.thread_safe_save_emoji_data)
        except Exception as e:
            print(f"Error processing hotkey on main thread: {e}")
            traceback.print_exc()
        finally:
            # Always release the processing flag
            with self.global_lock:
                self.processing_hotkey = False

    def update_listbox_item(self, index, text, abs_idx):
        """Helper method to update a listbox item safely from the main thread"""
        try:
            self.listbox.delete(index)
            self.listbox.insert(index, text)
            
            # If this was the selected item, keep it selected
            if self.selected_index == abs_idx:
                self.listbox.selection_clear(0, tk.END)
                self.listbox.selection_set(index)
        except Exception as e:
            print(f"Error updating listbox item: {e}")
                
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
        
        # Immediately unregister or register hotkeys
        try:
            if hasattr(self, 'hotkey_hwnd'):
                if self.keybinds_enabled:
                    # Clear any existing hotkeys first
                    self.unregister_all_hotkeys()
                    # Register hotkeys
                    self.register_system_wide_hotkeys()
                    self.update_status("Keybinds enabled")
                else:
                    # Unregister all hotkeys if disabled
                    self.unregister_all_hotkeys()
                    self.update_status("Keybinds disabled")
        except Exception as e:
            print(f"Error toggling hotkeys: {e}")
            traceback.print_exc()
            self.update_status("Error toggling keybinds")

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
                 "Using built-in emoji dataset instead.\n"
                 "For best results, download the font from:\n"
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
        """
        Load emojis with improved async loading and fallback mechanism.
        Displays loading indicators and prevents UI freezing.
        """
        # Clear existing data first
        self.emoji_data.clear()
        self.keybind_to_emoji.clear()  # Clear keybind mapping at start
        
        # Flag to prevent re-showing loading messages
        self.loading_complete = False
        
        # Display loading message in listbox
        self.listbox.delete(0, tk.END)
        loading_messages = [
            "🦄 Summoning unicorn magic...",
            "🔍 Gathering unicorn-enchanted emojis...",
            "✨ Performing unicorn magic...",
            "🚀 Completing unicorn magic..."
        ]
        
        def cycle_loading_message():
            """Cycle through loading messages to show app is alive"""
            if not hasattr(self, 'loading_complete') or not self.loading_complete:
                if not hasattr(self, 'loading_index'):
                    self.loading_index = 0
                
                # Only update if we're not done loading
                if not hasattr(self, 'loading_complete') or not self.loading_complete:
                    self.listbox.delete(0, tk.END)
                    self.listbox.insert(tk.END, loading_messages[self.loading_index])
                    self.loading_index = (self.loading_index + 1) % len(loading_messages)
                    
                    # Schedule next message if not complete
                    if not hasattr(self, 'loading_complete') or not self.loading_complete:
                        self.root.after(1000, cycle_loading_message)
        
        # Start cycling loading messages
        cycle_loading_message()
        
        def load_emojis_thread():
            """Background thread to load emojis without freezing UI"""
            try:
                # Try to get emojis from various sources with timeouts
                emoji_sources = [
                    (lambda: fetch_emojis(), "Primary API"),
                    (lambda: load_emoji_cache(), "Local Cache"),
                    (lambda: [{'name': emoji['name'], 'unicode': emoji['unicode'], 'category': emoji['category']} 
                            for emoji in BUILT_IN_EMOJI_DATA], "Built-in Dataset")
                ]
                
                raw_emojis = None
                source_name = "Unknown"
                
                for source_func, source_desc in emoji_sources:
                    try:
                        raw_emojis = source_func()
                        if raw_emojis and len(raw_emojis) > 100:
                            source_name = source_desc
                            break
                    except Exception as e:
                        print(f"Error loading from {source_desc}: {e}")
                
                if not raw_emojis:
                    raw_emojis = BUILT_IN_EMOJI_DATA
                    source_name = "Fallback Built-in Dataset"
                
                # Get saved data for persistent information
                saved_data = load_data()
                
                # Track loading statistics
                loaded_count = 0
                renamed_count = 0
                filtered_count = 0
                error_count = 0
                
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
                        
                        # Check if the emoji is supported by Noto Emoji
                        is_supported = True
                        if self.supported_chars:
                            for char in unicode_char:
                                if char not in self.supported_chars:
                                    is_supported = False
                                    break
                        
                        # Skip unsupported emojis
                        if not is_supported:
                            filtered_count += 1
                            continue
                        
                        # Get saved data if available
                        raw_name = emoji.get('name', '')
                        
                        # Check if the name is generic or missing
                        is_generic_name = raw_name == '' or raw_name.startswith(f"Emoji U+")
                        
                        # Get a proper name for the emoji
                        if is_generic_name:
                            proper_name = get_proper_emoji_name(unicode_char, raw_name)
                            renamed_count += 1
                        else:
                            proper_name = raw_name
                        
                        saved = saved_data.get(unicode_char, {})
                        count = saved.get('count', 0)
                        pinned = saved.get('pinned', False)
                        manual_order = saved.get('manual_order', float('inf'))
                        keybind = saved.get('keybind', None)
                        
                        # If the saved data has a name, use it (unless it's generic)
                        saved_name = saved.get('name', '')
                        if saved_name and not saved_name.startswith("Emoji U+"):
                            proper_name = saved_name
                        
                        # Create emoji data with proper name
                        self.emoji_data[unicode_char] = EmojiData(
                            proper_name, unicode_char, count, pinned, manual_order, keybind
                        )
                        
                        # Register keybind if it exists
                        if keybind:
                            # Check for duplicate keybinds
                            if keybind in self.keybind_to_emoji:
                                # Keep the last one, effectively overwriting earlier mappings
                                print(f"Warning: Duplicate keybind {keybind} found. Last one will be used.")
                            
                            # Map the keybind to the emoji's unicode
                            self.keybind_to_emoji[keybind] = unicode_char
                            
                        loaded_count += 1
                    except Exception as e:
                        error_count += 1
                        if error_count < 10:  # Limit error messages
                            print(f"Error processing emoji {emoji.get('name', 'unknown')}: {e}")
                
                # Prepare loading information
                loading_info = (
                    f"Loaded {loaded_count} emojis from {source_name}\n"
                    f"Renamed: {renamed_count} | Filtered: {filtered_count}"
                )
                
                # Schedule UI update on main thread
                self.root.after(0, lambda: self.finalize_emoji_loading(loading_info))
                
            except Exception as e:
                # Schedule error display on main thread
                self.root.after(0, lambda: self.display_loading_error(str(e)))
        
        # Start the loading in a separate thread
        threading.Thread(target=load_emojis_thread, daemon=True).start()

    def finalize_emoji_loading(self, loading_info):
        """
        Finalize emoji loading on the main thread.
        Manages UI update after background loading.
        """
        try:
            # Mark loading as complete to stop loading messages
            self.loading_complete = True
            
            # Log the loading information
            print(loading_info)
            
            # Filter and sort emojis
            self.filter_and_sort_emojis()
            
            # Ensure page is set and displayed
            def update_display():
                try:
                    # Explicitly set loading complete flag
                    self.loading_complete = True
                    
                    # Reset to first page
                    self.current_page = 0
                    
                    # Display the page using existing method
                    self.display_page()
                    
                    # Ensure UI is responsive
                    self.root.update_idletasks()
                    
                    # IMPORTANT: Re-register hotkeys after loading is complete
                    if self.keybinds_enabled:
                        self.unregister_all_hotkeys()
                        self.register_system_wide_hotkeys()
                    
                except Exception as e:
                    print(f"Error in update_display: {e}")
            
            # Schedule update on main thread
            self.root.after(0, update_display)
            
            # Update status with loading information
            total_pages = (len(self.emoji_data) + PAGE_SIZE - 1) // PAGE_SIZE
            self.update_status(f"Emoji loading complete. Total pages: {total_pages}")
            
        except Exception as e:
            print(f"Error finalizing emoji loading: {e}")
            self.update_status("Error loading emojis")

    def apply_theme(self, theme_name):
        """Apply the selected theme to all UI elements"""
        try:
            if theme_name == "Default":
                # Reset to system defaults
                self.root.configure(bg=self.root.cget("bg"))
                
                # Reset specific widgets to system defaults
                default_bg = self.root.cget("bg")
                default_fg = "black"
                
                # Reset scrollbar colors to default
                try:
                    self.scrollbar.configure(
                        troughcolor=default_bg,
                        bg=default_bg,
                        activebackground=default_fg,
                        highlightbackground=default_bg,
                        highlightcolor=default_fg
                    )
                except Exception as e:
                    print(f"Error resetting scrollbar: {e}")
                
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
                        
                # Theme the scrollbar with an extremely minimal style
                try:
                    # Determine thumb color based on theme
                    if theme_name in ["Dark Mode", "Cyberpunk Neon", "Forest Green", "Ocean Blue"]:
                        thumb_color = "#555555"  # Slightly lighter dark gray
                        bg_color = theme.get("listbox_bg", theme["background"])
                    else:
                        thumb_color = theme.get("primary", "#888888")  # Use primary color or fallback
                        bg_color = theme.get("listbox_bg", theme["background"])
                    
                    # Configure scrollbar with extreme minimalism
                    self.scrollbar.configure(
                        # Make trough completely transparent by matching background
                        troughcolor=bg_color,
                        # Thumb color
                        bg=thumb_color,
                        # Active (dragging) color
                        activebackground=thumb_color,
                        
                        # Aggressive border and highlight removal
                        relief='flat',
                        borderwidth=0,
                        highlightthickness=0,
                        
                        # Additional style settings to ensure transparency
                        bd=0,  # Border width
                        highlightbackground=bg_color,
                        highlightcolor=bg_color,
                        
                        # Keep it narrow
                        width=10
                    )
                except Exception as e:
                    print(f"Error theming scrollbar: {e}")
            
            # Save the current theme
            self.current_theme = theme_name
            self.settings.theme = theme_name
            self.settings.save()
        
            # Apply title bar colors after a short delay to ensure window is ready
            def apply_title_bar_colors():
                try:
                    # Get the window handle using Win32 API
                    import win32gui
                    
                    # Find our window by title
                    all_windows = []
                    
                    def enum_windows_callback(hwnd, results):
                        text = win32gui.GetWindowText(hwnd)
                        if "Crystal Realms EmojiPicker" in text:
                            results.append(hwnd)
                        return True
                    
                    win32gui.EnumWindows(enum_windows_callback, all_windows)
                    
                    if all_windows:
                        hwnd = all_windows[0]
                        print(f"Found application window: {hwnd}")
                        
                        if theme_name != "Default":
                            theme = THEMES[theme_name]
                            bg_color = theme.get("background", "#FFFFFF")
                            text_color = theme.get("text", "#000000")
                            
                            # Call the standalone function to set title bar colors
                            set_window_title_bar_colors(hwnd, bg_color, text_color)
                    else:
                        print("Could not find application window by title")
                        
                        # Fallback to using winfo_id
                        hwnd = self.root.winfo_id()
                        print(f"Using Tkinter window ID as fallback: {hwnd}")
                        
                        if theme_name != "Default":
                            theme = THEMES[theme_name]
                            bg_color = theme.get("background", "#FFFFFF")
                            text_color = theme.get("text", "#000000")
                            
                            # Call the standalone function to set title bar colors
                            set_window_title_bar_colors(hwnd, bg_color, text_color)
                    
                except Exception as e:
                    print(f"Error setting window title bar colors: {e}")
                    traceback.print_exc()
            
            # Always update title bar colors when theme changes
            self.root.after(1000, apply_title_bar_colors)  # Reduced delay for better responsiveness
            self.title_colors_applied = True  # Keep this to track that we've attempted to set colors at least once
        
        except Exception as e:
            print(f"Error applying theme {theme_name}: {e}")
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
        """Handle action when an emoji is selected (copy or direct paste) with immediate UI update"""
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

        # Update the specific item immediately without resorting and redisplaying page
        if update_ui:
            # Check if the emoji is on the current page
            start_idx = self.current_page * PAGE_SIZE
            end_idx = start_idx + PAGE_SIZE
            
            # Find the emoji's position in filtered_emojis
            try:
                abs_idx = next(i for i, e in enumerate(self.filtered_emojis) 
                              if e.unicode_key == emoji_data.unicode_key)
                              
                # If it's on the current page, update it in place
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
            except StopIteration:
                pass  # Emoji not found in filtered list

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
        """Save emoji data with names to AppData file"""
        data = {
            unicode_key: {
                'name': emoji.name,  # Added name to saved data
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

    def test_title_bar_colors(self):
        """
        Test function to try different color combinations for the title bar.
        This helps diagnose if title bar coloring is working on your current system.
        """
        
        # Create a test window
        test_win = tk.Toplevel(self.root)
        test_win.title("Title Bar Color Test")
        test_win.geometry("500x400")
        test_win.transient(self.root)
        
        # Create a frame for the controls
        frame = Frame(test_win, padx=20, pady=20)
        frame.pack(fill="both", expand=True)
        
        # Add a label
        Label(frame, text="Test different title bar colors", font=("Arial", 14)).pack(pady=10)
        
        # Force the window to be shown and processed
        test_win.update_idletasks()
        test_win.deiconify()
        test_win.focus_force()
        test_win.update()
        
        # Give Windows a moment to fully create the window
        time.sleep(0.5)
        
        # Get the window handle using Win32 API (more reliable)
        def get_win_handle():
            try:
                hwnd = test_win.winfo_id()
                print(f"Tkinter window ID: {hwnd}")
                
                # Try to find the window by title
                all_windows = []
                
                def enum_windows_callback(hwnd, results):
                    text = win32gui.GetWindowText(hwnd)
                    if "Title Bar Color Test" in text:
                        results.append(hwnd)
                    return True
                
                win32gui.EnumWindows(enum_windows_callback, all_windows)
                
                if all_windows:
                    found_hwnd = all_windows[0]
                    print(f"Found window by title: {found_hwnd}")
                    return found_hwnd
                
                # As a fallback, use the Tkinter ID
                return hwnd
            except Exception as e:
                print(f"Error getting window handle: {e}")
                return None
        
        # Function to set colors
        def set_color(bg, text):
            try:
                hwnd = get_win_handle()
                if not hwnd:
                    status_label.config(text="Error: Could not get window handle")
                    return
                
                status_label.config(text=f"Using handle: {hwnd} - Setting colors...")
                test_win.update()
                
                success = set_window_title_bar_colors(hwnd, bg, text)
                status_label.config(text=f"Set colors: BG={bg}, Text={text} - Success: {success}")
            except Exception as e:
                status_label.config(text=f"Error: {str(e)}")
                print(f"Error in test: {e}")
                import traceback
                traceback.print_exc()
        
        # Create color buttons
        color_sets = [
            ("Dark Theme", "#121212", "#FFFFFF"),
            ("Light Theme", "#F8F9FA", "#121212"),
            ("Blue Theme", "#1B2F40", "#FFFFFF"),
            ("Green Theme", "#1B2B1A", "#FFFFFF"),
            ("Red Theme", "#8B0000", "#FFFFFF"),
            ("Yellow Theme", "#FFD700", "#000000"),
            ("Reset to Default", None, None)
        ]
        
        for label, bg, text in color_sets:
            Button(
                frame, 
                text=label,
                command=lambda b=bg, t=text: set_color(b, t),
                width=20,
                height=1,
                font=("Arial", 12)
            ).pack(pady=5)
        
        # Status label
        status_label = Label(frame, text="Click a button to test", font=("Arial", 12))
        status_label.pack(pady=10)
        
        # Windows Version Info
        ver_info = f"Windows Version: {platform.platform()}"
        ver_label = Label(frame, text=ver_info, font=("Arial", 10))
        ver_label.pack(pady=2)
        
        # Close button
        Button(
            frame,
            text="Close",
            command=test_win.destroy,
            width=20,
            height=1,
            font=("Arial", 12)
        ).pack(pady=10)
        
        # Center the window
        width = test_win.winfo_width()
        height = test_win.winfo_height()
        x = (test_win.winfo_screenwidth() // 2) - (width // 2)
        y = (test_win.winfo_screenheight() // 2) - (height // 2)
        test_win.geometry(f"{width}x{height}+{x}+{y}")
        
        # Set initial colors after a slight delay
        test_win.after(500, lambda: set_color("#121212", "#FFFFFF"))
    
    def create_splash_screen(self):
        """Create a simple splash screen to show while loading"""
        # Create a frame that covers the whole window
        splash_frame = Frame(self.root, bg="#121212", width=600, height=450)
        splash_frame.place(x=0, y=0, relwidth=1, relheight=1)
        
        # Add a decorative header
        Label(
            splash_frame,
            text="Crystal Realms",
            font=("Arial", 24, "bold"),
            fg="#BB86FC",
            bg="#121212"
        ).pack(pady=(80, 5))
        
        Label(
            splash_frame,
            text="EmojiPicker",
            font=("Arial", 32, "bold"),
            fg="#03DAC6",
            bg="#121212"
        ).pack(pady=(0, 40))
        
        # Add a loading animation (simple text for now)
        loading_label = Label(
            splash_frame,
            text="Loading...",
            font=("Arial", 14),
            fg="#E0E0E0",
            bg="#121212"
        )
        loading_label.pack(pady=20)
        
        # Add a unicorn emoji for fun
        Label(
            splash_frame,
            text="🦄",
            font=("Arial", 48),
            bg="#121212"
        ).pack(pady=20)
        
        # Add credit
        Label(
            splash_frame,
            text="Made By SavageTheUnicorn",
            font=("Arial", 10, "italic"),
            fg="#BB86FC",
            bg="#121212"
        ).pack(pady=(30, 0))
        
        # Show the splash frame and update the display immediately
        self.root.update()
        
        # Schedule removal of splash screen
        self.root.after(2500, lambda: splash_frame.destroy())
    
    def async_load_font(self):
        """Load font in a background thread"""
        try:
            # Load font data
            self.supported_chars = get_supported_characters()
        except Exception as e:
            print(f"Error in async font loading: {e}")
            traceback.print_exc()
            # Fallback
            self.supported_chars = get_fallback_supported_chars()

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
    
        self.scrollbar = Scrollbar(self.list_frame)
        self.scrollbar.grid(row=0, column=1, sticky="ns")
        
        # Increased emoji font size by 2px and decreased name font size by 2px
        self.listbox = Listbox(
            self.list_frame,
            yscrollcommand=self.scrollbar.set,
            font=("Arial", 20),  # Increased from 18 to 20
            activestyle='none',  # Removes underline from active selection
            height=4  # Reduced to show exactly 4 items as in screenshot 1
        )
        self.listbox.grid(row=0, column=0, sticky="nsew")
        
        self.scrollbar.config(command=self.listbox.yview)
        
        # Bind to single clicks for the improved click behavior
        self.listbox.bind("<Button-1>", self.handle_click)
        self.listbox.bind("<Button-3>", self.show_context_menu)
        
        # Status label - reduced padding
        self.status_label = Label(
            self.container,
            text="Starting unicorn shenanigans...",
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
            text="Enable keybinds",
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
    # Check for multiple instances FIRST before doing anything else
    if not prevent_multiple_instances():
        print("Application is already running. Exiting.")
        # Exit immediately
        sys.exit(0)
    
    # If we get here, this is the only instance running
    try:
        app = EmojiPicker()
        app.run()
    except Exception as e:
        print(f"Fatal error: {e}")
        import traceback
        traceback.print_exc()
        # Try to clean up if possible
        try:
            if 'app' in locals() and hasattr(app, 'tray_icon') and app.tray_icon:
                app.tray_icon.stop()
        except:
            pass