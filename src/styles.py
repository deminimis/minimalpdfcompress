import tkinter as tk
from tkinter import ttk
import logging
import sys

if sys.platform == "win32":
    FONT_FAMILY = "Segoe UI"
elif sys.platform == "darwin":
    FONT_FAMILY = "Helvetica"
else:
    FONT_FAMILY = "DejaVu Sans"

PALETTES = {
    'light': {
        # Core application colors (baby blue accent)
        "BG": "#f8f6f2",
        "WIDGET_BG": "#ffffff",
        "WIDGET_HOVER": "#f0f0f0",
        "ACCENT": "#b9c5d0",
        "ACCENT_HOVER": "#a8b8c8",
        "TEXT": "#1c1e21",
        "DISABLED": "#a8a8a8",
        "SCROLL_THUMB": "#c1c1c1",
        "SCROLL_THUMB_HOVER": "#a8a8a8",
        "BORDER": "#ced4da",

        # Semantic colors for custom widgets
        "primary": "#2780e3",
        "secondary": "#7E8081",
        "success": "#3fb618",
        "info": "#9954bb",
        "warning": "#ff7518",
        "danger": "#ff0039",
    },
    'dark': {
        # Core application colors (nice green accent)
        "BG": "#2D3033",
        "WIDGET_BG": "#383c40",
        "WIDGET_HOVER": "#44484c",
        "ACCENT": "#39463e",
        "ACCENT_HOVER": "#4a5e53",
        "TEXT": "#e8eaed",
        "DISABLED": "#76797c",
        "SCROLL_THUMB": "#5a5f63",
        "SCROLL_THUMB_HOVER": "#6b7176",
        "BORDER": "#495057",

        # Semantic colors for custom widgets
        "primary": "#2780e3",
        "secondary": "#6c757d",
        "success": "#28a745",
        "info": "#17a2b8",
        "warning": "#ffc107",
        "danger": "#dc3545",
    }
}

def apply_theme(root, mode='dark'):
    style = ttk.Style(root)
    colors = PALETTES.get(mode, PALETTES['dark'])

    try:
        style.theme_use('clam')
    except tk.TclError:
        logging.warning("The 'clam' theme is not available. Custom theme may not apply correctly.")
        return colors

    style.configure('.',
                    background=colors["BG"],
                    foreground=colors["TEXT"],
                    borderwidth=0,
                    fieldbackground=colors["WIDGET_BG"],
                    font=(FONT_FAMILY, 10),
                    relief='flat',
                    highlightthickness=0)
    
    style.map('.',
              foreground=[('disabled', colors["DISABLED"])],
              fieldbackground=[('disabled', colors["BG"])],
              background=[('disabled', colors["BG"])])

    style.layout('TButton', [
        ('Button.padding', {'sticky': 'nswe', 'children': [
            ('Button.label', {'sticky': 'nswe'})
        ]})
    ])
    
    style.configure('TButton', 
                    padding=(10, 5), 
                    background=colors["ACCENT"], 
                    foreground=colors["TEXT"],
                    font=(FONT_FAMILY, 10, 'bold'),
                    highlightthickness=0)
                    
    style.map('TButton',
              background=[('pressed', colors["WIDGET_BG"]), ('active', colors["ACCENT_HOVER"])],
              foreground=[('pressed', colors["TEXT"])])

    style.configure('TFrame', background=colors["BG"])
    style.configure('TLabel', background=colors["BG"], foreground=colors["TEXT"])
    style.configure('TLabelFrame', relief='flat', borderwidth=1, bordercolor=colors["BORDER"], padding=(12, 12))
    style.configure('TLabelFrame.Label', 
                    relief='flat', 
                    background=colors["BG"], 
                    foreground=colors["TEXT"], 
                    font=(FONT_FAMILY, 11, 'bold'))

    style.configure('TEntry',
                    relief='flat',
                    borderwidth=1,
                    bordercolor=colors["BORDER"],
                    insertcolor=colors["TEXT"],
                    padding=6)
                    
    style.map('TEntry',
              bordercolor=[('focus', colors["ACCENT"])],
              fieldbackground=[('readonly', colors["BG"])])
    
    style.configure('TCombobox',
                    relief='flat',
                    borderwidth=1,
                    bordercolor=colors["BORDER"],
                    arrowcolor=colors["TEXT"],
                    arrowsize=18,
                    padding=6)
                    
    style.map('TCombobox',
              bordercolor=[('focus', colors["ACCENT"])],
              fieldbackground=[('active', colors["WIDGET_HOVER"]), ('readonly', colors["BG"])],
              background=[('active', colors["WIDGET_BG"])])

    style.configure('TCheckbutton',
                    indicatorrelief='flat',
                    indicatordiameter=16,
                    padding=5,
                    highlightthickness=0)
                    
    style.map('TCheckbutton',
              background=[('active', colors["BG"])],
              indicatorbackground=[
                  ('selected', colors["ACCENT"]),
                  ('!selected', colors["WIDGET_BG"])
              ],
              indicatorforeground=[('selected', colors["TEXT"])])

    style.configure('TScrollbar',
                    relief='flat',
                    borderwidth=0,
                    background=colors["BG"],
                    troughcolor=colors["BG"],
                    arrowsize=16,
                    arrowcolor=colors["TEXT"])
    style.map('TScrollbar',
              arrowcolor=[('disabled', colors["DISABLED"])],
              background=[('disabled', colors["BG"])])

    style.configure('Vertical.TScrollbar', background=colors["SCROLL_THUMB"], width=12)
    style.map('Vertical.TScrollbar', background=[
        ('pressed', colors["SCROLL_THUMB_HOVER"]),
        ('active', colors["SCROLL_THUMB_HOVER"])
    ])

    style.configure('Horizontal.TScrollbar', background=colors["SCROLL_THUMB"], height=12)
    style.map('Horizontal.TScrollbar', background=[
        ('pressed', colors["SCROLL_THUMB_HOVER"]),
        ('active', colors["SCROLL_THUMB_HOVER"])
    ])
    
    style.layout('TProgressbar', [
        ('Progressbar.trough', {'children': [
            ('Progressbar.pbar', {'sticky': 'nswe'})
        ], 'sticky': 'nswe'})
    ])
    style.configure('TProgressbar',
                    relief='flat',
                    pbarrelief='flat',
                    borderwidth=0,
                    troughcolor=colors["WIDGET_BG"],
                    background=colors["ACCENT"])
                    
    root.option_add('*TCombobox*Listbox.background', colors["WIDGET_BG"])
    root.option_add('*TCombobox*Listbox.foreground', colors["TEXT"])
    root.option_add('*TCombobox*Listbox.selectBackground', colors["ACCENT"])
    root.option_add('*TCombobox*Listbox.selectForeground', colors["TEXT"])
    root.option_add('*TCombobox*Listbox.font', (FONT_FAMILY, 10))
    root.option_add('*TCombobox*Listbox.relief', 'flat')
    root.option_add('*Menu.background', colors["WIDGET_BG"])
    root.option_add('*Menu.foreground', colors["TEXT"])
    root.option_add('*Menu.selectBackground', colors["ACCENT"])
    root.option_add('*Menu.activeBackground', colors["WIDGET_HOVER"])
    root.option_add('*Menu.relief', 'flat')

    return colors
