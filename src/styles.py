import tkinter as tk
from tkinter import ttk
import logging

# --- Color Palettes ---
DARK_PALETTE = {
    "BG": "#2D3033",
    "WIDGET_BG": "#383c40",
    "WIDGET_HOVER": "#44484c",
    "ACCENT": "#32503c",
    "ACCENT_HOVER": "#3c6048",
    "TEXT": "#e8eaed",
    "DISABLED": "#76797c",
    "SCROLL_THUMB": "#5a5f63",
    "SCROLL_THUMB_HOVER": "#6b7176",
}

LIGHT_PALETTE = {
    "BG": "#f8f6f2",
    "WIDGET_BG": "#ffffff",
    "WIDGET_HOVER": "#f0f0f0",
    "ACCENT": "#a8c5e0",
    "ACCENT_HOVER": "#b8d5f0",
    "TEXT": "#1c1e21",
    "DISABLED": "#a8a8a8",
    "SCROLL_THUMB": "#c1c1c1",
    "SCROLL_THUMB_HOVER": "#a8a8a8",
}

def apply_theme(root, mode='dark'):
    # --- A Modern, Dual-Mode Theme ---
    style = ttk.Style(root)
    colors = DARK_PALETTE if mode == 'dark' else LIGHT_PALETTE

    # --- Theme and Global Defaults ---
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
                    font=('Segoe UI', 10),
                    relief='flat',
                    highlightthickness=0)
    
    style.map('.',
              foreground=[('disabled', colors["DISABLED"])],
              fieldbackground=[('disabled', colors["BG"])],
              background=[('disabled', colors["BG"])])

    # --- Button ---
    style.layout('TButton', [
        ('Button.padding', {'sticky': 'nswe', 'children': [
            ('Button.label', {'sticky': 'nswe'})
        ]})
    ])
    
    style.configure('TButton', 
                    padding=(10, 5), 
                    background=colors["ACCENT"], 
                    foreground=colors["TEXT"],
                    font=('Segoe UI', 10, 'bold'),
                    highlightthickness=0)
                    
    style.map('TButton',
              background=[('pressed', colors["WIDGET_BG"]), ('active', colors["ACCENT_HOVER"])],
              foreground=[('pressed', colors["TEXT"])])

    # --- Frames and Labels ---
    style.configure('TFrame', background=colors["BG"])
    style.configure('TLabel', background=colors["BG"], foreground=colors["TEXT"])
    style.configure('TLabelFrame', relief='flat', borderwidth=1, bordercolor=colors["WIDGET_BG"], padding=(12, 12))
    style.configure('TLabelFrame.Label', 
                    relief='flat', 
                    background=colors["BG"], 
                    foreground=colors["TEXT"], 
                    font=('Segoe UI', 11, 'bold'))

    # --- Entry and Combobox ---
    style.configure('TEntry',
                    relief='flat',
                    borderwidth=1,
                    bordercolor=colors["WIDGET_BG"],
                    insertcolor=colors["TEXT"],
                    padding=6)
                    
    style.map('TEntry',
              bordercolor=[('focus', colors["ACCENT"])],
              fieldbackground=[('readonly', colors["BG"])])
    
    style.configure('TCombobox',
                    relief='flat',
                    borderwidth=1,
                    bordercolor=colors["WIDGET_BG"],
                    arrowcolor=colors["TEXT"],
                    arrowsize=18,
                    padding=6)
                    
    style.map('TCombobox',
              bordercolor=[('focus', colors["ACCENT"])],
              fieldbackground=[('active', colors["WIDGET_HOVER"]), ('readonly', colors["BG"])],
              background=[('active', colors["WIDGET_BG"])])

    # --- Checkbutton ---
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

    # --- Scrollbar ---
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
    
    # --- Progress Bar ---
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
                    
    # --- Global styling for popups ---
    root.option_add('*TCombobox*Listbox.background', colors["WIDGET_BG"])
    root.option_add('*TCombobox*Listbox.foreground', colors["TEXT"])
    root.option_add('*TCombobox*Listbox.selectBackground', colors["ACCENT"])
    root.option_add('*TCombobox*Listbox.selectForeground', colors["TEXT"])
    root.option_add('*TCombobox*Listbox.font', ('Segoe UI', 10))
    root.option_add('*TCombobox*Listbox.relief', 'flat')
    root.option_add('*Menu.background', colors["WIDGET_BG"])
    root.option_add('*Menu.foreground', colors["TEXT"])
    root.option_add('*Menu.selectBackground', colors["ACCENT"])
    root.option_add('*Menu.activeBackground', colors["WIDGET_HOVER"])
    root.option_add('*Menu.relief', 'flat')

    return colors