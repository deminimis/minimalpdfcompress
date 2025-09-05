# styles.py
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
        "BG": "#f0f2f5", "WIDGET_BG": "#ffffff", "WIDGET_HOVER": "#e4e6eb",
        "ACCENT": "#1877f2", "ACCENT_HOVER": "#166fe5", "TEXT": "#050505",
        "DISABLED": "#bcc0c4", "DISABLED_BG": "#e4e6eb", "SCROLL_THUMB": "#bcc0c4", "SCROLL_THUMB_HOVER": "#a8a8a8",
        "BORDER": "#ced0d4", "SUCCESS": "#42b72a", "ERROR": "#fa383e"
    },
    'dark': {
        "BG": "#18191a", "WIDGET_BG": "#242526", "WIDGET_HOVER": "#3a3b3c",
        "ACCENT": "#2d88ff", "ACCENT_HOVER": "#4296ff", "TEXT": "#e4e6eb",
        "DISABLED": "#76797c", "DISABLED_BG": "#3a3b3c", "SCROLL_THUMB": "#5a5f63", "SCROLL_THUMB_HOVER": "#6b7176",
        "BORDER": "#3e4042", "SUCCESS": "#45bd62", "ERROR": "#ff5252"
    }
}

def apply_theme(root, mode='dark'):
    style = ttk.Style(root)
    colors = PALETTES.get(mode, PALETTES['dark'])

    try: style.theme_use('clam')
    except tk.TclError: logging.warning("The 'clam' theme is not available."); return colors

    style.configure('.', background=colors["BG"], foreground=colors["TEXT"], borderwidth=0, fieldbackground=colors["WIDGET_BG"], font=(FONT_FAMILY, 10), relief='flat', highlightthickness=0)
    style.map('.', foreground=[('disabled', colors["DISABLED"])], fieldbackground=[('disabled', colors["BG"])], background=[('disabled', colors["BG"])])
    style.layout('TButton', [('Button.padding', {'sticky': 'nswe', 'children': [('Button.label', {'sticky': 'nswe'})]})])
    style.configure('TButton', padding=(10, 5), background=colors["WIDGET_HOVER"], foreground=colors["TEXT"], font=(FONT_FAMILY, 10), highlightthickness=0, borderwidth=0, relief='flat')
    style.map('TButton', background=[('pressed', colors["WIDGET_BG"]), ('active', colors["WIDGET_HOVER"])])
    style.configure('Accent.TButton', background=colors["ACCENT"], foreground="#ffffff", font=(FONT_FAMILY, 12, 'bold'))
    style.map('Accent.TButton',
        background=[
            ('disabled', colors["DISABLED_BG"]),
            ('pressed', colors["ACCENT"]),
            ('active', colors["ACCENT_HOVER"])
        ],
        foreground=[
            ('disabled', colors["DISABLED"]),
            ('pressed', "#ffffff"),
            ('active', "#ffffff")
        ]
    )
    style.configure('Outline.TButton', background=colors["WIDGET_BG"], foreground=colors["ACCENT"], borderwidth=1, bordercolor=colors["ACCENT"], highlightthickness=0, padding=(8,4), font=(FONT_FAMILY, 9))
    style.map('Outline.TButton', background=[('active', colors["WIDGET_HOVER"])], bordercolor=[('active', colors["ACCENT_HOVER"])], foreground=[('active', colors["ACCENT_HOVER"])])
    style.configure('Large.Outline.TButton', background=colors["WIDGET_BG"], foreground=colors["ACCENT"], borderwidth=1, bordercolor=colors["ACCENT"], highlightthickness=0, padding=(10, 6), font=(FONT_FAMILY, 11))
    style.map('Large.Outline.TButton', background=[('active', colors["WIDGET_HOVER"])], bordercolor=[('active', colors["ACCENT_HOVER"])], foreground=[('active', colors["ACCENT_HOVER"])])
    style.configure('TFrame', background=colors["BG"])
    style.configure('Card.TFrame', background=colors["WIDGET_BG"], relief='solid', borderwidth=1, bordercolor=colors["BORDER"])
    style.configure('AccentBG.TFrame', background=colors["ACCENT_HOVER"])
    style.configure('TLabel', background=colors["BG"], foreground=colors["TEXT"])
    style.configure('Card.TLabel', background=colors["WIDGET_BG"])
    style.configure('TEntry', relief='solid', borderwidth=1, bordercolor=colors["BORDER"], insertcolor=colors["TEXT"], padding=8)
    style.map('TEntry', bordercolor=[('focus', colors["ACCENT"])], fieldbackground=[('readonly', colors["BG"])])
    style.layout('TProgressbar', [('Progressbar.trough', {'children': [('Progressbar.pbar', {'sticky': 'nswe'})], 'sticky': 'nswe'})])
    style.configure('TProgressbar', relief='flat', pbarrelief='flat', borderwidth=0, troughcolor=colors["WIDGET_BG"], background=colors["ACCENT"])
    style.configure('TNotebook', background=colors["BG"], borderwidth=0)

    inactive_tab_fg = "#666666" if mode == 'light' else colors["DISABLED"]
    style.configure('TNotebook.Tab', background=colors["BG"], foreground=inactive_tab_fg, padding=[12, 6], borderwidth=0, font=(FONT_FAMILY, 10, 'bold'))
    style.map('TNotebook.Tab', background=[('selected', colors["WIDGET_BG"]), ('active', colors["WIDGET_HOVER"])], foreground=[('selected', colors["ACCENT"]), ('active', colors["TEXT"])])

    style.configure('TCombobox', relief='flat', borderwidth=1, bordercolor=colors["BORDER"], arrowcolor=colors["TEXT"], arrowsize=18, padding=6)
    style.map('TCombobox',
        bordercolor=[('focus', colors["ACCENT"])],
        fieldbackground=[('readonly', colors["WIDGET_BG"])],
        foreground=[('readonly', colors["TEXT"])]
    )

    style.configure('Horizontal.TScale', background=colors["WIDGET_BG"], troughcolor=colors["BORDER"], sliderrelief='flat', sliderlength=20)
    style.map('Horizontal.TScale', background=[('active', colors["ACCENT_HOVER"]), ('!active', colors["ACCENT"])])

    style.configure('Card.TRadiobutton', background=colors["WIDGET_BG"], foreground=colors["TEXT"], indicatorrelief='flat', indicatormargin=5, indicatordiameter=-1, padding=10, font=(FONT_FAMILY, 10))
    style.map('Card.TRadiobutton', background=[('active', colors["WIDGET_HOVER"]), ('selected', colors["WIDGET_BG"])], foreground=[('selected', colors["ACCENT"])])

    style.layout('Segmented.TRadiobutton', [('Radiobutton.padding', {'sticky': 'nswe', 'children': [('Radiobutton.label', {'sticky': 'nswe'})]})])
    style.configure('Segmented.TRadiobutton', indicatoron=0, relief='flat', padding=(10, 8), background=colors["WIDGET_HOVER"], foreground=colors["TEXT"], borderwidth=1, bordercolor=colors["BORDER"])
    style.map('Segmented.TRadiobutton',
        background=[('selected', colors["ACCENT"]), ('active', colors["WIDGET_HOVER"])],
        foreground=[('selected', "#ffffff")],
        bordercolor=[('disabled', colors["BORDER"])]
    )

    style.layout('Position.TRadiobutton', [('Radiobutton.padding', {'sticky': 'nswe', 'children': [('Radiobutton.indicator', {'sticky': 'nswe'})]})])
    style.configure('Position.TRadiobutton', indicatoron=0, relief='flat', padding=15, background=colors["WIDGET_BG"], borderwidth=1, bordercolor=colors["BORDER"])
    style.map('Position.TRadiobutton',
        background=[('selected', colors["ACCENT"]), ('active', colors["WIDGET_HOVER"])],
        bordercolor=[('selected', colors["ACCENT_HOVER"])]
    )

    style.configure('Treeview', rowheight=25, background=colors["WIDGET_BG"], fieldbackground=colors["WIDGET_BG"], foreground=colors["TEXT"], relief='solid', borderwidth=1, bordercolor=colors["BORDER"])
    style.map('Treeview', background=[('selected', colors["ACCENT"])], foreground=[('selected', "#ffffff")])
    style.configure('Treeview.Heading', background=colors["BG"], foreground=colors["TEXT"], font=(FONT_FAMILY, 10, 'bold'), relief='flat', borderwidth=0)
    style.map('Treeview.Heading', background=[('active', colors["WIDGET_HOVER"])])

    root.option_add('*TCombobox*Listbox.background', colors["WIDGET_BG"])
    root.option_add('*TCombobox*Listbox.foreground', colors["TEXT"])
    root.option_add('*TCombobox*Listbox.selectBackground', colors["ACCENT"])
    root.option_add('*TCombobox*Listbox.selectForeground', "#ffffff")
    root.option_add('*TCombobox*Listbox.relief', 'flat')
    root.option_add('*TCombobox*Listbox.highlightThickness', 0)
    root.option_add('*TCombobox*Listbox.borderwidth', 0)

    return colors

