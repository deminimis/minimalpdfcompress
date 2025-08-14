# ui_components.py
import tkinter as tk
from tkinter import ttk

class Tooltip:
    def __init__(self, widget, text):
        self.widget, self.text, self.tooltip_window = widget, text, None
        self.widget.bind("<Enter>", self.show_tooltip)
        self.widget.bind("<Leave>", self.hide_tooltip)

    def show_tooltip(self, event=None):
        if not self.widget.winfo_exists() or not self.text: return
        x, y, _, _ = self.widget.bbox("insert")
        x += self.widget.winfo_rootx() + 25
        y += self.widget.winfo_rooty() + 25
        self.tooltip_window = tk.Toplevel(self.widget)
        self.tooltip_window.wm_overrideredirect(True)
        self.tooltip_window.wm_geometry(f"+{x}+{y}")
        label = tk.Label(self.tooltip_window, text=self.text, justify='left', background="#383c40",
                         foreground="#e8eaed", relief='solid', borderwidth=1, wraplength=250, padx=8, pady=5)
        label.pack(ipadx=1)

    def hide_tooltip(self, event=None):
        if self.tooltip_window: self.tooltip_window.destroy()
        self.tooltip_window = None

class FileSelector(ttk.Frame):
    def __init__(self, parent, label_text, textvariable, browse_command, entry_tooltip, btn_tooltip):
        super().__init__(parent)
        self.columnconfigure(1, weight=1)

        label = ttk.Label(self, text=label_text)
        label.grid(row=0, column=0, sticky="w", padx=(0, 5))

        entry = ttk.Entry(self, textvariable=textvariable)
        entry.grid(row=0, column=1, sticky="we", padx=5)
        Tooltip(entry, entry_tooltip)

        button = ttk.Button(self, text="Browse", command=browse_command)
        button.grid(row=0, column=2, sticky="e", padx=(5, 0))
        Tooltip(button, btn_tooltip)