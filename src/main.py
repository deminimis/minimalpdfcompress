# main.py
import sys
import ctypes
import tkinter as tk
import logging
from pathlib import Path
from gui import GhostscriptGUI

def setup_logging():
    logging.basicConfig(
        level=logging.WARNING,
        format='%(asctime)s - %(levelname)s - [%(module)s.%(funcName)s:%(lineno)d] - %(message)s',
        encoding='utf-8'
    )

def set_dpi_awareness():
    if sys.platform == 'win32':
        try:
            ctypes.windll.shcore.SetProcessDpiAwareness(2)
        except (AttributeError, OSError):
            try:
                ctypes.windll.user32.SetProcessDPIAware()
            except Exception as e:
                logging.warning(f"Failed to set DPI awareness: {e}")

def main():
    setup_logging()
    set_dpi_awareness()
    try:
        root = tk.Tk()
        app = GhostscriptGUI(root)
        root.mainloop()
    except Exception as e:
        logging.critical("An unhandled exception occurred in the main application.", exc_info=True)

if __name__ == "__main__":
    main()
