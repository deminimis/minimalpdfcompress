import tkinter as tk
from gui import GhostscriptGUI
import sys
import ctypes

if __name__ == "__main__":
    if sys.platform == 'win32':
        try:
            ctypes.windll.shcore.SetProcessDpiAwareness(2)
        except (AttributeError, OSError):
            try:
                ctypes.windll.user32.SetProcessDPIAware(True)
            except:
                pass

    root = tk.Tk()
    app = GhostscriptGUI(root)
    root.mainloop()
