import tkinter as tk
from gui import GhostscriptGUI

if __name__ == "__main__":
    """
    Main entry point for the Minimal PDF Compress application.
    Initializes the Tkinter root window and the main application GUI.
    """
    root = tk.Tk()
    app = GhostscriptGUI(root)
    root.mainloop()