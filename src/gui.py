import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path
import logging
import threading
import webbrowser
import os
import sys
import json

import backend
import styles

IS_WINDOWS = sys.platform == "win32"
if IS_WINDOWS:
    try:
        import windnd
    except ImportError:
        print("Warning: windnd library not found. Drag and drop will be disabled.")
        IS_WINDOWS = False

OP_COMPRESS_SCREEN = "Compress (Screen - Smallest Size)"
OP_COMPRESS_EBOOK = "Compress (Ebook - Medium Size)"
OP_COMPRESS_PRINTER = "Compress (Printer - High Quality)"
OP_COMPRESS_PREPRESS = "Compress (Prepress - Highest Quality)"
OP_CONVERT_PDFA = "Convert to PDF/A"

OPERATIONS = [
    OP_COMPRESS_SCREEN, OP_COMPRESS_EBOOK, OP_COMPRESS_PRINTER,
    OP_COMPRESS_PREPRESS, OP_CONVERT_PDFA
]

ROTATION_MAP = {
    "No Rotation": 0, "90° Right (Clockwise)": 90,
    "180°": 180, "90° Left (Counter-Clockwise)": 270
}

class GhostscriptGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Minimal PDF Compress")
        self.root.minsize(600, 550)

        self.settings_file = Path("settings.json")

        self.input_path = tk.StringVar()
        self.output_path = tk.StringVar()
        self.operation = tk.StringVar()
        self.status = tk.StringVar(value="Ready")
        self.is_folder = False
        self.show_advanced = tk.BooleanVar()
        self.dark_mode_enabled = tk.BooleanVar()
        self.use_final_compression = tk.BooleanVar()
        self.overwrite_originals = tk.BooleanVar()
        self.adv_options = {
            'resolution': tk.StringVar(), 'downscale_factor': tk.StringVar(), 'pdfa_compression': tk.BooleanVar(),
            'color_strategy': tk.StringVar(), 'downsample_type': tk.StringVar(), 'fast_web_view': tk.BooleanVar(),
            'subset_fonts': tk.BooleanVar(), 'compress_fonts': tk.BooleanVar(), 'rotation': tk.StringVar(),
            'strip_metadata': tk.BooleanVar()
        }
        
        self.palette = {}
        self.load_settings()
        self.toggle_theme()

        self.icon_path = backend.resource_path("pdf.ico")
        try:
            if self.icon_path.exists():
                self.root.iconbitmap(self.icon_path)
        except Exception as e:
            logging.warning(f"Failed to set icon: {e}")

        self.main_frame = ttk.Frame(self.root, padding="20")
        self.main_frame.grid(row=0, column=0, sticky="nsew")
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        self.main_frame.columnconfigure(0, weight=1)

        logging.basicConfig(filename="ghostscript_gui.log", level=logging.INFO,
                           format="%(asctime)s - %(levelname)s - %(message)s")

        self.build_gui()
        self.setup_drag_and_drop()
        self.check_ghostscript()
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

    def on_closing(self):
        self.save_settings()
        self.root.destroy()

    def save_settings(self):
        settings = {
            'operation': self.operation.get(), 'show_advanced': self.show_advanced.get(),
            'dark_mode_enabled': self.dark_mode_enabled.get(), 'use_final_compression': self.use_final_compression.get(),
            'overwrite_originals': self.overwrite_originals.get(),
            'adv_options': {key: var.get() for key, var in self.adv_options.items()}
        }
        try:
            with open(self.settings_file, 'w') as f:
                json.dump(settings, f, indent=4)
        except Exception as e:
            logging.error(f"Failed to save settings: {e}")

    def load_settings(self):
        defaults = {
            'operation': OP_COMPRESS_SCREEN, 'show_advanced': False, 'dark_mode_enabled': True,
            'use_final_compression': True, 'overwrite_originals': False,
            'adv_options': {
                'resolution': "150", 'downscale_factor': "1", 'pdfa_compression': False,
                'color_strategy': "LeaveColorUnchanged", 'downsample_type': "Bicubic", 'fast_web_view': False,
                'subset_fonts': True, 'compress_fonts': True, 'rotation': "No Rotation", 'strip_metadata': False
            }
        }
        try:
            with open(self.settings_file, 'r') as f:
                settings = json.load(f)
        except (FileNotFoundError, json.JSONDecodeError):
            settings = defaults
        self.operation.set(settings.get('operation', defaults['operation']))
        self.show_advanced.set(settings.get('show_advanced', defaults['show_advanced']))
        self.dark_mode_enabled.set(settings.get('dark_mode_enabled', defaults['dark_mode_enabled']))
        self.use_final_compression.set(settings.get('use_final_compression', defaults['use_final_compression']))
        self.overwrite_originals.set(settings.get('overwrite_originals', defaults['overwrite_originals']))
        loaded_adv = settings.get('adv_options', defaults['adv_options'])
        for key, var in self.adv_options.items():
            var.set(loaded_adv.get(key, defaults['adv_options'][key]))

    def setup_drag_and_drop(self):
        if IS_WINDOWS:
            windnd.hook_dropfiles(self.root.winfo_id(), func=self.handle_drop)

    def handle_drop(self, files):
        if not files:
            return
        path_str = files[0].decode('utf-8')
        dropped_path = Path(path_str)
        if not dropped_path.exists():
            self.status.set(f"Drop failed: Path does not exist.")
            return
        if dropped_path.is_dir():
            self.input_path.set(str(dropped_path))
            self.is_folder = True
            self.output_path.set(str(dropped_path))
            self.status.set(f"Folder selected: {dropped_path.name}")
        elif dropped_path.is_file() and dropped_path.suffix.lower() == '.pdf':
            self.input_path.set(str(dropped_path))
            self.is_folder = False
            self.output_path.set(str(dropped_path.with_stem(f"{dropped_path.stem}_processed")))
            self.status.set(f"File selected: {dropped_path.name}")
        else:
            self.status.set("Drop failed: Please drop a single PDF file or a folder.")

    def _setup_toplevel(self, toplevel, title, geometry):
        toplevel.withdraw()
        toplevel.title(title)
        try:
            if self.icon_path.exists():
                toplevel.iconbitmap(self.icon_path)
        except Exception:
            pass
        toplevel.transient(self.root)
        self.root.update_idletasks()
        main_x = self.root.winfo_x()
        main_y = self.root.winfo_y()
        main_width = self.root.winfo_width()
        main_height = self.root.winfo_height()
        popup_width, popup_height = map(int, geometry.split('x'))
        x = main_x + (main_width // 2) - (popup_width // 2)
        y = main_y + (main_height // 2) - (popup_height // 2)
        toplevel.geometry(f'{geometry}+{x}+{y}')
        toplevel.deiconify()

    def check_ghostscript(self):
        try:
            self.gs_path = backend.find_ghostscript()
            self.status.set("Ready")
        except backend.GhostscriptNotFound:
            self.gs_path = None
            self.status.set("Error: Ghostscript not found")
            self.show_ghostscript_download_popup()

    def show_ghostscript_download_popup(self):
        popup = tk.Toplevel(self.root)
        ttk.Label(popup, text="Ghostscript not found. Please place it in a local 'bin'/'lib' folder\nor install it system-wide, then restart the application.", wraplength=430).pack(pady=10, padx=10)
        link = ttk.Label(popup, text="Download Ghostscript (v10.05.1)", foreground="#007bff", cursor="hand2")
        link.pack()
        link.bind("<Button-1>", lambda e: webbrowser.open("https://github.com/ArtifexSoftware/ghostpdl-downloads/releases/tag/gs10051"))
        ttk.Button(popup, text="OK", command=popup.destroy).pack(pady=10)
        self._setup_toplevel(popup, "Ghostscript Required", "450x150")
        popup.grab_set()
        self.root.wait_window(popup)

    def build_gui(self):
        input_frame = ttk.LabelFrame(self.main_frame, text="Input")
        input_frame.grid(row=0, column=0, sticky="nsew", pady=5)
        input_frame.columnconfigure(1, weight=1)
        ttk.Button(input_frame, text="Input File or Folder", command=self.select_input).grid(row=0, column=0, padx=5, pady=5)
        ttk.Entry(input_frame, textvariable=self.input_path).grid(row=0, column=1, sticky="ew", padx=5, pady=5)

        output_frame = ttk.LabelFrame(self.main_frame, text="Output")
        output_frame.grid(row=1, column=0, sticky="nsew", pady=5)
        output_frame.columnconfigure(1, weight=1)
        ttk.Button(output_frame, text="Output Location", command=self.browse_output).grid(row=0, column=0, padx=5, pady=5)
        ttk.Entry(output_frame, textvariable=self.output_path).grid(row=0, column=1, sticky="ew", padx=5, pady=5)

        op_frame = ttk.LabelFrame(self.main_frame, text="Operation")
        op_frame.grid(row=2, column=0, sticky="nsew", pady=5)
        op_frame.columnconfigure(0, weight=1)
        op_combo = ttk.Combobox(op_frame, textvariable=self.operation, values=OPERATIONS, state="readonly")
        op_combo.grid(row=0, column=0, padx=5, pady=5, sticky="ew")
        ttk.Checkbutton(op_frame, text="Advanced Options", variable=self.show_advanced, command=self.toggle_advanced).grid(row=0, column=1, padx=5, pady=5)
        
        action_frame = ttk.Frame(self.main_frame)
        action_frame.grid(row=3, column=0, sticky="ew", pady=(10, 5), padx=5)
        action_frame.columnconfigure(0, weight=1)
        ttk.Checkbutton(action_frame, text="Dark Mode", variable=self.dark_mode_enabled, command=self.toggle_theme).grid(row=0, column=0, sticky="w")
        self.process_button = ttk.Button(action_frame, text="Process", command=self.process)
        self.process_button.grid(row=0, column=1, sticky="e")

        self.progress_bar = ttk.Progressbar(self.main_frame)
        self.progress_bar.grid(row=4, column=0, sticky="ew", pady=(5, 10), padx=5)
        self.progress_bar.grid_remove() 
        self.advanced_frame = ttk.LabelFrame(self.main_frame, text="Advanced Options")
        self.advanced_frame.grid(row=5, column=0, sticky="nsew", pady=5) 
        self.advanced_frame.columnconfigure(1, weight=1)
        self.build_advanced_options()
        self.toggle_advanced()
        self.status_label = ttk.Label(self.main_frame, textvariable=self.status, anchor='center')
        self.status_label.grid(row=6, column=0, sticky="ew", pady=5, padx=5)

    def build_advanced_options(self):
        frame = self.advanced_frame
        options = self.adv_options
        gs_frame = ttk.LabelFrame(frame, text="Ghostscript Settings")
        gs_frame.grid(row=0, column=0, sticky="nsew", padx=5, pady=5)
        gs_frame.columnconfigure(1, weight=1)
        ttk.Label(gs_frame, text="Resolution (dpi):").grid(row=0, column=0, sticky="w", padx=5, pady=2)
        ttk.Combobox(gs_frame, textvariable=options['resolution'], values=["72", "150", "300"], state="readonly").grid(row=0, column=1, sticky="ew", padx=5, pady=2)
        ttk.Label(gs_frame, text="Downscaling Factor:").grid(row=1, column=0, sticky="w", padx=5, pady=2)
        ttk.Combobox(gs_frame, textvariable=options['downscale_factor'], values=["1", "2", "3"], state="readonly").grid(row=1, column=1, sticky="ew", padx=5, pady=2)
        ttk.Label(gs_frame, text="Color Conversion Strategy:").grid(row=2, column=0, sticky="w", padx=5, pady=2)
        ttk.Combobox(gs_frame, textvariable=options['color_strategy'], values=["LeaveColorUnchanged", "Gray", "RGB", "CMYK"], state="readonly").grid(row=2, column=1, sticky="ew", padx=5, pady=2)
        ttk.Label(gs_frame, text="Downsample Method:").grid(row=3, column=0, sticky="w", padx=5, pady=2)
        ttk.Combobox(gs_frame, textvariable=options['downsample_type'], values=["Subsample", "Average", "Bicubic"], state="readonly").grid(row=3, column=1, sticky="ew", padx=5, pady=2)
        ttk.Checkbutton(gs_frame, text="Enable Fast Web View", variable=options['fast_web_view']).grid(row=4, column=0, columnspan=2, sticky="w", padx=5, pady=2)
        ttk.Checkbutton(gs_frame, text="Subset Fonts", variable=options['subset_fonts']).grid(row=5, column=0, columnspan=2, sticky="w", padx=5, pady=2)
        ttk.Checkbutton(gs_frame, text="Compress Fonts", variable=options['compress_fonts']).grid(row=6, column=0, columnspan=2, sticky="w", padx=5, pady=2)
        self.pdfa_compression_check = ttk.Checkbutton(gs_frame, text="Compress PDF/A Output", variable=options['pdfa_compression'])
        self.pdfa_compression_check.grid(row=7, column=0, columnspan=2, sticky="w", padx=5, pady=2)
        
        pike_frame = ttk.LabelFrame(frame, text="Final Processing")
        pike_frame.grid(row=0, column=1, sticky="nsew", padx=5, pady=5)
        pike_frame.columnconfigure(1, weight=1)
        ttk.Label(pike_frame, text="Page Rotation:").grid(row=0, column=0, sticky="w", padx=5, pady=2)
        ttk.Combobox(pike_frame, textvariable=options['rotation'], values=list(ROTATION_MAP.keys()), state="readonly").grid(row=0, column=1, sticky="ew", padx=5, pady=2)
        ttk.Checkbutton(pike_frame, text="Remove All Metadata", variable=options['strip_metadata']).grid(row=1, column=0, columnspan=2, sticky="w", padx=5, pady=2)
        ttk.Checkbutton(pike_frame, text="Apply traditional compression", variable=self.use_final_compression).grid(row=2, column=0, columnspan=2, sticky="w", padx=5, pady=2)
        ttk.Checkbutton(pike_frame, text="Overwrite original files", variable=self.overwrite_originals).grid(row=3, column=0, columnspan=2, sticky="w", padx=5, pady=2)
        
        self.operation.trace_add("write", self.update_advanced_options_visibility)
        self.update_advanced_options_visibility()

    def toggle_theme(self):
        mode = 'dark' if self.dark_mode_enabled.get() else 'light'
        self.palette = styles.apply_theme(self.root, mode)
        self.root.configure(background=self.palette["BG"])

    def update_advanced_options_visibility(self, *args):
        is_pdfa = self.operation.get() == OP_CONVERT_PDFA
        self.pdfa_compression_check.grid() if is_pdfa else self.pdfa_compression_check.grid_remove()
        
    def toggle_advanced(self):
        if self.show_advanced.get():
            self.advanced_frame.grid()
        else:
            self.advanced_frame.grid_remove()

    def select_input(self):
        dialog = tk.Toplevel(self.root)
        dialog.configure(background=self.palette["BG"])
        ttk.Label(dialog, text="Choose input type:").pack(pady=10)
        ttk.Button(dialog, text="File", command=lambda: [self.browse_file(), dialog.destroy()]).pack(pady=5, padx=20, fill='x')
        ttk.Button(dialog, text="Folder", command=lambda: [self.browse_folder(), dialog.destroy()]).pack(pady=5, padx=20, fill='x')
        self._setup_toplevel(dialog, "Select Input", "300x150")
        dialog.grab_set()
        self.root.wait_window(dialog)

    def browse_file(self):
        file_path = filedialog.askopenfilename(parent=self.root, filetypes=[("PDF files", "*.pdf")])
        if file_path:
            self.is_folder = False
            self.input_path.set(file_path)
            p = Path(file_path)
            self.output_path.set(str(p.with_stem(f"{p.stem}_processed")))

    def browse_folder(self):
        folder_path = filedialog.askdirectory(parent=self.root)
        if folder_path:
            self.is_folder = True
            self.input_path.set(folder_path)
            self.output_path.set(folder_path)

    def browse_output(self):
        if self.is_folder:
            path = filedialog.askdirectory(parent=self.root)
        else:
            path = filedialog.asksaveasfilename(parent=self.root, defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
        if path:
            self.output_path.set(path)

    def process(self):
        if not self.gs_path:
            self.check_ghostscript()
            return
        input_p_str = self.input_path.get()
        if not input_p_str:
            messagebox.showwarning("Warning", "Please select an input file or folder.", parent=self.root)
            return
        is_overwrite = self.overwrite_originals.get()
        output_p_str = ""
        if is_overwrite:
            output_p_str = input_p_str
            warn_msg = f"You are about to process and OVERWRITE your original file(s) at:\n\n{input_p_str}\n\nThis cannot be undone. Are you sure you want to continue?"
            if not messagebox.askyesno("Confirm Overwrite", warn_msg, icon='warning', parent=self.root):
                return
        else:
            if self.is_folder:
                output_p_str = self.output_path.get()
                if not output_p_str or not Path(output_p_str).is_dir():
                     messagebox.showwarning("Warning", "Please select a valid output folder.", parent=self.root)
                     return
            else:
                input_path_obj = Path(input_p_str)
                output_p_str = str(input_path_obj.with_stem(f"{input_path_obj.stem}_processed"))
        if self.is_folder and not is_overwrite:
            if not messagebox.askyesno("Confirm Batch Processing", "This will process every .pdf file in the selected folder and its subfolders. Continue?", parent=self.root):
                return
        params = {
            'gs_path': self.gs_path, 'input_path': input_p_str, 'output_path': output_p_str,
            'operation': self.operation.get(), 'options': {k: v.get() for k, v in self.adv_options.items()},
            'use_final_compression': self.use_final_compression.get(), 'overwrite': is_overwrite
        }
        params['options']['rotation'] = ROTATION_MAP.get(self.adv_options['rotation'].get(), 0)
        self.process_button.config(state="disabled")
        self.progress_bar.grid() 
        self.progress_bar.start() 
        threading.Thread(
            target=backend.run_processing_task,
            args=(params, self.is_folder, self.status, self.on_task_complete),
            daemon=True
        ).start()

    def on_task_complete(self):
        self.process_button.config(state="normal")
        self.progress_bar.stop()
        self.progress_bar.grid_remove()