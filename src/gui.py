import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path
import logging
import threading

import backend
import styles

class GhostscriptGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Minimal PDF Compress")
        self.root.minsize(600, 550)

        # --- Variables ---
        self.input_path = tk.StringVar()
        self.output_path = tk.StringVar()
        self.operation = tk.StringVar(value="Compress (Screen - Smallest Size)")
        self.status = tk.StringVar(value="Ready")
        self.is_folder = False
        self.show_advanced = tk.BooleanVar(value=False)
        self.dark_mode_enabled = tk.BooleanVar(value=True)
        self.use_final_compression = tk.BooleanVar(value=True)
        self.adv_options = {
            'resolution': tk.StringVar(value="150"),
            'downscale_factor': tk.StringVar(value="1"),
            'pdfa_compression': tk.BooleanVar(value=False),
            'color_strategy': tk.StringVar(value="LeaveColorUnchanged"),
            'downsample_type': tk.StringVar(value="Bicubic"),
            'fast_web_view': tk.BooleanVar(value=False),
            'subset_fonts': tk.BooleanVar(value=True),
            'compress_fonts': tk.BooleanVar(value=True),
            'rotation': tk.StringVar(value="No Rotation"),
            'strip_metadata': tk.BooleanVar(value=False)
        }
        
        self.palette = {}
        self.toggle_theme()

        try:
            icon_path = backend.resource_path("pdf.ico")
            self.root.iconbitmap(icon_path)
        except Exception as e:
            logging.warning(f"Failed to set icon: {e}")

        # --- Main Frame ---
        self.main_frame = ttk.Frame(self.root, padding="20")
        self.main_frame.grid(row=0, column=0, sticky="nsew")
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        self.main_frame.columnconfigure(0, weight=1)

        logging.basicConfig(filename="ghostscript_gui.log", level=logging.INFO,
                           format="%(asctime)s - %(levelname)s - %(message)s")

        self.build_gui()

        self.gs_path = backend.find_ghostscript()
        if not self.gs_path:
            backend.show_ghostscript_download_popup(self.root)
            self.status.set("Error: Ghostscript not found")

    def build_gui(self):
        # --- Input Section ---
        input_frame = ttk.LabelFrame(self.main_frame, text="Input")
        input_frame.grid(row=0, column=0, sticky="nsew", pady=5)
        input_frame.columnconfigure(1, weight=1)
        ttk.Button(input_frame, text="Input File or Folder", command=self.select_input).grid(row=0, column=0, padx=5, pady=5)
        ttk.Entry(input_frame, textvariable=self.input_path).grid(row=0, column=1, sticky="ew", padx=5, pady=5)

        # --- Output Section ---
        output_frame = ttk.LabelFrame(self.main_frame, text="Output")
        output_frame.grid(row=1, column=0, sticky="nsew", pady=5)
        output_frame.columnconfigure(1, weight=1)
        ttk.Button(output_frame, text="Output Location", command=self.browse_output).grid(row=0, column=0, padx=5, pady=5)
        ttk.Entry(output_frame, textvariable=self.output_path).grid(row=0, column=1, sticky="ew", padx=5, pady=5)

        # --- Operation Section ---
        op_frame = ttk.LabelFrame(self.main_frame, text="Operation")
        op_frame.grid(row=2, column=0, sticky="nsew", pady=5)
        op_frame.columnconfigure(0, weight=1)
        
        operations = [
            "Compress (Screen - Smallest Size)",
            "Compress (Ebook - Medium Size)",
            "Compress (Printer - High Quality)",
            "Compress (Prepress - Highest Quality)",
            "Convert to PDF/A"
        ]
        
        op_combo = ttk.Combobox(op_frame, textvariable=self.operation, values=operations, state="readonly")
        op_combo.grid(row=0, column=0, padx=5, pady=5, sticky="ew")
        
        ttk.Checkbutton(op_frame, text="Advanced Options", variable=self.show_advanced, command=self.toggle_advanced).grid(row=0, column=1, padx=5, pady=5)
        
        action_frame = ttk.Frame(self.main_frame)
        action_frame.grid(row=3, column=0, sticky="ew", pady=(10, 5), padx=5)
        action_frame.columnconfigure(0, weight=1)

        ttk.Checkbutton(action_frame, text="Dark Mode", variable=self.dark_mode_enabled, command=self.toggle_theme).grid(row=0, column=0, sticky="w")
        
        self.process_button = ttk.Button(action_frame, text="Process", command=self.process)
        self.process_button.grid(row=0, column=1, sticky="e")

        # --- Progress Bar ---
        self.progress_bar = ttk.Progressbar(self.main_frame)
        self.progress_bar.grid(row=4, column=0, sticky="ew", pady=(5, 10), padx=5)
        self.progress_bar.grid_remove() 

        # --- Advanced Options Frame ---
        self.advanced_frame = ttk.LabelFrame(self.main_frame, text="Advanced Options")
        self.advanced_frame.grid(row=5, column=0, sticky="nsew", pady=5) 
        self.advanced_frame.columnconfigure(1, weight=1)
        self.build_advanced_options()
        self.advanced_frame.grid_remove()

        # --- Status Section ---
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
        rotation_values = ["No Rotation", "90° Right (Clockwise)", "90° Left (Counter-Clockwise)", "180°"]
        ttk.Combobox(pike_frame, textvariable=options['rotation'], values=rotation_values, state="readonly").grid(row=0, column=1, sticky="ew", padx=5, pady=2)
        ttk.Checkbutton(pike_frame, text="Remove All Metadata", variable=options['strip_metadata']).grid(row=1, column=0, columnspan=2, sticky="w", padx=5, pady=2)
        ttk.Checkbutton(pike_frame, text="Apply traditional compression", variable=self.use_final_compression).grid(row=2, column=0, columnspan=2, sticky="w", padx=5, pady=2)
        
        self.operation.trace_add("write", self.update_advanced_options_visibility)
        self.update_advanced_options_visibility()

    def toggle_theme(self):
        mode = 'dark' if self.dark_mode_enabled.get() else 'light'
        self.palette = styles.apply_theme(self.root, mode)
        self.root.configure(background=self.palette["BG"])

    def update_advanced_options_visibility(self, *args):
        is_pdfa = self.operation.get() == "Convert to PDF/A"
        self.pdfa_compression_check.grid() if is_pdfa else self.pdfa_compression_check.grid_remove()
        
        if is_pdfa:
            if self.adv_options['downscale_factor'].get() > "2":
                self.adv_options['downscale_factor'].set("2")
                messagebox.showinfo("Info", "Downscaling factor limited to 2 for PDF/A to ensure quality.")
            if self.adv_options['color_strategy'].get() != "RGB":
                self.adv_options['color_strategy'].set("RGB")
                messagebox.showinfo("Info", "Color conversion strategy set to RGB for PDF/A compliance.")

    def toggle_advanced(self):
        self.advanced_frame.grid() if self.show_advanced.get() else self.advanced_frame.grid_remove()

    def select_input(self):
        dialog = tk.Toplevel(self.root)
        dialog.title("Select Input")
        dialog.geometry("300x150")
        dialog.configure(background=self.palette["BG"])
        
        ttk.Label(dialog, text="Choose input type:").pack(pady=10)
        ttk.Button(dialog, text="File", command=lambda: [dialog.destroy(), self.browse_file()]).pack(pady=5, padx=20, fill='x')
        ttk.Button(dialog, text="Folder", command=lambda: [dialog.destroy(), self.browse_folder()]).pack(pady=5, padx=20, fill='x')

    def browse_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if file_path:
            self.input_path.set(file_path)
            self.is_folder = False
            self.output_path.set(Path(file_path).with_stem(Path(file_path).stem + "_out"))

    def browse_folder(self):
        folder_path = filedialog.askdirectory()
        if folder_path:
            self.input_path.set(folder_path)
            self.is_folder = True
            self.output_path.set(folder_path)

    def browse_output(self):
        if self.is_folder:
            path = filedialog.askdirectory()
        else:
            path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
        if path:
            self.output_path.set(path)

    def process(self):
        if not self.gs_path:
            backend.show_ghostscript_download_popup(self.root)
            self.status.set("Error: Ghostscript not found")
            return

        input_p = self.input_path.get()
        output_p = self.output_path.get()
        if not input_p or not output_p:
            messagebox.showwarning("Warning", "Please select input and output locations.")
            return

        if self.is_folder:
            if not messagebox.askyesno("Confirm Batch Processing", "This will process every .pdf file in the folder and its subfolders. Continue?"):
                return
        
        rotation_map = {
            "No Rotation": 0,
            "90° Right (Clockwise)": 90,
            "180°": 180,
            "90° Left (Counter-Clockwise)": 270
        }
        rotation_angle = rotation_map.get(self.adv_options['rotation'].get(), 0)

        params = {
            'gs_path': self.gs_path,
            'input_path': input_p,
            'output_path': str(output_p),
            'operation': self.operation.get(),
            'options': {k: v.get() for k, v in self.adv_options.items()},
            'use_final_compression': self.use_final_compression.get()
        }
        params['options']['rotation'] = rotation_angle
        
        self.process_button.config(state="disabled")
        self.progress_bar.grid() 
        self.progress_bar.start() 
        
        thread = threading.Thread(
            target=backend.run_processing_task,
            args=(params, self.is_folder, self.status, self.on_task_complete),
            daemon=True
        )
        thread.start()

    def on_task_complete(self):
        self.process_button.config(state="normal")
        self.progress_bar.stop()
        self.progress_bar.grid_remove()