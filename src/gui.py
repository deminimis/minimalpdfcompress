#region: gui.py
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, colorchooser
from pathlib import Path
import logging
import threading
import sys
import json

import backend
import styles
import tooltips

IS_WINDOWS = sys.platform == "win32"
if IS_WINDOWS:
    try:
        import windnd
    except ImportError:
        print("Warning: windnd library not found. Drag and drop will be disabled.")
        IS_WINDOWS = False

#region: Constants
APP_VERSION = "1.5"
OP_COMPRESS_SCREEN = "Compress (Screen - Smallest Size)"
OP_COMPRESS_EBOOK = "Compress (Ebook - Medium Size)"
OP_COMPRESS_PRINTER = "Compress (Printer - High Quality)"
OP_COMPRESS_PREPRESS = "Compress (Prepress - Highest Quality)"

OPERATIONS = [ OP_COMPRESS_SCREEN, OP_COMPRESS_EBOOK, OP_COMPRESS_PRINTER, OP_COMPRESS_PREPRESS ]

ROTATION_MAP = { "No Rotation": 0, "90° Right (Clockwise)": 90, "180°": 180, "90° Left (Counter-Clockwise)": 270 }

PDF_FONTS = [ "Times-Roman", "Times-Bold", "Times-Italic", "Times-BoldItalic",
              "Helvetica", "Helvetica-Bold", "Helvetica-Oblique", "Helvetica-BoldOblique",
              "Courier", "Courier-Bold", "Courier-Oblique", "Courier-BoldOblique",
              "Symbol", "ZapfDingbats" ]

DPI_PRESETS = { OP_COMPRESS_SCREEN: "72", OP_COMPRESS_EBOOK: "150", OP_COMPRESS_PRINTER: "300", OP_COMPRESS_PREPRESS: "300" }
#endregion

#region: Tooltip Class
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
#endregion

class GhostscriptGUI:
    def __init__(self, root):
        self.root = root
        self.root.title(f"MinimalPDF Compress v{APP_VERSION}")
        self.root.minsize(720, 750)
        self.settings_file = Path("settings.json")
        
        #region: Variable Declarations
        self.input_path = tk.StringVar()
        self.output_path = tk.StringVar()
        self.operation = tk.StringVar()
        self.is_folder = False
        self.show_advanced = tk.BooleanVar()
        self.dark_mode_enabled = tk.BooleanVar()
        self.overwrite_originals = tk.BooleanVar()
        self.pdf_to_img_input = tk.StringVar()
        self.pdf_to_img_output_dir = tk.StringVar()
        self.pdf_to_img_format = tk.StringVar(value="png")
        self.pdf_to_img_dpi = tk.StringVar(value="300")
        self.merge_files = []
        self.merge_output_path = tk.StringVar()
        self.rotate_input = tk.StringVar()
        self.rotate_output = tk.StringVar()
        self.rotate_angle = tk.StringVar(value="90° Right (Clockwise)")
        self.delete_pages_input = tk.StringVar()
        self.delete_pages_output = tk.StringVar()
        self.delete_pages_range = tk.StringVar(value="1, 3-5")
        self.split_input_path = tk.StringVar()
        self.split_output_dir = tk.StringVar()
        self.split_mode = tk.StringVar(value="Split to Single Pages")
        self.split_value = tk.StringVar(value="2")
        self.stamp_pdf_in = tk.StringVar()
        self.stamp_pdf_out = tk.StringVar()
        self.stamp_mode = tk.StringVar(value="Image")
        self.stamp_image_in = tk.StringVar()
        self.stamp_image_width = tk.StringVar()
        self.stamp_image_height = tk.StringVar()
        self.stamp_text_in = tk.StringVar(value="CONFIDENTIAL")
        self.stamp_font = tk.StringVar(value="Helvetica-Bold")
        self.stamp_font_size = tk.StringVar(value="48")
        self.stamp_font_color_var = tk.StringVar(value="#ff0000")
        self.stamp_pos = tk.StringVar(value="Center")
        self.stamp_opacity = tk.DoubleVar(value=0.5)
        self.stamp_on_top = tk.BooleanVar(value=True)
        self.stamp_dynamic_filename = tk.BooleanVar()
        self.stamp_dynamic_datetime = tk.BooleanVar()
        self.stamp_dynamic_bates = tk.BooleanVar()
        self.stamp_bates_start = tk.StringVar(value="1")
        self.meta_pdf_path = tk.StringVar()
        self.meta_title, self.meta_author, self.meta_subject, self.meta_keywords = tk.StringVar(), tk.StringVar(), tk.StringVar(), tk.StringVar()
        self.utility_pdf_in = tk.StringVar()
        self.pdfa_input = tk.StringVar()
        self.pdfa_output = tk.StringVar()
        self.status = tk.StringVar(value="Ready")
        self.adv_options = {
            'image_resolution': tk.StringVar(), 'downscale_factor': tk.StringVar(), 'color_strategy': tk.StringVar(),
            'downsample_type': tk.StringVar(), 'fast_web_view': tk.BooleanVar(), 'subset_fonts': tk.BooleanVar(), 'compress_fonts': tk.BooleanVar(),
            'rotation': tk.StringVar(), 'strip_metadata': tk.BooleanVar(), 'remove_interactive': tk.BooleanVar(), 'pikepdf_compression_level': tk.IntVar(),
            'decimal_precision': tk.StringVar(), 'use_cpdf_squeeze': tk.BooleanVar(), 'darken_text': tk.BooleanVar(), 'use_fast_processing': tk.BooleanVar(),
            'user_password': tk.StringVar(), 'owner_password': tk.StringVar(), 'show_passwords': tk.BooleanVar(),
        }
        #endregion
        
        self.palette, self.gs_path, self.cpdf_path = {}, None, None
        self.active_process_button = None
        
        self.icon_path = backend.resource_path("pdf.ico")
        try:
            if self.icon_path.exists(): self.root.iconbitmap(self.icon_path)
        except Exception as e: logging.warning(f"Failed to set icon: {e}")

        self.main_frame = ttk.Frame(self.root, padding="10"); self.main_frame.grid(row=0, column=0, sticky="nsew")
        self.root.columnconfigure(0, weight=1); self.root.rowconfigure(0, weight=1)
        self.main_frame.columnconfigure(0, weight=1); self.main_frame.rowconfigure(0, weight=1)

        logging.basicConfig(filename="app.log", level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")

        self.load_settings()
        self.build_gui()
        self.toggle_theme() 
        self.setup_drag_and_drop()
        self.check_tools()
        self.adv_options['user_password'].trace_add("write", self.update_final_processor_state)
        self.adv_options['owner_password'].trace_add("write", self.update_final_processor_state)
        self.adv_options['use_cpdf_squeeze'].trace_add("write", self.update_final_processor_state)
        self.adv_options['darken_text'].trace_add("write", self.update_final_processor_state)
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

    #region: Settings Management
    def on_closing(self):
        self.save_settings()
        self.root.destroy()

    def save_settings(self):
        settings = {
            'input_path': self.input_path.get(), 'output_path': self.output_path.get(), 'operation': self.operation.get(),
            'dark_mode_enabled': self.dark_mode_enabled.get(), 'overwrite_originals': self.overwrite_originals.get(),
            'show_advanced': self.show_advanced.get(),
            'adv_options': {key: var.get() for key, var in self.adv_options.items() if key != 'show_passwords'}
        }
        try:
            with open(self.settings_file, 'w') as f:
                json.dump(settings, f, indent=4)
        except Exception as e:
            logging.error(f"Failed to save settings: {e}")

    def load_settings(self):
        defaults = {
            'input_path': '', 'output_path': '', 'operation': OP_COMPRESS_SCREEN, 'dark_mode_enabled': True,
            'overwrite_originals': False, 'show_advanced': False,
            'adv_options': { 'image_resolution': "72", 'downscale_factor': "1", 'color_strategy': "LeaveColorUnchanged", 'downsample_type': "Bicubic",
                             'fast_web_view': True, 'subset_fonts': True, 'compress_fonts': True, 'rotation': "No Rotation", 'strip_metadata': False, 'remove_interactive': False,
                             'pikepdf_compression_level': 6, 'decimal_precision': "Default", 'use_cpdf_squeeze': False, 'darken_text': False, 'use_fast_processing': False,
                             'user_password': '', 'owner_password': '' }
        }
        try:
            with open(self.settings_file, 'r') as f: settings = json.load(f)
        except (FileNotFoundError, json.JSONDecodeError): settings = defaults

        for key, default_val in defaults.items():
            if key == 'adv_options': continue
            if hasattr(self, key):
                var = getattr(self, key)
                if hasattr(var, 'set'):
                    var.set(settings.get(key, default_val))

        loaded_adv = settings.get('adv_options', defaults['adv_options'])
        for key, default_val in defaults['adv_options'].items():
            if key in self.adv_options:
                self.adv_options[key].set(loaded_adv.get(key, default_val))
    #endregion

    #region: Tool Checks & Popups
    def setup_drag_and_drop(self):
        if IS_WINDOWS:
            windnd.hook_dropfiles(self.root.winfo_id(), func=self.handle_drop)

    def handle_drop(self, files):
        if not files: return
        path_str = files[0].decode('utf-8')
        p = Path(path_str)
        if not p.exists(): self.status.set(f"Drop failed: Path does not exist."); return

        active_tab_idx = self.notebook.index(self.notebook.select())
        if p.is_dir():
            if active_tab_idx == 0:
                self.input_path.set(str(p)); self.is_folder = True; self.output_path.set(str(p))
                self.status.set(f"Folder selected: {p.name}")
        elif p.is_file() and p.suffix.lower() == '.pdf':
            self.status.set(f"File selected: {p.name}")
            if active_tab_idx == 0:
                self.input_path.set(str(p)); self.is_folder = False; self.output_path.set(str(p.with_stem(f"{p.stem}_processed")))
            elif active_tab_idx == 1:
                self.rotate_input.set(str(p)); self.rotate_output.set(str(p.with_stem(f"{p.stem}_rotated")))
                self.delete_pages_input.set(str(p)); self.delete_pages_output.set(str(p.with_stem(f"{p.stem}_deleted")))
                self.split_input_path.set(str(p)); self.split_output_dir.set(str(p.parent))
            elif active_tab_idx == 2:
                self.pdfa_input.set(str(p)); self.pdfa_output.set(str(p.with_stem(f"{p.stem}_pdfa")))
                self.stamp_pdf_in.set(str(p)); self.stamp_pdf_out.set(str(p.with_stem(f"{p.stem}_stamped")))
                self.meta_pdf_path.set(str(p)); self.load_metadata()
                self.utility_pdf_in.set(str(p))
        else:
            self.status.set("Drop failed: Please drop a single PDF file or a folder.")

    def check_tools(self):
        try: self.gs_path = backend.find_ghostscript()
        except backend.GhostscriptNotFound: self.gs_path = None; self.status.set("Error: Ghostscript not found in 'bin' folder")
        try: self.cpdf_path = backend.find_cpdf()
        except backend.CpdfNotFound: self.cpdf_path = None; self.status.set("Error: cpdf not found in 'bin' folder")
    #endregion

    #region: GUI Building
    def build_gui(self):
        self.notebook = ttk.Notebook(self.main_frame); self.notebook.grid(row=0, column=0, sticky="nsew", pady=(0, 10))
        self.compress_tab = ttk.Frame(self.notebook, padding="10")
        self.page_tools_tab = ttk.Frame(self.notebook, padding="10")
        self.tools_tab = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(self.compress_tab, text='Compress & Optimize')
        self.notebook.add(self.page_tools_tab, text='Page Tools')
        self.notebook.add(self.tools_tab, text='PDF Tools')
        self.build_compress_tab(self.compress_tab)
        self.build_page_tools_tab(self.page_tools_tab)
        self.build_tools_tab(self.tools_tab)
        status_frame = ttk.Frame(self.main_frame); status_frame.grid(row=1, column=0, sticky="nsew"); status_frame.columnconfigure(0, weight=1)
        self.progress_bar = ttk.Progressbar(status_frame, mode='determinate'); self.progress_bar.grid(row=0, column=0, sticky="ew"); self.progress_bar.grid_remove()
        self.status_label = ttk.Label(status_frame, textvariable=self.status, anchor='center'); self.status_label.grid(row=1, column=0, sticky="ew", pady=(5,0))

    def build_compress_tab(self, parent):
        parent.columnconfigure(0, weight=1)
        input_frame = ttk.LabelFrame(parent, text="Input"); input_frame.grid(row=0, column=0, sticky="nsew", pady=5); input_frame.columnconfigure(1, weight=1)
        in_btn = ttk.Button(input_frame, text="Input File or Folder", command=self.select_input); in_btn.grid(row=0, column=0, padx=5, pady=5)
        Tooltip(in_btn, tooltips.TOOLTIP_TEXT["compress_input_btn"])
        in_entry = ttk.Entry(input_frame, textvariable=self.input_path); in_entry.grid(row=0, column=1, sticky="ew", padx=5, pady=5)
        Tooltip(in_entry, tooltips.TOOLTIP_TEXT["compress_input_entry"])
        
        output_frame = ttk.LabelFrame(parent, text="Output"); output_frame.grid(row=1, column=0, sticky="nsew", pady=5); output_frame.columnconfigure(1, weight=1)
        out_btn = ttk.Button(output_frame, text="Output Location", command=self.browse_output); out_btn.grid(row=0, column=0, padx=5, pady=5)
        Tooltip(out_btn, tooltips.TOOLTIP_TEXT["compress_output_btn"])
        out_entry = ttk.Entry(output_frame, textvariable=self.output_path); out_entry.grid(row=0, column=1, sticky="ew", padx=5, pady=5)
        Tooltip(out_entry, tooltips.TOOLTIP_TEXT["compress_output_entry"])

        op_frame = ttk.LabelFrame(parent, text="Operation"); op_frame.grid(row=2, column=0, sticky="nsew", pady=5); op_frame.columnconfigure(0, weight=1)
        op_combo = ttk.Combobox(op_frame, textvariable=self.operation, values=OPERATIONS, state="readonly"); op_combo.grid(row=0, column=0, padx=5, pady=5, sticky="ew")
        op_combo.bind("<<ComboboxSelected>>", self.update_image_resolution)
        Tooltip(op_combo, tooltips.TOOLTIP_TEXT["compress_op_combo"])
        dpi_entry = ttk.Entry(op_frame, textvariable=self.adv_options['image_resolution'], width=5); dpi_entry.grid(row=0, column=1, padx=5, pady=5)
        Tooltip(dpi_entry, tooltips.TOOLTIP_TEXT["compress_dpi_entry"])
        ttk.Label(op_frame, text="DPI").grid(row=0, column=2, padx=(0, 10))
        
        adv_check = ttk.Checkbutton(op_frame, text="Advanced Options", variable=self.show_advanced, command=self.toggle_advanced); adv_check.grid(row=0, column=3, padx=5, pady=5)
        Tooltip(adv_check, tooltips.TOOLTIP_TEXT["compress_adv_check"])
        
        action_frame = ttk.Frame(parent); action_frame.grid(row=3, column=0, sticky="ew", pady=(10, 5), padx=5); action_frame.columnconfigure(0, weight=1)
        dark_mode_check = ttk.Checkbutton(action_frame, text="Dark Mode", variable=self.dark_mode_enabled, command=self.toggle_theme); dark_mode_check.grid(row=0, column=0, sticky="w")
        Tooltip(dark_mode_check, tooltips.TOOLTIP_TEXT["compress_dark_mode_check"])
        self.compress_button = ttk.Button(action_frame, text="Process", command=self.process_conversion); self.compress_button.grid(row=0, column=1, sticky="e")
        Tooltip(self.compress_button, tooltips.TOOLTIP_TEXT["compress_process_btn"])
        
        self.advanced_frame = ttk.LabelFrame(parent, text="Advanced Options"); self.advanced_frame.grid(row=5, column=0, sticky="nsew", pady=5); self.advanced_frame.columnconfigure(1, weight=1)
        self.build_advanced_options(self.advanced_frame)
        self.toggle_advanced()
        self.update_image_resolution()
        
    def build_advanced_options(self, parent):
        gs_frame = ttk.LabelFrame(parent, text="Ghostscript Settings"); gs_frame.grid(row=0, column=0, sticky="nsew", padx=5, pady=5, ipady=5); gs_frame.columnconfigure(1, weight=1)
        ttk.Label(gs_frame, text="Downscaling Factor:").grid(row=0, column=0, sticky="w", padx=5, pady=2)
        ds_factor_combo = ttk.Combobox(gs_frame, textvariable=self.adv_options['downscale_factor'], values=["1", "2", "3", "4"], state="readonly"); ds_factor_combo.grid(row=0, column=1, sticky="ew", padx=5, pady=2)
        Tooltip(ds_factor_combo, tooltips.TOOLTIP_TEXT["adv_gs_downscale_factor"])
        ttk.Label(gs_frame, text="Color Conversion:").grid(row=1, column=0, sticky="w", padx=5, pady=2)
        color_strat_combo = ttk.Combobox(gs_frame, textvariable=self.adv_options['color_strategy'], values=["LeaveColorUnchanged", "Gray", "RGB", "CMYK"], state="readonly"); color_strat_combo.grid(row=1, column=1, sticky="ew", padx=5, pady=2)
        Tooltip(color_strat_combo, tooltips.TOOLTIP_TEXT["adv_gs_color_conversion"])
        ttk.Label(gs_frame, text="Downsample Method:").grid(row=2, column=0, sticky="w", padx=5, pady=2)
        ds_method_combo = ttk.Combobox(gs_frame, textvariable=self.adv_options['downsample_type'], values=["Subsample", "Average", "Bicubic"], state="readonly"); ds_method_combo.grid(row=2, column=1, sticky="ew", padx=5, pady=2)
        Tooltip(ds_method_combo, tooltips.TOOLTIP_TEXT["adv_gs_downsample_method"])
        fwv_check = ttk.Checkbutton(gs_frame, text="Enable Fast Web View", variable=self.adv_options['fast_web_view']); fwv_check.grid(row=3, column=0, columnspan=2, sticky="w", padx=5, pady=2)
        Tooltip(fwv_check, tooltips.TOOLTIP_TEXT["adv_gs_fast_web_view"])
        subset_check = ttk.Checkbutton(gs_frame, text="Subset Fonts", variable=self.adv_options['subset_fonts']); subset_check.grid(row=4, column=0, columnspan=2, sticky="w", padx=5, pady=2)
        Tooltip(subset_check, tooltips.TOOLTIP_TEXT["adv_gs_subset_fonts"])
        compress_fonts_check = ttk.Checkbutton(gs_frame, text="Compress Fonts", variable=self.adv_options['compress_fonts']); compress_fonts_check.grid(row=5, column=0, columnspan=2, sticky="w", padx=5, pady=2)
        Tooltip(compress_fonts_check, tooltips.TOOLTIP_TEXT["adv_gs_compress_fonts"])
        remove_interactive_check = ttk.Checkbutton(gs_frame, text="Remove Annotations & Forms", variable=self.adv_options['remove_interactive']); remove_interactive_check.grid(row=6, column=0, columnspan=2, sticky="w", padx=5, pady=2)
        Tooltip(remove_interactive_check, tooltips.TOOLTIP_TEXT["adv_gs_remove_interactive"])

        final_processing_frame = ttk.LabelFrame(parent, text="Final Processing & Security"); final_processing_frame.grid(row=0, column=1, rowspan=2, sticky="nsew", padx=5, pady=5, ipady=5); final_processing_frame.columnconfigure(1, weight=1)
        
        ttk.Label(final_processing_frame, text="Page Rotation:").grid(row=0, column=0, sticky="w", padx=5, pady=2)
        rotation_combo = ttk.Combobox(final_processing_frame, textvariable=self.adv_options['rotation'], values=list(ROTATION_MAP.keys()), state="readonly"); rotation_combo.grid(row=0, column=1, sticky="ew", padx=5, pady=2)
        Tooltip(rotation_combo, tooltips.TOOLTIP_TEXT["adv_final_rotation"])
        ttk.Label(final_processing_frame, text="Decimal Precision:").grid(row=1, column=0, sticky="w", padx=5, pady=2)
        precision_combo = ttk.Combobox(final_processing_frame, textvariable=self.adv_options['decimal_precision'], values=["Default", "6", "5", "4", "3", "2"], state="readonly"); precision_combo.grid(row=1, column=1, sticky="ew", padx=5, pady=2)
        Tooltip(precision_combo, tooltips.TOOLTIP_TEXT["adv_final_precision"])
        strip_meta_check = ttk.Checkbutton(final_processing_frame, text="Remove All Metadata", variable=self.adv_options['strip_metadata']); strip_meta_check.grid(row=2, column=0, columnspan=2, sticky="w", padx=5, pady=2)
        Tooltip(strip_meta_check, tooltips.TOOLTIP_TEXT["adv_final_strip_meta"])
        overwrite_check = ttk.Checkbutton(final_processing_frame, text="Overwrite original files", variable=self.overwrite_originals); overwrite_check.grid(row=3, column=0, columnspan=2, sticky="w", padx=5, pady=2)
        Tooltip(overwrite_check, tooltips.TOOLTIP_TEXT["adv_final_overwrite"])
        cpdf_squeeze_check = ttk.Checkbutton(final_processing_frame, text="Squeeze with cpdf", variable=self.adv_options['use_cpdf_squeeze'], command=self.update_final_processor_state); cpdf_squeeze_check.grid(row=4, column=0, columnspan=2, sticky="w", padx=5, pady=2)
        Tooltip(cpdf_squeeze_check, tooltips.TOOLTIP_TEXT["adv_final_cpdf_squeeze"])
        darken_text_check = ttk.Checkbutton(final_processing_frame, text="Darken Text (cpdf)", variable=self.adv_options['darken_text'], command=self.update_final_processor_state); darken_text_check.grid(row=5, column=0, columnspan=2, sticky="w", padx=5, pady=2)
        Tooltip(darken_text_check, tooltips.TOOLTIP_TEXT["adv_final_cpdf_darken"])
        fast_proc_check = ttk.Checkbutton(final_processing_frame, text="Fast Processing (cpdf)", variable=self.adv_options['use_fast_processing'], command=self.update_final_processor_state); fast_proc_check.grid(row=6, column=0, columnspan=2, sticky="w", padx=5, pady=2)
        Tooltip(fast_proc_check, tooltips.TOOLTIP_TEXT["adv_final_cpdf_fast"])
        
        ttk.Separator(final_processing_frame, orient='horizontal').grid(row=7, column=0, columnspan=2, sticky='ew', pady=10)
        
        ttk.Label(final_processing_frame, text="User Password:").grid(row=8, column=0, sticky="w", padx=5, pady=2)
        self.user_pass_entry = ttk.Entry(final_processing_frame, textvariable=self.adv_options['user_password'], show="*"); self.user_pass_entry.grid(row=8, column=1, sticky="ew", padx=5, pady=2)
        Tooltip(self.user_pass_entry, tooltips.TOOLTIP_TEXT["adv_sec_user_pass"])
        ttk.Label(final_processing_frame, text="Owner Password:").grid(row=9, column=0, sticky="w", padx=5, pady=2)
        self.owner_pass_entry = ttk.Entry(final_processing_frame, textvariable=self.adv_options['owner_password'], show="*"); self.owner_pass_entry.grid(row=9, column=1, sticky="ew", padx=5, pady=2)
        Tooltip(self.owner_pass_entry, tooltips.TOOLTIP_TEXT["adv_sec_owner_pass"])
        show_pass_check = ttk.Checkbutton(final_processing_frame, text="Show Passwords", variable=self.adv_options['show_passwords'], command=self.toggle_password_visibility); show_pass_check.grid(row=10, column=1, sticky="w", padx=5, pady=2)
        Tooltip(show_pass_check, tooltips.TOOLTIP_TEXT["adv_sec_show_pass"])
        
        compression_frame = ttk.LabelFrame(parent, text="Pikepdf Compression"); compression_frame.grid(row=1, column=0, sticky="nsew", padx=5, pady=5); compression_frame.columnconfigure(1, weight=1)
        ttk.Label(compression_frame, text="Level:").grid(row=0, column=0, sticky="w", padx=5, pady=2)
        self.compression_label = ttk.Label(compression_frame, text="6"); self.compression_label.grid(row=0, column=2, sticky="e", padx=5, pady=2)
        self.comp_slider = ttk.Scale(compression_frame, from_=0, to=9, orient="horizontal", variable=self.adv_options['pikepdf_compression_level'], command=self.update_compression_label); self.comp_slider.grid(row=1, column=0, columnspan=3, sticky="ew", padx=5)
        Tooltip(self.comp_slider, tooltips.TOOLTIP_TEXT["adv_pike_slider"])
        self.pikepdf_warning_label = ttk.Label(compression_frame, text="(Ignored when cpdf is active)", foreground="orange"); self.pikepdf_warning_label.grid(row=2, column=0, columnspan=3, sticky="w", padx=5)

        self.update_final_processor_state()
        self.update_compression_label()
    
    def build_page_tools_tab(self, parent):
        parent.columnconfigure(0, weight=1); parent.columnconfigure(1, weight=1)
        left_col = ttk.Frame(parent); left_col.grid(row=0, column=0, sticky="nsew", padx=(0, 5)); left_col.columnconfigure(0, weight=1)
        right_col = ttk.Frame(parent); right_col.grid(row=0, column=1, sticky="nsew", padx=(5, 0)); right_col.columnconfigure(0, weight=1)

        merge_frame = self.build_merge_frame(left_col); merge_frame.grid(row=0, column=0, sticky="nsew", pady=(0,5))
        split_frame = self.build_split_frame(left_col); split_frame.grid(row=1, column=0, sticky="nsew", pady=5)
        delete_frame = self.build_delete_pages_frame(right_col); delete_frame.grid(row=0, column=0, sticky="nsew", pady=(0,5))
        rotate_frame = self.build_rotate_frame(right_col); rotate_frame.grid(row=1, column=0, sticky="nsew", pady=5)
    
    def build_tools_tab(self, parent):
        parent.columnconfigure(0, weight=1); parent.columnconfigure(1, weight=1)
        left_col = ttk.Frame(parent); left_col.grid(row=0, column=0, sticky="n", padx=(0, 5)); left_col.columnconfigure(0, weight=1)
        right_col = ttk.Frame(parent); right_col.grid(row=0, column=1, sticky="nsew", padx=(5, 0)); right_col.columnconfigure(0, weight=1); right_col.rowconfigure(0, weight=1)
        
        convert_frame = self.build_convert_frame(left_col); convert_frame.grid(row=0, column=0, sticky="new", pady=(0,5))
        pdfa_frame = self.build_pdfa_frame(left_col); pdfa_frame.grid(row=1, column=0, sticky="new", pady=5)
        util_frame = self.build_utility_frame(left_col); util_frame.grid(row=2, column=0, sticky="new", pady=5)

        meta_frame = self.build_metadata_frame(right_col); meta_frame.grid(row=0, column=0, sticky="nsew", pady=(0,5))
        stamp_frame = self.build_stamp_frame(right_col); stamp_frame.grid(row=1, column=0, sticky="nsew", pady=5)

    #region: Tool Sub-Builders
    def build_merge_frame(self, parent):
        frame = ttk.LabelFrame(parent, text="Merge PDFs"); frame.columnconfigure(0, weight=1)
        list_frame = ttk.Frame(frame); list_frame.grid(row=0, column=0, columnspan=2, sticky="nsew", padx=5, pady=5); list_frame.columnconfigure(0, weight=1); list_frame.rowconfigure(0, weight=1)
        self.merge_listbox = tk.Listbox(list_frame, selectmode=tk.SINGLE, height=4); self.merge_listbox.grid(row=0, column=0, sticky="nsew")
        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=self.merge_listbox.yview); scrollbar.grid(row=0, column=1, sticky="ns"); self.merge_listbox.config(yscrollcommand=scrollbar.set)
        btn_frame = ttk.Frame(frame); btn_frame.grid(row=1, column=0, columnspan=2, pady=5)
        add_btn = ttk.Button(btn_frame, text="Add", command=self.add_files_to_merge_list); add_btn.pack(side="left", padx=2); Tooltip(add_btn, tooltips.TOOLTIP_TEXT["merge_add_btn"])
        rem_btn = ttk.Button(btn_frame, text="Remove", command=self.remove_selected_from_list); rem_btn.pack(side="left", padx=2); Tooltip(rem_btn, tooltips.TOOLTIP_TEXT["merge_remove_btn"])
        up_btn = ttk.Button(btn_frame, text="Up", command=lambda: self.move_in_list(-1)); up_btn.pack(side="left", padx=2); Tooltip(up_btn, tooltips.TOOLTIP_TEXT["merge_up_btn"])
        down_btn = ttk.Button(btn_frame, text="Down", command=lambda: self.move_in_list(1)); down_btn.pack(side="left", padx=2); Tooltip(down_btn, tooltips.TOOLTIP_TEXT["merge_down_btn"])
        out_frame = ttk.Frame(frame); out_frame.grid(row=2, column=0, columnspan=2, sticky="ew", padx=5, pady=5); out_frame.columnconfigure(1, weight=1)
        merge_out_btn = ttk.Button(out_frame, text="Output", command=self.browse_merge_output); merge_out_btn.grid(row=0, column=0); Tooltip(merge_out_btn, tooltips.TOOLTIP_TEXT["merge_output_btn"])
        merge_out_entry = ttk.Entry(out_frame, textvariable=self.merge_output_path); merge_out_entry.grid(row=0, column=1, sticky="ew", padx=5); Tooltip(merge_out_entry, tooltips.TOOLTIP_TEXT["merge_output_entry"])
        self.merge_button = ttk.Button(frame, text="Merge Files", command=self.process_merge); self.merge_button.grid(row=3, column=0, columnspan=2, pady=5)
        Tooltip(self.merge_button, tooltips.TOOLTIP_TEXT["merge_process_btn"])
        return frame

    def build_split_frame(self, parent):
        frame = ttk.LabelFrame(parent, text="Split PDF"); frame.columnconfigure(1, weight=1)
        split_in_btn = ttk.Button(frame, text="Input", command=self.browse_split_input); split_in_btn.grid(row=0, column=0, padx=5, pady=5); Tooltip(split_in_btn, tooltips.TOOLTIP_TEXT["split_input_btn"])
        split_in_entry = ttk.Entry(frame, textvariable=self.split_input_path); split_in_entry.grid(row=0, column=1, sticky="ew", padx=5); Tooltip(split_in_entry, tooltips.TOOLTIP_TEXT["split_input_entry"])
        ttk.Label(frame, text="Mode:").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.split_combo = ttk.Combobox(frame, textvariable=self.split_mode, values=["Split to Single Pages", "Split Every N Pages", "Custom Range(s)"], state="readonly"); self.split_combo.grid(row=1, column=1, sticky="ew", padx=5); Tooltip(self.split_combo, tooltips.TOOLTIP_TEXT["split_mode_combo"])
        self.split_combo.bind("<<ComboboxSelected>>", self.toggle_split_value_entry)
        self.split_val_entry = ttk.Entry(frame, textvariable=self.split_value); self.split_val_entry.grid(row=2, column=1, sticky="ew", padx=5, pady=5); self.toggle_split_value_entry()
        Tooltip(self.split_val_entry, tooltips.TOOLTIP_TEXT["split_value_entry"])
        split_out_btn = ttk.Button(frame, text="Output Dir", command=self.browse_split_output_dir); split_out_btn.grid(row=3, column=0, padx=5, pady=5); Tooltip(split_out_btn, tooltips.TOOLTIP_TEXT["split_output_dir_btn"])
        split_out_entry = ttk.Entry(frame, textvariable=self.split_output_dir); split_out_entry.grid(row=3, column=1, sticky="ew", padx=5); Tooltip(split_out_entry, tooltips.TOOLTIP_TEXT["split_output_dir_entry"])
        self.split_button = ttk.Button(frame, text="Split File", command=self.process_split); self.split_button.grid(row=4, column=1, pady=5, sticky="e", padx=5)
        Tooltip(self.split_button, tooltips.TOOLTIP_TEXT["split_process_btn"])
        return frame
        
    def build_rotate_frame(self, parent):
        frame = ttk.LabelFrame(parent, text="Rotate Pages"); frame.columnconfigure(1, weight=1)
        rot_in_btn = ttk.Button(frame, text="Input PDF", command=lambda: self.browse_file_for_var(self.rotate_input, self.rotate_output, "_rotated")); rot_in_btn.grid(row=0, column=0, padx=5, pady=5, sticky="ew"); Tooltip(rot_in_btn, tooltips.TOOLTIP_TEXT["rotate_input_btn"])
        rot_in_entry = ttk.Entry(frame, textvariable=self.rotate_input); rot_in_entry.grid(row=0, column=1, sticky="ew", padx=5, pady=5); Tooltip(rot_in_entry, tooltips.TOOLTIP_TEXT["rotate_input_entry"])
        ttk.Label(frame, text="Angle:").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        rot_angle_combo = ttk.Combobox(frame, textvariable=self.rotate_angle, values=list(ROTATION_MAP.keys()), state="readonly"); rot_angle_combo.grid(row=1, column=1, sticky="ew", padx=5, pady=5); Tooltip(rot_angle_combo, tooltips.TOOLTIP_TEXT["rotate_angle_combo"])
        rot_out_btn = ttk.Button(frame, text="Output PDF", command=lambda: self.browse_save_as_for_var(self.rotate_output)); rot_out_btn.grid(row=2, column=0, padx=5, pady=5, sticky="ew"); Tooltip(rot_out_btn, tooltips.TOOLTIP_TEXT["rotate_output_btn"])
        rot_out_entry = ttk.Entry(frame, textvariable=self.rotate_output); rot_out_entry.grid(row=2, column=1, sticky="ew", padx=5, pady=5); Tooltip(rot_out_entry, tooltips.TOOLTIP_TEXT["rotate_output_entry"])
        self.rotate_button = ttk.Button(frame, text="Rotate PDF", command=self.process_rotate); self.rotate_button.grid(row=3, column=1, pady=5, sticky="e", padx=5); Tooltip(self.rotate_button, tooltips.TOOLTIP_TEXT["rotate_process_btn"])
        return frame

    def build_delete_pages_frame(self, parent):
        frame = ttk.LabelFrame(parent, text="Delete Pages"); frame.columnconfigure(1, weight=1)
        del_in_btn = ttk.Button(frame, text="Input PDF", command=lambda: self.browse_file_for_var(self.delete_pages_input, self.delete_pages_output, "_deleted")); del_in_btn.grid(row=0, column=0, padx=5, pady=5, sticky="ew"); Tooltip(del_in_btn, tooltips.TOOLTIP_TEXT["delete_input_btn"])
        del_in_entry = ttk.Entry(frame, textvariable=self.delete_pages_input); del_in_entry.grid(row=0, column=1, sticky="ew", padx=5, pady=5); Tooltip(del_in_entry, tooltips.TOOLTIP_TEXT["delete_input_entry"])
        ttk.Label(frame, text="Pages to Delete:").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        del_range_entry = ttk.Entry(frame, textvariable=self.delete_pages_range); del_range_entry.grid(row=1, column=1, sticky="ew", padx=5, pady=5); Tooltip(del_range_entry, tooltips.TOOLTIP_TEXT["delete_pages_entry"])
        del_out_btn = ttk.Button(frame, text="Output PDF", command=lambda: self.browse_save_as_for_var(self.delete_pages_output)); del_out_btn.grid(row=2, column=0, padx=5, pady=5, sticky="ew"); Tooltip(del_out_btn, tooltips.TOOLTIP_TEXT["delete_output_btn"])
        del_out_entry = ttk.Entry(frame, textvariable=self.delete_pages_output); del_out_entry.grid(row=2, column=1, sticky="ew", padx=5, pady=5); Tooltip(del_out_entry, tooltips.TOOLTIP_TEXT["delete_output_entry"])
        self.delete_button = ttk.Button(frame, text="Delete Pages", command=self.process_delete_pages); self.delete_button.grid(row=3, column=1, pady=5, sticky="e", padx=5); Tooltip(self.delete_button, tooltips.TOOLTIP_TEXT["delete_process_btn"])
        return frame
        
    def build_convert_frame(self, parent):
        frame = ttk.LabelFrame(parent, text="Convert PDF to Image"); frame.columnconfigure(1, weight=1)
        conv_in_btn = ttk.Button(frame, text="Input PDF", command=lambda: self.browse_file_for_var(self.pdf_to_img_input)); conv_in_btn.grid(row=0, column=0, sticky="ew", padx=5, pady=5); Tooltip(conv_in_btn, tooltips.TOOLTIP_TEXT["convert_input_btn"])
        conv_in_entry = ttk.Entry(frame, textvariable=self.pdf_to_img_input); conv_in_entry.grid(row=0, column=1, columnspan=3, sticky="ew", padx=5, pady=5); Tooltip(conv_in_entry, tooltips.TOOLTIP_TEXT["convert_input_entry"])
        conv_out_btn = ttk.Button(frame, text="Output Folder", command=lambda: self.browse_dir_for_var(self.pdf_to_img_output_dir)); conv_out_btn.grid(row=1, column=0, sticky="ew", padx=5, pady=5); Tooltip(conv_out_btn, tooltips.TOOLTIP_TEXT["convert_output_btn"])
        conv_out_entry = ttk.Entry(frame, textvariable=self.pdf_to_img_output_dir); conv_out_entry.grid(row=1, column=1, columnspan=3, sticky="ew", padx=5, pady=5); Tooltip(conv_out_entry, tooltips.TOOLTIP_TEXT["convert_output_entry"])
        ttk.Label(frame, text="Format:").grid(row=2, column=0, padx=5, pady=5)
        conv_format_combo = ttk.Combobox(frame, textvariable=self.pdf_to_img_format, values=["png", "jpeg", "tiff"], state="readonly"); conv_format_combo.grid(row=2, column=1, padx=5, pady=5, sticky="ew"); Tooltip(conv_format_combo, tooltips.TOOLTIP_TEXT["convert_format_combo"])
        ttk.Label(frame, text="DPI:").grid(row=2, column=2, padx=5, pady=5)
        conv_dpi_entry = ttk.Entry(frame, textvariable=self.pdf_to_img_dpi, width=7); conv_dpi_entry.grid(row=2, column=3, padx=5, pady=5, sticky="w"); Tooltip(conv_dpi_entry, tooltips.TOOLTIP_TEXT["convert_dpi_entry"])
        self.convert_button = ttk.Button(frame, text="Convert to Images", command=self.process_pdf_to_image); self.convert_button.grid(row=3, column=1, columnspan=3, sticky="e", padx=5, pady=10); Tooltip(self.convert_button, tooltips.TOOLTIP_TEXT["convert_process_btn"])
        return frame
        
    def build_pdfa_frame(self, parent):
        frame = ttk.LabelFrame(parent, text="Convert to PDF/A"); frame.columnconfigure(1, weight=1)
        pdfa_in_btn = ttk.Button(frame, text="Input PDF", command=lambda: self.browse_file_for_var(self.pdfa_input, self.pdfa_output, "_pdfa")); pdfa_in_btn.grid(row=0, column=0, padx=5, pady=5, sticky="ew"); Tooltip(pdfa_in_btn, tooltips.TOOLTIP_TEXT["pdfa_input_btn"])
        pdfa_in_entry = ttk.Entry(frame, textvariable=self.pdfa_input); pdfa_in_entry.grid(row=0, column=1, sticky="ew", padx=5, pady=5); Tooltip(pdfa_in_entry, tooltips.TOOLTIP_TEXT["pdfa_input_entry"])
        pdfa_out_btn = ttk.Button(frame, text="Output PDF", command=lambda: self.browse_save_as_for_var(self.pdfa_output)); pdfa_out_btn.grid(row=1, column=0, padx=5, pady=5, sticky="ew"); Tooltip(pdfa_out_btn, tooltips.TOOLTIP_TEXT["pdfa_output_btn"])
        pdfa_out_entry = ttk.Entry(frame, textvariable=self.pdfa_output); pdfa_out_entry.grid(row=1, column=1, sticky="ew", padx=5, pady=5); Tooltip(pdfa_out_entry, tooltips.TOOLTIP_TEXT["pdfa_output_entry"])
        self.pdfa_button = ttk.Button(frame, text="Convert to PDF/A", command=self.process_pdfa_conversion); self.pdfa_button.grid(row=2, column=1, pady=5, sticky="e", padx=5); Tooltip(self.pdfa_button, tooltips.TOOLTIP_TEXT["pdfa_process_btn"])
        return frame

    def build_utility_frame(self, parent):
        frame = ttk.LabelFrame(parent, text="Utility"); frame.columnconfigure(1, weight=1)
        util_in_btn = ttk.Button(frame, text="Input PDF", command=lambda: self.browse_file_for_var(self.utility_pdf_in)); util_in_btn.grid(row=0, column=0, padx=5, pady=5, sticky="ew"); Tooltip(util_in_btn, tooltips.TOOLTIP_TEXT["util_input_btn"])
        util_in_entry = ttk.Entry(frame, textvariable=self.utility_pdf_in); util_in_entry.grid(row=0, column=1, sticky="ew", padx=5); Tooltip(util_in_entry, tooltips.TOOLTIP_TEXT["util_input_entry"])
        self.util_button = ttk.Button(frame, text="Remove Opening Action", command=self.process_remove_open_action); self.util_button.grid(row=1, column=1, pady=5, sticky="e", padx=5)
        Tooltip(self.util_button, tooltips.TOOLTIP_TEXT["util_process_btn"])
        return frame

    def build_metadata_frame(self, parent):
        frame = ttk.LabelFrame(parent, text="PDF Metadata Editor"); frame.columnconfigure(1, weight=1)
        in_frame = ttk.Frame(frame); in_frame.grid(row=0, column=0, columnspan=2, sticky="ew", padx=5, pady=5); in_frame.columnconfigure(1, weight=1)
        meta_in_btn = ttk.Button(in_frame, text="PDF Input", command=self.browse_meta_input); meta_in_btn.grid(row=0, column=0, padx=(0,5)); Tooltip(meta_in_btn, tooltips.TOOLTIP_TEXT["meta_input_btn"])
        meta_in_entry = ttk.Entry(in_frame, textvariable=self.meta_pdf_path); meta_in_entry.grid(row=0, column=1, sticky="ew"); Tooltip(meta_in_entry, tooltips.TOOLTIP_TEXT["meta_input_entry"])
        meta_load_btn = ttk.Button(in_frame, text="Load", command=self.load_metadata, width=5); meta_load_btn.grid(row=0, column=2, padx=(5,0)); Tooltip(meta_load_btn, tooltips.TOOLTIP_TEXT["meta_load_btn"])
        
        meta_fields = {
            "Title": (self.meta_title, tooltips.TOOLTIP_TEXT["meta_title"]),
            "Author": (self.meta_author, tooltips.TOOLTIP_TEXT["meta_author"]),
            "Subject": (self.meta_subject, tooltips.TOOLTIP_TEXT["meta_subject"]),
            "Keywords": (self.meta_keywords, tooltips.TOOLTIP_TEXT["meta_keywords"]),
        }
        for i, (label_text, (var, tooltip_text)) in enumerate(meta_fields.items(), start=1):
            ttk.Label(frame, text=f"{label_text}:").grid(row=i, column=0, sticky="w", padx=5, pady=2)
            entry = ttk.Entry(frame, textvariable=var)
            entry.grid(row=i, column=1, sticky="ew", padx=5, pady=2)
            Tooltip(entry, tooltip_text)

        self.meta_save_button = ttk.Button(frame, text="Save Metadata (overwrite)", command=self.save_metadata); self.meta_save_button.grid(row=len(meta_fields) + 1, column=1, pady=10, sticky="e", padx=5)
        Tooltip(self.meta_save_button, tooltips.TOOLTIP_TEXT["meta_save_btn"])
        return frame

    def build_stamp_frame(self, parent):
        frame = ttk.LabelFrame(parent, text="Stamp / Watermark"); frame.columnconfigure(1, weight=1)
        stamp_in_btn = ttk.Button(frame, text="PDF Input", command=lambda: self.browse_file_for_var(self.stamp_pdf_in, self.stamp_pdf_out, "_stamped")); stamp_in_btn.grid(row=0, column=0, padx=5, pady=5, sticky="ew"); Tooltip(stamp_in_btn, tooltips.TOOLTIP_TEXT["stamp_input_btn"])
        stamp_in_entry = ttk.Entry(frame, textvariable=self.stamp_pdf_in); stamp_in_entry.grid(row=0, column=1, sticky="ew", padx=5); Tooltip(stamp_in_entry, tooltips.TOOLTIP_TEXT["stamp_input_entry"])
        mode_frame = ttk.Frame(frame); mode_frame.grid(row=1, column=0, columnspan=2, pady=2)
        img_radio = ttk.Radiobutton(mode_frame, text="Image Stamp", variable=self.stamp_mode, value="Image", command=self.toggle_stamp_mode); img_radio.pack(side="left", padx=5); Tooltip(img_radio, tooltips.TOOLTIP_TEXT["stamp_image_radio"])
        txt_radio = ttk.Radiobutton(mode_frame, text="Text Stamp", variable=self.stamp_mode, value="Text", command=self.toggle_stamp_mode); txt_radio.pack(side="left", padx=5); Tooltip(txt_radio, tooltips.TOOLTIP_TEXT["stamp_text_radio"])
        
        self.stamp_image_frame = ttk.Frame(frame); self.stamp_image_frame.grid(row=2, column=0, columnspan=2, sticky="ew")
        self.stamp_image_frame.columnconfigure(1, weight=1)
        stamp_img_btn = ttk.Button(self.stamp_image_frame, text="Image", command=self.browse_stamp_image); stamp_img_btn.grid(row=0, column=0, padx=5, pady=2, sticky="ew"); Tooltip(stamp_img_btn, tooltips.TOOLTIP_TEXT["stamp_image_btn"])
        stamp_img_entry = ttk.Entry(self.stamp_image_frame, textvariable=self.stamp_image_in); stamp_img_entry.grid(row=0, column=1, sticky="ew", padx=5); Tooltip(stamp_img_entry, tooltips.TOOLTIP_TEXT["stamp_image_entry"])
        ttk.Label(self.stamp_image_frame, text="Use transparent PNG", font=(None, 8)).grid(row=1, column=1, sticky="w", padx=5)
        size_frame = ttk.Frame(self.stamp_image_frame); size_frame.grid(row=2, column=0, columnspan=2, sticky='w', padx=5, pady=2)
        ttk.Label(size_frame, text="Width:").pack(side='left'); w_entry = ttk.Entry(size_frame, textvariable=self.stamp_image_width, width=5); w_entry.pack(side='left'); Tooltip(w_entry, tooltips.TOOLTIP_TEXT["stamp_image_width"])
        ttk.Label(size_frame, text="Height:").pack(side='left', padx=(10,0)); h_entry = ttk.Entry(size_frame, textvariable=self.stamp_image_height, width=5); h_entry.pack(side='left'); Tooltip(h_entry, tooltips.TOOLTIP_TEXT["stamp_image_height"])
        ttk.Label(size_frame, text="(pixels)").pack(side='left', padx=5)

        self.stamp_text_frame = ttk.Frame(frame); self.stamp_text_frame.grid(row=2, column=0, columnspan=2, sticky="ew")
        self.stamp_text_frame.columnconfigure(1, weight=1)
        ttk.Label(self.stamp_text_frame, text="Text:").grid(row=0, column=0, padx=5, pady=2, sticky="w"); text_entry = ttk.Entry(self.stamp_text_frame, textvariable=self.stamp_text_in); text_entry.grid(row=0, column=1, columnspan=3, sticky="ew", padx=5); Tooltip(text_entry, tooltips.TOOLTIP_TEXT["stamp_text_entry"])
        font_frame = ttk.Frame(self.stamp_text_frame); font_frame.grid(row=1, column=0, columnspan=4, sticky='ew')
        ttk.Label(font_frame, text="Font:").grid(row=0, column=0, padx=5, pady=2, sticky="w"); font_combo = ttk.Combobox(font_frame, textvariable=self.stamp_font, values=PDF_FONTS, state="readonly", width=18); font_combo.grid(row=0, column=1, sticky="ew", padx=5); Tooltip(font_combo, tooltips.TOOLTIP_TEXT["stamp_font_combo"])
        ttk.Label(font_frame, text="Size:").grid(row=0, column=2, padx=5, pady=2, sticky="w"); size_entry = ttk.Entry(font_frame, textvariable=self.stamp_font_size, width=5); size_entry.grid(row=0, column=3, sticky="w", padx=5); Tooltip(size_entry, tooltips.TOOLTIP_TEXT["stamp_font_size"])
        color_frame = ttk.Frame(self.stamp_text_frame); color_frame.grid(row=2, column=0, columnspan=4, sticky='ew')
        ttk.Label(color_frame, text="Color:").grid(row=0, column=0, padx=5, pady=2, sticky="w"); color_btn = ttk.Button(color_frame, text="Choose Color", command=self.choose_color); color_btn.grid(row=0, column=1, padx=5); Tooltip(color_btn, tooltips.TOOLTIP_TEXT["stamp_color_btn"])
        self.color_swatch = tk.Label(color_frame, text="    ", relief="sunken", borderwidth=1); self.color_swatch.grid(row=0, column=2, padx=5)
        
        dyn_frame = ttk.LabelFrame(self.stamp_text_frame, text="Dynamic Fields"); dyn_frame.grid(row=3, column=0, columnspan=4, sticky='ew', padx=5, pady=5)
        fn_check = ttk.Checkbutton(dyn_frame, text="Filename", variable=self.stamp_dynamic_filename, command=self.update_stamp_text); fn_check.pack(side="left", padx=5); Tooltip(fn_check, tooltips.TOOLTIP_TEXT["stamp_dyn_filename"])
        dt_check = ttk.Checkbutton(dyn_frame, text="Date/Time", variable=self.stamp_dynamic_datetime, command=self.update_stamp_text); dt_check.pack(side="left", padx=5); Tooltip(dt_check, tooltips.TOOLTIP_TEXT["stamp_dyn_datetime"])
        bates_frame = ttk.Frame(dyn_frame)
        bates_frame.pack(side="left", padx=5)
        bates_check = ttk.Checkbutton(bates_frame, text="Bates No.", variable=self.stamp_dynamic_bates, command=self.update_stamp_text); bates_check.pack(side="left"); Tooltip(bates_check, tooltips.TOOLTIP_TEXT["stamp_dyn_bates_check"])
        bates_entry = ttk.Entry(bates_frame, textvariable=self.stamp_bates_start, width=5); bates_entry.pack(side="left"); Tooltip(bates_entry, tooltips.TOOLTIP_TEXT["stamp_dyn_bates_entry"])

        self.toggle_stamp_mode()
        pos_frame = ttk.Frame(frame); pos_frame.grid(row=3, column=0, columnspan=2, sticky="ew", padx=5, pady=5); pos_frame.columnconfigure(1, weight=1)
        ttk.Label(pos_frame, text="Position:").grid(row=0, column=0); pos_combo = ttk.Combobox(pos_frame, textvariable=self.stamp_pos, state="readonly", values=["Center", "Bottom-Left", "Bottom-Right"]); pos_combo.grid(row=0, column=1, sticky="ew", padx=5); Tooltip(pos_combo, tooltips.TOOLTIP_TEXT["stamp_pos_combo"])
        op_frame = ttk.Frame(frame); op_frame.grid(row=4, column=0, columnspan=2, sticky="ew", padx=5, pady=5); op_frame.columnconfigure(1, weight=1)
        ttk.Label(op_frame, text="Opacity:").grid(row=0, column=0); op_slider = ttk.Scale(op_frame, from_=0.1, to=1.0, variable=self.stamp_opacity, orient="horizontal"); op_slider.grid(row=0, column=1, sticky="ew", padx=(5,0)); Tooltip(op_slider, tooltips.TOOLTIP_TEXT["stamp_opacity_slider"])
        ontop_check = ttk.Checkbutton(frame, text="Stamp on top of content", variable=self.stamp_on_top); ontop_check.grid(row=5, column=0, columnspan=2, padx=5, pady=5, sticky="w"); Tooltip(ontop_check, tooltips.TOOLTIP_TEXT["stamp_ontop_check"])
        stamp_out_btn = ttk.Button(frame, text="PDF Output", command=lambda: self.browse_save_as_for_var(self.stamp_pdf_out)); stamp_out_btn.grid(row=6, column=0, padx=5, pady=5, sticky="ew"); Tooltip(stamp_out_btn, tooltips.TOOLTIP_TEXT["stamp_output_btn"])
        stamp_out_entry = ttk.Entry(frame, textvariable=self.stamp_pdf_out); stamp_out_entry.grid(row=6, column=1, sticky="ew", padx=5); Tooltip(stamp_out_entry, tooltips.TOOLTIP_TEXT["stamp_output_entry"])
        self.stamp_button = ttk.Button(frame, text="Stamp File", command=self.process_stamp); self.stamp_button.grid(row=7, column=1, pady=5, sticky="e", padx=5); Tooltip(self.stamp_button, tooltips.TOOLTIP_TEXT["stamp_process_btn"])
        return frame
    #endregion

    #region: GUI Callbacks and Updates
    def update_final_processor_state(self, *args):
        use_any_cpdf = (self.adv_options['use_cpdf_squeeze'].get() or
                        self.adv_options['darken_text'].get() or
                        self.adv_options['user_password'].get() or
                        self.adv_options['owner_password'].get())
        
        if use_any_cpdf:
            self.pikepdf_warning_label.grid()
        else:
            self.pikepdf_warning_label.grid_remove()
        
        self.update_compression_label()

    def choose_color(self):
        color_code = colorchooser.askcolor(title="Choose color")
        if color_code and color_code[1]:
            self.stamp_font_color_var.set(color_code[1])
            self.update_color_swatch()
            
    def update_color_swatch(self):
        self.color_swatch.config(background=self.stamp_font_color_var.get())

    def update_stamp_text(self):
        base_text = self.stamp_text_in.get().split(" | ")[0]
        dyn_parts = []
        if self.stamp_dynamic_filename.get(): dyn_parts.append("%filename")
        if self.stamp_dynamic_datetime.get(): dyn_parts.append("%Y-%m-%d %H:%M")
        if self.stamp_dynamic_bates.get(): dyn_parts.append(f"%Bates")
        if dyn_parts: self.stamp_text_in.set(f"{base_text} | {' '.join(dyn_parts)}")
        else: self.stamp_text_in.set(base_text)
        
    def toggle_password_visibility(self):
        show = self.adv_options['show_passwords'].get()
        self.user_pass_entry.config(show="" if show else "*")
        self.owner_pass_entry.config(show="" if show else "*")

    def update_image_resolution(self, *args):
        self.adv_options['image_resolution'].set(DPI_PRESETS.get(self.operation.get(), "150"))

    def update_compression_label(self, *args):
        level = int(self.adv_options['pikepdf_compression_level'].get())
        text = f"{level}"
        if level == 0: text += " (None)"
        elif level == 9: text += " (Max)"
        self.compression_label.config(text=text)

    def toggle_theme(self):
        mode = 'dark' if self.dark_mode_enabled.get() else 'light'
        self.palette = styles.apply_theme(self.root, mode)
        self.root.configure(background=self.palette["BG"])
        self.update_color_swatch()

    def toggle_advanced(self):
        self.advanced_frame.grid() if self.show_advanced.get() else self.advanced_frame.grid_remove()
    
    def toggle_split_value_entry(self, event=None):
        if self.split_mode.get() == "Split to Single Pages": self.split_val_entry.grid_remove()
        else: self.split_val_entry.grid()
        
    def toggle_stamp_mode(self):
        if self.stamp_mode.get() == "Image":
            self.stamp_text_frame.grid_remove(); self.stamp_image_frame.grid()
        else:
            self.stamp_image_frame.grid_remove(); self.stamp_text_frame.grid()
    #endregion

    #region: File Browse Logic
    def select_input(self):
        dialog = tk.Toplevel(self.root); dialog.configure(background=self.palette["BG"])
        ttk.Label(dialog, text="Choose input type:").pack(pady=10)
        ttk.Button(dialog, text="File", command=lambda: [self.browse_file(), dialog.destroy()]).pack(pady=5, padx=20, fill='x')
        ttk.Button(dialog, text="Folder", command=lambda: [self.browse_folder(), dialog.destroy()]).pack(pady=5, padx=20, fill='x')
        x = self.root.winfo_x() + (self.root.winfo_width() // 2) - 150
        y = self.root.winfo_y() + (self.root.winfo_height() // 2) - 75
        dialog.geometry(f'300x150+{x}+{y}'); dialog.transient(self.root); dialog.grab_set(); self.root.wait_window(dialog)
        
    def browse_file(self):
        path = filedialog.askopenfilename(parent=self.root, filetypes=[("PDF files", "*.pdf")])
        if path: self.is_folder = False; self.input_path.set(path); p = Path(path); self.output_path.set(str(p.with_stem(f"{p.stem}_processed")))
    def browse_folder(self):
        path = filedialog.askdirectory(parent=self.root)
        if path: self.is_folder = True; self.input_path.set(path); self.output_path.set(path)
    def browse_output(self):
        if self.is_folder: path = filedialog.askdirectory(parent=self.root)
        else: path = filedialog.asksaveasfilename(parent=self.root, defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
        if path: self.output_path.set(path)
    
    def browse_file_for_var(self, in_var, out_var=None, suffix=""):
        path = filedialog.askopenfilename(parent=self.root, filetypes=[("PDF files", "*.pdf")])
        if path:
            in_var.set(path) 
            if out_var:
                out_var.set(str(Path(path).with_stem(f"{Path(path).stem}{suffix}")))
    def browse_dir_for_var(self, dir_var):
        path = filedialog.askdirectory(parent=self.root)
        if path: dir_var.set(path)
    def browse_save_as_for_var(self, string_var):
        path = filedialog.asksaveasfilename(parent=self.root, defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
        if path: string_var.set(path)
    
    def add_files_to_merge_list(self, files=None):
        if not files: files = filedialog.askopenfilenames(parent=self.root, title="Select PDF files to merge", filetypes=[("PDF files", "*.pdf")])
        for f in files: self.merge_files.append(Path(f))
        self.update_merge_listbox()
    def update_merge_listbox(self):
        self.merge_listbox.delete(0, tk.END)
        for f in self.merge_files: self.merge_listbox.insert(tk.END, f.name)
    def remove_selected_from_list(self):
        selected = self.merge_listbox.curselection()
        if not selected: return
        del self.merge_files[selected[0]]; self.update_merge_listbox()
    def move_in_list(self, direction):
        selected_idx = self.merge_listbox.curselection()
        if not selected_idx: return
        idx = selected_idx[0]
        new_idx = idx + direction
        if 0 <= new_idx < len(self.merge_files):
            self.merge_files[idx], self.merge_files[new_idx] = self.merge_files[new_idx], self.merge_files[idx]
            self.update_merge_listbox(); self.merge_listbox.selection_set(new_idx)
            
    def browse_merge_output(self): self.browse_save_as_for_var(self.merge_output_path)
    def browse_split_input(self):
        path = filedialog.askopenfilename(parent=self.root, filetypes=[("PDF files", "*.pdf")])
        if path: self.split_input_path.set(path); self.split_output_dir.set(str(Path(path).parent))
    def browse_split_output_dir(self): self.browse_dir_for_var(self.split_output_dir)
    def browse_stamp_image(self):
        path = filedialog.askopenfilename(parent=self.root, filetypes=[("Image files", "*.png *.jpg *.jpeg")])
        if path: self.stamp_image_in.set(path)
    def browse_meta_input(self):
        path = filedialog.askopenfilename(parent=self.root, filetypes=[("PDF files", "*.pdf")])
        if path: self.meta_pdf_path.set(path); self.load_metadata()
    #endregion

    #region: Metadata Logic
    def load_metadata(self):
        pdf_path = self.meta_pdf_path.get()
        if not pdf_path: messagebox.showwarning("Warning", "Please select a PDF file first.", parent=self.root); return
        try:
            metadata = backend.run_metadata_task('load', pdf_path, self.cpdf_path)
            self.meta_title.set(metadata.get('title', ''))
            self.meta_author.set(metadata.get('author', ''))
            self.meta_subject.set(metadata.get('subject', ''))
            self.meta_keywords.set(metadata.get('keywords', ''))
            self.status.set("Metadata loaded successfully.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load metadata: {e}", parent=self.root)
            self.status.set(f"Error loading metadata.")

    def save_metadata(self):
        pdf_path = self.meta_pdf_path.get()
        if not pdf_path: messagebox.showwarning("Warning", "Please select a PDF file first.", parent=self.root); return
        metadata = {
            'title': self.meta_title.get(), 'author': self.meta_author.get(),
            'subject': self.meta_subject.get(), 'keywords': self.meta_keywords.get()
        }
        try:
            backend.run_metadata_task('save', pdf_path, self.cpdf_path, metadata)
            messagebox.showinfo("Success", "Metadata saved successfully.", parent=self.root)
            self.status.set("Metadata saved.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save metadata: {e}", parent=self.root)
            self.status.set(f"Error saving metadata.")
    #endregion

    #region: Processing Logic
    def start_task(self, button, target_func, args):
        if self.active_process_button: messagebox.showwarning("Busy", "A process is already running.", parent=self.root); return
        self.active_process_button = button; self.active_process_button.config(state="disabled")
        self.progress_bar.grid(); self.progress_bar['value'] = 0
        threading.Thread(target=target_func, args=args, daemon=True).start()

    def on_task_complete(self, final_status):
        self.progress_bar.grid_remove(); self.status.set(final_status)
        if self.active_process_button: self.active_process_button.config(state="normal"); self.active_process_button = None

    def process_conversion(self):
        if not self.gs_path: self.check_tools(); return
        input_p_str = self.input_path.get()
        if not input_p_str: messagebox.showwarning("Warning", "Please select an input.", parent=self.root); return
        is_overwrite = self.overwrite_originals.get()
        if is_overwrite:
            if not messagebox.askyesno("Confirm Overwrite", "This will OVERWRITE your original file(s). This cannot be undone. Continue?", icon='warning', parent=self.root): return
            output_p_str = input_p_str
        else: output_p_str = self.output_path.get()
        if not output_p_str: messagebox.showwarning("Warning", "Please select an output location.", parent=self.root); return
        params = { 'gs_path': self.gs_path, 'cpdf_path': self.cpdf_path, 'input_path': input_p_str, 'output_path': output_p_str,
                   'operation': self.operation.get(), 'options': {k: v.get() for k, v in self.adv_options.items()}, 'overwrite': is_overwrite }
        self.start_task(self.compress_button, backend.run_conversion_task, (params, self.is_folder, self.status, self.progress_bar, self.on_task_complete))

    def process_pdfa_conversion(self):
        in_pdf, out_pdf = self.pdfa_input.get(), self.pdfa_output.get()
        if not (in_pdf and out_pdf): messagebox.showwarning("Warning", "Input and Output PDF files are required.", parent=self.root); return
        if not self.gs_path: messagebox.showerror("Error", "Ghostscript not found.", parent=self.root); return
        self.start_task(self.pdfa_button, backend.run_pdfa_conversion_task, (in_pdf, out_pdf, self.gs_path, self.status, self.progress_bar, self.on_task_complete))

    def process_pdf_to_image(self):
        in_pdf, out_dir = self.pdf_to_img_input.get(), self.pdf_to_img_output_dir.get()
        if not (in_pdf and out_dir): messagebox.showwarning("Warning", "Input PDF and Output Folder are required.", parent=self.root); return
        if not self.gs_path: messagebox.showerror("Error", "Ghostscript not found.", parent=self.root); return
        options = {'format': self.pdf_to_img_format.get(), 'dpi': self.pdf_to_img_dpi.get()}
        self.start_task(self.convert_button, backend.run_pdf_to_image_task, (in_pdf, out_dir, options, self.gs_path, self.status, self.progress_bar, self.on_task_complete))

    def process_merge(self):
        out_path = self.merge_output_path.get()
        if not self.merge_files or not out_path: messagebox.showwarning("Warning", "Please add files and set an output path.", parent=self.root); return
        self.start_task(self.merge_button, backend.run_merge_task, (self.merge_files, out_path, self.status, self.progress_bar, self.on_task_complete))

    def process_rotate(self):
        in_pdf, out_pdf, angle = self.rotate_input.get(), self.rotate_output.get(), ROTATION_MAP[self.rotate_angle.get()]
        if not (in_pdf and out_pdf): messagebox.showwarning("Warning", "Input and Output PDFs are required.", parent=self.root); return
        self.start_task(self.rotate_button, backend.run_rotate_task, (in_pdf, out_pdf, angle, self.status, self.progress_bar, self.on_task_complete))

    def process_delete_pages(self):
        in_pdf, out_pdf, page_range = self.delete_pages_input.get(), self.delete_pages_output.get(), self.delete_pages_range.get()
        if not (in_pdf and out_pdf and page_range): messagebox.showwarning("Warning", "All fields are required to delete pages.", parent=self.root); return
        self.start_task(self.delete_button, backend.run_delete_pages_task, (in_pdf, out_pdf, page_range, self.status, self.progress_bar, self.on_task_complete))
    
    def process_split(self):
        in_path, out_dir = self.split_input_path.get(), self.split_output_dir.get()
        if not (in_path and out_dir): messagebox.showwarning("Warning", "Input file and output directory are required.", parent=self.root); return
        mode, value = self.split_mode.get(), self.split_value.get()
        self.start_task(self.split_button, backend.run_split_task, (in_path, out_dir, mode, value, self.status, self.progress_bar, self.on_task_complete))
        
    def process_stamp(self):
        if not self.cpdf_path: self.check_tools(); return
        in_pdf, out_pdf = self.stamp_pdf_in.get(), self.stamp_pdf_out.get()
        if not (in_pdf and out_pdf): messagebox.showwarning("Warning", "Input and Output PDF paths are required.", parent=self.root); return
        
        mode = self.stamp_mode.get()
        stamp_opts = {'pos': self.stamp_pos.get(), 'opacity': self.stamp_opacity.get(), 'on_top': self.stamp_on_top.get()}
        mode_opts = {}
        if mode == "Image":
            img_path = self.stamp_image_in.get()
            if not img_path: messagebox.showwarning("Warning", "Please select an image file to stamp.", parent=self.root); return
            mode_opts = {'image_path': img_path, 'width': self.stamp_image_width.get(), 'height': self.stamp_image_height.get()}
        else: 
            rgb = self.root.winfo_rgb(self.stamp_font_color_var.get())
            color_str = f"{rgb[0]/65535:.2f} {rgb[1]/65535:.2f} {rgb[2]/65535:.2f}"
            self.update_stamp_text()
            mode_opts = {'text': self.stamp_text_in.get(), 'font': self.stamp_font.get(), 'size': self.stamp_font_size.get(), 'color': color_str,
                         'bates_start': self.stamp_bates_start.get() if self.stamp_dynamic_bates.get() else None}
        
        self.start_task(self.stamp_button, backend.run_stamp_task, (in_pdf, out_pdf, stamp_opts, self.cpdf_path, self.status, self.progress_bar, self.on_task_complete, mode, mode_opts))
    
    def process_remove_open_action(self):
        pdf_in = self.utility_pdf_in.get()
        if not pdf_in: messagebox.showwarning("Warning", "Please select an input PDF.", parent=self.root); return
        self.start_task(self.util_button, backend.run_remove_open_action_task, (pdf_in, self.cpdf_path, self.status, self.progress_bar, self.on_task_complete))
    #endregion