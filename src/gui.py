# gui.py
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, colorchooser
from pathlib import Path
import logging
import threading
import sys
import json
import io
from datetime import datetime
from PIL import Image, ImageTk

import backend
import styles
import tooltips
from constants import (APP_VERSION, OPERATIONS, ROTATION_MAP, PDF_FONTS,
                       DPI_PRESETS, OP_COMPRESS_EBOOK)
from ui_components import FileSelector, Tooltip

IS_WINDOWS = sys.platform == "win32"
if IS_WINDOWS:
    try:
        import windnd
    except ImportError:
        print("Warning: windnd library not found. Drag and drop will be disabled.")
        IS_WINDOWS = False

#region: ScrolledFrame Class
class ScrolledFrame(ttk.Frame):
    def __init__(self, parent, *args, **kw):
        super().__init__(parent, *args, **kw)
        self.canvas = tk.Canvas(self, borderwidth=0, highlightthickness=0)
        self.scrollbar = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        self.scrollbar.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollable_frame = ttk.Frame(self.canvas, padding=10)
        self.canvas_window = self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.scrollable_frame.bind("<Configure>", self.on_frame_configure)
        self.canvas.bind("<Configure>", self.on_canvas_configure)
        self.scrollable_frame.bind("<Enter>", self.bind_mouse_wheel)
        self.scrollable_frame.bind("<Leave>", self.unbind_mouse_wheel)
    def on_frame_configure(self, event): self.canvas.configure(scrollregion=self.canvas.bbox("all"))
    def on_canvas_configure(self, event): self.canvas.itemconfig(self.canvas_window, width=event.width)
    def bind_mouse_wheel(self, event):
        self.canvas.bind_all("<MouseWheel>", self._on_mouse_wheel)
        self.canvas.bind_all("<Button-4>", self._on_mouse_wheel)
        self.canvas.bind_all("<Button-5>", self._on_mouse_wheel)
    def unbind_mouse_wheel(self, event):
        self.canvas.unbind_all("<MouseWheel>")
        self.canvas.unbind_all("<Button-4>")
        self.canvas.unbind_all("<Button-5>")
    def _on_mouse_wheel(self, event):
        if sys.platform == 'win32': self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        elif sys.platform == 'darwin': self.canvas.yview_scroll(int(-event.delta), "units")
        else:
            if event.num == 4: self.canvas.yview_scroll(-1, "units")
            elif event.num == 5: self.canvas.yview_scroll(1, "units")
#endregion

class GhostscriptGUI:
    def __init__(self, root):
        self.root = root
        self.root.title(f"MinimalPDF Compress v{APP_VERSION}")
        self.root.minsize(860, 600)
        self.settings_file = Path("settings.json")

        #region: Variable Declarations
        self.input_path = tk.StringVar()
        self.output_path = tk.StringVar()
        self.operation = tk.StringVar()
        self.is_folder = tk.BooleanVar(value=False)
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
        self.progress_status = tk.StringVar()
        self.optimize_images = tk.BooleanVar(value=False)
        self.adv_options = { 'image_resolution': tk.StringVar(), 'downscale_factor': tk.StringVar(), 'color_strategy': tk.StringVar(), 'downsample_type': tk.StringVar(), 'fast_web_view': tk.BooleanVar(), 'subset_fonts': tk.BooleanVar(), 'compress_fonts': tk.BooleanVar(), 'rotation': tk.StringVar(), 'strip_metadata': tk.BooleanVar(), 'remove_interactive': tk.BooleanVar(), 'pikepdf_compression_level': tk.IntVar(), 'decimal_precision': tk.StringVar(), 'use_cpdf_squeeze': tk.BooleanVar(), 'darken_text': tk.BooleanVar(), 'use_fast_processing': tk.BooleanVar(), 'user_password': tk.StringVar(), 'owner_password': tk.StringVar(), 'show_passwords': tk.BooleanVar(), }
        #endregion

        self.palette, self.gs_path, self.cpdf_path = {}, None, None
        self.active_process_button = None
        self.tab_frames = {}
        self.rotate_preview_image_tk = None
        self.stamp_preview_image_tk = None
        self.rotate_debounce_timer = None
        self.stamp_debounce_timer = None

        self.icon_path = backend.resource_path("pdf.ico")
        try:
            if self.icon_path.exists(): self.root.iconbitmap(self.icon_path)
        except Exception as e: logging.warning(f"Failed to set icon: {e}")

        self.main_frame = ttk.Frame(self.root, padding="10")
        self.main_frame.grid(row=0, column=0, sticky="nsew")
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)

        logging.basicConfig(filename="app.log", level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")

        self.load_settings()
        self.build_gui()
        self.toggle_theme()
        self.setup_drag_and_drop()
        self.check_tools()
        self.setup_traces()
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

    def on_closing(self):
        self.save_settings()
        self.root.destroy()

    def save_settings(self):
        settings = {
            'general': {
                'dark_mode_enabled': self.dark_mode_enabled.get(),
                'show_advanced': self.show_advanced.get(),
            },
            'compress': {
                'input_path': self.input_path.get(),
                'output_path': self.output_path.get(),
                'operation': self.operation.get(),
                'overwrite_originals': self.overwrite_originals.get(),
                'optimize_images': self.optimize_images.get(),
            },
            'adv_options': {key: var.get() for key, var in self.adv_options.items() if key != 'show_passwords'},
            'merge': {
                'files': self.merge_files,
                'output_path': self.merge_output_path.get(),
            },
            'split': {
                'input_path': self.split_input_path.get(),
                'output_dir': self.split_output_dir.get(),
                'mode': self.split_mode.get(),
                'value': self.split_value.get(),
            },
            'rotate': {
                'input': self.rotate_input.get(),
                'output': self.rotate_output.get(),
                'angle': self.rotate_angle.get(),
            },
            'delete': {
                'input': self.delete_pages_input.get(),
                'output': self.delete_pages_output.get(),
                'range': self.delete_pages_range.get(),
            },
            'convert': {
                'input': self.pdf_to_img_input.get(),
                'output_dir': self.pdf_to_img_output_dir.get(),
                'format': self.pdf_to_img_format.get(),
                'dpi': self.pdf_to_img_dpi.get(),
            },
            'pdfa': {
                'input': self.pdfa_input.get(),
                'output': self.pdfa_output.get(),
            },
            'utility': {
                'input': self.utility_pdf_in.get(),
            },
            'metadata': {
                'input': self.meta_pdf_path.get(),
            },
            'stamp': {
                'input': self.stamp_pdf_in.get(),
                'output': self.stamp_pdf_out.get(),
                'mode': self.stamp_mode.get(),
                'image_in': self.stamp_image_in.get(),
                'image_w': self.stamp_image_width.get(),
                'image_h': self.stamp_image_height.get(),
                'text': self.stamp_text_in.get(),
                'font': self.stamp_font.get(),
                'font_size': self.stamp_font_size.get(),
                'font_color': self.stamp_font_color_var.get(),
                'pos': self.stamp_pos.get(),
                'opacity': self.stamp_opacity.get(),
                'on_top': self.stamp_on_top.get(),
                'dyn_filename': self.stamp_dynamic_filename.get(),
                'dyn_datetime': self.stamp_dynamic_datetime.get(),
                'dyn_bates': self.stamp_dynamic_bates.get(),
                'bates_start': self.stamp_bates_start.get(),
            }
        }
        try:
            with open(self.settings_file, 'w') as f:
                json.dump(settings, f, indent=4)
        except Exception as e:
            logging.warning(f"Failed to save settings: {e}")

    def load_settings(self):
        if not self.settings_file.exists(): return
        try:
            with open(self.settings_file, 'r') as f:
                s = json.load(f)

            gen = s.get('general', {})
            self.dark_mode_enabled.set(gen.get('dark_mode_enabled', True))
            self.show_advanced.set(gen.get('show_advanced', False))

            comp = s.get('compress', {})
            self.input_path.set(comp.get('input_path', ''))
            self.output_path.set(comp.get('output_path', ''))
            self.operation.set(comp.get('operation', OP_COMPRESS_EBOOK))
            self.overwrite_originals.set(comp.get('overwrite_originals', False))
            self.optimize_images.set(comp.get('optimize_images', False))
            
            adv = s.get('adv_options', {})
            for key, value in adv.items():
                if key in self.adv_options:
                    self.adv_options[key].set(value)

            mer = s.get('merge', {})
            self.merge_files = mer.get('files', [])
            self.update_merge_listbox()
            self.merge_output_path.set(mer.get('output_path', ''))

            spl = s.get('split', {})
            self.split_input_path.set(spl.get('input_path', ''))
            self.split_output_dir.set(spl.get('output_dir', ''))
            self.split_mode.set(spl.get('mode', 'Split to Single Pages'))
            self.split_value.set(spl.get('value', '2'))

            rot = s.get('rotate', {})
            self.rotate_input.set(rot.get('input', ''))
            self.rotate_output.set(rot.get('output', ''))
            self.rotate_angle.set(rot.get('angle', '90° Right (Clockwise)'))

            dlt = s.get('delete', {})
            self.delete_pages_input.set(dlt.get('input', ''))
            self.delete_pages_output.set(dlt.get('output', ''))
            self.delete_pages_range.set(dlt.get('range', '1, 3-5'))

            cnv = s.get('convert', {})
            self.pdf_to_img_input.set(cnv.get('input', ''))
            self.pdf_to_img_output_dir.set(cnv.get('output_dir', ''))
            self.pdf_to_img_format.set(cnv.get('format', 'png'))
            self.pdf_to_img_dpi.set(cnv.get('dpi', '300'))
            
            pdfa = s.get('pdfa', {})
            self.pdfa_input.set(pdfa.get('input', ''))
            self.pdfa_output.set(pdfa.get('output', ''))

            util = s.get('utility', {})
            self.utility_pdf_in.set(util.get('input', ''))

            meta = s.get('metadata', {})
            self.meta_pdf_path.set(meta.get('input', ''))
            
            stmp = s.get('stamp', {})
            self.stamp_pdf_in.set(stmp.get('input', ''))
            self.stamp_pdf_out.set(stmp.get('output', ''))
            self.stamp_mode.set(stmp.get('mode', 'Image'))
            self.stamp_image_in.set(stmp.get('image_in', ''))
            self.stamp_image_width.set(stmp.get('image_w', ''))
            self.stamp_image_height.set(stmp.get('image_h', ''))
            self.stamp_text_in.set(stmp.get('text', 'CONFIDENTIAL'))
            self.stamp_font.set(stmp.get('font', 'Helvetica-Bold'))
            self.stamp_font_size.set(stmp.get('font_size', '48'))
            self.stamp_font_color_var.set(stmp.get('font_color', '#ff0000'))
            self.stamp_pos.set(stmp.get('pos', 'Center'))
            self.stamp_opacity.set(stmp.get('opacity', 0.5))
            self.stamp_on_top.set(stmp.get('on_top', True))
            self.stamp_dynamic_filename.set(stmp.get('dyn_filename', False))
            self.stamp_dynamic_datetime.set(stmp.get('dyn_datetime', False))
            self.stamp_dynamic_bates.set(stmp.get('dyn_bates', False))
            self.stamp_bates_start.set(stmp.get('bates_start', '1'))

        except Exception as e:
            logging.warning(f"Failed to load settings: {e}")

    def _create_preview_pane(self, parent):
        preview_frame = ttk.LabelFrame(parent, text="Preview", padding=10)
        preview_width = 300
        preview_height = 425

        sized_frame = ttk.Frame(preview_frame, width=preview_width, height=preview_height)
        sized_frame.pack(pady=5, expand=True, fill="both")
        sized_frame.pack_propagate(False)

        preview_label = ttk.Label(sized_frame, text="Select a file to see a preview.", anchor="center", justify="center", wraplength=preview_width-10)
        preview_label.pack(expand=True, fill="both")

        parent.rowconfigure(0, weight=1)
        return preview_frame, preview_label

    def build_gui(self):
        self.main_frame.rowconfigure(0, weight=1)
        self.main_frame.columnconfigure(0, weight=1)
        self.notebook = ttk.Notebook(self.main_frame)
        self.notebook.grid(row=0, column=0, sticky="nsew")

        def add_tab(name, text):
            scrolled_frame = ScrolledFrame(self.notebook)
            inner_frame = scrolled_frame.scrollable_frame
            self.notebook.add(scrolled_frame, text=text)
            self.tab_frames[name] = scrolled_frame
            return inner_frame

        #region: Compression Tab
        self.compress_frame = add_tab("compress", "Compress")
        self.compress_frame.columnconfigure(0, weight=1)
        FileSelector(self.compress_frame, "Input File/Folder:", self.input_path, self.browse_input, tooltips.TOOLTIP_TEXT['compress_input_entry'], tooltips.TOOLTIP_TEXT['compress_input_btn']).grid(row=0, column=0, sticky="we", pady=(0, 5))
        FileSelector(self.compress_frame, "Output File/Folder:", self.output_path, self.browse_output, tooltips.TOOLTIP_TEXT['compress_output_entry'], tooltips.TOOLTIP_TEXT['compress_output_btn']).grid(row=1, column=0, sticky="we", pady=5)

        options_frame = ttk.Frame(self.compress_frame)
        options_frame.grid(row=2, column=0, sticky="we", pady=5)
        options_frame.columnconfigure(1, weight=1)
        ttk.Label(options_frame, text="Compression Level:").grid(row=0, column=0, sticky="w")
        ttk.Combobox(options_frame, textvariable=self.operation, values=OPERATIONS, state="readonly").grid(row=0, column=1, padx=5, sticky="we")
        Tooltip(options_frame.winfo_children()[-1], tooltips.TOOLTIP_TEXT['compress_op_combo'])
        ttk.Label(options_frame, text="Image Resolution (DPI):").grid(row=1, column=0, sticky="w", pady=(5,0))
        ttk.Entry(options_frame, textvariable=self.adv_options['image_resolution'], width=10).grid(row=1, column=1, sticky="w", padx=5, pady=(5,0))
        Tooltip(options_frame.winfo_children()[-1], tooltips.TOOLTIP_TEXT['compress_dpi_entry'])

        ttk.Checkbutton(self.compress_frame, text="Show Advanced Settings", variable=self.show_advanced, command=self.toggle_advanced).grid(row=3, column=0, sticky="w", pady=5)
        Tooltip(self.compress_frame.winfo_children()[-1], tooltips.TOOLTIP_TEXT['compress_adv_check'])
        ttk.Checkbutton(self.compress_frame, text="Dark Mode", variable=self.dark_mode_enabled, command=self.toggle_theme).grid(row=4, column=0, sticky="w", pady=5)
        Tooltip(self.compress_frame.winfo_children()[-1], tooltips.TOOLTIP_TEXT['compress_dark_mode_check'])
        ttk.Checkbutton(self.compress_frame, text="Overwrite Original Files", variable=self.overwrite_originals).grid(row=5, column=0, sticky="w", pady=5)
        Tooltip(self.compress_frame.winfo_children()[-1], tooltips.TOOLTIP_TEXT['adv_final_overwrite'])
        ttk.Checkbutton(self.compress_frame, text="Optimize Images (zlib, sam2p)", variable=self.optimize_images).grid(row=6, column=0, sticky="w", pady=5)
        Tooltip(self.compress_frame.winfo_children()[-1], tooltips.TOOLTIP_TEXT['adv_optimize_images'])
        self.compress_button = ttk.Button(self.compress_frame, text="Process", command=self.process_conversion)
        self.compress_button.grid(row=7, column=0, pady=10)
        Tooltip(self.compress_button, tooltips.TOOLTIP_TEXT['compress_process_btn'])
        self.adv_frame = ttk.LabelFrame(self.compress_frame, text="Advanced Settings", padding=10)
        self.adv_frame.grid(row=8, column=0, sticky="nsew", pady=5)
        self.adv_frame.columnconfigure(1, weight=1)
        self.adv_frame.grid_remove()
        ttk.Label(self.adv_frame, text="Downscale Factor (1-8):").grid(row=0, column=0, sticky="w")
        ttk.Entry(self.adv_frame, textvariable=self.adv_options['downscale_factor'], width=10).grid(row=0, column=1, sticky="w", padx=5)
        Tooltip(self.adv_frame.winfo_children()[-1], tooltips.TOOLTIP_TEXT['adv_gs_downscale_factor'])
        ttk.Label(self.adv_frame, text="Color Conversion:").grid(row=1, column=0, sticky="w")
        ttk.Combobox(self.adv_frame, textvariable=self.adv_options['color_strategy'], values=["LeaveColorUnchanged", "Gray", "CMYK", "RGB"], state="readonly").grid(row=1, column=1, sticky="w", padx=5)
        Tooltip(self.adv_frame.winfo_children()[-1], tooltips.TOOLTIP_TEXT['adv_gs_color_conversion'])
        ttk.Label(self.adv_frame, text="Downsample Method:").grid(row=2, column=0, sticky="w")
        ttk.Combobox(self.adv_frame, textvariable=self.adv_options['downsample_type'], values=["Subsample", "Average", "Bicubic"], state="readonly").grid(row=2, column=1, sticky="w", padx=5)
        Tooltip(self.adv_frame.winfo_children()[-1], tooltips.TOOLTIP_TEXT['adv_gs_downsample_method'])
        ttk.Checkbutton(self.adv_frame, text="Fast Web View", variable=self.adv_options['fast_web_view']).grid(row=3, column=0, sticky="w")
        Tooltip(self.adv_frame.winfo_children()[-1], tooltips.TOOLTIP_TEXT['adv_gs_fast_web_view'])
        ttk.Checkbutton(self.adv_frame, text="Subset Fonts", variable=self.adv_options['subset_fonts']).grid(row=4, column=0, sticky="w")
        Tooltip(self.adv_frame.winfo_children()[-1], tooltips.TOOLTIP_TEXT['adv_gs_subset_fonts'])
        ttk.Checkbutton(self.adv_frame, text="Compress Fonts", variable=self.adv_options['compress_fonts']).grid(row=5, column=0, sticky="w")
        Tooltip(self.adv_frame.winfo_children()[-1], tooltips.TOOLTIP_TEXT['adv_gs_compress_fonts'])
        ttk.Checkbutton(self.adv_frame, text="Remove Interactive Elements", variable=self.adv_options['remove_interactive']).grid(row=6, column=0, sticky="w")
        Tooltip(self.adv_frame.winfo_children()[-1], tooltips.TOOLTIP_TEXT['adv_gs_remove_interactive'])
        ttk.Label(self.adv_frame, text="Final Processing", font=(None, 10, "bold")).grid(row=7, column=0, columnspan=2, pady=5)
        ttk.Label(self.adv_frame, text="Rotation:").grid(row=8, column=0, sticky="w")
        ttk.Combobox(self.adv_frame, textvariable=self.adv_options['rotation'], values=list(ROTATION_MAP.keys()), state="readonly").grid(row=8, column=1, sticky="w", padx=5)
        Tooltip(self.adv_frame.winfo_children()[-1], tooltips.TOOLTIP_TEXT['adv_final_rotation'])
        ttk.Label(self.adv_frame, text="Decimal Precision (1-10):").grid(row=9, column=0, sticky="w")
        ttk.Entry(self.adv_frame, textvariable=self.adv_options['decimal_precision'], width=10).grid(row=9, column=1, sticky="w", padx=5)
        Tooltip(self.adv_frame.winfo_children()[-1], tooltips.TOOLTIP_TEXT['adv_final_precision'])
        ttk.Checkbutton(self.adv_frame, text="Strip Metadata", variable=self.adv_options['strip_metadata']).grid(row=10, column=0, sticky="w")
        Tooltip(self.adv_frame.winfo_children()[-1], tooltips.TOOLTIP_TEXT['adv_final_strip_meta'])
        ttk.Checkbutton(self.adv_frame, text="Use cpdf Squeeze", variable=self.adv_options['use_cpdf_squeeze']).grid(row=11, column=0, sticky="w")
        Tooltip(self.adv_frame.winfo_children()[-1], tooltips.TOOLTIP_TEXT['adv_final_cpdf_squeeze'])
        ttk.Checkbutton(self.adv_frame, text="Darken Text", variable=self.adv_options['darken_text']).grid(row=12, column=0, sticky="w")
        Tooltip(self.adv_frame.winfo_children()[-1], tooltips.TOOLTIP_TEXT['adv_final_cpdf_darken'])
        ttk.Checkbutton(self.adv_frame, text="Use cpdf Fast Processing", variable=self.adv_options['use_fast_processing']).grid(row=13, column=0, sticky="w")
        Tooltip(self.adv_frame.winfo_children()[-1], tooltips.TOOLTIP_TEXT['adv_final_cpdf_fast'])
        ttk.Label(self.adv_frame, text="Security", font=(None, 10, "bold")).grid(row=14, column=0, columnspan=2, pady=5)
        ttk.Label(self.adv_frame, text="User Password:").grid(row=15, column=0, sticky="w")
        self.user_pass_entry = ttk.Entry(self.adv_frame, textvariable=self.adv_options['user_password'], show="*", width=20)
        self.user_pass_entry.grid(row=15, column=1, sticky="w", padx=5)
        Tooltip(self.user_pass_entry, tooltips.TOOLTIP_TEXT['adv_sec_user_pass'])
        ttk.Label(self.adv_frame, text="Owner Password:").grid(row=16, column=0, sticky="w")
        self.owner_pass_entry = ttk.Entry(self.adv_frame, textvariable=self.adv_options['owner_password'], show="*", width=20)
        self.owner_pass_entry.grid(row=16, column=1, sticky="w", padx=5)
        Tooltip(self.owner_pass_entry, tooltips.TOOLTIP_TEXT['adv_sec_owner_pass'])
        ttk.Checkbutton(self.adv_frame, text="Show Passwords", variable=self.adv_options['show_passwords'], command=self.toggle_password_visibility).grid(row=17, column=0, sticky="w")
        Tooltip(self.adv_frame.winfo_children()[-1], tooltips.TOOLTIP_TEXT['adv_sec_show_pass'])
        ttk.Label(self.adv_frame, text="Pikepdf Compression Level:").grid(row=18, column=0, sticky="w")
        ttk.Scale(self.adv_frame, from_=0, to=9, orient="horizontal", variable=self.adv_options['pikepdf_compression_level']).grid(row=18, column=1, sticky="we", padx=5)
        Tooltip(self.adv_frame.winfo_children()[-1], tooltips.TOOLTIP_TEXT['adv_pike_slider'])
        #endregion

        #region: Merge Tab
        self.merge_frame = add_tab("merge", "Merge")
        self.merge_frame.columnconfigure(0, weight=1)
        self.merge_frame.rowconfigure(1, weight=1)

        list_buttons_frame = ttk.Frame(self.merge_frame)
        list_buttons_frame.grid(row=0, column=0, sticky="w")
        ttk.Button(list_buttons_frame, text="Add Files", command=self.add_merge_files).pack(side="left", padx=(0,5))
        Tooltip(list_buttons_frame.winfo_children()[-1], tooltips.TOOLTIP_TEXT['merge_add_btn'])
        ttk.Button(list_buttons_frame, text="Remove File", command=self.remove_merge_file).pack(side="left", padx=5)
        Tooltip(list_buttons_frame.winfo_children()[-1], tooltips.TOOLTIP_TEXT['merge_remove_btn'])
        ttk.Button(list_buttons_frame, text="Move Up", command=self.move_merge_up).pack(side="left", padx=5)
        Tooltip(list_buttons_frame.winfo_children()[-1], tooltips.TOOLTIP_TEXT['merge_up_btn'])
        ttk.Button(list_buttons_frame, text="Move Down", command=self.move_merge_down).pack(side="left", padx=5)
        Tooltip(list_buttons_frame.winfo_children()[-1], tooltips.TOOLTIP_TEXT['merge_down_btn'])

        self.merge_listbox = tk.Listbox(self.merge_frame, height=10)
        self.merge_listbox.grid(row=1, column=0, pady=5, sticky="nsew")
        Tooltip(self.merge_listbox, "List of PDF files to merge.")

        FileSelector(self.merge_frame, "Output File:", self.merge_output_path, self.browse_merge_output, tooltips.TOOLTIP_TEXT['merge_output_entry'], tooltips.TOOLTIP_TEXT['merge_output_btn']).grid(row=2, column=0, sticky="we", pady=5)
        self.merge_button = ttk.Button(self.merge_frame, text="Merge", command=self.process_merge)
        self.merge_button.grid(row=3, column=0, pady=10)
        Tooltip(self.merge_button, tooltips.TOOLTIP_TEXT['merge_process_btn'])
        #endregion

        #region: Split Tab
        self.split_frame = add_tab("split", "Split")
        self.split_frame.columnconfigure(0, weight=1)
        FileSelector(self.split_frame, "Input PDF:", self.split_input_path, self.browse_split_input, tooltips.TOOLTIP_TEXT['split_input_entry'], tooltips.TOOLTIP_TEXT['split_input_btn']).grid(row=0, column=0, sticky="we", pady=(0, 5))

        split_options_frame = ttk.Frame(self.split_frame)
        split_options_frame.grid(row=1, column=0, sticky="we", pady=5)
        split_options_frame.columnconfigure(1, weight=1)
        ttk.Label(split_options_frame, text="Split Mode:").grid(row=0, column=0, sticky="w")
        ttk.Combobox(split_options_frame, textvariable=self.split_mode, values=["Split to Single Pages", "Split Every N Pages", "Custom Range(s)"], state="readonly").grid(row=0, column=1, padx=5, sticky="we")
        Tooltip(split_options_frame.winfo_children()[-1], tooltips.TOOLTIP_TEXT['split_mode_combo'])
        ttk.Label(split_options_frame, text="Value:").grid(row=1, column=0, sticky="w", pady=(5,0))
        ttk.Entry(split_options_frame, textvariable=self.split_value, width=20).grid(row=1, column=1, sticky="w", padx=5, pady=(5,0))
        Tooltip(split_options_frame.winfo_children()[-1], tooltips.TOOLTIP_TEXT['split_value_entry'])

        FileSelector(self.split_frame, "Output Directory:", self.split_output_dir, self.browse_split_output, tooltips.TOOLTIP_TEXT['split_output_dir_entry'], tooltips.TOOLTIP_TEXT['split_output_dir_btn']).grid(row=2, column=0, sticky="we", pady=5)
        self.split_button = ttk.Button(self.split_frame, text="Split", command=self.process_split)
        self.split_button.grid(row=3, column=0, pady=10)
        Tooltip(self.split_button, tooltips.TOOLTIP_TEXT['split_process_btn'])
        #endregion

        #region: Rotate Tab
        self.rotate_frame = add_tab("rotate", "Rotate")
        self.rotate_frame.columnconfigure(0, weight=1, minsize=350)
        self.rotate_frame.columnconfigure(1, weight=1)

        rotate_controls_frame = ttk.Frame(self.rotate_frame)
        rotate_controls_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        rotate_controls_frame.columnconfigure(0, weight=1)

        FileSelector(rotate_controls_frame, "Input PDF:", self.rotate_input, self.browse_rotate_input, tooltips.TOOLTIP_TEXT['rotate_input_entry'], tooltips.TOOLTIP_TEXT['rotate_input_btn']).grid(row=0, column=0, sticky="we", pady=(0, 5))

        rotate_options_frame = ttk.Frame(rotate_controls_frame)
        rotate_options_frame.grid(row=1, column=0, sticky="we", pady=5)
        rotate_options_frame.columnconfigure(1, weight=1)
        ttk.Label(rotate_options_frame, text="Rotation Angle:").grid(row=0, column=0, sticky="w")
        ttk.Combobox(rotate_options_frame, textvariable=self.rotate_angle, values=list(ROTATION_MAP.keys()), state="readonly").grid(row=0, column=1, columnspan=2, padx=5, sticky="we")
        Tooltip(rotate_options_frame.winfo_children()[-1], tooltips.TOOLTIP_TEXT['rotate_angle_combo'])

        FileSelector(rotate_controls_frame, "Output PDF:", self.rotate_output, self.browse_rotate_output, tooltips.TOOLTIP_TEXT['rotate_output_entry'], tooltips.TOOLTIP_TEXT['rotate_output_btn']).grid(row=2, column=0, sticky="we", pady=5)
        self.rotate_button = ttk.Button(rotate_controls_frame, text="Rotate", command=self.process_rotate)
        self.rotate_button.grid(row=3, column=0, pady=10)
        Tooltip(self.rotate_button, tooltips.TOOLTIP_TEXT['rotate_process_btn'])

        preview_pane, self.rotate_preview_label = self._create_preview_pane(self.rotate_frame)
        preview_pane.grid(row=0, column=1, sticky="nsew", rowspan=4)
        #endregion

        #region: Delete Pages Tab
        self.delete_pages_frame = add_tab("delete", "Delete Pages")
        self.delete_pages_frame.columnconfigure(0, weight=1)
        FileSelector(self.delete_pages_frame, "Input PDF:", self.delete_pages_input, self.browse_delete_input, tooltips.TOOLTIP_TEXT['delete_input_entry'], tooltips.TOOLTIP_TEXT['delete_input_btn']).grid(row=0, column=0, sticky="we", pady=(0,5))

        delete_options_frame = ttk.Frame(self.delete_pages_frame)
        delete_options_frame.grid(row=1, column=0, sticky="we", pady=5)
        delete_options_frame.columnconfigure(1, weight=1)
        ttk.Label(delete_options_frame, text="Pages to Delete:").grid(row=0, column=0, sticky="w")
        ttk.Entry(delete_options_frame, textvariable=self.delete_pages_range, width=20).grid(row=0, column=1, sticky="w", padx=5)
        Tooltip(delete_options_frame.winfo_children()[-1], tooltips.TOOLTIP_TEXT['delete_pages_entry'])

        FileSelector(self.delete_pages_frame, "Output PDF:", self.delete_pages_output, self.browse_delete_output, tooltips.TOOLTIP_TEXT['delete_output_entry'], tooltips.TOOLTIP_TEXT['delete_output_btn']).grid(row=2, column=0, sticky="we", pady=5)
        self.delete_button = ttk.Button(self.delete_pages_frame, text="Delete Pages", command=self.process_delete_pages)
        self.delete_button.grid(row=3, column=0, pady=10)
        Tooltip(self.delete_button, tooltips.TOOLTIP_TEXT['delete_process_btn'])
        #endregion

        #region: Convert to Images Tab
        self.convert_frame = add_tab("convert", "PDF to Images")
        self.convert_frame.columnconfigure(0, weight=1)
        FileSelector(self.convert_frame, "Input PDF:", self.pdf_to_img_input, self.browse_convert_input, tooltips.TOOLTIP_TEXT['convert_input_entry'], tooltips.TOOLTIP_TEXT['convert_input_btn']).grid(row=0, column=0, sticky="we", pady=(0,5))
        FileSelector(self.convert_frame, "Output Directory:", self.pdf_to_img_output_dir, self.browse_convert_output, tooltips.TOOLTIP_TEXT['convert_output_entry'], tooltips.TOOLTIP_TEXT['convert_output_btn']).grid(row=1, column=0, sticky="we", pady=5)

        convert_opts_frame = ttk.Frame(self.convert_frame)
        convert_opts_frame.grid(row=2, column=0, sticky="we", pady=5)
        convert_opts_frame.columnconfigure(1, weight=1)
        ttk.Label(convert_opts_frame, text="Image Format:").grid(row=0, column=0, sticky="w")
        ttk.Combobox(convert_opts_frame, textvariable=self.pdf_to_img_format, values=["png", "jpeg", "tiff"], state="readonly").grid(row=0, column=1, sticky="w", padx=5)
        Tooltip(convert_opts_frame.winfo_children()[-1], tooltips.TOOLTIP_TEXT['convert_format_combo'])
        ttk.Label(convert_opts_frame, text="DPI:").grid(row=1, column=0, sticky="w", pady=(5,0))
        ttk.Entry(convert_opts_frame, textvariable=self.pdf_to_img_dpi, width=10).grid(row=1, column=1, sticky="w", padx=5, pady=(5,0))
        Tooltip(convert_opts_frame.winfo_children()[-1], tooltips.TOOLTIP_TEXT['convert_dpi_entry'])

        self.convert_button = ttk.Button(self.convert_frame, text="Convert", command=self.process_convert)
        self.convert_button.grid(row=3, column=0, pady=10)
        Tooltip(self.convert_button, tooltips.TOOLTIP_TEXT['convert_process_btn'])
        #endregion

        #region: PDF/A Tab
        self.pdfa_frame = add_tab("pdfa", "PDF/A")
        self.pdfa_frame.columnconfigure(0, weight=1)
        FileSelector(self.pdfa_frame, "Input PDF:", self.pdfa_input, self.browse_pdfa_input, tooltips.TOOLTIP_TEXT['pdfa_input_entry'], tooltips.TOOLTIP_TEXT['pdfa_input_btn']).grid(row=0, column=0, sticky="we", pady=(0,5))
        FileSelector(self.pdfa_frame, "Output PDF:", self.pdfa_output, self.browse_pdfa_output, tooltips.TOOLTIP_TEXT['pdfa_output_entry'], tooltips.TOOLTIP_TEXT['pdfa_output_btn']).grid(row=1, column=0, sticky="we", pady=5)
        self.pdfa_button = ttk.Button(self.pdfa_frame, text="Convert to PDF/A", command=self.process_pdfa)
        self.pdfa_button.grid(row=2, column=0, pady=10)
        Tooltip(self.pdfa_button, tooltips.TOOLTIP_TEXT['pdfa_process_btn'])
        #endregion

        #region: Utility Tab
        self.utility_frame = add_tab("utility", "Utilities")
        self.utility_frame.columnconfigure(0, weight=1)
        FileSelector(self.utility_frame, "Input PDF:", self.utility_pdf_in, self.browse_utility_input, tooltips.TOOLTIP_TEXT['util_input_entry'], tooltips.TOOLTIP_TEXT['util_input_btn']).grid(row=0, column=0, sticky="we", pady=(0,5))
        self.utility_button = ttk.Button(self.utility_frame, text="Remove Open Action", command=self.process_remove_open_action)
        self.utility_button.grid(row=1, column=0, pady=10)
        Tooltip(self.utility_button, tooltips.TOOLTIP_TEXT['util_process_btn'])
        #endregion

        #region: Metadata Tab
        self.meta_frame = add_tab("metadata", "Metadata")
        self.meta_frame.columnconfigure(0, weight=1)
        FileSelector(self.meta_frame, "Input PDF:", self.meta_pdf_path, self.browse_meta_input, tooltips.TOOLTIP_TEXT['meta_input_entry'], tooltips.TOOLTIP_TEXT['meta_input_btn']).grid(row=0, column=0, sticky="we", pady=(0,5))

        meta_buttons = ttk.Frame(self.meta_frame)
        meta_buttons.grid(row=1, column=0, sticky="w", pady=5)
        ttk.Button(meta_buttons, text="Load Metadata", command=self.load_metadata).pack(side="left")
        Tooltip(meta_buttons.winfo_children()[-1], tooltips.TOOLTIP_TEXT['meta_load_btn'])

        meta_fields_frame = ttk.LabelFrame(self.meta_frame, text="Fields", padding=10)
        meta_fields_frame.grid(row=2, column=0, sticky="we", pady=5)
        meta_fields_frame.columnconfigure(1, weight=1)
        ttk.Label(meta_fields_frame, text="Title:").grid(row=0, column=0, sticky="w")
        ttk.Entry(meta_fields_frame, textvariable=self.meta_title).grid(row=0, column=1, padx=5, pady=2, sticky="we")
        Tooltip(meta_fields_frame.winfo_children()[-1], tooltips.TOOLTIP_TEXT['meta_title'])
        ttk.Label(meta_fields_frame, text="Author:").grid(row=1, column=0, sticky="w")
        ttk.Entry(meta_fields_frame, textvariable=self.meta_author).grid(row=1, column=1, padx=5, pady=2, sticky="we")
        Tooltip(meta_fields_frame.winfo_children()[-1], tooltips.TOOLTIP_TEXT['meta_author'])
        ttk.Label(meta_fields_frame, text="Subject:").grid(row=2, column=0, sticky="w")
        ttk.Entry(meta_fields_frame, textvariable=self.meta_subject).grid(row=2, column=1, padx=5, pady=2, sticky="we")
        Tooltip(meta_fields_frame.winfo_children()[-1], tooltips.TOOLTIP_TEXT['meta_subject'])
        ttk.Label(meta_fields_frame, text="Keywords:").grid(row=3, column=0, sticky="w")
        ttk.Entry(meta_fields_frame, textvariable=self.meta_keywords).grid(row=3, column=1, padx=5, pady=2, sticky="we")
        Tooltip(meta_fields_frame.winfo_children()[-1], tooltips.TOOLTIP_TEXT['meta_keywords'])

        self.meta_button = ttk.Button(self.meta_frame, text="Save Metadata", command=self.save_metadata)
        self.meta_button.grid(row=3, column=0, pady=10)
        Tooltip(self.meta_button, tooltips.TOOLTIP_TEXT['meta_save_btn'])
        #endregion

        #region: Stamp/Watermark Tab
        self.stamp_frame = add_tab("stamp", "Stamp/Watermark")
        self.stamp_frame.columnconfigure(0, weight=1, minsize=420)
        self.stamp_frame.columnconfigure(1, weight=1)

        stamp_controls_frame = ttk.Frame(self.stamp_frame)
        stamp_controls_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        stamp_controls_frame.columnconfigure(0, weight=1)

        FileSelector(stamp_controls_frame, "Input PDF:", self.stamp_pdf_in, self.browse_stamp_input, tooltips.TOOLTIP_TEXT['stamp_input_entry'], tooltips.TOOLTIP_TEXT['stamp_input_btn']).grid(row=0, column=0, sticky="we", pady=(0,5))

        mode_frame = ttk.Frame(stamp_controls_frame)
        mode_frame.grid(row=1, column=0, sticky="w", pady=5)
        ttk.Radiobutton(mode_frame, text="Image Stamp", variable=self.stamp_mode, value="Image", command=self.toggle_stamp_mode).pack(side="left", padx=(0,10))
        Tooltip(mode_frame.winfo_children()[-1], tooltips.TOOLTIP_TEXT['stamp_image_radio'])
        ttk.Radiobutton(mode_frame, text="Text Stamp", variable=self.stamp_mode, value="Text", command=self.toggle_stamp_mode).pack(side="left")
        Tooltip(mode_frame.winfo_children()[-1], tooltips.TOOLTIP_TEXT['stamp_text_radio'])

        self.stamp_image_frame = ttk.Frame(stamp_controls_frame)
        self.stamp_image_frame.grid(row=2, column=0, sticky="we", pady=5)
        self.stamp_image_frame.columnconfigure(0, weight=1)
        FileSelector(self.stamp_image_frame, "Image File:", self.stamp_image_in, self.browse_stamp_image, tooltips.TOOLTIP_TEXT['stamp_image_entry'], tooltips.TOOLTIP_TEXT['stamp_image_btn']).grid(row=0, column=0, sticky="we", pady=(0,5))
        size_frame = ttk.Frame(self.stamp_image_frame)
        size_frame.grid(row=1, column=0, sticky="w")
        ttk.Label(size_frame, text="Width (px):").pack(side="left")
        ttk.Entry(size_frame, textvariable=self.stamp_image_width, width=10).pack(side="left", padx=5)
        Tooltip(size_frame.winfo_children()[-1], tooltips.TOOLTIP_TEXT['stamp_image_width'])
        ttk.Label(size_frame, text="Height (px):").pack(side="left", padx=(10,0))
        ttk.Entry(size_frame, textvariable=self.stamp_image_height, width=10).pack(side="left", padx=5)
        Tooltip(size_frame.winfo_children()[-1], tooltips.TOOLTIP_TEXT['stamp_image_height'])

        self.stamp_text_frame = ttk.Frame(stamp_controls_frame)
        self.stamp_text_frame.grid(row=2, column=0, sticky="we", pady=5)
        self.stamp_text_frame.columnconfigure(1, weight=1)
        self.stamp_text_frame.grid_remove()
        ttk.Label(self.stamp_text_frame, text="Text:").grid(row=0, column=0, sticky="w")
        ttk.Entry(self.stamp_text_frame, textvariable=self.stamp_text_in).grid(row=0, column=1, columnspan=2, padx=5, pady=2, sticky="we")
        Tooltip(self.stamp_text_frame.winfo_children()[-1], tooltips.TOOLTIP_TEXT['stamp_text_entry'])
        ttk.Label(self.stamp_text_frame, text="Font:").grid(row=1, column=0, sticky="w")
        ttk.Combobox(self.stamp_text_frame, textvariable=self.stamp_font, values=PDF_FONTS, state="readonly").grid(row=1, column=1, sticky="w", padx=5, pady=2)
        Tooltip(self.stamp_text_frame.winfo_children()[-1], tooltips.TOOLTIP_TEXT['stamp_font_combo'])
        ttk.Label(self.stamp_text_frame, text="Font Size:").grid(row=2, column=0, sticky="w")
        ttk.Entry(self.stamp_text_frame, textvariable=self.stamp_font_size, width=10).grid(row=2, column=1, sticky="w", padx=5, pady=2)
        Tooltip(self.stamp_text_frame.winfo_children()[-1], tooltips.TOOLTIP_TEXT['stamp_font_size'])
        ttk.Button(self.stamp_text_frame, text="Pick Color", command=self.pick_stamp_color).grid(row=3, column=0, padx=5, pady=2)
        Tooltip(self.stamp_text_frame.winfo_children()[-1], tooltips.TOOLTIP_TEXT['stamp_color_btn'])
        dyn_frame = ttk.Frame(self.stamp_text_frame)
        dyn_frame.grid(row=4, column=0, columnspan=2, sticky="w", pady=2)
        ttk.Checkbutton(dyn_frame, text="Add Filename", variable=self.stamp_dynamic_filename).pack(side="left", padx=(0,10))
        Tooltip(dyn_frame.winfo_children()[-1], tooltips.TOOLTIP_TEXT['stamp_dyn_filename'])
        ttk.Checkbutton(dyn_frame, text="Add Date/Time", variable=self.stamp_dynamic_datetime).pack(side="left")
        Tooltip(dyn_frame.winfo_children()[-1], tooltips.TOOLTIP_TEXT['stamp_dyn_datetime'])
        bates_frame = ttk.Frame(self.stamp_text_frame)
        bates_frame.grid(row=5, column=0, columnspan=2, sticky="w", pady=2)
        ttk.Checkbutton(bates_frame, text="Add Bates Numbering", variable=self.stamp_dynamic_bates).pack(side="left")
        Tooltip(bates_frame.winfo_children()[-1], tooltips.TOOLTIP_TEXT['stamp_dyn_bates_check'])
        ttk.Label(bates_frame, text="Bates Start:").pack(side="left", padx=(10,0))
        ttk.Entry(bates_frame, textvariable=self.stamp_bates_start, width=10).pack(side="left", padx=5)
        Tooltip(bates_frame.winfo_children()[-1], tooltips.TOOLTIP_TEXT['stamp_dyn_bates_entry'])

        common_stamp_frame = ttk.Frame(stamp_controls_frame)
        common_stamp_frame.grid(row=3, column=0, sticky="we", pady=5)
        common_stamp_frame.columnconfigure(1, weight=1)
        ttk.Label(common_stamp_frame, text="Position:").grid(row=0, column=0, sticky="w")
        ttk.Combobox(common_stamp_frame, textvariable=self.stamp_pos, values=["Center", "Bottom-Left", "Bottom-Right"], state="readonly").grid(row=0, column=1, sticky="w", padx=5)
        Tooltip(common_stamp_frame.winfo_children()[-1], tooltips.TOOLTIP_TEXT['stamp_pos_combo'])
        ttk.Label(common_stamp_frame, text="Opacity:").grid(row=1, column=0, sticky="w", pady=(5,0))
        ttk.Scale(common_stamp_frame, from_=0.1, to=1.0, orient="horizontal", variable=self.stamp_opacity).grid(row=1, column=1, sticky="we", padx=5, pady=(5,0))
        Tooltip(common_stamp_frame.winfo_children()[-1], tooltips.TOOLTIP_TEXT['stamp_opacity_slider'])
        ttk.Checkbutton(common_stamp_frame, text="Stamp on Top", variable=self.stamp_on_top).grid(row=2, column=0, columnspan=2, sticky="w", pady=(5,0))
        Tooltip(common_stamp_frame.winfo_children()[-1], tooltips.TOOLTIP_TEXT['stamp_ontop_check'])

        FileSelector(stamp_controls_frame, "Output PDF:", self.stamp_pdf_out, self.browse_stamp_output, tooltips.TOOLTIP_TEXT['stamp_output_entry'], tooltips.TOOLTIP_TEXT['stamp_output_btn']).grid(row=4, column=0, sticky="we", pady=5)
        self.stamp_button = ttk.Button(stamp_controls_frame, text="Apply Stamp", command=self.process_stamp)
        self.stamp_button.grid(row=5, column=0, pady=10)
        Tooltip(self.stamp_button, tooltips.TOOLTIP_TEXT['stamp_process_btn'])

        preview_pane, self.stamp_preview_label = self._create_preview_pane(self.stamp_frame)
        preview_pane.grid(row=0, column=1, sticky="nsew", rowspan=6)
        #endregion

        #region: Status Bar & Progress Popup
        self.status_bar = ttk.Label(self.main_frame, textvariable=self.status, anchor="w")
        self.status_bar.grid(row=1, column=0, sticky="we", pady=(5,0))

        self.progress_popup = tk.Toplevel(self.root)
        self.progress_popup.title("Processing...")
        self.progress_popup.transient(self.root)
        self.progress_popup.resizable(False, False)
        self.progress_popup.protocol("WM_DELETE_WINDOW", lambda: None)

        popup_frame = ttk.Frame(self.progress_popup, padding=20)
        popup_frame.pack(expand=True, fill="both")

        ttk.Label(popup_frame, textvariable=self.progress_status, anchor="w").pack(pady=5, fill="x", expand=True)
        ttk.Label(popup_frame, text="Current File Progress:").pack(pady=(5,0), anchor="w")
        self.file_progress_bar = ttk.Progressbar(popup_frame, mode="determinate", maximum=100, length=300)
        self.file_progress_bar.pack(pady=5, fill="x", expand=True)
        ttk.Label(popup_frame, text="Overall Progress:").pack(pady=(5,0), anchor="w")
        self.overall_progress_bar = ttk.Progressbar(popup_frame, mode="determinate", maximum=100, length=300)
        self.overall_progress_bar.pack(pady=5, fill="x", expand=True)

        self.progress_popup.withdraw()
        #endregion

    def setup_traces(self):
        self.adv_options['user_password'].trace_add("write", self.update_final_processor_state)
        self.adv_options['owner_password'].trace_add("write", self.update_final_processor_state)
        self.adv_options['use_cpdf_squeeze'].trace_add("write", self.update_final_processor_state)
        self.adv_options['darken_text'].trace_add("write", self.update_final_processor_state)
        self.toggle_advanced()
        self.toggle_stamp_mode()
        self.operation.trace_add("write", self.update_dpi_default)
        self.input_path.trace_add("write", self.update_is_folder)

        self.rotate_angle.trace_add("write", self.trigger_rotate_preview_debounce)

        stamp_vars_to_trace = [
            self.stamp_mode, self.stamp_image_in, self.stamp_image_width,
            self.stamp_image_height, self.stamp_text_in, self.stamp_font,
            self.stamp_font_size, self.stamp_pos, self.stamp_opacity,
            self.stamp_on_top, self.stamp_dynamic_filename, self.stamp_dynamic_datetime,
            self.stamp_dynamic_bates, self.stamp_bates_start
        ]
        for var in stamp_vars_to_trace:
            var.trace_add("write", self.trigger_stamp_preview_debounce)

    def trigger_rotate_preview_debounce(self, *args):
        if self.rotate_debounce_timer:
            self.root.after_cancel(self.rotate_debounce_timer)
        self.rotate_debounce_timer = self.root.after(500, self.update_rotate_preview)

    def trigger_stamp_preview_debounce(self, *args):
        if self.stamp_debounce_timer:
            self.root.after_cancel(self.stamp_debounce_timer)
        self.stamp_debounce_timer = self.root.after(500, self.update_stamp_preview)

    def _center_popup(self):
        self.root.update_idletasks()
        self.progress_popup.update_idletasks()
        popup_width = self.progress_popup.winfo_reqwidth()
        popup_height = self.progress_popup.winfo_reqheight()
        root_x = self.root.winfo_x()
        root_y = self.root.winfo_y()
        root_width = self.root.winfo_width()
        root_height = self.root.winfo_height()
        center_x = root_x + (root_width // 2) - (popup_width // 2)
        center_y = root_y + (root_height // 2) - (popup_height // 2)
        self.progress_popup.geometry(f"+{center_x}+{center_y}")

    def update_is_folder(self, *args):
        path_str = self.input_path.get()
        self.is_folder.set(Path(path_str).is_dir() if path_str else False)

    def browse_input(self):
        if messagebox.askyesno("Select Input Type", "Do you want to select a folder?\n\n(Choose 'No' for a single file.)", parent=self.root):
            path = filedialog.askdirectory(mustexist=True)
        else:
            path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if path: self.input_path.set(path)

    def browse_output(self):
        if self.is_folder.get():
            path = filedialog.askdirectory(mustexist=True)
        else:
            path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
        if path: self.output_path.set(path)

    def browse_merge_output(self):
        path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
        if path: self.merge_output_path.set(path)

    def browse_split_input(self):
        path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if path: self.split_input_path.set(path)

    def browse_split_output(self):
        path = filedialog.askdirectory(mustexist=True)
        if path: self.split_output_dir.set(path)

    def browse_rotate_input(self):
        path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if path:
            self.rotate_input.set(path)
            self.trigger_rotate_preview_debounce()

    def browse_rotate_output(self):
        path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
        if path: self.rotate_output.set(path)

    def browse_delete_input(self):
        path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if path: self.delete_pages_input.set(path)

    def browse_delete_output(self):
        path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
        if path: self.delete_pages_output.set(path)

    def browse_convert_input(self):
        path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if path: self.pdf_to_img_input.set(path)

    def browse_convert_output(self):
        path = filedialog.askdirectory(mustexist=True)
        if path: self.pdf_to_img_output_dir.set(path)

    def browse_pdfa_input(self):
        path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if path: self.pdfa_input.set(path)

    def browse_pdfa_output(self):
        path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
        if path: self.pdfa_output.set(path)

    def browse_utility_input(self):
        path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if path: self.utility_pdf_in.set(path)

    def browse_meta_input(self):
        path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if path: self.meta_pdf_path.set(path)

    def browse_stamp_input(self):
        path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if path:
            self.stamp_pdf_in.set(path)
            self.trigger_stamp_preview_debounce()

    def browse_stamp_output(self):
        path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
        if path: self.stamp_pdf_out.set(path)

    def browse_stamp_image(self):
        path = filedialog.askopenfilename(filetypes=[("Image files", "*.png *.jpg *.jpeg")])
        if path: self.stamp_image_in.set(path)

    def toggle_advanced(self):
        self.adv_frame.grid() if self.show_advanced.get() else self.adv_frame.grid_remove()

    def toggle_theme(self):
        colors = styles.apply_theme(self.root, 'dark' if self.dark_mode_enabled.get() else 'light')
        bg_color = colors['BG']
        for frame in self.tab_frames.values():
            if hasattr(frame, 'canvas'):
                frame.canvas.config(background=bg_color)

    def toggle_password_visibility(self):
        show = "" if self.adv_options['show_passwords'].get() else "*"
        self.user_pass_entry.config(show=show)
        self.owner_pass_entry.config(show=show)

    def toggle_stamp_mode(self):
        if self.stamp_mode.get() == "Image":
            self.stamp_image_frame.grid()
            self.stamp_text_frame.grid_remove()
        else:
            self.stamp_image_frame.grid_remove()
            self.stamp_text_frame.grid()

    def pick_stamp_color(self):
        color = colorchooser.askcolor(title="Select Text Color", initialcolor=self.stamp_font_color_var.get())
        if color[1]:
            self.stamp_font_color_var.set(color[1])
            self.trigger_stamp_preview_debounce()

    def setup_drag_and_drop(self):
        if IS_WINDOWS: windnd.hook_dropfiles(self.root, func=self.handle_drop)

    def handle_drop(self, files):
        decoded_files = [f.decode('utf-8') for f in files]
        if not decoded_files: return
        try:
            selected_tab_widget = self.notebook.nametowidget(self.notebook.select())
        except tk.TclError: return

        path = decoded_files[0]
        if selected_tab_widget == self.tab_frames.get("compress"): self.input_path.set(path)
        elif selected_tab_widget == self.tab_frames.get("merge"):
            self.merge_files.extend(decoded_files)
            self.update_merge_listbox()
        elif selected_tab_widget == self.tab_frames.get("split"): self.split_input_path.set(path)
        elif selected_tab_widget == self.tab_frames.get("rotate"):
            self.rotate_input.set(path)
            self.trigger_rotate_preview_debounce()
        elif selected_tab_widget == self.tab_frames.get("delete"): self.delete_pages_input.set(path)
        elif selected_tab_widget == self.tab_frames.get("convert"): self.pdf_to_img_input.set(path)
        elif selected_tab_widget == self.tab_frames.get("pdfa"): self.pdfa_input.set(path)
        elif selected_tab_widget == self.tab_frames.get("utility"): self.utility_pdf_in.set(path)
        elif selected_tab_widget == self.tab_frames.get("metadata"): self.meta_pdf_path.set(path)
        elif selected_tab_widget == self.tab_frames.get("stamp"):
            self.stamp_pdf_in.set(path)
            self.trigger_stamp_preview_debounce()

    def update_merge_listbox(self):
        self.merge_listbox.delete(0, tk.END)
        for file in self.merge_files: self.merge_listbox.insert(tk.END, file)

    def add_merge_files(self):
        files = filedialog.askopenfilenames(filetypes=[("PDF files", "*.pdf")])
        if files:
            self.merge_files.extend(files)
            self.update_merge_listbox()

    def remove_merge_file(self):
        selection = self.merge_listbox.curselection()
        if selection:
            self.merge_files.pop(selection[0])
            self.update_merge_listbox()

    def move_merge_up(self):
        selection = self.merge_listbox.curselection()
        if selection and selection[0] > 0:
            idx = selection[0]
            self.merge_files[idx], self.merge_files[idx-1] = self.merge_files[idx-1], self.merge_files[idx]
            self.update_merge_listbox()
            self.merge_listbox.selection_set(idx-1)

    def move_merge_down(self):
        selection = self.merge_listbox.curselection()
        if selection and selection[0] < len(self.merge_files) - 1:
            idx = selection[0]
            self.merge_files[idx], self.merge_files[idx+1] = self.merge_files[idx+1], self.merge_files[idx]
            self.update_merge_listbox()
            self.merge_listbox.selection_set(idx+1)

    def update_dpi_default(self, *args):
        if not self.adv_options['image_resolution'].get():
            self.adv_options['image_resolution'].set(DPI_PRESETS.get(self.operation.get(), "150"))

    def check_tools(self):
        try:
            self.gs_path = backend.find_ghostscript()
            self.cpdf_path = backend.find_cpdf()
        except backend.ToolNotFound as e:
            messagebox.showerror("Critical Error", f"{e}\n\nPlease ensure the 'bin' folder with all required tools is next to the application.\n\nThe application will now close.", parent=self.root)
            self.root.destroy()

    def update_final_processor_state(self, *args):
        use_cpdf = (self.adv_options['use_cpdf_squeeze'].get() or self.adv_options['darken_text'].get() or self.adv_options['user_password'].get() or self.adv_options['owner_password'].get())
        if use_cpdf and not self.cpdf_path:
            messagebox.showwarning("Warning", "cpdf is required for the selected advanced options but was not found.", parent=self.root)

    def start_task(self, button, target_func, args):
        if self.active_process_button:
            messagebox.showwarning("Busy", "A process is already running.", parent=self.root)
            return
        self.active_process_button = button

        self.file_progress_bar['value'] = 0
        self.overall_progress_bar['value'] = 0
        self._center_popup()
        self.progress_popup.deiconify()
        self.progress_popup.grab_set()

        threading.Thread(target=target_func, args=args, daemon=True).start()

    def on_task_complete(self, final_status):
        self.progress_popup.grab_release()
        self.progress_popup.withdraw()
        self.status.set(final_status)
        if "Error:" in final_status:
            messagebox.showerror("Task Failed", final_status, parent=self.root)

        if self.active_process_button:
            self.active_process_button = None

    def update_rotate_preview(self):
        pdf_path = self.rotate_input.get()
        if not pdf_path:
            return

        self.rotate_preview_label.config(image=None, text="Generating preview...")
        self.root.update_idletasks()

        options = {'angle': ROTATION_MAP.get(self.rotate_angle.get(), 0)}

        threading.Thread(
            target=self.run_preview_task,
            args=(pdf_path, 'rotate', options, self.rotate_preview_label, 'rotate_preview_image_tk'),
            daemon=True
        ).start()

    def update_stamp_preview(self):
        pdf_path = self.stamp_pdf_in.get()
        if not pdf_path:
            return

        self.stamp_preview_label.config(image=None, text="Generating preview...")
        self.root.update_idletasks()

        stamp_opts = {
            'pos': self.stamp_pos.get(), 'opacity': self.stamp_opacity.get(), 'on_top': self.stamp_on_top.get(),
        }

        mode_opts = {
            'image_path': self.stamp_image_in.get(), 'width': self.stamp_image_width.get(),
            'height': self.stamp_image_height.get(), 'text': self.stamp_text_in.get(),
            'font': self.stamp_font.get(), 'size': self.stamp_font_size.get(),
            'color': self.stamp_font_color_var.get(),
            'bates_start': self.stamp_bates_start.get() if self.stamp_dynamic_bates.get() else None
        }

        if self.stamp_mode.get() == "Text":
            current_text = self.stamp_text_in.get()
            if self.stamp_dynamic_filename.get() and pdf_path: current_text += f" {Path(pdf_path).name}"
            if self.stamp_dynamic_datetime.get(): current_text += f" {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
            mode_opts['text'] = current_text

        options = { 'stamp_opts': stamp_opts, 'mode_opts': mode_opts, 'mode': self.stamp_mode.get() }

        threading.Thread(
            target=self.run_preview_task,
            args=(pdf_path, 'stamp', options, self.stamp_preview_label, 'stamp_preview_image_tk'),
            daemon=True
        ).start()

    def run_preview_task(self, pdf_path, operation, options, target_label, image_holder_attr):
        try:
            image_bytes = backend.generate_preview_image(self.gs_path, self.cpdf_path, pdf_path, operation, options)
            img = Image.open(io.BytesIO(image_bytes))

            label_w = target_label.master.winfo_width()
            label_h = target_label.master.winfo_height()
            img.thumbnail((label_w -10, label_h -10), Image.LANCZOS)

            photo_image = ImageTk.PhotoImage(img)
            setattr(self, image_holder_attr, photo_image)
            target_label.config(image=photo_image, text="")

        except Exception as e:
            logging.error(f"Preview generation failed: {e}", exc_info=True)
            target_label.config(image=None, text=f"Preview failed:\n{e}", foreground="red")

    def process_conversion(self):
        try:
            input_p_str = self.input_path.get()
            if not input_p_str:
                messagebox.showwarning("Input Missing", "Please select an input file or folder.", parent=self.root)
                return
            is_overwrite = self.overwrite_originals.get()
            if is_overwrite and not messagebox.askyesno("Confirm Overwrite", "This will permanently replace your original file(s). This action cannot be undone. Are you sure?", icon='warning', parent=self.root):
                return
            output_p_str = self.output_path.get() if not is_overwrite else input_p_str
            if not output_p_str:
                messagebox.showwarning("Output Missing", "Please select an output location.", parent=self.root)
                return
            params = { 'gs_path': self.gs_path, 'cpdf_path': self.cpdf_path, 'input_path': input_p_str, 'output_path': output_p_str, 'operation': self.operation.get(), 'options': {k: v.get() for k, v in self.adv_options.items()}, 'overwrite': is_overwrite, 'optimize_images': self.optimize_images.get() }
            self.start_task(self.compress_button, backend.run_conversion_task, (params, self.is_folder.get(), self.progress_status, self.file_progress_bar, self.overall_progress_bar, self.on_task_complete))
        except Exception as e:
            messagebox.showerror("GUI Error", f"An unexpected error occurred: {e}", parent=self.root)
            self.on_task_complete("Ready")

    def process_merge(self):
        try:
            if not self.merge_files: messagebox.showwarning("Input Missing", "Please add files to the merge list.", parent=self.root); return
            output_path = self.merge_output_path.get()
            if not output_path: messagebox.showwarning("Output Missing", "Please select an output file.", parent=self.root); return
            self.start_task(self.merge_button, backend.run_merge_task, (self.merge_files, output_path, self.progress_status, self.file_progress_bar, self.overall_progress_bar, self.on_task_complete))
        except Exception as e:
            messagebox.showerror("GUI Error", f"An unexpected error occurred: {e}", parent=self.root)
            self.on_task_complete("Ready")

    def process_split(self):
        try:
            input_path = self.split_input_path.get()
            output_dir = self.split_output_dir.get()
            if not input_path or not output_dir: messagebox.showwarning("Input Missing", "Please select an input PDF and output directory.", parent=self.root); return
            self.start_task(self.split_button, backend.run_split_task, (input_path, output_dir, self.split_mode.get(), self.split_value.get(), self.progress_status, self.file_progress_bar, self.overall_progress_bar, self.on_task_complete))
        except Exception as e:
            messagebox.showerror("GUI Error", f"An unexpected error occurred: {e}", parent=self.root)
            self.on_task_complete("Ready")

    def process_rotate(self):
        try:
            input_path = self.rotate_input.get()
            output_path = self.rotate_output.get()
            if not input_path or not output_path: messagebox.showwarning("Input Missing", "Please select an input and output PDF.", parent=self.root); return
            angle = ROTATION_MAP.get(self.rotate_angle.get(), 0)
            self.start_task(self.rotate_button, backend.run_rotate_task, (input_path, output_path, angle, self.progress_status, self.file_progress_bar, self.overall_progress_bar, self.on_task_complete))
        except Exception as e:
            messagebox.showerror("GUI Error", f"An unexpected error occurred: {e}", parent=self.root)
            self.on_task_complete("Ready")

    def process_delete_pages(self):
        try:
            input_path = self.delete_pages_input.get()
            output_path = self.delete_pages_output.get()
            page_range = self.delete_pages_range.get()
            if not input_path or not output_path or not page_range: messagebox.showwarning("Input Missing", "Please provide an input PDF, output PDF, and pages to delete.", parent=self.root); return
            self.start_task(self.delete_button, backend.run_delete_pages_task, (input_path, output_path, page_range, self.progress_status, self.file_progress_bar, self.overall_progress_bar, self.on_task_complete))
        except Exception as e:
            messagebox.showerror("GUI Error", f"An unexpected error occurred: {e}", parent=self.root)
            self.on_task_complete("Ready")

    def process_convert(self):
        try:
            input_path = self.pdf_to_img_input.get()
            output_dir = self.pdf_to_img_output_dir.get()
            if not input_path or not output_dir: messagebox.showwarning("Input Missing", "Please select an input PDF and output directory.", parent=self.root); return
            options = {'format': self.pdf_to_img_format.get(), 'dpi': self.pdf_to_img_dpi.get()}
            params = {'gs_path': self.gs_path, 'input_path': input_path, 'output_path': output_dir, 'operation': 'PDF to Images', 'options': options, 'overwrite': False}
            self.start_task(self.convert_button, backend.run_conversion_task, (params, False, self.progress_status, self.file_progress_bar, self.overall_progress_bar, self.on_task_complete))
        except Exception as e:
            messagebox.showerror("GUI Error", f"An unexpected error occurred: {e}", parent=self.root)
            self.on_task_complete("Ready")

    def process_pdfa(self):
        try:
            input_path = self.pdfa_input.get()
            output_path = self.pdfa_output.get()
            if not input_path or not output_path: messagebox.showwarning("Input Missing", "Please select an input and output PDF.", parent=self.root); return
            params = {'gs_path': self.gs_path, 'input_path': input_path, 'output_path': output_path, 'operation': 'PDF/A', 'options': {}, 'overwrite': False}
            self.start_task(self.pdfa_button, backend.run_conversion_task, (params, False, self.progress_status, self.file_progress_bar, self.overall_progress_bar, self.on_task_complete))
        except Exception as e:
            messagebox.showerror("GUI Error", f"An unexpected error occurred: {e}", parent=self.root)
            self.on_task_complete("Ready")

    def process_remove_open_action(self):
        try:
            input_path = self.utility_pdf_in.get()
            if not input_path: messagebox.showwarning("Input Missing", "Please select an input PDF.", parent=self.root); return
            self.start_task(self.utility_button, backend.run_remove_open_action_task, (input_path, self.cpdf_path, self.progress_status, self.file_progress_bar, self.overall_progress_bar, self.on_task_complete))
        except Exception as e:
            messagebox.showerror("GUI Error", f"An unexpected error occurred: {e}", parent=self.root)
            self.on_task_complete("Ready")

    def load_metadata(self):
        pdf_path = self.meta_pdf_path.get()
        if not pdf_path: messagebox.showwarning("Input Missing", "Please select a PDF file.", parent=self.root); return
        try:
            info = backend.run_metadata_task('load', pdf_path, self.cpdf_path)
            self.meta_title.set(info.get('title', ''))
            self.meta_author.set(info.get('author', ''))
            self.meta_subject.set(info.get('subject', ''))
            self.meta_keywords.set(info.get('keywords', ''))
            self.status.set("Metadata loaded successfully.")
        except Exception as e: messagebox.showerror("Error", f"Failed to load metadata: {e}", parent=self.root)

    def save_metadata(self):
        pdf_path = self.meta_pdf_path.get()
        if not pdf_path: messagebox.showwarning("Input Missing", "Please select a PDF file.", parent=self.root); return
        metadata = { 'title': self.meta_title.get(), 'author': self.meta_author.get(), 'subject': self.meta_subject.get(), 'keywords': self.meta_keywords.get() }
        try:
            backend.run_metadata_task('save', pdf_path, self.cpdf_path, metadata)
            self.status.set("Metadata saved successfully.")
        except Exception as e: messagebox.showerror("Error", f"Failed to save metadata: {e}", parent=self.root)

    def process_stamp(self):
        try:
            input_path = self.stamp_pdf_in.get()
            output_path = self.stamp_pdf_out.get()
            if not input_path or not output_path: messagebox.showwarning("Input Missing", "Please select an input and output PDF.", parent=self.root); return
            if self.stamp_mode.get() == "Image" and not self.stamp_image_in.get():
                messagebox.showwarning("Input Missing", "Please select an image file for stamping.", parent=self.root); return

            stamp_opts = { 'pos': self.stamp_pos.get(), 'opacity': self.stamp_opacity.get(), 'on_top': self.stamp_on_top.get() }
            mode_opts = { 'image_path': self.stamp_image_in.get(), 'width': self.stamp_image_width.get(), 'height': self.stamp_image_height.get(), 'text': self.stamp_text_in.get(), 'font': self.stamp_font.get(), 'size': self.stamp_font_size.get(), 'color': self.stamp_font_color_var.get(), 'bates_start': self.stamp_bates_start.get() if self.stamp_dynamic_bates.get() else None }

            if self.stamp_dynamic_filename.get() and input_path: mode_opts['text'] += f" {Path(input_path).name}"
            if self.stamp_dynamic_datetime.get():
                mode_opts['text'] += f" {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"

            self.start_task(self.stamp_button, backend.run_stamp_task, (input_path, output_path, stamp_opts, self.cpdf_path, self.progress_status, self.file_progress_bar, self.overall_progress_bar, self.on_task_complete, self.stamp_mode.get(), mode_opts))
        except Exception as e:
            messagebox.showerror("GUI Error", f"An unexpected error occurred: {e}", parent=self.root)
            self.on_task_complete("Ready")

if __name__ == "__main__":
    root = tk.Tk()
    app = GhostscriptGUI(root)
    root.mainloop()