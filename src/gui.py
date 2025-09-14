# gui.py
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, colorchooser
from pathlib import Path
import logging
import threading
import sys
import json
import queue
import webbrowser
import base64
from io import BytesIO
from dataclasses import dataclass, field
from PIL import Image, ImageTk, ImageOps
from datetime import datetime
import re

import backend
import styles
from constants import (APP_VERSION, ROTATION_MAP, PDF_FONTS, SPLIT_MODES, SPLIT_SINGLE,
                       SPLIT_EVERY_N, SPLIT_CUSTOM,
                       STAMP_IMAGE, STAMP_TEXT, STAMP_POSITIONS, POS_CENTER, IMAGE_FORMATS, META_LOAD, META_SAVE,
                       PAGE_NUMBER_POSITIONS)
from ui_components import (ScrolledFrame, FileSelector, Tooltip, ModernToggle,
                           CompressionGauge, DropZone, PositionSelector, CustomSlider)
from tooltips import TOOLTIP_TEXT

IS_WINDOWS = sys.platform == "win32"
if IS_WINDOWS:
    try:
        import windnd
    except ImportError:
        logging.warning("windnd library not found. Drag and drop will be disabled.")
        IS_WINDOWS = False

COFFEE_ICON_B64 = "iVBORw0KGgoAAAANSUhEUgAAADAAAAAwCAYAAABXAvmHAAAACXBIWXMAAA7EAAAOxAGVKw4bAAAE6ElEQVRogdWZX2hcRRTGf67LstQYQgghlBBCjHkIWm0oQfNQi4oU8UGLiNQipcQHixQtIr74ICKCoYhIn6QPElsJUYtIUYs+VFGordSoaatpTaRqtCk2prba/On68J2be/fm7mZ2906jH1zu3NmZb87MnHPmnFnwh3qgxyM/ABmP3M8B5zzye0UDMInfBQKPA9wLzAJXPPEvwtcEbvfIXQRfgzQCLUCdJ/5FpD2BXntPA1ngvpT5vWO3vTcDBWCEq6RKaaAZGLNyHXAeTaJ/xSSqEOuBBaDNvp9HE5gCmlZKqEpwExJ40L7rgF+sbnepTv8lZIAjSOD7ra7fvueAzhWSqyKsQ8L+jlxpFjiNJjHgY8BrU+b7FYURdyN7+Bip0l3IyF9LeTwvaAL+RisPsAbtQAFFqP8LfABcRiqUs3IBD3bg65A5DswA8yiom7b62bQHyqbA0QR0If//JfAj8CfwRWycS8hGQHbxVwpjV4Um5B6HCf188NxjbXYRxkGrkGf6KNJ/DtnIILCFq2QbLcAeQn1eAEaBN4BngYcjgmwnVM8ea7/FvruBITSBBfvtAnKz3ibSSbjah9EOtJRpn4+Ut6MYKZfQrg3YAZww7hN4Cjv22gDPVNF3mFC1SiGDwo0C8EIl5K5eKHB/uQr6gIz3K+CggxzBDrVXwO+MM4SGegypRatDvzzlJ9yFdnUswn8AqdYGlCAlqd4irnEQAuQWAd4EtiLPAnKZR4GTwDhyk9PIRc5bmywyzkZgNXADMuQewkWYBt5GtnWF4kmfBR4H3nWUNREXUFwPinUeRS4w6kkqeeaQB3sdeBAtSGvk9xHk3Q4Y/xwKFJfAdQembJDrEn6rRzbShjxTI3A94SE5awvwB/AbMAGcQgdbFB1oQT4F7ojUP4GCwH3AI47yLsEoWomGagkc0ItWfzhW32D1o0mdXD3KKWvbVa10Duiw98+x+rKXY64T+NrevWVb1Ya19h6J1ffZ+2Qt5HeibdxfC8kyOGxjxHf5E6vfXAt5HrnSi/i5bVuNbGw8Vv80YYhR9jxwwaCRba2VKAE7jXuXfdcBr1jdRVJS3fWEJ3GaiVA08X8RxURT9n0e5depIdDTh1Lk3EbyQbcXt3ClIgTGfIZ0zoQmdAWzgAQeQHlDcwrcS9ABPImMuQC8Q22qlAHeN65JtBOpC54BNgGfURzzBOVXqW4S0fg/+lxGrrqvdFd3tCPBA/LTwEtGvgZtfQFtfSWudRVhgjQJ3ApsRBMKOAvAW9SQYrYSppDH0H9e8ZXuBn4inNymhDZRZNCdaRD7jxtHFDnkpoP84wjF6akzAp+/h/JXL83AexTv0gAKkfuA26w8QHHSsp/yuW8D4e7vrGYC31vnRsf2G4FDLNXp+HPI2rpgg/UZKteo1OoeRzFJP/Cyw2Af2tOODp6bCT3KWeBbdNE74cAVyPWYlRPD6OXQTeguh/AbRsfRB3xuY49Rw3mzjlBvF9DN2jbK3wdVi06k68EfJAVkA8uexMullHl0A/FUjOwHdA86YuUJlNDPUPoCN0+Y2LejXb4FBWpR7m+Q0e/D4Z9+15w4i4zqAaTjnSS7zFngH5TvBrcSOeT78yTb3LwJfRCd7kcdZQLcJxBHM7oW6QZuRCvaglxjvQkcTeovoauTc2inJpCn+w5lezNVysG/s35Qp+p2ynIAAAAASUVORK5CYII="

@dataclass
class GeneralSettings:
    dark_mode_enabled: tk.BooleanVar = field(default_factory=lambda: tk.BooleanVar(value=True))
    logging_enabled: tk.BooleanVar = field(default_factory=lambda: tk.BooleanVar(value=False))
    window_geometry: tk.StringVar = field(default_factory=tk.StringVar)

@dataclass
class OutputSettings:
    use_default_folder: tk.BooleanVar = field(default_factory=lambda: tk.BooleanVar(value=False))
    default_folder: tk.StringVar = field(default_factory=tk.StringVar)
    prefix: tk.StringVar = field(default_factory=tk.StringVar)
    suffix: tk.StringVar = field(default_factory=lambda: tk.StringVar(value="_compressed"))
    add_date: tk.BooleanVar = field(default_factory=tk.BooleanVar)
    add_time: tk.BooleanVar = field(default_factory=tk.BooleanVar)

@dataclass
class CompressSettings:
    input_path: tk.StringVar = field(default_factory=tk.StringVar)
    output_path: tk.StringVar = field(default_factory=tk.StringVar)
    compress_mode: tk.StringVar = field(default_factory=lambda: tk.StringVar(value="Compression"))
    dpi: tk.IntVar = field(default_factory=lambda: tk.IntVar(value=72))
    use_bicubic: tk.BooleanVar = field(default_factory=tk.BooleanVar)
    downsample_threshold_enabled: tk.BooleanVar = field(default_factory=lambda: tk.BooleanVar(value=True))
    quantize_colors: tk.BooleanVar = field(default_factory=tk.BooleanVar)
    quantize_level: tk.IntVar = field(default_factory=lambda: tk.IntVar(value=4))
    convert_to_grayscale: tk.BooleanVar = field(default_factory=tk.BooleanVar)
    convert_to_cmyk: tk.BooleanVar = field(default_factory=tk.BooleanVar)
    darken_text: tk.BooleanVar = field(default_factory=tk.BooleanVar)
    strip_metadata: tk.BooleanVar = field(default_factory=tk.BooleanVar)
    remove_interactive: tk.BooleanVar = field(default_factory=tk.BooleanVar)
    remove_open_action: tk.BooleanVar = field(default_factory=tk.BooleanVar)
    fast_web_view: tk.BooleanVar = field(default_factory=tk.BooleanVar)
    only_if_smaller: tk.BooleanVar = field(default_factory=tk.BooleanVar)
    fast_mode: tk.BooleanVar = field(default_factory=tk.BooleanVar)
    true_lossless: tk.BooleanVar = field(default_factory=lambda: tk.BooleanVar(value=False))

@dataclass
class MergeSettings:
    files: list = field(default_factory=list)
    output_path: tk.StringVar = field(default_factory=tk.StringVar)

@dataclass
class SplitSettings:
    input_path: tk.StringVar = field(default_factory=tk.StringVar)
    output_dir: tk.StringVar = field(default_factory=tk.StringVar)
    mode: tk.StringVar = field(default_factory=lambda: tk.StringVar(value=SPLIT_SINGLE))
    value: tk.StringVar = field(default_factory=lambda: tk.StringVar(value="2"))

@dataclass
class RotateSettings:
    input_path: tk.StringVar = field(default_factory=tk.StringVar)
    output_path: tk.StringVar = field(default_factory=tk.StringVar)
    angle: tk.StringVar = field(default_factory=lambda: tk.StringVar(value=list(ROTATION_MAP.keys())[1]))

@dataclass
class DeleteSettings:
    input_path: tk.StringVar = field(default_factory=tk.StringVar)
    output_path: tk.StringVar = field(default_factory=tk.StringVar)
    page_range: tk.StringVar = field(default_factory=lambda: tk.StringVar(value="1, 3-5"))

@dataclass
class StampSettings:
    input_path: tk.StringVar = field(default_factory=tk.StringVar)
    output_path: tk.StringVar = field(default_factory=tk.StringVar)
    mode: tk.StringVar = field(default_factory=lambda: tk.StringVar(value=STAMP_IMAGE))
    image_path: tk.StringVar = field(default_factory=tk.StringVar)
    image_scale: tk.DoubleVar = field(default_factory=lambda: tk.DoubleVar(value=100.0))
    text: tk.StringVar = field(default_factory=lambda: tk.StringVar(value="CONFIDENTIAL"))
    font: tk.StringVar = field(default_factory=lambda: tk.StringVar(value=PDF_FONTS[0]))
    font_size: tk.StringVar = field(default_factory=lambda: tk.StringVar(value="48"))
    font_color: tk.StringVar = field(default_factory=lambda: tk.StringVar(value="#ff0000"))
    pos: tk.StringVar = field(default_factory=lambda: tk.StringVar(value=POS_CENTER))
    opacity: tk.DoubleVar = field(default_factory=lambda: tk.DoubleVar(value=0.5))
    on_top: tk.BooleanVar = field(default_factory=lambda: tk.BooleanVar(value=True))
    bates_enabled: tk.BooleanVar = field(default_factory=tk.BooleanVar)
    bates_start: tk.StringVar = field(default_factory=lambda: tk.StringVar(value="1"))

@dataclass
class PageNumberSettings:
    input_path: tk.StringVar = field(default_factory=tk.StringVar)
    output_path: tk.StringVar = field(default_factory=tk.StringVar)
    mode: tk.StringVar = field(default_factory=lambda: tk.StringVar(value="Page X of Y"))
    custom_text: tk.StringVar = field(default_factory=lambda: tk.StringVar(value="%File"))
    page_range: tk.StringVar = field(default_factory=tk.StringVar)
    pos: tk.StringVar = field(default_factory=lambda: tk.StringVar(value="Bottom Center"))
    font: tk.StringVar = field(default_factory=lambda: tk.StringVar(value=PDF_FONTS[0]))
    font_size: tk.StringVar = field(default_factory=lambda: tk.StringVar(value="12"))
    font_color: tk.StringVar = field(default_factory=lambda: tk.StringVar(value="#000000"))

@dataclass
class MetadataSettings:
    input_path: tk.StringVar = field(default_factory=tk.StringVar)
    title: tk.StringVar = field(default_factory=tk.StringVar)
    author: tk.StringVar = field(default_factory=tk.StringVar)
    subject: tk.StringVar = field(default_factory=tk.StringVar)
    keywords: tk.StringVar = field(default_factory=tk.StringVar)

@dataclass
class ConvertSettings:
    input_path: tk.StringVar = field(default_factory=tk.StringVar)
    output_dir: tk.StringVar = field(default_factory=tk.StringVar)
    format: tk.StringVar = field(default_factory=lambda: tk.StringVar(value=IMAGE_FORMATS[0]))
    dpi: tk.StringVar = field(default_factory=lambda: tk.StringVar(value="300"))
    dpi_slider: tk.IntVar = field(default_factory=lambda: tk.IntVar(value=300))

@dataclass
class RepairSettings:
    input_path: tk.StringVar = field(default_factory=tk.StringVar)
    output_path: tk.StringVar = field(default_factory=tk.StringVar)

@dataclass
class TocSettings:
    input_path: tk.StringVar = field(default_factory=tk.StringVar)
    output_path: tk.StringVar = field(default_factory=tk.StringVar)
    title: tk.StringVar = field(default_factory=lambda: tk.StringVar(value="Table of Contents"))
    font: tk.StringVar = field(default_factory=lambda: tk.StringVar(value=PDF_FONTS[0]))
    font_size: tk.StringVar = field(default_factory=lambda: tk.StringVar(value="12"))
    dot_leaders: tk.BooleanVar = field(default_factory=lambda: tk.BooleanVar(value=True))
    no_bookmark: tk.BooleanVar = field(default_factory=lambda: tk.BooleanVar(value=False))

@dataclass
class PasswordSettings:
    input_path: tk.StringVar = field(default_factory=tk.StringVar)
    output_path: tk.StringVar = field(default_factory=tk.StringVar)
    user_password: tk.StringVar = field(default_factory=tk.StringVar)
    owner_password: tk.StringVar = field(default_factory=tk.StringVar)
    decrypt_password: tk.StringVar = field(default_factory=tk.StringVar)
    show_passwords: tk.BooleanVar = field(default_factory=tk.BooleanVar)
    allow_printing: tk.BooleanVar = field(default_factory=lambda: tk.BooleanVar(value=True))
    allow_modification: tk.BooleanVar = field(default_factory=lambda: tk.BooleanVar(value=False))
    allow_copy_and_extract: tk.BooleanVar = field(default_factory=lambda: tk.BooleanVar(value=True))
    allow_annotations_and_forms: tk.BooleanVar = field(default_factory=lambda: tk.BooleanVar(value=True))


class GhostscriptGUI:
    def __init__(self, root):
        self.root = root
        self.root.title(f"MinimalPDF Compress v{APP_VERSION}")
        self.root.minsize(860, 640)
        self.settings_file = Path("settings.json")
        self.is_folder = tk.BooleanVar(value=False)
        self.status = tk.StringVar(value="Ready")
        self.compress_progress_status = tk.StringVar()
        self.progress_var = tk.DoubleVar()
        self.overall_progress_var = tk.DoubleVar()
        self.compression_ratio_var = tk.DoubleVar(value=0.0)

        self.general_settings = GeneralSettings()
        self.output_settings = OutputSettings()
        self.compress_settings = CompressSettings()
        self.merge_settings = MergeSettings()
        self.split_settings = SplitSettings()
        self.rotate_settings = RotateSettings()
        self.delete_settings = DeleteSettings()
        self.stamp_settings = StampSettings()
        self.page_number_settings = PageNumberSettings()
        self.meta_settings = MetadataSettings()
        self.convert_settings = ConvertSettings()
        self.repair_settings = RepairSettings()
        self.toc_settings = TocSettings()
        self.password_settings = PasswordSettings()

        self.palette = {}
        self.toggles = []
        self.gs_path, self.cpdf_path, self.pngquant_path, self.jpegoptim_path, self.ect_path, self.optipng_path = None, None, None, None, None, None
        self.active_process_button = None
        self.tab_frames = {}
        self.drop_zones = {}
        self.progress_queue = queue.Queue()
        self._preview_image_cache = {}
        self.tab_statuses = {}
        self.active_status_var = None
        self.log_handler = None
        self.coffee_img_label = None
        self.coffee_icon_light = None
        self.coffee_icon_dark = None
        self.merge_tree = None
        self._preview_job = None

        self.vcmd_int = (self.root.register(self._validate_integer), '%P')
        self.vcmd_pagerange = (self.root.register(self._validate_page_range), '%P')
        self.icon_path = backend.resource_path("pdf.ico")
        try:
            if self.icon_path.exists(): self.root.iconbitmap(self.icon_path)
        except Exception as e: logging.warning(f"Failed to set icon: {e}")

        self.main_frame = ttk.Frame(self.root, padding="10")
        self.main_frame.grid(row=0, column=0, sticky="nsew")
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)

        self.load_settings()

        if self.general_settings.window_geometry.get():
            try:
                self.root.geometry(self.general_settings.window_geometry.get())
            except tk.TclError:
                self.root.state('zoomed')
        else:
            try:
                self.root.state('zoomed')
            except tk.TclError:
                pass

        self.configure_logging()
        self.palette = styles.apply_theme(self.root, 'dark' if self.general_settings.dark_mode_enabled.get() else 'light')
        self.build_gui()
        self.toggle_password_visibility()
        self._update_widget_colors()
        self.setup_drag_and_drop()
        self.check_tools()
        self.setup_traces()
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        self.root.after(100, self.check_progress_queue)

    def _hex_to_cpdf_color(self, hex_color):
        try:
            hex_color = hex_color.lstrip('#')
            r, g, b = tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4))
            return f"{r/255:.3f} {g/255:.3f} {b/255:.3f}"
        except Exception:
            return "0 0 0"

    def _validate_integer(self, P):
        return P.isdigit() or P == ""

    def _validate_page_range(self, P):
        return all(c in "0123456789,-endEND " for c in P)

    def _clamp_dpi(self, event=None):
        try:
            value = self.compress_settings.dpi.get()
            if value > 600:
                self.compress_settings.dpi.set(600)
            elif value < 0:
                self.compress_settings.dpi.set(0)
        except (ValueError, tk.TclError):
            self.compress_settings.dpi.set(72)

    def _clamp_quantize_level(self, event=None):
        try:
            value = self.compress_settings.quantize_level.get()
            if value > 8: self.compress_settings.quantize_level.set(8)
            elif value < 2: self.compress_settings.quantize_level.set(2)
        except (ValueError, tk.TclError):
            self.compress_settings.quantize_level.set(4)

    def check_progress_queue(self):
        try:
            while not self.progress_queue.empty():
                message = self.progress_queue.get_nowait()
                msg_type, value = message
                if msg_type == 'overall':
                    self.overall_progress_var.set(value)
                elif msg_type == 'status':
                    max_chars = 55
                    status_text = value[:max_chars] + "..." if len(value) > max_chars else value
                    self.compress_progress_status.set(status_text)
                    if self.active_status_var:
                        self.active_status_var.set(status_text)
                elif msg_type == 'complete':
                    self.on_task_complete(value)
        except queue.Empty:
            pass
        finally:
            self.root.after(100, self.check_progress_queue)

    def on_closing(self):
        if self.root.state() != 'iconic':
            self.general_settings.window_geometry.set(self.root.winfo_geometry())
        self.save_settings()
        self.root.destroy()

    def _get_tk_vars_as_dict(self, obj):
        if isinstance(obj, (tk.Variable)): return obj.get()
        if isinstance(obj, list): return [self._get_tk_vars_as_dict(i) for i in obj]
        if hasattr(obj, '__dict__'):
            return {key: self._get_tk_vars_as_dict(value) for key, value in obj.__dict__ .items() if not key.startswith('_')}
        return obj

    def save_settings(self):
        settings = {
            'general': self._get_tk_vars_as_dict(self.general_settings),
            'output': self._get_tk_vars_as_dict(self.output_settings),
            'compress': self._get_tk_vars_as_dict(self.compress_settings),
            'merge': self._get_tk_vars_as_dict(self.merge_settings),
            'split': self._get_tk_vars_as_dict(self.split_settings),
            'rotate': self._get_tk_vars_as_dict(self.rotate_settings),
            'delete': self._get_tk_vars_as_dict(self.delete_settings),
            'stamp': self._get_tk_vars_as_dict(self.stamp_settings),
            'page_number': self._get_tk_vars_as_dict(self.page_number_settings),
            'metadata': self._get_tk_vars_as_dict(self.meta_settings),
            'convert': self._get_tk_vars_as_dict(self.convert_settings),
            'repair': self._get_tk_vars_as_dict(self.repair_settings),
            'toc': self._get_tk_vars_as_dict(self.toc_settings),
            'password': self._get_tk_vars_as_dict(self.password_settings),
        }
        try:
            with open(self.settings_file, 'w') as f: json.dump(settings, f, indent=4)
        except Exception as e: logging.warning(f"Failed to save settings: {e}")

    def _set_tk_vars_from_dict(self, obj, data):
        for key, value in data.items():
            if hasattr(obj, key):
                attr = getattr(obj, key)
                if isinstance(attr, tk.Variable):
                    try: attr.set(value)
                    except: pass
                elif isinstance(attr, list) and isinstance(value, list):
                    attr.clear(); attr.extend(value)

    def load_settings(self):
        if not self.settings_file.exists(): return
        try:
            with open(self.settings_file, 'r') as f: s = json.load(f)
            self._set_tk_vars_from_dict(self.general_settings, s.get('general', {}))
            self._set_tk_vars_from_dict(self.output_settings, s.get('output', {}))
            self._set_tk_vars_from_dict(self.compress_settings, s.get('compress', {}))
            self._set_tk_vars_from_dict(self.merge_settings, s.get('merge', {}))
            self._set_tk_vars_from_dict(self.split_settings, s.get('split', {}))
            self._set_tk_vars_from_dict(self.rotate_settings, s.get('rotate', {}))
            self._set_tk_vars_from_dict(self.delete_settings, s.get('delete', {}))
            self._set_tk_vars_from_dict(self.stamp_settings, s.get('stamp', {}))
            self._set_tk_vars_from_dict(self.page_number_settings, s.get('page_number', {}))
            self._set_tk_vars_from_dict(self.meta_settings, s.get('metadata', {}))
            self._set_tk_vars_from_dict(self.convert_settings, s.get('convert', {}))
            self._set_tk_vars_from_dict(self.repair_settings, s.get('repair', {}))
            self._set_tk_vars_from_dict(self.toc_settings, s.get('toc', {}))
            self._set_tk_vars_from_dict(self.password_settings, s.get('password', {}))
            self.update_merge_view()
        except Exception as e: logging.warning(f"Failed to load settings: {e}")

    def configure_logging(self):
        root_logger = logging.getLogger()
        if self.log_handler:
            root_logger.removeHandler(self.log_handler)
            self.log_handler.close()
            self.log_handler = None

        if self.general_settings.logging_enabled.get():
            log_file = Path("app.log")
            self.log_handler = logging.FileHandler(log_file, 'w', 'utf-8')
            formatter = logging.Formatter('%(asctime)s - %(levelname)s - [%(module)s.%(funcName)s:%(lineno)d] - %(message)s')
            self.log_handler.setFormatter(formatter)
            root_logger.addHandler(self.log_handler)
            root_logger.setLevel(logging.INFO)
            logging.info("File logging enabled.")
        else:
            root_logger.setLevel(logging.CRITICAL + 1)

    def _on_utility_tab_changed(self, event):
        for var in self.tab_statuses.values():
            var.set("")
        try:
            notebook = event.widget
            selected_tab_text = notebook.tab(notebook.select(), "text")

            if selected_tab_text == "Rotate":
                if self.rotate_settings.input_path.get():
                    self._trigger_rotate_preview()
            elif selected_tab_text == "Delete Pages":
                if self.delete_settings.input_path.get():
                    self._update_preview(self.delete_settings.input_path.get(), self.delete_preview_canvas)
            elif selected_tab_text == "Stamp/Watermark":
                if self.stamp_settings.input_path.get():
                    self._trigger_stamp_preview()
        except tk.TclError:
            pass
        except Exception as e:
            logging.warning(f"Error handling tab change: {e}")

    def build_gui(self):
        self.main_frame.rowconfigure(0, weight=1)
        self.main_frame.columnconfigure(0, weight=1)

        main_notebook = ttk.Notebook(self.main_frame)
        main_notebook.grid(row=0, column=0, sticky="nsew")

        compress_tab = ScrolledFrame(main_notebook)
        main_notebook.add(compress_tab, text="Compress")
        self._build_compress_tab(compress_tab.scrollable_frame)

        utilities_notebook = ttk.Notebook(main_notebook)
        main_notebook.add(utilities_notebook, text="Utilities")
        utilities_notebook.bind("<<NotebookTabChanged>>", self._on_utility_tab_changed)

        utility_tab_configs = [
            ("merge", "Merge"), ("split", "Split/Extract"), ("rotate", "Rotate"),
            ("delete", "Delete Pages"), ("password", "Password"), ("stamp", "Stamp/Watermark"), 
            ("page_number", "Header/Footer"), ("toc", "Table of Contents"), 
            ("metadata", "Metadata"), ("convert", "PDF to Image"), ("repair", "PDF Repair")
        ]

        for name, text in utility_tab_configs:
            frame = ScrolledFrame(utilities_notebook)
            self.tab_frames[name] = frame
            utilities_notebook.add(frame, text=text)
            self.tab_statuses[name] = tk.StringVar()

        self._build_merge_tab(self.tab_frames["merge"].scrollable_frame)
        self._build_split_tab(self.tab_frames["split"].scrollable_frame)
        self._build_rotate_tab(self.tab_frames["rotate"].scrollable_frame)
        self._build_delete_tab(self.tab_frames["delete"].scrollable_frame)
        self._build_password_tab(self.tab_frames["password"].scrollable_frame)
        self._build_stamp_tab(self.tab_frames["stamp"].scrollable_frame)
        self._build_page_number_tab(self.tab_frames["page_number"].scrollable_frame)
        self._build_toc_tab(self.tab_frames["toc"].scrollable_frame)
        self._build_metadata_tab(self.tab_frames["metadata"].scrollable_frame)
        self._build_convert_tab(self.tab_frames["convert"].scrollable_frame)
        self._build_repair_tab(self.tab_frames["repair"].scrollable_frame)

        settings_tab = ScrolledFrame(main_notebook)
        main_notebook.add(settings_tab, text="Settings")
        self.tab_frames["settings"] = settings_tab
        self._build_settings_tab(settings_tab.scrollable_frame)

        status_bar = ttk.Frame(self.main_frame, style="Card.TFrame", padding=(5, 2))
        status_bar.grid(row=1, column=0, sticky="ew", pady=(10,0))
        ttk.Label(status_bar, textvariable=self.status, anchor="w", style="Card.TLabel").pack(fill="x")

    def setup_traces(self):
        s = self.stamp_settings
        self.compress_settings.input_path.trace_add("write", self.update_is_folder)
        self.compress_settings.input_path.trace_add("write", lambda *a: self._update_output_path(self.compress_settings.input_path, self.compress_settings.output_path, is_folder=self.is_folder.get()))
        self.compress_settings.compress_mode.trace_add("write", self._update_compress_options)
        self.split_settings.input_path.trace_add("write", lambda *a: self._update_output_path(self.split_settings.input_path, self.split_settings.output_dir, is_dir=True))

        self.rotate_settings.input_path.trace_add("write", lambda *a: self._update_output_path(self.rotate_settings.input_path, self.rotate_settings.output_path))
        self.rotate_settings.input_path.trace_add("write", self._trigger_rotate_preview)
        self.rotate_settings.angle.trace_add("write", self._trigger_rotate_preview)

        self.delete_settings.input_path.trace_add("write", lambda *a: self._update_output_path(self.delete_settings.input_path, self.delete_settings.output_path))
        self.delete_settings.input_path.trace_add("write", lambda *a: self._update_preview(self.delete_settings.input_path.get(), self.delete_preview_canvas))

        self.password_settings.input_path.trace_add("write", lambda *a: self._update_output_path(self.password_settings.input_path, self.password_settings.output_path))
        
        stamp_vars_to_trace = [s.input_path, s.output_path, s.mode, s.image_path, s.image_scale, s.text, s.font, s.font_size, s.font_color, s.pos, s.opacity, s.on_top, s.bates_enabled, s.bates_start]
        for var in stamp_vars_to_trace:
            var.trace_add("write", self._trigger_stamp_preview)
        s.input_path.trace_add("write", lambda *a: self._update_output_path(s.input_path, s.output_path))

        self.page_number_settings.input_path.trace_add("write", lambda *a: self._update_output_path(self.page_number_settings.input_path, self.page_number_settings.output_path))

        self.toc_settings.input_path.trace_add("write", lambda *a: self._update_output_path(self.toc_settings.input_path, self.toc_settings.output_path))

        self.convert_settings.input_path.trace_add("write", lambda *a: self._update_output_path(self.convert_settings.input_path, self.convert_settings.output_dir, is_dir=True))
        self.repair_settings.input_path.trace_add("write", lambda *a: self._update_output_path(self.repair_settings.input_path, self.repair_settings.output_path))

        self.split_settings.mode.trace_add("write", self._update_split_validation)
        self.convert_settings.dpi_slider.trace_add("write", lambda *a: self.convert_settings.dpi.set(self.convert_settings.dpi_slider.get()))
        self.meta_settings.input_path.trace_add("write", self.load_metadata)

        os = self.output_settings
        vars_to_trace = [
            os.use_default_folder, os.default_folder,
            os.prefix, os.suffix, os.add_date, os.add_time
        ]
        for var in vars_to_trace:
            var.trace_add("write", self._re_evaluate_all_paths)

    def _re_evaluate_all_paths(self, *args):
        path_vars = [
            (self.compress_settings.input_path, self.compress_settings.output_path, {'is_folder': self.is_folder.get()}),
            (self.split_settings.input_path, self.split_settings.output_dir, {'is_dir': True}),
            (self.rotate_settings.input_path, self.rotate_settings.output_path, {}),
            (self.delete_settings.input_path, self.delete_settings.output_path, {}),
            (self.password_settings.input_path, self.password_settings.output_path, {}),
            (self.stamp_settings.input_path, self.stamp_settings.output_path, {}),
            (self.page_number_settings.input_path, self.page_number_settings.output_path, {}),
            (self.toc_settings.input_path, self.toc_settings.output_path, {}),
            (self.convert_settings.input_path, self.convert_settings.output_dir, {'is_dir': True}),
            (self.repair_settings.input_path, self.repair_settings.output_path, {}),
        ]
        for in_var, out_var, kwargs in path_vars:
            if in_var.get():
                self._update_output_path(in_var, out_var, **kwargs)

    def browse_file(self, var): var.set(filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")]))
    def browse_files(self, file_list):
        is_first_addition = not file_list
        files = filedialog.askopenfilenames(filetypes=[("PDF files", "*.pdf")])
        if files:
            file_list.extend(files)
            self.update_merge_view()
            if is_first_addition:
                first_file = Path(files[0])
                output_path = first_file.parent / f"{first_file.stem}_merged.pdf"
                self.merge_settings.output_path.set(str(output_path))
    def browse_save_file(self, var): var.set(filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")]))
    def browse_dir(self, var): var.set(filedialog.askdirectory(mustexist=True))
    def browse_image(self, var): var.set(filedialog.askopenfilename(filetypes=[("Image files", "*.png *.jpg *.jpeg")]))

    def _create_drop_zone(self, parent, input_var, is_folder=False):
        cmd = (lambda: self.browse_dir(input_var)) if is_folder else (lambda: self.browse_file(input_var))
        dz = DropZone(parent, height=100, browse_file_cmd=cmd, browse_folder_cmd=lambda: self.browse_dir(input_var), palette=self.palette)
        dz.grid(row=0, column=0, sticky="nsew", pady=(0, 15))
        Tooltip(dz, TOOLTIP_TEXT.get("drop_zone"))
        return dz

    def _create_process_button(self, parent, text, command, row, columnspan=1, tooltip_key=None):
        btn_shadow = ttk.Frame(parent, style="AccentBG.TFrame")
        btn_shadow.grid(row=row, column=0, columnspan=columnspan, sticky="ew", pady=(15, 12), padx=3)
        btn = ttk.Button(btn_shadow, text=text, command=command, style="Accent.TButton")
        btn.pack(fill="x", ipady=10, pady=(0,3), padx=(0,3))
        if tooltip_key:
            Tooltip(btn, TOOLTIP_TEXT.get(tooltip_key))
        return btn

    def _update_color_conversion_checks(self, changed_var):
        cs = self.compress_settings
        if changed_var == 'cmyk' and cs.convert_to_cmyk.get():
            cs.convert_to_grayscale.set(False)
        elif changed_var == 'grayscale' and cs.convert_to_grayscale.get():
            cs.convert_to_cmyk.set(False)

    def _build_compress_tab(self, parent):
        cs = self.compress_settings
        parent.columnconfigure(0, weight=1)

        io_frame = ttk.Frame(parent, style="Card.TFrame", padding=20)
        io_frame.grid(row=0, column=0, sticky="nsew", pady=(0, 10))
        io_frame.columnconfigure(0, weight=1)
        self.drop_zones['compress'] = DropZone(io_frame, height=150, browse_file_cmd=lambda: self.browse_file(cs.input_path), browse_folder_cmd=lambda: self.browse_dir(cs.input_path), palette=self.palette)
        self.drop_zones['compress'].grid(row=0, column=0, sticky="nsew", pady=(0,15))
        Tooltip(self.drop_zones['compress'], TOOLTIP_TEXT.get("drop_zone"))
        FileSelector(io_frame, cs.input_path, cs.output_path, lambda: self.browse_output())

        control_panel = ttk.Frame(parent, style="Card.TFrame", padding=20)
        control_panel.grid(row=1, column=0, sticky="nsew", pady=10)
        control_panel.columnconfigure(0, weight=1, minsize=200)
        control_panel.columnconfigure(1, weight=2, minsize=300)
        control_panel.rowconfigure(0, weight=1)

        gauge_frame = ttk.Frame(control_panel, style="Card.TFrame")
        gauge_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 20))
        gauge_frame.rowconfigure(0, weight=1)
        self.gauge = CompressionGauge(gauge_frame, variable=self.compression_ratio_var, palette=self.palette)
        self.gauge.grid(row=0, column=0, sticky="nsew")

        options_frame = ttk.Frame(control_panel, style="Card.TFrame")
        options_frame.grid(row=0, column=1, sticky="nsew")
        options_frame.columnconfigure(0, weight=1)

        ttk.Label(options_frame, text="Mode", font=(styles.FONT_FAMILY, 14, "bold"), style="Card.TLabel").grid(row=0, column=0, sticky="w", pady=(0, 15))
        mode_frame = ttk.Frame(options_frame, style="TFrame")
        mode_frame.grid(row=1, column=0, sticky="ew", pady=5)

        rb_comp = ttk.Radiobutton(mode_frame, text="Compression", variable=cs.compress_mode, value="Compression", style="Segmented.TRadiobutton"); rb_comp.pack(side="left", fill="x", expand=True)
        Tooltip(rb_comp, TOOLTIP_TEXT.get("compress_mode_compression"))
        ttk.Separator(mode_frame, orient='vertical').pack(side="left", fill='y', padx=2)
        rb_lossless = ttk.Radiobutton(mode_frame, text="Lossless", variable=cs.compress_mode, value="Lossless", style="Segmented.TRadiobutton"); rb_lossless.pack(side="left", fill="x", expand=True)
        Tooltip(rb_lossless, TOOLTIP_TEXT.get("compress_mode_lossless"))
        ttk.Separator(mode_frame, orient='vertical').pack(side="left", fill='y', padx=2)
        rb_pdfa = ttk.Radiobutton(mode_frame, text="PDF/A", variable=cs.compress_mode, value="PDF/A", style="Segmented.TRadiobutton"); rb_pdfa.pack(side="left", fill="x", expand=True)
        Tooltip(rb_pdfa, TOOLTIP_TEXT.get("compress_mode_pdfa"))
        ttk.Separator(mode_frame, orient='vertical').pack(side="left", fill='y', padx=2)
        rb_rem_img = ttk.Radiobutton(mode_frame, text="Remove Images", variable=cs.compress_mode, value="Remove Images", style="Segmented.TRadiobutton"); rb_rem_img.pack(side="left", fill="x", expand=True)
        Tooltip(rb_rem_img, TOOLTIP_TEXT.get("compress_mode_remove_images"))

        dpi_frame = ttk.Frame(options_frame, style="TFrame")
        dpi_frame.grid(row=2, column=0, sticky="ew", pady=(20, 0))
        dpi_frame.columnconfigure(0, weight=1)

        dpi_label_frame = ttk.Frame(dpi_frame, style="TFrame")
        dpi_label_frame.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0,5))
        ttk.Label(dpi_label_frame, text="Image Quality", font=(styles.FONT_FAMILY, 14, "bold"), style="Card.TLabel").pack(side="left")
        self.dpi_label = ttk.Label(dpi_label_frame, text=f"({cs.dpi.get()} DPI)", style="Card.TLabel")
        self.dpi_label.pack(side="left", padx=5)

        def update_dpi_label(*args): self.dpi_label.config(text=f"({cs.dpi.get()} DPI)")
        cs.dpi.trace_add("write", update_dpi_label)

        self.dpi_slider = CustomSlider(dpi_frame, from_=50, to=150, variable=cs.dpi, palette=self.palette)
        self.dpi_slider.grid(row=1, column=0, sticky="ew", padx=(0, 10))

        self.dpi_entry = ttk.Entry(dpi_frame, textvariable=cs.dpi, width=5, validate='key', validatecommand=self.vcmd_int)
        self.dpi_entry.grid(row=1, column=1, sticky="w")
        self.dpi_entry.bind("<FocusOut>", self._clamp_dpi)
        Tooltip(dpi_frame, TOOLTIP_TEXT.get("compress_dpi_slider"))

        progress_frame = ttk.Frame(options_frame, style="TFrame")
        progress_frame.grid(row=3, column=0, sticky="ew", pady=(20,0))
        progress_frame.columnconfigure(0, weight=1)
        ttk.Label(progress_frame, textvariable=self.compress_progress_status).grid(row=0, column=0, sticky="w")
        self.main_progress_bar = ttk.Progressbar(progress_frame, mode='indeterminate')
        self.main_progress_bar.grid(row=1, column=0, sticky="ew", pady=2)
        self.main_overall_progress_bar = ttk.Progressbar(progress_frame, variable=self.overall_progress_var, maximum=100)
        self.main_overall_progress_bar.grid(row=2, column=0, sticky="ew", pady=2)

        self.compress_button = self._create_process_button(control_panel, "PROCESS PDF", self.process_compression, row=1, columnspan=2, tooltip_key="compress_process_btn")

        adv_frame = ttk.LabelFrame(parent, text="Advanced Settings", padding=10)
        adv_frame.grid(row=3, column=0, sticky="nsew", pady=5); adv_frame.columnconfigure(1, weight=1)
        f1 = ttk.Frame(adv_frame); f1.grid(row=0, column=0, sticky="nsew");
        f2 = ttk.Frame(adv_frame); f2.grid(row=0, column=1, sticky="nsew", padx=20);

        cb1 = ttk.Checkbutton(f1, text="Remove Interactive (Compression)", variable=cs.remove_interactive); cb1.pack(anchor="w"); Tooltip(cb1, TOOLTIP_TEXT.get("output_remove_interactive"))
        cb2 = ttk.Checkbutton(f1, text="Strip All Metadata", variable=cs.strip_metadata); cb2.pack(anchor="w"); Tooltip(cb2, TOOLTIP_TEXT.get("output_strip_metadata"))
        self.bicubic_check = ttk.Checkbutton(f1, text="Use Bicubic Sampling (Higher Quality)", variable=cs.use_bicubic)
        self.bicubic_check.pack(anchor="w"); Tooltip(self.bicubic_check, TOOLTIP_TEXT.get("output_bicubic"))
        self.downsample_threshold_check = ttk.Checkbutton(f1, text="Only Downsample Larger Images (Recommended)", variable=cs.downsample_threshold_enabled)
        self.downsample_threshold_check.pack(anchor="w"); Tooltip(self.downsample_threshold_check, TOOLTIP_TEXT.get("compress_downsample_threshold"))
        cb4 = ttk.Checkbutton(f1, text="Linearize for Fast Web View", variable=cs.fast_web_view); cb4.pack(anchor="w"); Tooltip(cb4, TOOLTIP_TEXT.get("output_fast_web"))
        cb5 = ttk.Checkbutton(f1, text="Darken Text", variable=cs.darken_text); cb5.pack(anchor="w"); Tooltip(cb5, TOOLTIP_TEXT.get("output_darken_text"))
        cb6 = ttk.Checkbutton(f1, text="Remove Open Action", variable=cs.remove_open_action); cb6.pack(anchor="w"); Tooltip(cb6, TOOLTIP_TEXT.get("output_remove_openaction"))

        self.grayscale_check = ttk.Checkbutton(f2, text="Convert to Grayscale", variable=cs.convert_to_grayscale, command=lambda: self._update_color_conversion_checks('grayscale')); self.grayscale_check.pack(anchor="w"); Tooltip(self.grayscale_check, TOOLTIP_TEXT.get("compress_grayscale_check"))
        self.cmyk_check = ttk.Checkbutton(f2, text="Convert to CMYK (for printing)", variable=cs.convert_to_cmyk, command=lambda: self._update_color_conversion_checks('cmyk')); self.cmyk_check.pack(anchor="w"); Tooltip(self.cmyk_check, TOOLTIP_TEXT.get("compress_cmyk_check"))

        quantize_frame = ttk.Frame(f2)
        quantize_frame.pack(anchor="w", fill="x", pady=(10,0))
        self.quantize_toggle = ModernToggle(quantize_frame, text="Quantize Colors (Posterize)", variable=cs.quantize_colors, palette=self.palette)
        self.quantize_toggle.pack(side="left")
        self.toggles.append(self.quantize_toggle)
        Tooltip(self.quantize_toggle, TOOLTIP_TEXT.get("compress_quantize"))

        self.quantize_level_entry = ttk.Entry(quantize_frame, textvariable=cs.quantize_level, width=3, validate='key', validatecommand=self.vcmd_int)
        self.quantize_level_entry.pack(side="left", padx=5)
        ttk.Label(quantize_frame, text="Levels (2-8)").pack(side="left")
        self.quantize_level_entry.bind("<FocusOut>", self._clamp_quantize_level)

        def toggle_quantize_entry_state(*args):
            state = "normal" if cs.quantize_colors.get() else "disabled"
            if hasattr(self, 'quantize_level_entry'):
                self.quantize_level_entry.config(state=state)
        cs.quantize_colors.trace_add("write", toggle_quantize_entry_state)
        toggle_quantize_entry_state()

        cb8 = ttk.Checkbutton(f2, text="Don't save if larger than original", variable=cs.only_if_smaller); cb8.pack(anchor="w", pady=(10,0)); Tooltip(cb8, TOOLTIP_TEXT.get("compress_only_if_smaller"))
        cb9 = ttk.Checkbutton(f2, text="Fast Mode (lower compression)", variable=cs.fast_mode); cb9.pack(anchor="w"); Tooltip(cb9, TOOLTIP_TEXT.get("compress_fast_mode"))
        self.true_lossless_check = ttk.Checkbutton(f2, text="True Lossless (preserves JPEGs)", variable=cs.true_lossless)
        self.true_lossless_check.pack(anchor="w", pady=(5,0))
        Tooltip(self.true_lossless_check, TOOLTIP_TEXT.get("compress_true_lossless"))

        self._update_compress_options()

    def _build_merge_tab(self, parent):
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(0, weight=1)
        s = self.merge_settings

        main_card = ttk.Frame(parent, style="Card.TFrame", padding=20)
        main_card.grid(row=0, column=0, sticky="nsew")
        main_card.columnconfigure(0, weight=1)
        main_card.rowconfigure(1, weight=1)

        ttk.Label(main_card, text="Merge Files", font=(styles.FONT_FAMILY, 14, "bold"), style="Card.TLabel").grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 15))

        tree_frame = ttk.Frame(main_card)
        tree_frame.grid(row=1, column=0, sticky="nsew", padx=(0, 10))
        tree_frame.rowconfigure(0, weight=1)
        tree_frame.columnconfigure(0, weight=1)

        self.merge_tree = ttk.Treeview(tree_frame, columns=('name', 'pages', 'size'), show='headings')
        self.merge_tree.grid(row=0, column=0, sticky='nsew')
        Tooltip(self.merge_tree, TOOLTIP_TEXT.get("merge_tree"))
        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.merge_tree.yview)
        vsb.grid(row=0, column=1, sticky='ns')
        self.merge_tree.configure(yscrollcommand=vsb.set)

        btn_frame = ttk.Frame(main_card, style="Card.TFrame")
        btn_frame.grid(row=1, column=1, sticky="ns")
        btn1 = ttk.Button(btn_frame, text="Add Files", style="Outline.TButton", command=lambda: self.browse_files(s.files)); btn1.pack(fill="x", pady=2); Tooltip(btn1, TOOLTIP_TEXT.get("merge_add_btn"))
        btn2 = ttk.Button(btn_frame, text="Remove", style="Outline.TButton", command=self.remove_merge_file); btn2.pack(fill="x", pady=2); Tooltip(btn2, TOOLTIP_TEXT.get("merge_remove_btn"))
        btn3 = ttk.Button(btn_frame, text="Move Up", style="Outline.TButton", command=self.move_merge_up); btn3.pack(fill="x", pady=(10, 2)); Tooltip(btn3, TOOLTIP_TEXT.get("merge_move_up_btn"))
        btn4 = ttk.Button(btn_frame, text="Move Down", style="Outline.TButton", command=self.move_merge_down); btn4.pack(fill="x", pady=2); Tooltip(btn4, TOOLTIP_TEXT.get("merge_move_down_btn"))

        self.merge_tree.heading('name', text='File Name')
        self.merge_tree.heading('pages', text='Pages')
        self.merge_tree.heading('size', text='Size')
        self.merge_tree.column('name', stretch=True, minwidth=250)
        self.merge_tree.column('pages', width=60, anchor='center', stretch=False)
        self.merge_tree.column('size', width=100, anchor='e', stretch=False)

        FileSelector(main_card, None, s.output_path, lambda: self.browse_save_file(s.output_path)).grid(row=2, column=0, columnspan=2, sticky="ew", pady=(15, 0))

        ttk.Label(parent, textvariable=self.tab_statuses['merge'], anchor="center").grid(row=1, column=0, sticky="ew", pady=(10, 0))
        self.merge_button = self._create_process_button(parent, "MERGE PDFS", self.process_merge, row=2, tooltip_key="merge_process_btn")
        self.update_merge_view()

    def _build_split_tab(self, parent):
        parent.columnconfigure(0, weight=1)
        s = self.split_settings

        io_frame = ttk.Frame(parent, style="Card.TFrame", padding=20)
        io_frame.grid(row=0, column=0, sticky="ew")
        io_frame.columnconfigure(0, weight=1)
        self.drop_zones['split'] = self._create_drop_zone(io_frame, s.input_path)
        FileSelector(io_frame, s.input_path, s.output_dir, lambda: self.browse_dir(s.output_dir))

        options_card = ttk.Frame(parent, style="Card.TFrame", padding=20)
        options_card.grid(row=1, column=0, sticky="ew", pady=10)
        ttk.Label(options_card, text="Split Mode", font=(styles.FONT_FAMILY, 14, "bold"), style="Card.TLabel").pack(anchor="w", pady=(0, 15))

        mode_frame = ttk.Frame(options_card, style="Card.TFrame")
        mode_frame.pack(fill="x")
        rb1 = ttk.Radiobutton(mode_frame, text="Split to Single Pages", variable=s.mode, value=SPLIT_SINGLE, style="Card.TRadiobutton"); rb1.pack(fill="x"); Tooltip(rb1, TOOLTIP_TEXT.get("split_mode_single"))

        n_frame = ttk.Frame(options_card, style="Card.TFrame"); n_frame.pack(fill="x")
        rb2 = ttk.Radiobutton(n_frame, text="Split Every N Pages. N =", variable=s.mode, value=SPLIT_EVERY_N, style="Card.TRadiobutton"); rb2.pack(side="left");
        self.split_value_entry = ttk.Entry(n_frame, textvariable=s.value, width=5)
        self.split_value_entry.pack(side="left", ipady=1)
        Tooltip(n_frame, TOOLTIP_TEXT.get("split_mode_every_n"))

        custom_frame = ttk.Frame(options_card, style="Card.TFrame"); custom_frame.pack(fill="x")
        rb3 = ttk.Radiobutton(custom_frame, text="Custom Range(s):", variable=s.mode, value=SPLIT_CUSTOM, style="Card.TRadiobutton"); rb3.pack(side="left", fill="x")
        self.split_custom_entry = ttk.Entry(custom_frame, textvariable=s.value)
        self.split_custom_entry.pack(side="left", expand=True, fill="x", ipady=1, padx=5)
        Tooltip(custom_frame, TOOLTIP_TEXT.get("split_mode_custom"))

        ttk.Label(parent, textvariable=self.tab_statuses['split'], anchor="center").grid(row=2, column=0, sticky="ew", pady=(10, 0))
        self.split_button = self._create_process_button(parent, "SPLIT PDF", self.process_split, row=3, tooltip_key="split_process_btn")
        self._update_split_validation()

    def _update_split_validation(self, *args):
        if not hasattr(self, 'split_value_entry'): return
        mode = self.split_settings.mode.get()
        self.split_value_entry.config(state="disabled")
        self.split_custom_entry.config(state="disabled")
        if mode == SPLIT_EVERY_N:
            self.split_value_entry.config(state="normal", validate='key', validatecommand=self.vcmd_int)
        elif mode == SPLIT_CUSTOM:
            self.split_custom_entry.config(state="normal", validate='key', validatecommand=self.vcmd_pagerange)
        else:
            self.split_value_entry.config(validate='none')
            self.split_custom_entry.config(validate='none')

    def _build_rotate_tab(self, parent):
        parent.columnconfigure(0, weight=1)
        s = self.rotate_settings

        main_card = ttk.Frame(parent, style="Card.TFrame", padding=20)
        main_card.grid(row=0, column=0, sticky="nsew")
        main_card.columnconfigure(0, weight=1)
        main_card.columnconfigure(1, weight=1)

        left_panel = ttk.Frame(main_card, style="Card.TFrame")
        left_panel.grid(row=0, column=0, sticky="nsew", padx=(0, 20))
        left_panel.columnconfigure(0, weight=1)

        self.drop_zones['rotate'] = self._create_drop_zone(left_panel, s.input_path)
        FileSelector(left_panel, s.input_path, s.output_path, lambda: self.browse_save_file(s.output_path)).grid(row=1, column=0, sticky="ew")

        options_card = ttk.Frame(left_panel, style="Card.TFrame", padding=(0, 20))
        options_card.grid(row=2, column=0, sticky="ew", pady=(20,0))
        options_card.columnconfigure(0, weight=1)
        ttk.Label(options_card, text="Rotation Angle", font=(styles.FONT_FAMILY, 14, "bold"), style="Card.TLabel").pack(anchor="w", pady=(0, 15))

        btn_frame = ttk.Frame(options_card, style="Card.TFrame")
        btn_frame.pack(fill="both", expand=True)
        btn_frame.columnconfigure(0, weight=1); btn_frame.columnconfigure(1, weight=1); btn_frame.columnconfigure(2, weight=1)
        Tooltip(btn_frame, TOOLTIP_TEXT.get("rotate_angle_btns"))

        angles = list(ROTATION_MAP.keys())[1:]
        for i, angle_text in enumerate(angles):
            btn = ttk.Radiobutton(btn_frame, text=angle_text, variable=s.angle, value=angle_text, style="Card.TRadiobutton")
            btn.grid(row=0, column=i, sticky="ew", padx=5)

        preview_frame = ttk.Frame(main_card, style="Card.TFrame")
        preview_frame.grid(row=0, column=1, sticky="nsew")
        preview_frame.rowconfigure(0, weight=1); preview_frame.columnconfigure(0, weight=1)
        self.rotate_preview_canvas = tk.Canvas(preview_frame, highlightthickness=1)
        self.rotate_preview_canvas.grid(row=0, column=0, sticky="nsew")
        Tooltip(self.rotate_preview_canvas, TOOLTIP_TEXT.get("rotate_preview"))

        ttk.Label(parent, textvariable=self.tab_statuses['rotate'], anchor="center").grid(row=1, column=0, sticky="ew", pady=(10, 0))
        self.rotate_button = self._create_process_button(parent, "ROTATE PDF", self.process_rotate, row=2, tooltip_key="rotate_process_btn")

    def _build_delete_tab(self, parent):
        parent.columnconfigure(0, weight=1)
        s = self.delete_settings

        main_card = ttk.Frame(parent, style="Card.TFrame", padding=20)
        main_card.grid(row=0, column=0, sticky="nsew")
        main_card.columnconfigure(0, weight=1)
        main_card.columnconfigure(1, weight=1)

        left_panel = ttk.Frame(main_card, style="Card.TFrame")
        left_panel.grid(row=0, column=0, sticky="nsew", padx=(0, 20))
        left_panel.columnconfigure(0, weight=1)

        self.drop_zones['delete'] = self._create_drop_zone(left_panel, s.input_path)
        FileSelector(left_panel, s.input_path, s.output_path, lambda: self.browse_save_file(s.output_path)).grid(row=1, column=0, sticky="ew")

        options_card = ttk.Frame(left_panel, style="Card.TFrame", padding=(0, 20))
        options_card.grid(row=2, column=0, sticky="ew", pady=(20,0))
        options_card.columnconfigure(0, weight=1)
        ttk.Label(options_card, text="Pages to Delete", font=(styles.FONT_FAMILY, 14, "bold"), style="Card.TLabel").grid(row=0, column=0, sticky="w", pady=(0,5))
        ttk.Label(options_card, text="e.g., 1, 3-5, 8-end", style="Card.TLabel").grid(row=1, column=0, sticky="w", pady=(0, 15))

        entry = ttk.Entry(options_card, textvariable=s.page_range, validate='key', validatecommand=self.vcmd_pagerange, font=(styles.FONT_FAMILY, 12)); entry.grid(row=2, column=0, sticky="ew")
        Tooltip(entry, TOOLTIP_TEXT.get("delete_pages_entry"))

        preview_frame = ttk.Frame(main_card, style="Card.TFrame")
        preview_frame.grid(row=0, column=1, sticky="nsew")
        preview_frame.rowconfigure(0, weight=1); preview_frame.columnconfigure(0, weight=1)
        self.delete_preview_canvas = tk.Canvas(preview_frame, highlightthickness=1)
        self.delete_preview_canvas.grid(row=0, column=0, sticky="nsew")
        Tooltip(self.delete_preview_canvas, TOOLTIP_TEXT.get("delete_preview"))

        ttk.Label(parent, textvariable=self.tab_statuses['delete'], anchor="center").grid(row=1, column=0, sticky="ew", pady=(10, 0))
        self.delete_button = self._create_process_button(parent, "DELETE PAGES", self.process_delete_pages, row=2, tooltip_key="delete_process_btn")

    def _get_stamp_text_content(self):
        s = self.stamp_settings
        text_content = s.text.get()
        if s.bates_enabled.get() and "%Bates" not in text_content:
            if text_content:
                text_content = f"%Bates\\n{text_content}"
            else:
                text_content = "%Bates"
        return text_content.replace('\n', '\\n')
        
    def _build_password_tab(self, parent):
        parent.columnconfigure(0, weight=1)
        s = self.password_settings

        io_frame = ttk.Frame(parent, style="Card.TFrame", padding=20)
        io_frame.grid(row=0, column=0, sticky="ew")
        io_frame.columnconfigure(0, weight=1)
        self.drop_zones['password'] = self._create_drop_zone(io_frame, s.input_path)
        FileSelector(io_frame, s.input_path, s.output_path, lambda: self.browse_save_file(s.output_path))

        # --- Encrypt Frame ---
        encrypt_frame = ttk.LabelFrame(parent, text="Encrypt PDF", padding=15)
        encrypt_frame.grid(row=1, column=0, sticky="ew", pady=(10, 5))
        encrypt_frame.columnconfigure(1, weight=1)

        ttk.Label(encrypt_frame, text="User Password:", style="Card.TLabel").grid(row=0, column=0, sticky="w", padx=(0, 10), pady=4)
        self.pass_encrypt_user_entry = ttk.Entry(encrypt_frame, textvariable=s.user_password, show="*")
        self.pass_encrypt_user_entry.grid(row=0, column=1, sticky="ew", pady=4)
        Tooltip(self.pass_encrypt_user_entry, TOOLTIP_TEXT.get("password_user_entry"))
        
        ttk.Label(encrypt_frame, text="Owner Password:", style="Card.TLabel").grid(row=1, column=0, sticky="w", padx=(0, 10), pady=4)
        self.pass_encrypt_owner_entry = ttk.Entry(encrypt_frame, textvariable=s.owner_password, show="*")
        self.pass_encrypt_owner_entry.grid(row=1, column=1, sticky="ew", pady=4)
        Tooltip(self.pass_encrypt_owner_entry, TOOLTIP_TEXT.get("password_owner_entry"))

        self.permissions_frame = ttk.LabelFrame(encrypt_frame, text="Permissions (when using User Password)", padding=10)
        self.permissions_frame.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(10,0))
        cb_print = ttk.Checkbutton(self.permissions_frame, text="Allow Printing", variable=s.allow_printing); cb_print.pack(anchor="w"); Tooltip(cb_print, TOOLTIP_TEXT.get("password_allow_printing"))
        cb_modify = ttk.Checkbutton(self.permissions_frame, text="Allow Modification", variable=s.allow_modification); cb_modify.pack(anchor="w"); Tooltip(cb_modify, TOOLTIP_TEXT.get("password_allow_modification"))
        cb_copy = ttk.Checkbutton(self.permissions_frame, text="Allow Copy & Extract", variable=s.allow_copy_and_extract); cb_copy.pack(anchor="w"); Tooltip(cb_copy, TOOLTIP_TEXT.get("password_allow_copy"))
        cb_annotate = ttk.Checkbutton(self.permissions_frame, text="Allow Annotations & Forms", variable=s.allow_annotations_and_forms); cb_annotate.pack(anchor="w"); Tooltip(cb_annotate, TOOLTIP_TEXT.get("password_allow_annotations"))
        
        self.encrypt_button = self._create_process_button(encrypt_frame, "ENCRYPT PDF", self.process_encrypt, row=3, columnspan=2, tooltip_key="password_encrypt_btn")

        # --- Decrypt Frame ---
        decrypt_frame = ttk.LabelFrame(parent, text="Decrypt PDF", padding=15)
        decrypt_frame.grid(row=2, column=0, sticky="ew", pady=5)
        decrypt_frame.columnconfigure(1, weight=1)

        ttk.Label(decrypt_frame, text="Current Password:", style="Card.TLabel").grid(row=0, column=0, sticky="w", padx=(0, 10), pady=4)
        self.pass_decrypt_entry = ttk.Entry(decrypt_frame, textvariable=s.decrypt_password, show="*")
        self.pass_decrypt_entry.grid(row=0, column=1, sticky="ew", pady=4)
        Tooltip(self.pass_decrypt_entry, TOOLTIP_TEXT.get("password_decrypt_entry"))
        
        self.decrypt_button = self._create_process_button(decrypt_frame, "DECRYPT PDF", self.process_decrypt, row=1, columnspan=2, tooltip_key="password_decrypt_btn")

        # --- Shared Controls ---
        show_pass_cb = ttk.Checkbutton(parent, text="Show Passwords", variable=s.show_passwords, command=self.toggle_password_visibility)
        show_pass_cb.grid(row=3, column=0, sticky="w", pady=(5, 10), padx=5)
        
        ttk.Label(parent, textvariable=self.tab_statuses['password'], anchor="center").grid(row=4, column=0, sticky="ew", pady=(10, 0))

    def _build_stamp_tab(self, parent):
        parent.columnconfigure(0, weight=1)
        s = self.stamp_settings

        main_card = ttk.Frame(parent, style="Card.TFrame", padding=20)
        main_card.grid(row=0, column=0, sticky="nsew")
        main_card.columnconfigure(0, weight=1)
        main_card.columnconfigure(1, weight=1)

        left_panel = ttk.Frame(main_card, style="Card.TFrame")
        left_panel.grid(row=0, column=0, sticky="nsew", padx=(0, 20))
        left_panel.columnconfigure(0, weight=1)

        self.drop_zones['stamp'] = self._create_drop_zone(left_panel, s.input_path)
        FileSelector(left_panel, s.input_path, s.output_path, lambda: self.browse_save_file(s.output_path)).grid(row=1, column=0, sticky="ew")

        notebook = ttk.Notebook(left_panel)
        notebook.grid(row=2, column=0, sticky="ew", pady=(20,0))
        notebook.bind("<<NotebookTabChanged>>", lambda e, nb=notebook: self._on_stamp_mode_change(nb))

        image_tab = ttk.Frame(notebook, padding=15); text_tab = ttk.Frame(notebook, padding=15)
        notebook.add(image_tab, text="Image"); notebook.add(text_tab, text="Text")

        image_tab.columnconfigure(1, weight=1)
        ttk.Label(image_tab, text="Image File:").grid(row=0, column=0, sticky="w", padx=(0,10), pady=4)
        ttk.Entry(image_tab, textvariable=s.image_path).grid(row=0, column=1, sticky="ew", pady=4)
        btn_img_browse = ttk.Button(image_tab, text="Browse", command=lambda: self.browse_image(s.image_path)); btn_img_browse.grid(row=0, column=2, sticky="e", padx=(10,0), pady=4)
        Tooltip(btn_img_browse, TOOLTIP_TEXT.get("stamp_image_browse"))

        size_frame = ttk.Frame(image_tab)
        size_frame.grid(row=1, column=0, columnspan=3, sticky="ew", pady=4)
        size_frame.columnconfigure(1, weight=1)
        ttk.Label(size_frame, text="Scale:").grid(row=0, column=0, sticky="w", padx=(0,10))
        scale_slider = ttk.Scale(size_frame, from_=10, to=200, orient="horizontal", variable=s.image_scale, style="Horizontal.TScale")
        scale_slider.grid(row=0, column=1, sticky="ew")
        scale_label = ttk.Label(size_frame, text="100%")
        scale_label.grid(row=0, column=2, sticky="e", padx=(10,0))
        def update_scale_label(*args):
            scale_label.config(text=f"{s.image_scale.get():.0f}%")
        s.image_scale.trace_add("write", update_scale_label)
        update_scale_label()
        Tooltip(size_frame, TOOLTIP_TEXT.get("stamp_image_scale"))

        text_tab.columnconfigure(1, weight=1)
        ttk.Label(text_tab, text="Text:").grid(row=0, column=0, sticky="nw", pady=4, padx=(0,10))
        text_entry = tk.Text(text_tab, height=4, width=30, wrap="word", bg=self.palette.get("WIDGET_BG"), fg=self.palette.get("TEXT"),
                             relief='solid', bd=1, highlightbackground=self.palette.get("BORDER"),
                             highlightcolor=self.palette.get("ACCENT"), insertbackground=self.palette.get("TEXT"),
                             font=(styles.FONT_FAMILY, 10))
        text_entry.grid(row=0, column=1, columnspan=2, sticky="ew", pady=4)
        Tooltip(text_entry, TOOLTIP_TEXT.get("stamp_text_entry"))
        if s.text.get(): text_entry.insert("1.0", s.text.get())
        text_entry.bind("<<Modified>>", lambda e: self._on_text_modified(e, text_entry))

        bates_frame = ttk.Frame(text_tab)
        bates_frame.grid(row=1, column=1, columnspan=2, sticky="w")
        cb_bates = ttk.Checkbutton(bates_frame, text="Enable Bates Numbering (%Bates)", variable=s.bates_enabled); cb_bates.pack(side="left"); Tooltip(cb_bates, TOOLTIP_TEXT.get("stamp_bates_check"))
        ttk.Label(bates_frame, text="Start at:").pack(side="left", padx=(10, 5))
        entry_bates = ttk.Entry(bates_frame, textvariable=s.bates_start, width=8); entry_bates.pack(side="left"); Tooltip(entry_bates, TOOLTIP_TEXT.get("stamp_bates_start"))

        ttk.Label(text_tab, text="Font:").grid(row=2, column=0, sticky="w", pady=4, padx=(0,10))
        font_combo = ttk.Combobox(text_tab, textvariable=s.font, values=PDF_FONTS, state="readonly"); font_combo.grid(row=2, column=1, sticky="ew", pady=4); Tooltip(font_combo, TOOLTIP_TEXT.get("stamp_font_combo"))
        ttk.Label(text_tab, text="Size:").grid(row=3, column=0, sticky="w", pady=4, padx=(0,10))
        font_size_entry = ttk.Entry(text_tab, textvariable=s.font_size, width=5); font_size_entry.grid(row=3, column=1, sticky="w", pady=4); Tooltip(font_size_entry, TOOLTIP_TEXT.get("stamp_font_size"))
        btn_color = ttk.Button(text_tab, text="Pick Color", command=self.pick_stamp_color); btn_color.grid(row=3, column=2, sticky="e", padx=(10,0), pady=4); Tooltip(btn_color, TOOLTIP_TEXT.get("stamp_font_color"))

        options_card = ttk.Frame(left_panel, style="Card.TFrame")
        options_card.grid(row=3, column=0, sticky="ew", pady=(20,0))
        options_card.columnconfigure(1, weight=1)
        ttk.Label(options_card, text="Position:").grid(row=0, column=0, sticky="nw", padx=(0,10), pady=5)
        pos_selector = PositionSelector(options_card, variable=s.pos, positions=STAMP_POSITIONS)
        pos_selector.grid(row=0, column=1, sticky="w", pady=5); Tooltip(pos_selector, TOOLTIP_TEXT.get("stamp_position"))

        ttk.Label(options_card, text="Opacity:").grid(row=1, column=0, sticky="w", padx=(0,10), pady=5)
        opacity_slider = ttk.Scale(options_card, from_=0.1, to=1.0, orient="horizontal", variable=s.opacity, style="Horizontal.TScale"); opacity_slider.grid(row=1, column=1, sticky="ew", pady=5); Tooltip(opacity_slider, TOOLTIP_TEXT.get("stamp_opacity"))

        stamp_on_top_toggle = ModernToggle(options_card, text="Stamp on Top", variable=s.on_top, palette=self.palette)
        stamp_on_top_toggle.grid(row=2, column=0, columnspan=2, sticky="w", pady=10); Tooltip(stamp_on_top_toggle, TOOLTIP_TEXT.get("stamp_on_top"))
        self.toggles.append(stamp_on_top_toggle)

        preview_frame = ttk.Frame(main_card, style="Card.TFrame")
        preview_frame.grid(row=0, column=1, sticky="nsew", rowspan=4)
        preview_frame.rowconfigure(0, weight=1); preview_frame.columnconfigure(0, weight=1)
        self.stamp_preview_canvas = tk.Canvas(preview_frame, highlightthickness=1)
        self.stamp_preview_canvas.grid(row=0, column=0, sticky="nsew")

        ttk.Label(parent, textvariable=self.tab_statuses['stamp'], anchor="center").grid(row=1, column=0, sticky="ew", pady=(10, 0))
        self.stamp_button = self._create_process_button(parent, "APPLY STAMP", self.process_stamp, row=2, tooltip_key="stamp_process_btn")

    def _on_text_modified(self, event, widget):
        widget.edit_modified(False)
        self.stamp_settings.text.set(widget.get("1.0", "end-1c"))

    def _on_stamp_mode_change(self, notebook):
        try:
            selected_tab_text = notebook.tab(notebook.select(), "text")
            current_mode = self.stamp_settings.mode.get()
            new_mode = STAMP_IMAGE if selected_tab_text == "Image" else STAMP_TEXT
            if current_mode != new_mode:
                self.stamp_settings.mode.set(new_mode)
        except tk.TclError:
            pass

    def _build_page_number_tab(self, parent):
        parent.columnconfigure(0, weight=1)
        s = self.page_number_settings

        io_frame = ttk.Frame(parent, style="Card.TFrame", padding=20)
        io_frame.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        io_frame.columnconfigure(0, weight=1)
        self.drop_zones['page_number'] = self._create_drop_zone(io_frame, s.input_path)
        FileSelector(io_frame, s.input_path, s.output_path, lambda: self.browse_save_file(s.output_path))

        options_card = ttk.Frame(parent, style="Card.TFrame", padding=20)
        options_card.grid(row=1, column=0, sticky="ew", pady=10)
        options_card.columnconfigure(1, weight=1)

        ttk.Label(options_card, text="Header/Footer Options", font=(styles.FONT_FAMILY, 14, "bold"), style="Card.TLabel").grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 15))

        mode_frame = ttk.Frame(options_card, style="Card.TFrame")
        mode_frame.grid(row=1, column=0, columnspan=3, sticky="ew", pady=5)

        MODES = ["Page Number", "Page X of Y", "Custom"]

        def toggle_custom_entry_state(*args):
            state = "normal" if s.mode.get() == "Custom" else "disabled"
            if hasattr(self, 'pn_custom_entry'):
                self.pn_custom_entry.config(state=state)

        s.mode.trace_add("write", toggle_custom_entry_state)

        rb_pn1 = ttk.Radiobutton(mode_frame, text="Page Number (e.g., 1)", variable=s.mode, value=MODES[0], style="Card.TRadiobutton"); rb_pn1.pack(anchor="w"); Tooltip(rb_pn1, TOOLTIP_TEXT.get("hf_mode_page_num"))
        rb_pn2 = ttk.Radiobutton(mode_frame, text="Page X of Y (e.g., 1 of 10)", variable=s.mode, value=MODES[1], style="Card.TRadiobutton"); rb_pn2.pack(anchor="w"); Tooltip(rb_pn2, TOOLTIP_TEXT.get("hf_mode_page_x_of_y"))

        custom_frame = ttk.Frame(mode_frame, style="Card.TFrame")
        custom_frame.pack(fill="x", expand=True)
        rb_pn3 = ttk.Radiobutton(custom_frame, text="Custom:", variable=s.mode, value=MODES[2], style="Card.TRadiobutton"); rb_pn3.pack(side="left", anchor="n");
        self.pn_custom_entry = ttk.Entry(custom_frame, textvariable=s.custom_text)
        self.pn_custom_entry.pack(side="left", fill="x", expand=True)
        Tooltip(custom_frame, TOOLTIP_TEXT.get("hf_mode_custom"))

        ttk.Label(options_card, text="Position:", style="Card.TLabel").grid(row=2, column=0, sticky="nw", pady=5, padx=(0,10))
        pos_selector = PositionSelector(options_card, variable=s.pos, positions=PAGE_NUMBER_POSITIONS)
        pos_selector.grid(row=2, column=1, sticky="w", pady=5); Tooltip(pos_selector, TOOLTIP_TEXT.get("hf_position"))

        ttk.Label(options_card, text="Font:", style="Card.TLabel").grid(row=3, column=0, sticky="w", pady=5, padx=(0,10))
        font_frame = ttk.Frame(options_card, style="Card.TFrame")
        font_frame.grid(row=3, column=1, sticky="ew", pady=5)
        font_frame.columnconfigure(0, weight=1)
        combo_font = ttk.Combobox(font_frame, textvariable=s.font, values=PDF_FONTS, state="readonly"); combo_font.grid(row=0, column=0, sticky="ew"); Tooltip(combo_font, TOOLTIP_TEXT.get("hf_font_combo"))

        ttk.Label(font_frame, text="Size:", style="Card.TLabel").grid(row=0, column=1, sticky="w", padx=(10,5))
        entry_size = ttk.Entry(font_frame, textvariable=s.font_size, width=5, validate='key', validatecommand=self.vcmd_int); entry_size.grid(row=0, column=2, sticky="w"); Tooltip(entry_size, TOOLTIP_TEXT.get("hf_font_size"))
        btn_color = ttk.Button(font_frame, text="Color", command=self.pick_page_number_color); btn_color.grid(row=0, column=3, sticky="e", padx=(10,0)); Tooltip(btn_color, TOOLTIP_TEXT.get("hf_font_color"))

        ttk.Label(options_card, text="Page Range:", style="Card.TLabel").grid(row=4, column=0, sticky="w", pady=5, padx=(0,10))
        page_range_entry = ttk.Entry(options_card, textvariable=s.page_range, validate='key', validatecommand=self.vcmd_pagerange)
        page_range_entry.grid(row=4, column=1, sticky="ew", pady=5)
        Tooltip(page_range_entry, TOOLTIP_TEXT.get("hf_page_range"))

        toggle_custom_entry_state()

        ttk.Label(parent, textvariable=self.tab_statuses['page_number'], anchor="center").grid(row=2, column=0, sticky="ew", pady=(10, 0))
        self.page_number_button = self._create_process_button(parent, "ADD HEADER/FOOTER", self.process_page_number, row=3, tooltip_key="hf_process_btn")

    def _build_toc_tab(self, parent):
        parent.columnconfigure(0, weight=1)
        s = self.toc_settings

        io_frame = ttk.Frame(parent, style="Card.TFrame", padding=20)
        io_frame.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        io_frame.columnconfigure(0, weight=1)
        self.drop_zones['toc'] = self._create_drop_zone(io_frame, s.input_path)
        FileSelector(io_frame, s.input_path, s.output_path, lambda: self.browse_save_file(s.output_path))

        options_card = ttk.Frame(parent, style="Card.TFrame", padding=20)
        options_card.grid(row=1, column=0, sticky="ew", pady=10)
        options_card.columnconfigure(1, weight=1)

        ttk.Label(options_card, text="Table of Contents", font=(styles.FONT_FAMILY, 14, "bold"), style="Card.TLabel").grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 15))

        ttk.Label(options_card, text="Title:", style="Card.TLabel").grid(row=1, column=0, sticky="w", pady=5, padx=(0,10))
        entry_title = ttk.Entry(options_card, textvariable=s.title); entry_title.grid(row=1, column=1, sticky="ew", pady=5); Tooltip(entry_title, TOOLTIP_TEXT.get("toc_title"))
        ttk.Label(options_card, text="Font:", style="Card.TLabel").grid(row=2, column=0, sticky="w", pady=5, padx=(0,10))
        combo_font = ttk.Combobox(options_card, textvariable=s.font, values=PDF_FONTS, state="readonly"); combo_font.grid(row=2, column=1, sticky="ew", pady=5); Tooltip(combo_font, TOOLTIP_TEXT.get("toc_font"))
        ttk.Label(options_card, text="Font Size:", style="Card.TLabel").grid(row=3, column=0, sticky="w", pady=5, padx=(0,10))
        entry_size = ttk.Entry(options_card, textvariable=s.font_size, width=8, validate='key', validatecommand=self.vcmd_int); entry_size.grid(row=3, column=1, sticky="w", pady=5); Tooltip(entry_size, TOOLTIP_TEXT.get("toc_font_size"))

        toggle_frame = ttk.Frame(options_card, style="Card.TFrame")
        toggle_frame.grid(row=4, column=0, columnspan=2, sticky="ew", pady=(15,0))
        dot_leader_toggle = ModernToggle(toggle_frame, text="Use Dot Leaders", variable=s.dot_leaders, palette=self.palette)
        dot_leader_toggle.pack(anchor="w", pady=2); Tooltip(dot_leader_toggle, TOOLTIP_TEXT.get("toc_dot_leaders"))
        self.toggles.append(dot_leader_toggle)
        no_bookmark_toggle = ModernToggle(toggle_frame, text="Don't Bookmark the ToC", variable=s.no_bookmark, palette=self.palette)
        no_bookmark_toggle.pack(anchor="w", pady=2); Tooltip(no_bookmark_toggle, TOOLTIP_TEXT.get("toc_no_bookmark"))
        self.toggles.append(no_bookmark_toggle)

        ttk.Label(parent, textvariable=self.tab_statuses['toc'], anchor="center").grid(row=2, column=0, sticky="ew", pady=(10, 0))
        self.toc_button = self._create_process_button(parent, "GENERATE TOC", self.process_toc, row=3, tooltip_key="toc_process_btn")

    def _build_metadata_tab(self, parent):
        parent.columnconfigure(0, weight=1)
        s = self.meta_settings

        io_frame = ttk.Frame(parent, style="Card.TFrame", padding=20)
        io_frame.grid(row=0, column=0, sticky="ew")
        io_frame.columnconfigure(1, weight=1)
        ttk.Label(io_frame, text="Input", style="Card.TLabel").grid(row=0, column=0, sticky="e", padx=(0, 10), pady=(0,10))
        in_entry = ttk.Entry(io_frame, textvariable=s.input_path); in_entry.grid(row=0, column=1, sticky="we", pady=(0,10))
        btn_browse = ttk.Button(io_frame, text="Browse", command=lambda: self.browse_file(s.input_path), width=8); btn_browse.grid(row=0, column=2, padx=(5,0), pady=(0,10))
        Tooltip(btn_browse, TOOLTIP_TEXT.get("meta_browse_btn"))

        fields_card = ttk.Frame(parent, style="Card.TFrame", padding=20)
        fields_card.grid(row=1, column=0, sticky="ew", pady=10)
        fields_card.columnconfigure(1, weight=1)

        fields = [("Title:", s.title, "meta_title"), ("Author:", s.author, "meta_author"), ("Subject:", s.subject, "meta_subject"), ("Keywords:", s.keywords, "meta_keywords")]
        for i, (label, var, key) in enumerate(fields):
            ttk.Label(fields_card, text=label, style="Card.TLabel").grid(row=i, column=0, sticky="w", padx=(0,10), pady=4)
            entry = ttk.Entry(fields_card, textvariable=var); entry.grid(row=i, column=1, sticky="ew", pady=4)
            Tooltip(entry, TOOLTIP_TEXT.get(key))

        ttk.Label(parent, textvariable=self.tab_statuses['metadata'], anchor="center").grid(row=2, column=0, sticky="ew", pady=(10, 0))
        self.meta_button = self._create_process_button(parent, "SAVE METADATA (OVERWRITE)", self.save_metadata, row=3, tooltip_key="meta_process_btn")

    def _build_convert_tab(self, parent):
        parent.columnconfigure(0, weight=1)
        s = self.convert_settings

        io_frame = ttk.Frame(parent, style="Card.TFrame", padding=20)
        io_frame.grid(row=0, column=0, sticky="ew")
        io_frame.columnconfigure(0, weight=1)
        self.drop_zones['convert'] = self._create_drop_zone(io_frame, s.input_path)
        FileSelector(io_frame, s.input_path, s.output_dir, lambda: self.browse_dir(s.output_dir))

        options_card = ttk.Frame(parent, style="Card.TFrame", padding=20)
        options_card.grid(row=1, column=0, sticky="ew", pady=10)
        options_card.columnconfigure(1, weight=1)

        ttk.Label(options_card, text="Output Format", font=(styles.FONT_FAMILY, 14, "bold"), style="Card.TLabel").grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 15))
        format_frame = ttk.Frame(options_card, style="Card.TFrame")
        format_frame.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(0, 15))
        Tooltip(format_frame, TOOLTIP_TEXT.get("convert_format_btns"))
        for fmt in IMAGE_FORMATS:
            ttk.Radiobutton(format_frame, text=fmt.upper(), variable=s.format, value=fmt, style="Card.TRadiobutton").pack(side="left", expand=True, fill="x")

        ttk.Label(options_card, text="Resolution (DPI)", font=(styles.FONT_FAMILY, 14, "bold"), style="Card.TLabel").grid(row=2, column=0, columnspan=2, sticky="w", pady=(0, 5))
        dpi_frame = ttk.Frame(options_card, style="Card.TFrame")
        dpi_frame.grid(row=3, column=0, columnspan=2, sticky="ew")
        dpi_frame.columnconfigure(0, weight=1)
        Tooltip(dpi_frame, TOOLTIP_TEXT.get("convert_dpi_slider"))
        ttk.Scale(dpi_frame, from_=72, to=600, orient="horizontal", variable=s.dpi_slider, style="Horizontal.TScale").grid(row=0, column=0, sticky="ew", padx=(0, 15))
        ttk.Entry(dpi_frame, textvariable=s.dpi, width=5, validate='key', validatecommand=self.vcmd_int).grid(row=0, column=1, sticky="e")

        ttk.Label(parent, textvariable=self.tab_statuses['convert'], anchor="center").grid(row=2, column=0, sticky="ew", pady=(10, 0))
        self.convert_button = self._create_process_button(parent, "CONVERT TO IMAGES", self.process_convert, row=3, tooltip_key="convert_process_btn")

    def _build_repair_tab(self, parent):
        parent.columnconfigure(0, weight=1)
        s = self.repair_settings

        io_frame = ttk.Frame(parent, style="Card.TFrame", padding=20)
        io_frame.grid(row=0, column=0, sticky="ew")
        io_frame.columnconfigure(0, weight=1)
        self.drop_zones['repair'] = self._create_drop_zone(io_frame, s.input_path)
        FileSelector(io_frame, s.input_path, s.output_path, lambda: self.browse_save_file(s.output_path))

        info_card = ttk.Frame(parent, style="Card.TFrame", padding=20)
        info_card.grid(row=1, column=0, sticky="ew", pady=10)
        info_label = ttk.Label(info_card, text="This tool attempts to repair corrupted or damaged PDF files by rebuilding them. Results may vary.", style="Card.TLabel", wraplength=1500, justify="left")
        info_label.pack(fill="x")

        ttk.Label(parent, textvariable=self.tab_statuses['repair'], anchor="center").grid(row=2, column=0, sticky="ew", pady=(10, 0))
        self.repair_button = self._create_process_button(parent, "ATTEMPT REPAIR", self.process_repair, row=3, tooltip_key="repair_process_btn")

    def _build_settings_tab(self, parent):
        parent.columnconfigure(0, weight=1)
        s = self.general_settings
        os = self.output_settings

        appearance_card = ttk.Frame(parent, style="Card.TFrame", padding=20)
        appearance_card.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        ttk.Label(appearance_card, text="Appearance", font=(styles.FONT_FAMILY, 14, "bold"), style="Card.TLabel").pack(anchor="w", pady=(0, 15))
        dark_mode_toggle = ModernToggle(appearance_card, text="Dark Mode", variable=s.dark_mode_enabled, command=self.toggle_theme, palette=self.palette)
        dark_mode_toggle.pack(anchor="w"); Tooltip(dark_mode_toggle, TOOLTIP_TEXT.get("settings_dark_mode"))
        self.toggles.append(dark_mode_toggle)

        logging_card = ttk.Frame(parent, style="Card.TFrame", padding=20)
        logging_card.grid(row=1, column=0, sticky="ew", pady=10)
        ttk.Label(logging_card, text="Logging", font=(styles.FONT_FAMILY, 14, "bold"), style="Card.TLabel").pack(anchor="w", pady=(0, 15))
        logging_toggle = ModernToggle(logging_card, text="Enable File Logging (app.log)", variable=s.logging_enabled, command=self.configure_logging, palette=self.palette)
        logging_toggle.pack(anchor="w"); Tooltip(logging_toggle, TOOLTIP_TEXT.get("settings_logging"))
        self.toggles.append(logging_toggle)

        output_card = ttk.Frame(parent, style="Card.TFrame", padding=20)
        output_card.grid(row=2, column=0, sticky="ew", pady=10)
        output_card.columnconfigure(1, weight=1)

        ttk.Label(output_card, text="Default Output", font=(styles.FONT_FAMILY, 14, "bold"), style="Card.TLabel").grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 15))
        use_default_folder_toggle = ModernToggle(output_card, text="Use Default Output Folder", variable=os.use_default_folder, palette=self.palette)
        use_default_folder_toggle.grid(row=1, column=0, columnspan=3, sticky="w", pady=5); Tooltip(use_default_folder_toggle, TOOLTIP_TEXT.get("settings_use_default_folder"))
        self.toggles.append(use_default_folder_toggle)

        default_folder_frame = ttk.Frame(output_card, style="Card.TFrame")
        default_folder_frame.grid(row=2, column=0, columnspan=3, sticky="ew", pady=(0, 20), padx=(20,0))
        default_folder_frame.columnconfigure(1, weight=1)
        Tooltip(default_folder_frame, TOOLTIP_TEXT.get("settings_default_folder_entry"))
        ttk.Label(default_folder_frame, text="Folder:", style="Card.TLabel").grid(row=0, column=0, sticky="w", padx=(0, 10))
        self.default_folder_entry = ttk.Entry(default_folder_frame, textvariable=os.default_folder)
        self.default_folder_entry.grid(row=0, column=1, sticky="ew")
        ttk.Button(default_folder_frame, text="Browse", command=lambda: self.browse_dir(os.default_folder)).grid(row=0, column=2, sticky="e", padx=(10,0))

        def toggle_default_folder_state(*args):
            state = "normal" if os.use_default_folder.get() else "disabled"
            if hasattr(self, 'default_folder_entry'):
                self.default_folder_entry.config(state=state)
        os.use_default_folder.trace_add("write", toggle_default_folder_state)
        toggle_default_folder_state()

        ttk.Label(output_card, text="Output File Naming", font=(styles.FONT_FAMILY, 14, "bold"), style="Card.TLabel").grid(row=3, column=0, columnspan=3, sticky="w", pady=(10, 15))
        pattern_frame = ttk.Frame(output_card, style="Card.TFrame")
        pattern_frame.grid(row=4, column=0, columnspan=3, sticky="ew", padx=(20,0))

        ttk.Label(pattern_frame, text="Prefix:", style="Card.TLabel").grid(row=0, column=0, sticky="w", pady=2)
        entry_prefix = ttk.Entry(pattern_frame, textvariable=os.prefix, width=15); entry_prefix.grid(row=0, column=1, sticky="ew", pady=2, padx=5); Tooltip(entry_prefix, TOOLTIP_TEXT.get("settings_prefix"))
        ttk.Label(pattern_frame, text="Suffix:", style="Card.TLabel").grid(row=0, column=3, sticky="w", pady=2, padx=5)
        entry_suffix = ttk.Entry(pattern_frame, textvariable=os.suffix, width=15); entry_suffix.grid(row=0, column=4, sticky="ew", pady=2, padx=5); Tooltip(entry_suffix, TOOLTIP_TEXT.get("settings_suffix"))

        date_toggle = ModernToggle(pattern_frame, text="Add Date (YYYY-MM-DD)", variable=os.add_date, palette=self.palette)
        date_toggle.grid(row=1, column=0, columnspan=5, sticky="w", pady=5); Tooltip(date_toggle, TOOLTIP_TEXT.get("settings_add_date"))
        self.toggles.append(date_toggle)
        time_toggle = ModernToggle(pattern_frame, text="Add Time (HH-MM-SS)", variable=os.add_time, palette=self.palette)
        time_toggle.grid(row=2, column=0, columnspan=5, sticky="w", pady=5); Tooltip(time_toggle, TOOLTIP_TEXT.get("settings_add_time"))
        self.toggles.append(time_toggle)

        coffee_card = ttk.Frame(parent, style="Card.TFrame", padding=20)
        coffee_card.grid(row=3, column=0, sticky="ew", pady=10)
        coffee_card.columnconfigure(1, weight=1)

        try:
            img_data = base64.b64decode(COFFEE_ICON_B64.encode('ascii'))
            resample = Image.Resampling.LANCZOS if hasattr(Image, 'Resampling') else Image.LANCZOS

            pil_img_light = Image.open(BytesIO(img_data)).convert('RGBA')
            pil_img_light = pil_img_light.resize((48, 48), resample)
            self.coffee_icon_light = ImageTk.PhotoImage(pil_img_light)

            r, g, b, a = pil_img_light.split()
            rgb_image = Image.merge('RGB', (r, g, b))
            inverted_rgb = ImageOps.invert(rgb_image)
            r2, g2, b2 = inverted_rgb.split()
            pil_img_dark = Image.merge('RGBA', (r2, g2, b2, a))
            pil_img_dark = pil_img_dark.resize((48, 48), resample)
            self.coffee_icon_dark = ImageTk.PhotoImage(pil_img_dark)

            initial_icon = self.coffee_icon_dark if self.general_settings.dark_mode_enabled.get() else self.coffee_icon_light
            self.coffee_img_label = ttk.Label(coffee_card, image=initial_icon, style="Card.TLabel", cursor="hand2")
            self.coffee_img_label.grid(row=0, column=0, rowspan=2, padx=(0, 15))
            self.coffee_img_label.bind("<Button-1>", self.open_coffee_link)
            self.coffee_img_label.image = initial_icon
            Tooltip(self.coffee_img_label, TOOLTIP_TEXT.get("settings_buy_coffee"))

        except Exception as e:
            logging.warning(f"Could not load coffee icon: {e}")
            self.coffee_img_label = ttk.Label(coffee_card, text="☕", style="Card.TLabel", font=(styles.FONT_FAMILY, 24))
            self.coffee_img_label.grid(row=0, column=0, rowspan=2, padx=(0, 15))

        coffee_label = ttk.Label(coffee_card, text="If you enjoy using my app, please consider supporting its development.", style="Card.TLabel", wraplength=600)
        coffee_label.grid(row=0, column=1, sticky="w")

        link_label = ttk.Label(coffee_card, text="Buy me a coffee", style="Card.TLabel", foreground=self.palette.get("ACCENT"), cursor="hand2")
        link_label.grid(row=1, column=1, sticky="w")
        link_label.bind("<Button-1>", self.open_coffee_link)
        Tooltip(link_label, TOOLTIP_TEXT.get("settings_buy_coffee"))

        updates_card = ttk.Frame(parent, style="Card.TFrame", padding=20)
        updates_card.grid(row=4, column=0, sticky="ew", pady=10)
        updates_card.columnconfigure(0, weight=1)
        ttk.Label(updates_card, text="Updates", font=(styles.FONT_FAMILY, 14, "bold"), style="Card.TLabel").pack(anchor="w", pady=(0, 15))
        btn_updates = ttk.Button(updates_card, text="Check for Updates", command=self.check_for_updates, style="Outline.TButton"); btn_updates.pack(anchor="w")
        Tooltip(btn_updates, TOOLTIP_TEXT.get("settings_check_updates"))

    def open_coffee_link(self, event=None):
        webbrowser.open_new("https://www.buymeacoffee.com/deminimis")

    def check_for_updates(self, event=None):
        webbrowser.open_new("https://github.com/deminimis/minimalpdfcompress/releases")

    def _update_compress_options(self, *args):
        mode = self.compress_settings.compress_mode.get()
        is_lossy = mode == "Compression"
        is_lossless = mode == "Lossless"

        widget_state = "normal" if is_lossy else "disabled"

        if hasattr(self, 'dpi_slider'):
            self.dpi_slider.config(state=widget_state)
            self.dpi_entry.config(state=widget_state)
        if hasattr(self, 'bicubic_check'): self.bicubic_check.config(state=widget_state)
        if hasattr(self, 'downsample_threshold_check'): self.downsample_threshold_check.config(state=widget_state)
        if hasattr(self, 'grayscale_check'): self.grayscale_check.config(state=widget_state)
        if hasattr(self, 'cmyk_check'): self.cmyk_check.config(state=widget_state)
        if hasattr(self, 'quantize_toggle'):
            self.quantize_toggle.config(state=widget_state)
            if not is_lossy:
                self.compress_settings.quantize_colors.set(False)


        if hasattr(self, 'true_lossless_check'):
            state = "normal" if is_lossless else "disabled"
            self.true_lossless_check.config(state=state)
            if not is_lossless:
                self.compress_settings.true_lossless.set(False)

        if not is_lossy:
            self.compress_settings.use_bicubic.set(False)
            self.compress_settings.convert_to_grayscale.set(False)
            self.compress_settings.convert_to_cmyk.set(False)

    def toggle_password_visibility(self):
        show = "" if self.password_settings.show_passwords.get() else "*"
        if hasattr(self, 'pass_encrypt_user_entry'): self.pass_encrypt_user_entry.config(show=show)
        if hasattr(self, 'pass_encrypt_owner_entry'): self.pass_encrypt_owner_entry.config(show=show)
        if hasattr(self, 'pass_decrypt_entry'): self.pass_decrypt_entry.config(show=show)

    def update_is_folder(self, *args):
        path = self.compress_settings.input_path.get()
        self.is_folder.set(Path(path).is_dir() if path else False)

    def _update_output_path(self, in_var, out_var, is_folder=False, is_dir=False):
        in_path_str = in_var.get()
        if not in_path_str: return

        try:
            p = Path(in_path_str)
            os = self.output_settings

            if os.use_default_folder.get() and os.default_folder.get():
                output_dir = Path(os.default_folder.get())
            else:
                output_dir = p.parent

            if is_dir:
                output_p = output_dir / f"{p.stem}_output"
                out_var.set(str(output_p))
                return

            if is_folder or p.is_dir():
                suffix = os.suffix.get() or "_compressed"
                output_p = output_dir / f"{p.name}{suffix}"
                out_var.set(str(output_p))
                return

            if p.is_file():
                filename = ""
                if os.prefix.get():
                    filename += os.prefix.get()

                filename += p.stem

                if os.suffix.get():
                    filename += os.suffix.get()

                if os.add_date.get():
                    filename += f"_{datetime.now().strftime('%Y-%m-%d')}"

                if os.add_time.get():
                    filename += f"_{datetime.now().strftime('%H-%M-%S')}"

                filename += p.suffix

                output_p = output_dir / filename
                out_var.set(str(output_p))

        except Exception as e:
            logging.warning(f"Could not auto-set output path: {e}")

    def check_tools(self):
        try: self.gs_path = backend.find_ghostscript()
        except backend.GhostscriptNotFound as e: self.status.set(f"Error: {e}"); messagebox.showerror("Tool Not Found", f"{e}\nThe application may not function correctly.")
        try: self.cpdf_path = backend.find_cpdf()
        except backend.CpdfNotFound as e: logging.warning(f"cpdf not found: {e}")
        try: self.pngquant_path = backend.find_pngquant()
        except backend.PngquantNotFound as e: logging.warning(f"pngquant not found: {e}")
        try: self.jpegoptim_path = backend.find_jpegoptim()
        except backend.JpegoptimNotFound as e: logging.warning(f"jpegoptim not found: {e}")
        try: self.ect_path = backend.find_ect()
        except backend.EctNotFound as e: logging.warning(f"ECT not found: {e}")
        try: self.optipng_path = backend.find_optipng()
        except backend.OptipngNotFound as e: logging.warning(f"OptiPNG not found: {e}")


    def start_task(self, button, target_func, args, status_var):
        if self.active_process_button:
            messagebox.showwarning("Busy", "A process is already running.", parent=self.root)
            return
        self.active_process_button = button
        button.config(state="disabled")

        self.active_status_var = status_var
        self.active_status_var.set("Starting...")

        self.compress_progress_status.set("Starting...")
        self.main_progress_bar.start(10)
        self.overall_progress_var.set(0)

        thread = threading.Thread(target=target_func, args=args, daemon=True)
        thread.start()

    def on_task_complete(self, message):
        self.status.set("Ready")
        self.compress_progress_status.set(message)
        if self.active_status_var:
            self.active_status_var.set(message)
            self.active_status_var = None

        match = re.search(r'\((\d+\.\d+)%\)', message)
        if match:
            percent_saved = float(match.group(1))
            self.compression_ratio_var.set(percent_saved)

        if "Error" in message:
            messagebox.showerror("Error", message, parent=self.root)

        self.main_progress_bar.stop()
        self.main_progress_bar['value'] = 0

        if self.active_process_button:
            self.active_process_button.config(state="normal")
            self.active_process_button = None

    def _update_preview(self, pdf_path, canvas, page_num=0, rotate_angle=0):
        if not pdf_path:
            canvas.delete("all")
            return

        args = (pdf_path, page_num, rotate_angle)
        threading.Thread(target=self._render_thread_target, args=(args, canvas), daemon=True).start()

    def _render_thread_target(self, args, canvas):
        try:
            pdf_path, _, rotate_angle = args
            options = {'angle': rotate_angle}
            image = backend.generate_preview(self.gs_path, self.cpdf_path, pdf_path, 'rotate', options)
            self.root.after(0, self._display_preview_image, image, canvas)
        except Exception as e:
            logging.error(f"Error in render thread: {e}")
            self.root.after(0, self._display_preview_image, None, canvas)

    def _display_preview_image(self, pil_image, canvas):
        canvas.delete("all")
        canvas.update_idletasks()

        canvas_width = canvas.winfo_width()
        canvas_height = canvas.winfo_height()

        if pil_image:
            img_width, img_height = pil_image.size
            ratio = min(canvas_width / img_width, canvas_height / img_height)
            new_width = int(img_width * ratio)
            new_height = int(img_height * ratio)

            resample_method = Image.Resampling.LANCZOS if hasattr(Image, 'Resampling') else Image.ANTIALIAS
            resized_image = pil_image.resize((new_width, new_height), resample_method)

            photo_image = ImageTk.PhotoImage(resized_image)
            self._preview_image_cache[id(canvas)] = photo_image
            canvas.create_image(canvas_width / 2, canvas_height / 2, anchor=tk.CENTER, image=photo_image)
        else:
            if canvas_width > 1 and canvas_height > 1:
                 canvas.create_text(canvas_width / 2, canvas_height / 2, text="Preview N/A", anchor=tk.CENTER, fill=self.palette.get("DISABLED"))

    def _debounce_trigger(self, callback, delay=300):
        if self._preview_job is not None:
            self.root.after_cancel(self._preview_job)
        self._preview_job = self.root.after(delay, callback)

    def _trigger_rotate_preview(self, *args):
        self._debounce_trigger(self._execute_rotate_preview)

    def _execute_rotate_preview(self):
        s = self.rotate_settings
        angle = ROTATION_MAP.get(s.angle.get(), 0)
        self._update_preview(s.input_path.get(), self.rotate_preview_canvas, rotate_angle=angle)

    def _trigger_stamp_preview(self, *args):
        self._debounce_trigger(self._execute_stamp_preview)

    def _execute_stamp_preview(self):
        s = self.stamp_settings
        if not s.input_path.get():
            self.stamp_preview_canvas.delete("all")
            return

        stamp_opts = { 'pos': s.pos.get(), 'opacity': s.opacity.get(), 'on_top': s.on_top.get() }

        text_content = self._get_stamp_text_content()

        mode_opts = {
            'image_path': s.image_path.get(),
            'image_scale': s.image_scale.get() / 100.0,
            'text': text_content,
            'font': s.font.get(),
            'size': s.font_size.get(),
            'color': self._hex_to_cpdf_color(s.font_color.get()),
            'bates_start': s.bates_start.get() if s.bates_enabled.get() else None
        }
        args = (s.input_path.get(), stamp_opts, self.cpdf_path, s.mode.get(), mode_opts)
        threading.Thread(target=self._stamp_render_thread_target, args=(args, self.stamp_preview_canvas), daemon=True).start()

    def _stamp_render_thread_target(self, args, canvas):
        try:
            pdf_path, stamp_opts, _, mode, mode_opts = args
            options = {'stamp_opts': stamp_opts, 'mode_opts': mode_opts, 'mode': mode}
            image = backend.generate_preview(self.gs_path, self.cpdf_path, pdf_path, 'stamp', options)
            self.root.after(0, self._display_preview_image, image, canvas)
        except Exception as e:
            logging.error(f"Error in stamp render thread: {e}")
            self.root.after(0, self._display_preview_image, None, canvas)

    def process_compression(self):
        cs = self.compress_settings
        in_path, out_path = cs.input_path.get(), cs.output_path.get()
        if not in_path or not out_path: messagebox.showerror("Error", "Input and Output paths must be set.", parent=self.root); return

        self.compression_ratio_var.set(0.0)
        self.status.set("Starting compression...")
        params = {
            'input_path': in_path, 'output_path': out_path, 'mode': cs.compress_mode.get(),
            'convert_to_grayscale': cs.convert_to_grayscale.get(),
            'convert_to_cmyk': cs.convert_to_cmyk.get(),
            'dpi': cs.dpi.get(), 'gs_path': self.gs_path, 'cpdf_path': self.cpdf_path,
            'pngquant_path': self.pngquant_path,
            'jpegoptim_path': self.jpegoptim_path,
            'ect_path': self.ect_path,
            'optipng_path': self.optipng_path,
            'darken_text': cs.darken_text.get(),
            'strip_metadata': cs.strip_metadata.get(),
            'remove_interactive': cs.remove_interactive.get(),
            'remove_open_action': cs.remove_open_action.get(),
            'use_bicubic': cs.use_bicubic.get(),
            'downsample_threshold_enabled': cs.downsample_threshold_enabled.get(),
            'quantize_colors': cs.quantize_colors.get(),
            'quantize_level': cs.quantize_level.get(),
            'fast_web_view': cs.fast_web_view.get(),
            'only_if_smaller': cs.only_if_smaller.get(),
            'fast_mode': cs.fast_mode.get(),
            'true_lossless': cs.true_lossless.get()
        }
        self.start_task(self.compress_button, backend.run_compress_task, (params, self.is_folder.get(), self.progress_queue), status_var=self.compress_progress_status)

    def browse_output(self):
        if self.is_folder.get(): path = filedialog.askdirectory()
        else: path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
        if path: self.compress_settings.output_path.set(path)

    def update_merge_view(self):
        if not self.merge_tree: return
        for i in self.merge_tree.get_children():
            self.merge_tree.delete(i)
        if self.merge_settings.files:
            self.merge_tree.insert('', 'end', values=("", "Loading...", ""))
            threading.Thread(target=self._populate_merge_view_thread, daemon=True).start()

    def _populate_merge_view_thread(self):
        metadata_list = [backend.get_pdf_metadata(f) for f in self.merge_settings.files]
        self.root.after(0, self._insert_merge_data, metadata_list)

    def _insert_merge_data(self, metadata_list):
        if not self.merge_tree: return
        for i in self.merge_tree.get_children():
            self.merge_tree.delete(i)
        for data in metadata_list:
            self.merge_tree.insert('', 'end', values=(data['name'], data['pages'], data['size']))

    def remove_merge_file(self):
        selected_items = self.merge_tree.selection()
        if not selected_items: return

        filenames_to_remove = {self.merge_tree.item(item)['values'][0] for item in selected_items}

        new_files_list = [f for f in self.merge_settings.files if Path(f).name not in filenames_to_remove]
        self.merge_settings.files = new_files_list

        for item in selected_items:
            self.merge_tree.delete(item)

        if not self.merge_settings.files:
            self.merge_settings.output_path.set("")

    def move_merge_up(self):
        selected = self.merge_tree.selection()
        if not selected: return
        for item in selected:
            idx = self.merge_tree.index(item)
            if idx > 0:
                self.merge_tree.move(item, '', idx - 1)
                files = self.merge_settings.files
                files[idx], files[idx-1] = files[idx-1], files[idx]

    def move_merge_down(self):
        selected = self.merge_tree.selection()
        if not selected: return
        for item in reversed(selected):
            idx = self.merge_tree.index(item)
            if idx < len(self.merge_settings.files) - 1:
                self.merge_tree.move(item, '', idx + 1)
                files = self.merge_settings.files
                files[idx], files[idx+1] = files[idx+1], files[idx]

    def pick_stamp_color(self):
        color = colorchooser.askcolor(title="Select Text Color", initialcolor=self.stamp_settings.font_color.get())
        if color[1]: self.stamp_settings.font_color.set(color[1])

    def pick_page_number_color(self):
        color = colorchooser.askcolor(title="Select Text Color", initialcolor=self.page_number_settings.font_color.get())
        if color[1]: self.page_number_settings.font_color.set(color[1])

    def process_merge(self):
        s = self.merge_settings
        if not s.files or not s.output_path.get():
            messagebox.showerror("Error", "Please add files and set an output path.", parent=self.root)
            return
        args = (s.files, s.output_path.get(), self.progress_queue)
        self.start_task(self.merge_button, backend.run_merge_task, args, status_var=self.tab_statuses['merge'])

    def process_split(self):
        s = self.split_settings
        if not s.input_path.get() or not s.output_dir.get():
            messagebox.showerror("Error", "Please set input and output paths.", parent=self.root)
            return
        args = (s.input_path.get(), s.output_dir.get(), s.mode.get(), s.value.get(), self.progress_queue)
        self.start_task(self.split_button, backend.run_split_task, args, status_var=self.tab_statuses['split'])

    def process_rotate(self):
        s = self.rotate_settings
        if not s.input_path.get() or not s.output_path.get():
            messagebox.showerror("Error", "Please set input and output paths.", parent=self.root)
            return
        angle = ROTATION_MAP.get(s.angle.get(), 0)
        args = (s.input_path.get(), s.output_path.get(), angle, self.progress_queue)
        self.start_task(self.rotate_button, backend.run_rotate_task, args, status_var=self.tab_statuses['rotate'])

    def process_delete_pages(self):
        s = self.delete_settings
        page_range = s.page_range.get()
        if not s.input_path.get() or not s.output_path.get() or not page_range:
            messagebox.showerror("Error", "Please set input, output, and a page range.", parent=self.root)
            return
        args = (s.input_path.get(), s.output_path.get(), page_range, self.progress_queue)
        self.start_task(self.delete_button, backend.run_delete_pages_task, args, status_var=self.tab_statuses['delete'])

    def process_encrypt(self):
        s = self.password_settings
        in_path, out_path = s.input_path.get(), s.output_path.get()
        if not in_path or not out_path:
            messagebox.showerror("Error", "Input and Output paths must be set.", parent=self.root)
            return
        params = {
            'input_path': in_path, 'output_path': out_path, 'mode': 'add',
            'user_password': s.user_password.get(),
            'owner_password': s.owner_password.get(),
            'allow_printing': s.allow_printing.get(),
            'allow_modification': s.allow_modification.get(),
            'allow_copy_and_extract': s.allow_copy_and_extract.get(),
            'allow_annotations_and_forms': s.allow_annotations_and_forms.get()
        }
        self.start_task(self.encrypt_button, backend.run_password_task, (params, self.progress_queue), status_var=self.tab_statuses['password'])

    def process_decrypt(self):
        s = self.password_settings
        in_path, out_path = s.input_path.get(), s.output_path.get()
        if not in_path or not out_path:
            messagebox.showerror("Error", "Input and Output paths must be set.", parent=self.root)
            return
        params = {
            'input_path': in_path, 'output_path': out_path, 'mode': 'remove',
            'user_password': s.decrypt_password.get()
        }
        self.start_task(self.decrypt_button, backend.run_password_task, (params, self.progress_queue), status_var=self.tab_statuses['password'])

    def process_stamp(self):
        s = self.stamp_settings
        if not s.input_path.get() or not s.output_path.get():
            messagebox.showerror("Error", "Please set input and output paths.", parent=self.root)
            return
        if s.mode.get() == STAMP_IMAGE and not s.image_path.get():
            messagebox.showerror("Error", "Please select an image file for stamping.", parent=self.root)
            return

        text_content = self._get_stamp_text_content()

        stamp_opts = { 'pos': s.pos.get(), 'opacity': s.opacity.get(), 'on_top': s.on_top.get() }
        mode_opts = { 'image_path': s.image_path.get(),
                      'image_scale': s.image_scale.get() / 100.0,
                      'text': text_content,
                      'font': s.font.get(), 'size': s.font_size.get(), 'color': self._hex_to_cpdf_color(s.font_color.get()),
                      'bates_start': s.bates_start.get() if s.bates_enabled.get() else None }
        args = (s.input_path.get(), s.output_path.get(), stamp_opts, self.cpdf_path, self.progress_queue, s.mode.get(), mode_opts)
        self.start_task(self.stamp_button, backend.run_stamp_task, args, status_var=self.tab_statuses['stamp'])

    def process_page_number(self):
        s = self.page_number_settings
        if not s.input_path.get() or not s.output_path.get():
            messagebox.showerror("Error", "Please set input and output paths.", parent=self.root)
            return

        mode = s.mode.get()
        text_to_use = ""
        if mode == "Page Number":
            text_to_use = "%Page"
        elif mode == "Page X of Y":
            text_to_use = "%Page of %EndPage"
        elif mode == "Custom":
            text_to_use = s.custom_text.get()

        options = {
            'text': text_to_use,
            'pos': s.pos.get(),
            'font': s.font.get(),
            'font_size': s.font_size.get(),
            'color': self._hex_to_cpdf_color(s.font_color.get()),
            'page_range': s.page_range.get()
        }
        args = (s.input_path.get(), s.output_path.get(), self.cpdf_path, self.progress_queue, options)
        self.start_task(self.page_number_button, backend.run_page_number_task, args, status_var=self.tab_statuses['page_number'])

    def process_convert(self):
        s = self.convert_settings
        if not s.input_path.get() or not s.output_dir.get():
            messagebox.showerror("Error", "Please set input and output paths.", parent=self.root)
            return
        options = {'format': s.format.get(), 'dpi': s.dpi.get()}
        args = (self.gs_path, s.input_path.get(), s.output_dir.get(), options, self.progress_queue)
        self.start_task(self.convert_button, backend.run_pdf_to_image_task, args, status_var=self.tab_statuses['convert'])

    def process_repair(self):
        s = self.repair_settings
        if not s.input_path.get() or not s.output_path.get():
            messagebox.showerror("Error", "Please set input and output paths.", parent=self.root)
            return
        args = (s.input_path.get(), s.output_path.get(), self.progress_queue)
        self.start_task(self.repair_button, backend.run_repair_task, args, status_var=self.tab_statuses['repair'])

    def process_toc(self):
        s = self.toc_settings
        if not s.input_path.get() or not s.output_path.get():
            messagebox.showerror("Error", "Please set input and output paths.", parent=self.root)
            return
        options = {
            'title': s.title.get(), 'font': s.font.get(),
            'font_size': s.font_size.get(), 'dot_leaders': s.dot_leaders.get(),
            'no_bookmark': s.no_bookmark.get()
        }
        args = (self.cpdf_path, s.input_path.get(), s.output_path.get(), options, self.progress_queue)
        self.start_task(self.toc_button, backend.run_toc_task, args, status_var=self.tab_statuses['toc'])

    def load_metadata(self, *args):
        s = self.meta_settings
        pdf_path = s.input_path.get()
        if not pdf_path:
            return
        try:
            info = backend.run_metadata_task(META_LOAD, pdf_path, self.cpdf_path)
            s.title.set(info.get('title', ''))
            s.author.set(info.get('author', ''))
            s.subject.set(info.get('subject', ''))
            s.keywords.set(info.get('keywords', ''))
            self.status.set("Metadata loaded successfully.")
        except Exception as e:
            self.status.set(f"Failed to load metadata.")
            logging.error(f"Failed to load metadata: {e}")
            s.title.set('')
            s.author.set('')
            s.subject.set('')
            s.keywords.set('')

    def save_metadata(self):
        s = self.meta_settings
        pdf_path = s.input_path.get()
        if not pdf_path:
            messagebox.showwarning("Input Missing", "Please select a PDF file.", parent=self.root)
            return
        metadata = {
            'title': s.title.get(),
            'author': s.author.get(),
            'subject': s.subject.get(),
            'keywords': s.keywords.get()
        }
        try:
            if messagebox.askyesno("Confirm Overwrite", "This will permanently overwrite the metadata in the original file. Are you sure?", icon='warning', parent=self.root):
                self.tab_statuses['metadata'].set("Saving metadata...")
                self.root.update_idletasks()
                backend.run_metadata_task(META_SAVE, pdf_path, self.cpdf_path, metadata)
                self.status.set("Metadata saved successfully.")
                self.tab_statuses['metadata'].set("Save complete.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save metadata: {e}", parent=self.root)
            self.tab_statuses['metadata'].set("Error during save.")

    def toggle_theme(self):
        self.palette = styles.apply_theme(self.root, 'dark' if self.general_settings.dark_mode_enabled.get() else 'light')
        self._update_widget_colors()

    def _update_widget_colors(self):
        if hasattr(self, 'gauge'): self.gauge.update_colors(self.palette)
        if hasattr(self, 'dpi_slider'): self.dpi_slider.update_colors(self.palette)
        for dz in self.drop_zones.values(): dz.update_colors(self.palette)

        preview_canvases = [getattr(self, name, None) for name in ['rotate_preview_canvas', 'delete_preview_canvas', 'stamp_preview_canvas']]
        for canvas in preview_canvases:
            if canvas:
                canvas.config(bg=self.palette.get("BG"), highlightbackground=self.palette.get("BORDER"))
        for toggle in self.toggles: toggle.update_colors(self.palette)
        for frame in self.tab_frames.values():
            if hasattr(frame, 'update_colors'): frame.update_colors(self.palette)

        if self.coffee_img_label:
            is_dark = self.general_settings.dark_mode_enabled.get()
            new_icon = self.coffee_icon_dark if is_dark else self.coffee_icon_light
            self.coffee_img_label.config(image=new_icon)
            self.coffee_img_label.image = new_icon

    def setup_drag_and_drop(self):
        if IS_WINDOWS:
            for key, dz in self.drop_zones.items():
                if key == 'compress': windnd.hook_dropfiles(dz, lambda files: self._on_drop(files, self.compress_settings.input_path))
                elif key == 'split': windnd.hook_dropfiles(dz, lambda files: self._on_drop(files, self.split_settings.input_path))
                elif key == 'rotate': windnd.hook_dropfiles(dz, lambda files: self._on_drop(files, self.rotate_settings.input_path))
                elif key == 'delete': windnd.hook_dropfiles(dz, lambda files: self._on_drop(files, self.delete_settings.input_path))
                elif key == 'password': windnd.hook_dropfiles(dz, lambda files: self._on_drop(files, self.password_settings.input_path))
                elif key == 'stamp': windnd.hook_dropfiles(dz, lambda files: self._on_drop(files, self.stamp_settings.input_path))
                elif key == 'page_number': windnd.hook_dropfiles(dz, lambda files: self._on_drop(files, self.page_number_settings.input_path))
                elif key == 'toc': windnd.hook_dropfiles(dz, lambda files: self._on_drop(files, self.toc_settings.input_path))
                elif key == 'convert': windnd.hook_dropfiles(dz, lambda files: self._on_drop(files, self.convert_settings.input_path))
                elif key == 'repair': windnd.hook_dropfiles(dz, lambda files: self._on_drop(files, self.repair_settings.input_path))

    def _on_drop(self, files, target_var):
        try:
            decoded_files = [f.decode('utf-8') for f in files]
            if decoded_files: target_var.set(decoded_files[0]); self.root.attributes('-topmost', 1); self.root.attributes('-topmost', 0)
        except Exception as e: logging.error(f"Drag and drop failed: {e}"); self.status.set(f"Error during drop: {e}")

