import tkinter as tk
from tkinter import ttk
from PIL import Image, ImageDraw, ImageTk
import sys

class ScrolledFrame(ttk.Frame):
    def __init__(self, parent, *args, **kw):
        super().__init__(parent, *args, **kw)
        self.canvas = tk.Canvas(self, borderwidth=0, highlightthickness=0)
        self.scrollbar = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        self.scrollbar.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollable_frame = ttk.Frame(self.canvas, padding=10, style="TFrame")
        self.canvas_window = self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")

        self.scrollable_frame.bind("<Configure>", self._on_frame_configure)
        self.canvas.bind("<Configure>", self._on_canvas_configure)

        if sys.platform == "win32" or sys.platform == "darwin":
            self.bind_all("<MouseWheel>", self._on_mousewheel, add='+')
            self.bind_all("<KeyPress-Prior>", lambda e: self._on_page_key("up"), add='+')
            self.bind_all("<KeyPress-Next>", lambda e: self._on_page_key("down"), add='+')
        else:
            self.bind_all("<Button-4>", self._on_mousewheel, add='+')
            self.bind_all("<Button-5>", self._on_mousewheel, add='+')
            self.bind_all("<KeyPress-Prior>", lambda e: self._on_page_key("up"), add='+')
            self.bind_all("<KeyPress-Next>", lambda e: self._on_page_key("down"), add='+')

        self.bind("<Destroy>", self._unbind_scroll)

    def _on_frame_configure(self, event=None):
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def _on_canvas_configure(self, event=None):
        self.canvas.itemconfig(self.canvas_window, width=event.width)

    def update_colors(self, palette):
        self.canvas.configure(bg=palette.get("BG"))

    def _is_mouse_over(self):
        try:
            x_root, y_root = self.winfo_pointerxy()
            widget_x = self.winfo_rootx()
            widget_y = self.winfo_rooty()
            widget_width = self.winfo_width()
            widget_height = self.winfo_height()

            return (widget_x <= x_root < widget_x + widget_width and
                    widget_y <= y_root < widget_y + widget_height)
        except tk.TclError:
            return False


    def _on_mousewheel(self, event):
        if not self._is_mouse_over():
             return

        if sys.platform == 'win32':
            delta = -1 * (event.delta // 120)
        elif sys.platform == 'darwin':
             delta = -1 * event.delta
        else:
            delta = -1 if event.num == 4 else 1

        y_top, y_bottom = self.canvas.yview()
        if (delta < 0 and y_top > 0) or \
           (delta > 0 and y_bottom < 1.0):
             self.canvas.yview_scroll(delta, "units")
             return "break"

    def _on_page_key(self, direction):
        if not self._is_mouse_over():
            return

        if direction == "up":
            self.canvas.yview_scroll(-1, "pages")
        else:
            self.canvas.yview_scroll(1, "pages")
        return "break"


    def _unbind_scroll(self, event=None):
        try:
            if sys.platform == "win32" or sys.platform == "darwin":
                self.unbind_all("<MouseWheel>")
                self.unbind_all("<KeyPress-Prior>")
                self.unbind_all("<KeyPress-Next>")
            else:
                self.unbind_all("<Button-4>")
                self.unbind_all("<Button-5>")
                self.unbind_all("<KeyPress-Prior>")
                self.unbind_all("<KeyPress-Next>")
        except tk.TclError:
            pass

class Tooltip:
    def __init__(self, widget, text, delay=400):
        self.widget = widget
        self.text = text
        self.delay = delay
        self.tooltip_window = None
        self.schedule_id = None
        
        self.widget.bind("<Enter>", self.schedule_tooltip)
        self.widget.bind("<Leave>", self.hide_tooltip)
        self.widget.bind("<ButtonPress>", self.hide_tooltip)

    def schedule_tooltip(self, event=None):
        self.unschedule_tooltip()
        self.schedule_id = self.widget.after(self.delay, self.show_tooltip)

    def unschedule_tooltip(self):
        if self.schedule_id:
            self.widget.after_cancel(self.schedule_id)
            self.schedule_id = None

    def show_tooltip(self, event=None):
        self.unschedule_tooltip()
        if not self.widget.winfo_exists() or not self.text: 
            return
        if self.tooltip_window: 
            return # Prevent duplicate tooltips from spawning
            
        x = self.widget.winfo_rootx() + 25
        y = self.widget.winfo_rooty() + 25
        
        self.tooltip_window = tk.Toplevel(self.widget)
        self.tooltip_window.wm_overrideredirect(True)
        self.tooltip_window.wm_geometry(f"+{x}+{y}")
        self.tooltip_window.attributes('-topmost', True) # Keep tooltip above other windows
        
        tk.Label(self.tooltip_window, text=self.text, justify='left', 
                 background="#383c40", foreground="#e8eaed", relief='solid', 
                 borderwidth=1, wraplength=250, padx=8, pady=5).pack(ipadx=1)

    def hide_tooltip(self, event=None):
        self.unschedule_tooltip()
        if self.tooltip_window: 
            self.tooltip_window.destroy()
            self.tooltip_window = None

class FileSelector(ttk.Frame):
    def __init__(self, parent, input_var, output_var, browse_output_cmd):
        super().__init__(parent, style="Card.TFrame")
        self.columnconfigure(1, weight=1)
        if input_var:
            ttk.Label(self, text="Input", style="Card.TLabel").grid(row=0, column=0, sticky="e", padx=(0, 10))
            in_entry = ttk.Entry(self, textvariable=input_var, state="readonly"); in_entry.grid(row=0, column=1, sticky="we"); Tooltip(in_entry, "Selected input file.")

        output_row = 1 if input_var else 0
        ttk.Label(self, text="Output", style="Card.TLabel").grid(row=output_row, column=0, sticky="e", padx=(0, 10), pady=(5 if input_var else 0, 0))
        out_entry = ttk.Entry(self, textvariable=output_var); out_entry.grid(row=output_row, column=1, sticky="we", pady=(5 if input_var else 0, 0)); Tooltip(out_entry, "Destination for the processed file(s).")
        browse_btn = ttk.Button(self, text="...", command=browse_output_cmd, width=4); browse_btn.grid(row=output_row, column=2, sticky="e", padx=(5,0), pady=(5 if input_var else 0, 0)); Tooltip(browse_btn, "Browse for output location.")

class ModernToggle(ttk.Checkbutton):
    def __init__(self, parent, text, variable, palette, command=None, **kwargs):
        self.palette = palette
        super().__init__(parent, text=text, variable=variable, command=command, **kwargs)

    def update_colors(self, palette):
        self.palette = palette
        # With standard ttk widgets, the global style changes handle the colors, 
        # but we keep this method so gui.py doesn't throw an error when switching themes.
        pass

class CompressionGauge(ttk.Frame):
    def __init__(self, parent, variable=None, palette=None, **kwargs):
        super().__init__(parent, style="Card.TFrame", **kwargs)
        self.variable = variable if variable else tk.DoubleVar(value=0.0)
        self.palette = palette if palette else {}
        self.text_id = None
        self.state = "normal"

        self.canvas = tk.Canvas(self, highlightthickness=0)
        self.canvas.pack(expand=True, fill="both")

        self.canvas.bind("<Configure>", self.draw_dial)
        self.variable.trace_add("write", self.draw_dial)

    def update_colors(self, palette):
        self.palette = palette
        self.canvas.configure(bg=self.palette.get("WIDGET_BG"))
        self.draw_dial()

    def set_state(self, state):
        self.state = state
        self.draw_dial()

    def draw_dial(self, *args):
        self.canvas.delete("all")
        width, height = self.canvas.winfo_width(), self.canvas.winfo_height()
        if width < 10 or height < 10: return

        radius = min(width, height) / 2.8
        center_x, center_y = width / 2, height / 2

        bg_color = self.palette.get("BORDER")
        value = self.variable.get()

        if value > 0:
            fg_color = self.palette.get("SUCCESS")
            text_label = "SAVED"
        elif value < 0:
            fg_color = self.palette.get("ERROR")
            text_label = "GAINED"
        else:
            fg_color = self.palette.get("ACCENT")
            text_label = "SAVED"

        text_color = self.palette.get("TEXT")

        if self.state == 'disabled':
            bg_color = self.palette.get("DISABLED_BG")
            fg_color = self.palette.get("DISABLED")
            text_color = self.palette.get("DISABLED")

        start_angle, end_angle = 225, -45

        self.canvas.create_arc(center_x-radius, center_y-radius, center_x+radius, center_y+radius,
                               start=start_angle, extent=(end_angle-start_angle),
                               outline=bg_color, width=12, style=tk.ARC)

        angle = self._value_to_angle(value)
        if value != 0:
            self.canvas.create_arc(center_x-radius, center_y-radius, center_x+radius, center_y+radius,
                                   start=start_angle, extent=(angle-start_angle),
                                   outline=fg_color, width=14, style=tk.ARC)

        self.text_id = self.canvas.create_text(center_x, center_y, text=f"{abs(value):.1f}%",
                                               font=("Segoe UI", 28, "bold"), fill=text_color)
        self.canvas.create_text(center_x, center_y + 30, text=text_label,
                                font=("Segoe UI", 10), fill=text_color)

    def _value_to_angle(self, value):
        start_angle, end_angle = 225, -45
        clamped_value = max(0.0, min(100.0, abs(value)))
        ratio = clamped_value / 100.0
        return start_angle + ratio * (end_angle - start_angle)

class CustomSlider(ttk.Frame):
    def __init__(self, parent, from_=0, to=100, variable=None, palette=None, **kwargs):
        super().__init__(parent, style="Card.TFrame", **kwargs)
        self.from_ = from_
        self.to = to
        self.variable = variable
        self.palette = palette if palette else {}
        self.state = "normal"
        self.thumb_radius = 8

        self.canvas = tk.Canvas(self, height=self.thumb_radius * 2 + 4, highlightthickness=0)
        self.canvas.pack(expand=True, fill="x")

        self.canvas.bind("<Configure>", self._draw_slider)
        self.canvas.bind("<Button-1>", self._handle_mouse_event)
        self.canvas.bind("<B1-Motion>", self._handle_mouse_event)
        self.variable.trace_add("write", self._draw_slider)

    def update_colors(self, palette):
        self.palette = palette
        self.canvas.configure(bg=self.palette.get("WIDGET_BG"))
        self._draw_slider()

    def config(self, state=None):
        if state in ("normal", "disabled"):
            self.state = state
            self._draw_slider()

    def _draw_slider(self, *args):
        self.canvas.delete("all")
        width = self.canvas.winfo_width()
        height = self.canvas.winfo_height()
        if width < 20: return

        track_y = height / 2
        track_start_x = self.thumb_radius + 2
        track_end_x = width - self.thumb_radius - 2
        track_width = track_end_x - track_start_x

        track_color = self.palette.get("BORDER")
        thumb_color = self.palette.get("ACCENT")
        if self.state == "disabled":
            track_color = self.palette.get("DISABLED_BG")
            thumb_color = self.palette.get("DISABLED")

        self.canvas.create_line(track_start_x, track_y, track_end_x, track_y, fill=track_color, width=4, capstyle='round')

        try:
            value = self.variable.get()
        except (ValueError, tk.TclError):
            value = self.from_

        value_ratio = (value - self.from_) / (self.to - self.from_)
        value_ratio = max(0, min(1, value_ratio))

        thumb_x = track_start_x + value_ratio * track_width

        self.canvas.create_oval(
            thumb_x - self.thumb_radius, track_y - self.thumb_radius,
            thumb_x + self.thumb_radius, track_y + self.thumb_radius,
            fill=thumb_color, outline=""
        )

    def _handle_mouse_event(self, event):
        if self.state == "disabled": return

        width = self.canvas.winfo_width()
        track_start_x = self.thumb_radius + 2
        track_end_x = width - self.thumb_radius - 2
        track_width = track_end_x - track_start_x

        x = max(track_start_x, min(event.x, track_end_x))

        value_ratio = (x - track_start_x) / track_width
        new_value = self.from_ + value_ratio * (self.to - self.from_)

        if self.variable.get() != int(new_value):
            self.variable.set(int(new_value))

class DropZone(ttk.Frame):
    def __init__(self, parent, browse_file_cmd, browse_folder_cmd, palette, **kwargs):
        self.palette = palette if palette else {}
        # We drop the explicit height/width kwargs and use padding/grid to size it naturally
        kwargs.pop('height', None)
        kwargs.pop('width', None)
        super().__init__(parent, style="Card.TFrame", padding=20, **kwargs)
        
        self.columnconfigure(0, weight=1)
        self.columnconfigure(1, weight=1)
        self.rowconfigure(0, weight=1)
        
        self.lbl = ttk.Label(self, text="Drag & Drop PDF File or Folder Here", 
                             font=("Segoe UI", 12, "italic"), foreground=self.palette.get("DISABLED"))
        self.lbl.grid(row=0, column=0, columnspan=2, pady=(0, 15))

        self.file_btn = ttk.Button(self, text="Select File(s)", command=browse_file_cmd, style="Large.Outline.TButton")
        self.file_btn.grid(row=1, column=0, sticky="e", padx=(0, 5))
        
        self.folder_btn = ttk.Button(self, text="Select Folder", command=browse_folder_cmd, style="Large.Outline.TButton")
        self.folder_btn.grid(row=1, column=1, sticky="w", padx=(5, 0))

    def update_colors(self, palette):
        self.palette = palette
        self.lbl.configure(foreground=self.palette.get("DISABLED"))

class PositionSelector(ttk.Frame):
    def __init__(self, parent, variable, positions, **kwargs):
        super().__init__(parent, style="Card.TFrame", **kwargs)
        self.variable = variable

        for i, pos in enumerate(positions):
            row, col = divmod(i, 3)
            rb = ttk.Radiobutton(self, text="", variable=self.variable, value=pos, style="Position.TRadiobutton", width=-5)
            rb.grid(row=row, column=col, sticky="nsew", padx=1, pady=1)
            self.rowconfigure(row, weight=1)
            self.columnconfigure(col, weight=1)