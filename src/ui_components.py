# ui_components.py
import tkinter as tk
from tkinter import ttk
from PIL import Image, ImageDraw, ImageTk

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
        self.scrollable_frame.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.bind("<Configure>", lambda e: self.canvas.itemconfig(self.canvas_window, width=e.width))
        self.scrollable_frame.bind("<Enter>", self._bind_mouse)
        self.scrollable_frame.bind("<Leave>", self._unbind_mouse)

    def update_colors(self, palette):
        self.canvas.configure(bg=palette.get("BG"))

    def _bind_mouse(self, e): self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)
    def _unbind_mouse(self, e): self.canvas.unbind_all("<MouseWheel>")
    def _on_mousewheel(self, e): self.canvas.yview_scroll(int(-1*(e.delta/120)), "units")

class Tooltip:
    def __init__(self, widget, text):
        self.widget, self.text, self.tooltip_window = widget, text, None
        self.widget.bind("<Enter>", self.show_tooltip)
        self.widget.bind("<Leave>", self.hide_tooltip)
    def show_tooltip(self, event=None):
        if not self.widget.winfo_exists() or not self.text: return
        x = self.widget.winfo_rootx() + 25
        y = self.widget.winfo_rooty() + 25
        self.tooltip_window = tk.Toplevel(self.widget); self.tooltip_window.wm_overrideredirect(True); self.tooltip_window.wm_geometry(f"+{x}+{y}")
        tk.Label(self.tooltip_window, text=self.text, justify='left', background="#383c40", foreground="#e8eaed", relief='solid', borderwidth=1, wraplength=250, padx=8, pady=5).pack(ipadx=1)
    def hide_tooltip(self, event=None):
        if self.tooltip_window: self.tooltip_window.destroy(); self.tooltip_window = None

class FileSelector(ttk.Frame):
    def __init__(self, parent, input_var, output_var, browse_output_cmd):
        super().__init__(parent, style="Card.TFrame")
        self.columnconfigure(1, weight=1)
        if input_var:
            ttk.Label(self, text="Input", style="Card.TLabel").grid(row=0, column=0, sticky="e", padx=(0, 10))
            in_entry = ttk.Entry(self, textvariable=input_var, state="readonly"); in_entry.grid(row=0, column=1, sticky="we"); Tooltip(in_entry, "Selected input file or folder.")
        
        ttk.Label(self, text="Output", style="Card.TLabel").grid(row=1, column=0, sticky="e", padx=(0, 10), pady=(5,0))
        out_entry = ttk.Entry(self, textvariable=output_var); out_entry.grid(row=1, column=1, sticky="we", pady=(5,0)); Tooltip(out_entry, "Destination for the processed file(s).")
        browse_btn = ttk.Button(self, text="...", command=browse_output_cmd, width=4); browse_btn.grid(row=1, column=2, sticky="e", padx=(5,0), pady=(5,0)); Tooltip(browse_btn, "Browse for output location.")
        self.grid(row=1, column=0, sticky="nsew", pady=5)

class ModernToggle(ttk.Frame):
    def __init__(self, parent, text, variable, palette, command=None, **kwargs):
        super().__init__(parent, style="Card.TFrame", **kwargs)
        self.variable = variable
        self.command = command
        self.palette = palette
        self._image_cache = None
        self.state = 'normal'
        self.variable.trace_add("write", self._update_toggle)
        self.label = ttk.Label(self, text=text, style="Card.TLabel")
        self.label.pack(side="left", padx=(0, 10))
        self.canvas = tk.Canvas(self, width=50, height=26, highlightthickness=0)
        self.canvas.pack(side="left")
        self.canvas.bind("<Button-1>", self._toggle)
        self.update_colors(palette)

    def configure(self, cnf=None, **kw):
        if 'state' in kw:
            new_state = kw.pop('state')
            if new_state in ('normal', 'disabled'):
                self.state = new_state
                self._update_toggle()
        if cnf or kw:
            super().configure(cnf, **kw)
    config = configure

    def update_colors(self, palette):
        self.palette = palette
        self.configure(style="Card.TFrame")
        self.label.configure(style="Card.TLabel")
        self.canvas.configure(bg=self.palette.get("WIDGET_BG"))
        self._update_toggle()

    def _update_toggle(self, *args):
        scale = 4
        width, height = 50 * scale, 26 * scale
        radius = height / 2
        img = Image.new("RGBA", (width, height), (0, 0, 0, 0))
        draw = ImageDraw.Draw(img)
        colors = self.palette

        if self.state == 'disabled':
            bg_color = colors.get("DISABLED_BG")
            handle_color = colors.get("DISABLED")
        else:
            bg_color = colors.get("ACCENT") if self.variable.get() else colors.get("BORDER")
            handle_color = "#ffffff"

        draw.rounded_rectangle((0, 0, width, height), radius=radius, fill=bg_color)
        
        padding = 2 * scale
        handle_diameter = height - 2 * padding
        if self.variable.get():
            x0 = width - padding - handle_diameter
        else:
            x0 = padding
        y0 = padding
        x1 = x0 + handle_diameter
        y1 = y0 + handle_diameter
        draw.ellipse((x0, y0, x1, y1), fill=handle_color)

        resample_method = Image.Resampling.LANCZOS if hasattr(Image, 'Resampling') else Image.ANTIALIAS
        resized_img = img.resize((50, 26), resample_method)

        self.canvas.delete("all")
        self._image_cache = ImageTk.PhotoImage(resized_img)
        self.canvas.create_image(0, 0, anchor="nw", image=self._image_cache)

    def _toggle(self, event=None):
        if self.state == 'disabled':
            return
        self.variable.set(not self.variable.get())
        if self.command: self.command()

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
        fg_color = self.palette.get("SUCCESS") if value > 0 else self.palette.get("ACCENT")
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
        if value > 0:
            self.canvas.create_arc(center_x-radius, center_y-radius, center_x+radius, center_y+radius,
                                   start=start_angle, extent=(angle-start_angle),
                                   outline=fg_color, width=14, style=tk.ARC)
        
        self.text_id = self.canvas.create_text(center_x, center_y, text=f"{value:.1f}%", 
                                               font=("Segoe UI", 28, "bold"), fill=text_color)
        self.canvas.create_text(center_x, center_y + 30, text="SAVED", 
                                font=("Segoe UI", 10), fill=text_color)

    def _value_to_angle(self, value):
        start_angle, end_angle = 225, -45
        value = max(0.0, min(100.0, value))
        ratio = value / 100.0
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

class DropZone(tk.Canvas):
    def __init__(self, parent, browse_file_cmd, browse_folder_cmd, palette, **kwargs):
        self.palette = palette if palette else {}
        super().__init__(parent, highlightthickness=0, bg=self.palette.get("WIDGET_BG"), **kwargs)
        self.browse_file_cmd = browse_file_cmd; self.browse_folder_cmd = browse_folder_cmd
        self.is_hovering = False
        self.bind("<Configure>", self.draw)
    def update_colors(self, palette): 
        self.palette = palette; self.configure(bg=self.palette.get("WIDGET_BG")); self.draw()
    def draw(self, event=None):
        self.delete("all"); width, height = self.winfo_width(), self.winfo_height()
        if width < 10 or height < 10: return
        border = self.palette.get("ACCENT") if self.is_hovering else self.palette.get("BORDER")
        self.create_rectangle(2, 2, width-2, height-2, outline=border, width=2, dash=(6, 4))
        self.create_text(width/2, height/2-20, text="Drag & Drop PDF File or Folder Here", font=("Segoe UI", 12, "italic"), fill=self.palette.get("DISABLED"))
        if not hasattr(self, 'file_btn'):
            self.file_btn = ttk.Button(self, text="Select File", command=self.browse_file_cmd, style="Large.Outline.TButton")
            self.folder_btn = ttk.Button(self, text="Select Folder", command=self.browse_folder_cmd, style="Large.Outline.TButton")
        self.create_window(width/2 - 75, height/2 + 25, window=self.file_btn)
        self.create_window(width/2 + 75, height/2 + 25, window=self.folder_btn)

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