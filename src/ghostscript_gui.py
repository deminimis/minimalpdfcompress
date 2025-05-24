import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import subprocess
import sys
from pathlib import Path
import logging
import winreg
import webbrowser
import tempfile

class GhostscriptGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Minimal PDF Compress")
        self.root.minsize(600, 500)  # Minimum size to ensure content fits

        # Set window icon (requires pdf.ico in the script directory)
        try:
            icon_path = self.resource_path("pdf.ico")
            if os.path.exists(icon_path):
                self.root.iconbitmap(icon_path)
        except Exception as e:
            logging.warning(f"Failed to set icon: {e}")

        # Main frame for centering content
        self.main_frame = ttk.Frame(self.root, padding="20")
        self.main_frame.grid(row=0, column=0, sticky="nsew")
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        self.main_frame.columnconfigure(0, weight=1)

        # Variables
        self.input_path = tk.StringVar()
        self.output_path = tk.StringVar()
        self.operation = tk.StringVar(value="Compress PDF")
        self.status = tk.StringVar(value="Ready")
        self.is_folder = False  # Track if input is a folder
        self.show_advanced = tk.BooleanVar(value=False)  # Track advanced options visibility
        self.resolution = tk.StringVar(value="150")  # Default resolution
        self.downscale_factor = tk.StringVar(value="1")  # Default downscaling factor
        self.pdfa_compression = tk.BooleanVar(value=False)  # PDF/A compression toggle
        self.color_strategy = tk.StringVar(value="LeaveColorUnchanged")  # Default color conversion
        self.downsample_type = tk.StringVar(value="Bicubic")  # Default downsample method
        self.fast_web_view = tk.BooleanVar(value=False)  # Fast web view toggle
        self.subset_fonts = tk.BooleanVar(value=True)  # Subset fonts toggle (default True)
        self.compress_fonts = tk.BooleanVar(value=True)  # Compress fonts toggle (default True)

        # Setup logging
        logging.basicConfig(filename="ghostscript_gui.log", level=logging.DEBUG,
                           format="%(asctime)s - %(levelname)s - %(message)s")

        # Build GUI
        self.build_gui()

        # Find Ghostscript executable
        self.gs_path = self.find_ghostscript()
        if not self.gs_path:
            self.show_ghostscript_download_popup()
            self.status.set("Error: Ghostscript not installed")

    def resource_path(self, relative_path):
        """Get absolute path to resource, works for dev and PyInstaller."""
        try:
            base_path = Path(sys._MEIPASS) if hasattr(sys, '_MEIPASS') else Path(__file__).parent
            logging.debug(f"Resolved base path for {relative_path}: {base_path}")
            return str(base_path / relative_path)
        except Exception as e:
            logging.error(f"Error resolving resource path for {relative_path}: {e}")
            return relative_path

    def find_ghostscript(self):
        """Locate the Ghostscript executable."""
        exe_name = "gswin64c.exe"
        possible_paths = [
            Path(r"C:\Program Files\gs\gs10.05.1\bin") / exe_name,
            Path(r"C:\Program Files (x86)\gs\gs10.05.1\bin") / exe_name,
            Path(r"C:\Program Files\gs\gs10.04.0\bin") / exe_name,
            Path(r"C:\Program Files (x86)\gs\gs10.04.0\bin") / exe_name,
            Path(self.resource_path(exe_name))
        ]

        for gs_path in possible_paths:
            logging.debug(f"Checking for Ghostscript at: {gs_path}")
            if gs_path.exists():
                logging.info(f"Ghostscript found at: {gs_path}")
                return str(gs_path)

        try:
            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Artifex\Ghostscript") as key:
                for i in range(winreg.QueryInfoKey(key)[0]):
                    version = winreg.EnumKey(key, i)
                    with winreg.OpenKey(key, version) as subkey:
                        gs_path = Path(winreg.QueryValueEx(subkey, "GS_DLL")[0]).parent / exe_name
                        logging.debug(f"Checking registry path: {gs_path}")
                        if gs_path.exists():
                            logging.info(f"Ghostscript found via registry at: {gs_path}")
                            return str(gs_path)
        except Exception as e:
            logging.debug(f"Registry check failed: {e}")

        logging.error(f"Ghostscript executable not found in any known locations")
        return None

    def show_ghostscript_download_popup(self):
        """Show popup with clickable link to download Ghostscript."""
        popup = tk.Toplevel(self.root)
        popup.title("Ghostscript Required")
        popup.geometry("400x150")
        popup.resizable(False, False)

        ttk.Label(popup, text="Ghostscript is not installed. Please download and install it:").pack(pady=10)

        link = ttk.Label(popup, text="https://github.com/ArtifexSoftware/ghostpdl-downloads/releases/tag/gs10051",
                         foreground="#0000EE", cursor="hand2")
        link.pack()
        link.bind("<Button-1>", lambda e: webbrowser.open("https://github.com/ArtifexSoftware/ghostpdl-downloads/releases/tag/gs10051"))

        ttk.Button(popup, text="OK", command=popup.destroy).pack(pady=10)

    def show_batch_confirmation_popup(self):
        """Show confirmation popup for batch processing."""
        popup = tk.Toplevel(self.root)
        popup.title("Confirm Batch Processing")
        popup.geometry("400x150")
        popup.resizable(False, False)

        ttk.Label(popup, text="This will process every .pdf file in the folder and its subfolders,\nwhich may take some time. Continue?",
                  justify="center").pack(pady=20)

        button_frame = ttk.Frame(popup)
        button_frame.pack(pady=10)
        ttk.Button(button_frame, text="Yes", command=lambda: [popup.destroy(), self.process_batch()]).pack(side="left", padx=10)
        ttk.Button(button_frame, text="No", command=lambda: [popup.destroy(), self.status.set("Ready")]).pack(side="left", padx=10)

    def toggle_advanced(self):
        """Toggle visibility of advanced options."""
        if self.show_advanced.get():
            self.advanced_frame.grid(row=3, column=0, sticky="nsew", pady=5)
        else:
            self.advanced_frame.grid_remove()

    def build_gui(self):
        """Build the GUI with advanced options section."""
        # Input Section
        input_frame = ttk.LabelFrame(self.main_frame, text="Input", padding=10)
        input_frame.grid(row=0, column=0, sticky="nsew", pady=5)
        input_frame.columnconfigure(1, weight=1)

        ttk.Button(input_frame, text="Input File or Folder", command=self.select_input).grid(row=0, column=0, padx=5, pady=5)
        ttk.Entry(input_frame, textvariable=self.input_path).grid(row=0, column=1, sticky="ew", padx=5, pady=5)

        # Output Section
        output_frame = ttk.LabelFrame(self.main_frame, text="Output", padding=10)
        output_frame.grid(row=1, column=0, sticky="nsew", pady=5)
        output_frame.columnconfigure(1, weight=1)

        ttk.Label(output_frame, text="Output Location:").grid(row=0, column=0, sticky="w", padx=5, pady=5)
        ttk.Entry(output_frame, textvariable=self.output_path).grid(row=0, column=1, sticky="ew", padx=5, pady=5)
        ttk.Button(output_frame, text="Browse", command=self.browse_output).grid(row=0, column=2, padx=5, pady=5)

        # Operation Section
        op_frame = ttk.LabelFrame(self.main_frame, text="Operation", padding=10)
        op_frame.grid(row=2, column=0, sticky="nsew", pady=5)
        op_frame.columnconfigure(1, weight=1)

        operations = ["Compress PDF", "Convert to PDF/A"]
        ttk.Combobox(op_frame, textvariable=self.operation, values=operations, state="readonly").grid(row=0, column=0, padx=5, pady=5, sticky="ew")
        ttk.Button(op_frame, text="Process", command=self.process).grid(row=0, column=1, padx=5, pady=5)
        ttk.Checkbutton(op_frame, text="Advanced Options", variable=self.show_advanced, command=self.toggle_advanced).grid(row=0, column=2, padx=5, pady=5)

        # Advanced Options Section
        self.advanced_frame = ttk.LabelFrame(self.main_frame, text="Advanced Options", padding=10)
        self.advanced_frame.columnconfigure(1, weight=1)

        # Resolution Dropdown
        ttk.Label(self.advanced_frame, text="Resolution (dpi):").grid(row=0, column=0, sticky="w", padx=5, pady=2)
        ttk.Combobox(self.advanced_frame, textvariable=self.resolution, values=["72", "150", "300"], state="readonly").grid(row=0, column=1, sticky="ew", padx=5, pady=2)

        # Downscaling Factor Dropdown
        ttk.Label(self.advanced_frame, text="Downscaling Factor:").grid(row=1, column=0, sticky="w", padx=5, pady=2)
        downscale_combobox = ttk.Combobox(self.advanced_frame, textvariable=self.downscale_factor, values=["1", "2", "3"], state="readonly")
        downscale_combobox.grid(row=1, column=1, sticky="ew", padx=5, pady=2)

        # Color Conversion Strategy Dropdown
        ttk.Label(self.advanced_frame, text="Color Conversion Strategy:").grid(row=2, column=0, sticky="w", padx=5, pady=2)
        color_combobox = ttk.Combobox(self.advanced_frame, textvariable=self.color_strategy, values=["LeaveColorUnchanged", "Gray", "RGB", "CMYK"], state="readonly")
        color_combobox.grid(row=2, column=1, sticky="ew", padx=5, pady=2)

        # Downsample Method Dropdown
        ttk.Label(self.advanced_frame, text="Downsample Method:").grid(row=3, column=0, sticky="w", padx=5, pady=2)
        ttk.Combobox(self.advanced_frame, textvariable=self.downsample_type, values=["Subsample", "Average", "Bicubic"], state="readonly").grid(row=3, column=1, sticky="ew", padx=5, pady=2)

        # Fast Web View Checkbox
        ttk.Checkbutton(self.advanced_frame, text="Enable Fast Web View", variable=self.fast_web_view).grid(row=4, column=0, columnspan=2, sticky="w", padx=5, pady=2)

        # Subset Fonts Checkbox
        ttk.Checkbutton(self.advanced_frame, text="Subset Fonts", variable=self.subset_fonts).grid(row=5, column=0, columnspan=2, sticky="w", padx=5, pady=2)

        # Compress Fonts Checkbox
        ttk.Checkbutton(self.advanced_frame, text="Compress Fonts", variable=self.compress_fonts).grid(row=6, column=0, columnspan=2, sticky="w", padx=5, pady=2)

        # PDF/A Compression Checkbox (only if Convert to PDF/A is selected)
        self.pdfa_compression_check = ttk.Checkbutton(self.advanced_frame, text="Compress PDF/A Output", variable=self.pdfa_compression)
        self.pdfa_compression_check.grid(row=7, column=0, columnspan=2, sticky="w", padx=5, pady=2)

        # Update visibility and enforce constraints based on operation
        def update_advanced_options(*args):
            if self.operation.get() == "Convert to PDF/A":
                self.pdfa_compression_check.grid(row=7, column=0, columnspan=2, sticky="w", padx=5, pady=2)
                # Limit downscaling factor for PDF/A
                if self.downscale_factor.get() > "2":
                    self.downscale_factor.set("2")
                    messagebox.showinfo("Info", "Downscaling factor limited to 2 for PDF/A to ensure quality.")
                # Enforce RGB for PDF/A
                if self.color_strategy.get() != "RGB":
                    self.color_strategy.set("RGB")
                    messagebox.showinfo("Info", "Color conversion strategy set to RGB for PDF/A compliance.")
            else:
                self.pdfa_compression_check.grid_remove()

        self.operation.trace("w", update_advanced_options)

        # Initially hide advanced frame
        self.advanced_frame.grid_remove()

        # Status Section
        status_frame = ttk.Frame(self.main_frame)
        status_frame.grid(row=8, column=0, sticky="nsew", pady=5)
        ttk.Label(status_frame, textvariable=self.status).pack()

    def select_input(self):
        """Allow user to select either a file or a folder."""
        dialog = tk.Toplevel(self.root)
        dialog.title("Select Input")
        dialog.geometry("300x150")
        dialog.resizable(False, False)

        ttk.Label(dialog, text="Choose input type:").pack(pady=10)
        ttk.Button(dialog, text="File", command=lambda: [dialog.destroy(), self.browse_file()]).pack(pady=5)
        ttk.Button(dialog, text="Folder", command=lambda: [dialog.destroy(), self.browse_folder()]).pack(pady=5)

    def browse_file(self):
        """Open file dialog for selecting a PDF."""
        file_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if file_path:
            self.input_path.set(file_path)
            self.is_folder = False
            output_path = Path(file_path).with_stem(Path(file_path).stem + "_out")
            self.output_path.set(str(output_path))

    def browse_folder(self):
        """Open folder dialog for batch processing."""
        folder_path = filedialog.askdirectory()
        if folder_path:
            self.input_path.set(folder_path)
            self.is_folder = True
            self.output_path.set(folder_path)

    def browse_output(self):
        """Open folder dialog for selecting output location."""
        if self.is_folder:
            folder_path = filedialog.askdirectory()
            if folder_path:
                self.output_path.set(folder_path)
        else:
            file_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
            if file_path:
                self.output_path.set(file_path)

    def process(self):
        """Determine whether to process a single file or a batch based on input."""
        if self.is_folder:
            self.show_batch_confirmation_popup()
        else:
            self.process_single()

    def run_ghostscript(self, input_path, output_path, operation):
        """Run Ghostscript command for a single file with advanced options."""
        if not self.gs_path:
            self.show_ghostscript_download_popup()
            self.status.set("Error: Ghostscript not installed")
            return False

        input_path = os.path.abspath(input_path)
        output_path = os.path.abspath(output_path)

        cmd = [self.gs_path, "-sDEVICE=pdfwrite", "-dCompatibilityLevel=1.4", "-dNOPAUSE", "-dBATCH", "-dQUIET", "-dSAFER"]
        
        # Add advanced options
        cmd.append(f"-r{self.resolution.get()}")
        cmd.append(f"-dDownScaleFactor={self.downscale_factor.get()}")

        if self.fast_web_view.get():
            cmd.append("-dFastWebView=true")
        cmd.append(f"-dSubsetFonts={'true' if self.subset_fonts.get() else 'false'}")
        cmd.append(f"-dCompressFonts={'true' if self.compress_fonts.get() else 'false'}")

        if operation == "Compress PDF":
            cmd.extend(["-dPDFSETTINGS=/screen"])
            cmd.append(f"-sColorConversionStrategy={self.color_strategy.get()}")
        elif operation == "Convert to PDF/A":
            cmd.extend(["-dPDFA=1", "-dPDFACompatibilityPolicy=1", "-sColorConversionStrategy=RGB"])
            cmd.append("-sOutputICCProfile=srgb.icc")
            if self.pdfa_compression.get():
                cmd.extend(["-dPDFSETTINGS=/screen"])

        cmd.extend([f"-sOutputFile={output_path}", input_path])
        cmd_str = " ".join(f'"{arg}"' if " " in arg else arg for arg in cmd)
        logging.info(f"Executing Ghostscript command: {cmd_str}")

        try:
            result = subprocess.run(cmd_str, check=True, capture_output=True, text=True, shell=True, timeout=60)
            logging.info(f"Ghostscript success: {result.stdout}")
            logging.debug(f"Ghostscript stderr: {result.stderr}")
            return True
        except subprocess.TimeoutExpired:
            logging.error("Ghostscript timed out after 60 seconds")
            messagebox.showerror("Error", "Ghostscript took too long to process (timeout after 60 seconds).")
            self.status.set("Error: Timeout")
            return False
        except subprocess.CalledProcessError as e:
            logging.error(f"Ghostscript failed: {e.stderr}")
            messagebox.showerror("Error", f"Ghostscript failed: {e.stderr}")
            self.status.set("Error: Processing failed")
            return False
        except FileNotFoundError as e:
            logging.error(f"Ghostscript executable not found during execution: {self.gs_path}")
            self.show_ghostscript_download_popup()
            self.status.set("Error: Ghostscript not found")
            return False
        except Exception as e:
            logging.error(f"Unexpected error during Ghostscript execution: {str(e)}")
            if "740" in str(e):
                messagebox.showerror("Error", "Operation requires administrator privileges. Please accept the UAC prompt or run as administrator.")
                self.status.set("Error: Requires elevation")
            else:
                messagebox.showerror("Error", f"Unexpected error: {str(e)}")
                self.status.set("Error: Processing failed")
            return False

    def process_single(self):
        """Process a single PDF file."""
        input_path = self.input_path.get()
        output_path = self.output_path.get()
        operation = self.operation.get()

        if not input_path or not output_path:
            messagebox.showwarning("Warning", "Please select input and output locations.")
            self.status.set("Error: Missing input/output")
            logging.warning("Missing input or output location")
            return

        if not os.path.exists(input_path):
            messagebox.showerror("Error", "Input file does not exist.")
            self.status.set("Error: Input file not found")
            logging.error(f"Input file not found: {input_path}")
            return

        self.status.set("Processing... (Accept UAC prompt)")
        self.root.update()

        if self.run_ghostscript(input_path, output_path, operation):
            self.status.set(f"Complete: Processed {Path(input_path).name}")
            logging.info(f"Successfully processed: {input_path}")
        else:
            self.status.set("Error: Processing failed")

    def process_batch(self):
        """Process all PDFs in a folder and its subfolders."""
        input_folder = self.input_path.get()
        output_folder = self.output_path.get()
        operation = self.operation.get()

        if not input_folder:
            messagebox.showwarning("Warning", "Please select an input folder.")
            self.status.set("Error: Missing input folder")
            logging.warning("Missing input folder")
            return

        if not output_folder:
            messagebox.showwarning("Warning", "Please select an output folder.")
            self.status.set("Error: Missing output folder")
            logging.warning("Missing output folder")
            return

        if not os.path.isdir(input_folder):
            messagebox.showerror("Error", "Input folder does not exist.")
            self.status.set("Error: Input folder not found")
            logging.error(f"Input folder not found: {input_folder}")
            return

        if not os.path.isdir(output_folder):
            messagebox.showerror("Error", "Output folder does not exist.")
            self.status.set("Error: Output folder not found")
            logging.error(f"Output folder not found: {output_folder}")
            return

        # Collect all PDFs in folder and subfolders
        pdf_files = []
        for root, _, files in os.walk(input_folder):
            for file in files:
                if file.lower().endswith(".pdf"):
                    pdf_files.append(os.path.join(root, file))

        if not pdf_files:
            messagebox.showwarning("Warning", "No PDF files found in the folder or subfolders.")
            self.status.set("Error: No PDFs found")
            logging.warning(f"No PDF files found in: {input_folder}")
            return

        self.status.set("Processing batch... (Accept UAC prompt)")
        self.root.update()

        # Create a temporary batch file to run all commands with one UAC prompt
        commands = []
        for pdf in pdf_files:
            relative_path = os.path.relpath(pdf, input_folder)
            output_path = os.path.join(output_folder, f"out_{Path(relative_path).stem}_{operation.replace(' ', '_').lower()}.pdf")
            os.makedirs(os.path.dirname(output_path), exist_ok=True)
            cmd = [self.gs_path, "-sDEVICE=pdfwrite", "-dCompatibilityLevel=1.4", "-dNOPAUSE", "-dBATCH", "-dQUIET", "-dSAFER"]
            cmd.append(f"-r{self.resolution.get()}")
            cmd.append(f"-dDownScaleFactor={self.downscale_factor.get()}")
            if self.fast_web_view.get():
                cmd.append("-dFastWebView=true")
            cmd.append(f"-dSubsetFonts={'true' if self.subset_fonts.get() else 'false'}")
            cmd.append(f"-dCompressFonts={'true' if self.compress_fonts.get() else 'false'}")
            if operation == "Compress PDF":
                cmd.extend(["-dPDFSETTINGS=/screen"])
                cmd.append(f"-sColorConversionStrategy={self.color_strategy.get()}")
            elif operation == "Convert to PDF/A":
                cmd.extend(["-dPDFA=1", "-dPDFACompatibilityPolicy=1", "-sColorConversionStrategy=RGB"])
                cmd.append("-sOutputICCProfile=srgb.icc")
                if self.pdfa_compression.get():
                    cmd.extend(["-dPDFSETTINGS=/screen"])
            cmd.extend([f"-sOutputFile={output_path}", pdf])
            cmd_str = " ".join(f'"{arg}"' if " " in arg else arg for arg in cmd)
            commands.append(cmd_str)

        success_count = 0
        with tempfile.NamedTemporaryFile(mode='w', suffix='.bat', delete=False) as batch_file:
            batch_file.write("@echo off\n")
            for cmd in commands:
                batch_file.write(f"{cmd}\n")
                batch_file.write("if %ERRORLEVEL% NEQ 0 exit /b %ERRORLEVEL%\n")
            batch_file_path = batch_file.name

        try:
            result = subprocess.run(f'"{batch_file_path}"', check=True, capture_output=True, text=True, shell=True, timeout=60 * len(pdf_files))
            logging.info(f"Batch Ghostscript success: {result.stdout}")
            logging.debug(f"Batch Ghostscript stderr: {result.stderr}")
            success_count = len(pdf_files)
        except subprocess.TimeoutExpired:
            logging.error("Batch Ghostscript timed out")
            messagebox.showerror("Error", "Batch processing took too long (timeout).")
            self.status.set("Error: Timeout")
        except subprocess.CalledProcessError as e:
            logging.error(f"Batch Ghostscript failed: {e.stderr}")
            messagebox.showerror("Error", f"Batch processing failed: {e.stderr}")
            self.status.set("Error: Batch processing failed")
        except Exception as e:
            logging.error(f"Unexpected error during batch Ghostscript execution: {str(e)}")
            if "740" in str(e):
                messagebox.showerror("Error", "Operation requires administrator privileges. Please accept the UAC prompt or run as administrator.")
                self.status.set("Error: Requires elevation")
            else:
                messagebox.showerror("Error", f"Unexpected error: {str(e)}")
                self.status.set("Error: Batch processing failed")
        finally:
            try:
                os.unlink(batch_file_path)
            except Exception as e:
                logging.warning(f"Failed to delete temporary batch file: {e}")

        if success_count > 0:
            self.status.set(f"Complete: {success_count}/{len(pdf_files)} files processed")
            logging.info(f"Batch processing complete: {success_count}/{len(pdf_files)} files")
        else:
            self.status.set("Error: Batch processing failed")

if __name__ == "__main__":
    root = tk.Tk()
    app = GhostscriptGUI(root)
    root.mainloop()
