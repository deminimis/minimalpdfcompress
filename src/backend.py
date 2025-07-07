import os
import sys
import subprocess
import tempfile
import logging
import winreg
import webbrowser
from pathlib import Path
import tkinter as tk
from tkinter import ttk, messagebox

def resource_path(relative_path):
    """Get absolute path to resource, works for dev and PyInstaller."""
    try:
        base_path = Path(sys._MEIPASS) if hasattr(sys, '_MEIPASS') else Path(__file__).parent
        return base_path / relative_path
    except Exception as e:
        logging.error(f"Error resolving resource path: {e}")
        return Path(relative_path)

def find_ghostscript():
    # This function remains the same
    exe_name = "gswin64c.exe"
    app_root = resource_path('.')
    
    portable_exe_path = app_root / "bin" / exe_name
    portable_lib_path = app_root / "lib"
    if portable_exe_path.exists() and portable_lib_path.exists():
        logging.info(f"Found portable Ghostscript at: {portable_exe_path}")
        return str(portable_exe_path)

    for pf in [os.getenv("ProgramFiles"), os.getenv("ProgramFiles(x86)")]:
        if pf:
            for version_dir in ["gs10.05.1", "gs10.04.0", "gs"]:
                gs_path = Path(pf) / "gs" / version_dir / "bin" / exe_name
                if gs_path.exists():
                    logging.info(f"Found installed Ghostscript at: {gs_path}")
                    return str(gs_path)

    try:
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Artifex\Ghostscript") as key:
            for i in range(winreg.QueryInfoKey(key)[0]):
                version = winreg.EnumKey(key, i)
                with winreg.OpenKey(key, version) as subkey:
                    gs_dll_path = winreg.QueryValueEx(subkey, "GS_DLL")[0]
                    gs_path = Path(gs_dll_path).parent.parent / "bin" / exe_name
                    if gs_path.exists():
                        logging.info(f"Found Ghostscript via registry: {gs_path}")
                        return str(gs_path)
    except Exception:
        logging.debug("Ghostscript not found in registry.")

    logging.error("Ghostscript executable not found in any standard location.")
    return None

def show_ghostscript_download_popup(parent):
    # This function remains the same
    popup = tk.Toplevel(parent)
    popup.title("Ghostscript Required")
    popup.geometry("450x150")
    
    ttk.Label(popup, text="Ghostscript not found. Please place it in a local 'bin'/'lib' folder\nor install it system-wide, then restart the application.").pack(pady=10)
    link = ttk.Label(popup, text="Download Ghostscript", foreground="#0000EE", cursor="hand2")
    link.pack()
    link.bind("<Button-1>", lambda e: webbrowser.open("https://github.com/ArtifexSoftware/ghostpdl-downloads/releases/tag/gs10051"))
    ttk.Button(popup, text="OK", command=popup.destroy).pack(pady=10)

def build_gs_command(gs_path, input_path, output_path, operation, options):
    # This function remains the same
    cmd = [
        gs_path, "-sDEVICE=pdfwrite", "-dCompatibilityLevel=1.4", "-dNOPAUSE",
        "-dBATCH", "-dQUIET", "-dSAFER", f"-r{options['resolution']}",
        f"-dDownScaleFactor={options['downscale_factor']}",
        f"-dDownsampleType=/{options['downsample_type']}",
        f"-dSubsetFonts={'true' if options['subset_fonts'] else 'false'}",
        f"-dCompressFonts={'true' if options['compress_fonts'] else 'false'}"
    ]

    if options['fast_web_view']:
        cmd.append("-dFastWebView=true")
    
    preset_map = {
        "Compress (Screen - Smallest Size)": "/screen",
        "Compress (Ebook - Medium Size)": "/ebook",
        "Compress (Printer - High Quality)": "/printer",
        "Compress (Prepress - Highest Quality)": "/prepress"
    }
    
    if operation in preset_map:
        pdf_setting = preset_map[operation]
        cmd.extend([f"-dPDFSETTINGS={pdf_setting}", f"-sColorConversionStrategy={options['color_strategy']}"])
    
    elif operation == "Convert to PDF/A":
        srgb_profile = Path(gs_path).parent.parent / "lib" / "srgb.icc"
        if not srgb_profile.exists():
            raise FileNotFoundError(f"Could not find 'srgb.icc' in expected lib folder: {srgb_profile.parent}")
        
        cmd.extend([
            "-dPDFA=1", "-dPDFACompatibilityPolicy=1", "-sColorConversionStrategy=RGB",
            f"-sOutputICCProfile={str(srgb_profile)}"
        ])
        if options['pdfa_compression']:
            cmd.append("-dPDFSETTINGS=/screen")
    
    cmd.extend([f"-sOutputFile={output_path}", input_path])
    return cmd

def run_command(command):
    """Executes a command-line process with no timeout."""
    logging.info(f"Executing command: {command}")
    use_shell = isinstance(command, str)

    try:
        # The 'timeout' parameter has been removed from this call
        subprocess.run(
            command, check=True, capture_output=True, text=True,
            shell=use_shell, creationflags=subprocess.CREATE_NO_WINDOW
        )
    except subprocess.CalledProcessError as e:
        logging.error(f"Ghostscript failed: {e.stderr}")
        if "740" in e.stderr:
             raise Exception("Administrator privileges required.")
        raise Exception(f"Processing failed: {e.stderr}")
    except Exception as e:
        logging.error(f"An unexpected error occurred: {e}")
        raise

def process_single(params):
    """Process a single PDF file."""
    if not Path(params['input_path']).exists():
        raise FileNotFoundError("Input file does not exist.")
    
    command = build_gs_command(**params)
    run_command(command)

def process_batch(params):
    """Process all PDFs in a folder and its subfolders."""
    input_folder = Path(params['input_path'])
    output_folder = Path(params['output_path'])

    if not input_folder.is_dir():
        raise FileNotFoundError("Input folder does not exist.")
    if not output_folder.is_dir():
        os.makedirs(output_folder, exist_ok=True)

    pdf_files = list(input_folder.rglob("*.pdf"))
    if not pdf_files:
        raise FileNotFoundError("No PDF files found in the specified folder.")

    commands_to_write = []
    for pdf in pdf_files:
        out_name = f"{pdf.stem}_processed.pdf"
        batch_params = params.copy()
        batch_params['input_path'] = str(pdf)
        batch_params['output_path'] = str(output_folder / out_name)
        
        cmd_list = build_gs_command(**batch_params)
        cmd_str = " ".join(f'"{arg}"' if " " in str(arg) else str(arg) for arg in cmd_list)
        commands_to_write.append(cmd_str)

    with tempfile.NamedTemporaryFile(mode='w', suffix='.bat', delete=False, encoding='utf-8') as f:
        f.write("@echo off\n")
        for cmd in commands_to_write:
            f.write(f"{cmd}\n")
            f.write("if %ERRORLEVEL% NEQ 0 exit /b %ERRORLEVEL%\n")
        batch_file_path = f.name
    
    try:
        # Timeout removed from this call as well
        run_command(batch_file_path)
        return len(pdf_files)
    finally:
        os.unlink(batch_file_path)

def run_processing_task(params, is_folder, status_var, completion_callback):
    """A wrapper function to run the correct processing task and handle status updates."""
    status_var.set("Processing... (This may take a while)")
    try:
        if is_folder:
            processed_count = process_batch(params)
            status_var.set(f"Complete: Processed {processed_count} files.")
        else:
            process_single(params)
            status_var.set(f"Complete: Processed {Path(params['input_path']).name}")
    except Exception as e:
        status_var.set(f"Error: {e}")
        messagebox.showerror("Error", str(e))
    finally:
        completion_callback()