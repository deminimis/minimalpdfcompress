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
import pikepdf

def resource_path(relative_path):
    try:
        base_path = Path(sys._MEIPASS) if hasattr(sys, '_MEIPASS') else Path(__file__).parent
        return base_path / relative_path
    except Exception as e:
        logging.error(f"Error resolving resource path: {e}")
        return Path(relative_path)

def find_ghostscript():
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
    popup = tk.Toplevel(parent)
    popup.title("Ghostscript Required")
    popup.geometry("450x150")
    
    ttk.Label(popup, text="Ghostscript not found. Please place it in a local 'bin'/'lib' folder\nor install it system-wide, then restart the application.").pack(pady=10)
    link = ttk.Label(popup, text="Download Ghostscript", foreground="#0000EE", cursor="hand2")
    link.pack()
    link.bind("<Button-1>", lambda e: webbrowser.open("https://github.com/ArtifexSoftware/ghostpdl-downloads/releases/tag/gs10051"))
    ttk.Button(popup, text="OK", command=popup.destroy).pack(pady=10)

def build_gs_command(gs_path, input_path, output_path, operation, options):
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
    logging.info(f"Executing command: {command}")
    use_shell = isinstance(command, str)

    try:
        subprocess.run(
            command, check=True, capture_output=True, text=True,
            shell=use_shell, creationflags=subprocess.CREATE_NO_WINDOW
        )
    except subprocess.CalledProcessError as e:
        logging.error(f"Command failed: {e.stderr}")
        if "740" in e.stderr:
             raise Exception("Administrator privileges required.")
        raise Exception(f"Processing failed: {e.stderr}")
    except Exception as e:
        logging.error(f"An unexpected error occurred: {e}")
        raise

def apply_final_processing(file_path, options, use_compression):
    """
    Applies all selected pikepdf modifications in one go: rotation, metadata
    stripping, and final compression.
    """
    try:
        logging.info(f"Applying final processing to {file_path}")
        pdf = pikepdf.open(file_path, allow_overwriting_input=True)
        
        # 1. Apply Rotation
        angle = options.get('rotation', 0)
        if angle != 0:
            logging.info(f"Rotating all pages by {angle} degrees.")
            for page in pdf.pages:
                page.rotate(angle, relative=True)

        # 2. Strip Metadata
        if options.get('strip_metadata', False):
            logging.info("Stripping metadata.")
            # --- MODIFIED: More robust error handling for metadata removal ---
            try:
                del pdf.Info
            except Exception as e:
                logging.info(f"Could not remove Info dictionary (it may not have existed): {e}")
            try:
                del pdf.Root.Metadata
            except Exception as e:
                logging.info(f"Could not remove XMP Metadata (it may not have existed): {e}")
            # --- END MODIFICATION ---

        # 3. Save with optional compression
        save_kwargs = {}
        if use_compression:
            logging.info("Applying traditional compression.")
            save_kwargs['object_stream_mode'] = pikepdf.ObjectStreamMode.generate
            save_kwargs['compress_streams'] = True
        
        # Only save if changes were made
        if angle != 0 or options.get('strip_metadata', False) or use_compression:
            pdf.save(file_path, **save_kwargs)
            logging.info("Final processing successful.")
        else:
            logging.info("No final processing options selected. Skipping save.")

    except Exception as e:
        logging.error(f"pikepdf final processing failed: {e}")
        raise Exception(f"Final processing step failed: {e}")

def process_single(params):
    if not Path(params['input_path']).exists():
        raise FileNotFoundError("Input file does not exist.")
    
    gs_params = {k: v for k, v in params.items() if k not in ['use_final_compression']}
    command = build_gs_command(**gs_params)
    run_command(command)

def process_batch(params):
    input_folder = Path(params['input_path'])
    output_folder = Path(params['output_path'])

    if not input_folder.is_dir():
        raise FileNotFoundError("Input folder does not exist.")
    if not output_folder.is_dir():
        os.makedirs(output_folder, exist_ok=True)

    pdf_files = list(input_folder.rglob("*.pdf"))
    if not pdf_files:
        raise FileNotFoundError("No PDF files found in the specified folder.")

    output_files = []
    for pdf in pdf_files:
        out_name = f"{pdf.stem}_processed.pdf"
        output_file_path = output_folder / out_name
        
        gs_params = {
            'gs_path': params['gs_path'],
            'input_path': str(pdf),
            'output_path': str(output_file_path),
            'operation': params['operation'],
            'options': params['options']
        }
        
        command = build_gs_command(**gs_params)
        run_command(command)
        output_files.append(output_file_path)
        
    return len(pdf_files), output_files

def run_processing_task(params, is_folder, status_var, completion_callback):
    """A wrapper function to run the correct processing task and handle status updates."""
    status_var.set("Processing with Ghostscript...")
    try:
        options = params.get('options', {})
        use_final_compression = params.get('use_final_compression', False)
        final_processing_needed = use_final_compression or options.get('rotation', 0) != 0 or options.get('strip_metadata', False)

        if is_folder:
            processed_count, output_files = process_batch(params)
            
            if final_processing_needed:
                status_var.set(f"Applying final processing to {processed_count} files...")
                for i, file_path in enumerate(output_files):
                    status_var.set(f"Finalizing file {i+1}/{processed_count}...")
                    apply_final_processing(file_path, options, use_final_compression)
            
            status_var.set(f"Complete: Processed {processed_count} files.")
        else:
            process_single(params)
            
            if final_processing_needed:
                status_var.set("Applying final processing...")
                apply_final_processing(params['output_path'], options, use_final_compression)
                
            status_var.set(f"Complete: Processed {Path(params['input_path']).name}")
    except Exception as e:
        status_var.set(f"Error: {e}")
        messagebox.showerror("Error", str(e))
    finally:
        completion_callback()