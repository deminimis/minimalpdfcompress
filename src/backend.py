import os
import sys
import subprocess
import logging
import shutil
import tempfile
from pathlib import Path
import pikepdf
import tkinter as tk
from tkinter import messagebox

if sys.platform == "win32":
    import winreg

#region: Exceptions and Helpers
class GhostscriptNotFound(Exception):
    pass

class AdminPrivilegesError(Exception):
    pass

class ProcessingError(Exception):
    pass

def resource_path(relative_path):
    try:
        base_path = Path(sys._MEIPASS)
    except AttributeError:
        base_path = Path(__file__).parent
    return base_path / relative_path
#endregion

#region: Ghostscript Handling
def find_ghostscript():
    if sys.platform == "win32":
        exe_name = "gswin64c.exe"
        app_root = resource_path('.')
        
        portable_exe_path = app_root / "bin" / exe_name
        if portable_exe_path.exists():
            logging.info(f"Found portable Ghostscript: {portable_exe_path}")
            return str(portable_exe_path)

        program_files_paths = [os.getenv("ProgramFiles"), os.getenv("ProgramFiles(x86)")]
        for pf_path in filter(None, program_files_paths):
            gs_base_dir = Path(pf_path) / "gs"
            if gs_base_dir.is_dir():
                for version_dir in gs_base_dir.glob('gs*'):
                    gs_exe = version_dir / "bin" / exe_name
                    if gs_exe.exists():
                        logging.info(f"Found installed Ghostscript: {gs_exe}")
                        return str(gs_exe)
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
        except FileNotFoundError:
            logging.debug("Ghostscript not found in registry.")
    else:
        exe_name = "gs"
        gs_path = shutil.which(exe_name)
        if gs_path:
            logging.info(f"Found Ghostscript in PATH: {gs_path}")
            return str(gs_path)

    raise GhostscriptNotFound(f"Ghostscript executable ('{exe_name}') not found.")

def build_gs_command(gs_path, input_path, output_path, operation, options):
    cmd = [
        gs_path, "-sDEVICE=pdfwrite", "-dCompatibilityLevel=1.4", "-dNOPAUSE",
        "-dBATCH", "-dQUIET", "-dSAFER",
        f"-dDownScaleFactor={options['downscale_factor']}",
        f"-dDownsampleType=/{options['downsample_type']}",
        f"-dSubsetFonts={'true' if options['subset_fonts'] else 'false'}",
        f"-dCompressFonts={'true' if options['compress_fonts'] else 'false'}"
    ]

    try:
        img_res = int(options['image_resolution'])
        cmd.extend([
            f"-dColorImageResolution={img_res}",
            f"-dGrayImageResolution={img_res}"
        ])
    except (ValueError, TypeError):
        logging.warning(f"Invalid image resolution value: {options['image_resolution']}. Using defaults.")

    if options['fast_web_view']:
        cmd.append("-dFastWebView=true")
    
    if options.get('remove_interactive', False):
        cmd.extend(["-dShowAnnots=false", "-dShowAcroForm=false"])

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
        srgb_profile_path = Path(gs_path).parent.parent / "lib" / "srgb.icc"
        if sys.platform != "win32":
            srgb_profile_path = Path(gs_path).parent.parent / "share" / "ghostscript" / "iccprofiles" / "srgb.icc"

        if not srgb_profile_path.exists():
            raise FileNotFoundError(f"Could not find 'srgb.icc' in expected location: {srgb_profile_path}")
        
        cmd.extend([
            "-dPDFA=1", "-dPDFACompatibilityPolicy=1", "-sColorConversionStrategy=RGB",
            f"-sOutputICCProfile={str(srgb_profile_path)}"
        ])
        if options['pdfa_compression']:
            cmd.append("-dPDFSETTINGS=/screen")
    
    cmd.extend([f"-sOutputFile={output_path}", input_path])
    return cmd

def run_command(command):
    logging.info(f"Executing command: {' '.join(command)}")
    try:
        kwargs = {
            'stdin': subprocess.DEVNULL, 'check': True, 'capture_output': True, 'text': True,
            'encoding': 'utf-8', 'errors': 'ignore'
        }
        if sys.platform == "win32":
            kwargs['creationflags'] = subprocess.CREATE_NO_WINDOW
        
        subprocess.run(command, **kwargs)

    except subprocess.CalledProcessError as e:
        logging.error(f"Command failed: {e.stderr}")
        if sys.platform == "win32" and "740" in e.stderr:
            raise AdminPrivilegesError("Administrator privileges required to run Ghostscript.")
        raise ProcessingError(f"Ghostscript failed: {e.stderr[:200]}...")
    except FileNotFoundError as e:
        raise ProcessingError(f"Command not found: {e}. Is Ghostscript installed and in PATH?")
    except Exception as e:
        logging.error(f"An unexpected error occurred during command execution: {e}")
        raise ProcessingError(str(e))
#endregion

#region: Final Processing
def apply_final_processing(file_path, options, pikepdf_compression_level):
    try:
        logging.info(f"Applying final processing to {file_path}")
        pdf = pikepdf.open(file_path, allow_overwriting_input=True)
        
        angle = options.get('rotation', 0)
        if angle != 0:
            logging.info(f"Rotating all pages by {angle} degrees.")
            for page in pdf.pages:
                page.rotate(angle, relative=True)

        if options.get('strip_metadata', False):
            logging.info("Stripping metadata.")
            try:
                del pdf.docinfo
            except Exception as e:
                logging.info(f"Could not remove Info dictionary (it may not have existed): {e}")
            try:
                del pdf.Root.Metadata
            except Exception as e:
                logging.info(f"Could not remove XMP Metadata (it may not have existed): {e}")

        save_kwargs = {}
        use_compression = pikepdf_compression_level > 0
        if use_compression:
            logging.info("Applying pikepdf object stream compression.")
            save_kwargs['object_stream_mode'] = pikepdf.ObjectStreamMode.generate
            save_kwargs['compress_streams'] = True
        
        if angle != 0 or options.get('strip_metadata', False) or use_compression:
            pdf.save(file_path, **save_kwargs)
            logging.info("Final processing successful.")
        else:
            logging.info("No final processing options selected. Skipping save.")

    except Exception as e:
        logging.error(f"pikepdf final processing failed: {e}")
        raise ProcessingError(f"Final processing step failed: {e}")
#endregion

#region: Main Task Runner
def run_processing_task(params, is_folder, status_var, completion_callback):
    try:
        gs_path = params['gs_path']
        options = params.get('options', {})
        overwrite = params.get('overwrite', False)
        
        pikepdf_compression_level = options.get('pikepdf_compression_level', 0)
        pikepdf.settings.set_flate_compression_level(pikepdf_compression_level)

        decimal_precision_str = options.get('decimal_precision', "Default")
        if decimal_precision_str != "Default":
            try:
                pikepdf.settings.set_decimal_precision(int(decimal_precision_str))
            except (ValueError, TypeError):
                logging.warning(f"Invalid decimal precision value: {decimal_precision_str}")

        final_processing_needed = (pikepdf_compression_level > 0 or 
                                   options.get('rotation', 0) != 0 or 
                                   options.get('strip_metadata', False))
        
        def process_a_file(input_file, output_file_target):
            final_output_path = output_file_target
            processing_path = output_file_target

            if overwrite:
                with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as temp_f:
                    processing_path = temp_f.name
            
            cmd = build_gs_command(gs_path, str(input_file), processing_path, params['operation'], options)
            run_command(cmd)

            if final_processing_needed:
                apply_final_processing(processing_path, options, pikepdf_compression_level)
            
            if overwrite:
                shutil.move(processing_path, final_output_path)
            
            return final_output_path

        if is_folder:
            input_folder = Path(params['input_path'])
            output_folder = Path(params['output_path'])
            if not input_folder.is_dir():
                raise FileNotFoundError("Input folder does not exist.")
            if not overwrite:
                output_folder.mkdir(exist_ok=True)

            pdf_files = list(input_folder.rglob("*.pdf"))
            if not pdf_files:
                raise FileNotFoundError("No PDF files found in the specified folder.")

            total = len(pdf_files)
            for i, pdf in enumerate(pdf_files):
                status_var.set(f"Processing file {i+1}/{total}: {pdf.name}")
                
                if overwrite:
                    process_a_file(pdf, pdf)
                else:
                    out_name = f"{pdf.stem}_processed.pdf"
                    output_file_path = output_folder / out_name
                    process_a_file(pdf, output_file_path)
            
            status_var.set(f"Complete: Processed {total} files.")
        else:
            input_path = params['input_path']
            output_path = params['output_path']
            if not Path(input_path).exists():
                raise FileNotFoundError("Input file does not exist.")

            status_var.set("Processing with Ghostscript...")
            process_a_file(input_path, output_path)
            status_var.set(f"Complete: Processed {Path(input_path).name}")
            
    except Exception as e:
        status_var.set(f"Error: {e}")
        messagebox.showerror("Error", str(e))
    finally:
        completion_callback()
#endregion
