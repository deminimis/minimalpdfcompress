import os
import sys
import shutil
import logging
import tempfile
import subprocess
from pathlib import Path
import pikepdf

from constants import ToolNotFound, ProcessingError

def resource_path(relative_path):
    try:
        base_path = Path(sys._MEIPASS)
    except AttributeError:
        base_path = Path(__file__).parent
    return base_path / relative_path

def find_executable(name, tool_name):
    exe_name = f"{name}.exe" if sys.platform == "win32" else name
    local_bin_path = resource_path('bin') / exe_name
    if local_bin_path.exists(): return str(local_bin_path)
    if shutil.which(exe_name): return exe_name
    raise ToolNotFound(f"Bundled '{tool_name}' ({exe_name}) not found and not in system PATH.")

def find_ghostscript(): return find_executable("gswin64c" if sys.platform == "win32" else "gs", "Ghostscript")
def find_cpdf(): return find_executable("cpdf", "cpdf")
def find_pngquant(): return find_executable("pngquant", "pngquant")
def find_jpegoptim(): return find_executable("jpegoptim", "jpegoptim")
def find_ect(): return find_executable("ect", "ECT")
def find_oxipng(): return find_executable("oxipng", "oxipng")

def format_size(size_bytes, decimals=1):
    abs_size = abs(size_bytes)
    if abs_size > 1024 * 1024: return f"{size_bytes / (1024*1024):.{decimals}f} MB"
    if abs_size > 1024: return f"{size_bytes / 1024:.{decimals}f} KB"
    return f"{size_bytes} bytes" if abs_size != 1 else f"{size_bytes} byte"

def get_pdf_metadata(file_path):
    try:
        p = Path(file_path)
        size_str = format_size(p.stat().st_size, decimals=1)
        
        with pikepdf.open(p) as pdf:
            page_count = len(pdf.pages)
            
        return {'name': p.name, 'pages': page_count, 'size': size_str}
    except Exception as e:
        logging.warning(f"Could not get metadata for {file_path}: {e}")
        return {'name': Path(file_path).name, 'pages': 'N/A', 'size': 'N/A'}

def run_command(command, check=True):
    use_shell = isinstance(command, str)
    logging.info(f"Executing command: {command}")
    try:
        kwargs = { 'stdin': subprocess.DEVNULL, 'check': check, 'capture_output': True, 'text': True, 'encoding': 'utf-8', 'errors': 'ignore', 'shell': use_shell }
        if sys.platform == "win32": kwargs['creationflags'] = subprocess.CREATE_NO_WINDOW
        result = subprocess.run(command, **kwargs)
        if result.stderr:
            stderr_text = result.stderr.strip()
            if "wmic.exe" in stderr_text and "Failed to retrieve time" in stderr_text:
                pass
            elif "not permitted in PDF/A-2, overprint mode not set" in stderr_text:
                pass
            elif "invalid xref" in stderr_text.lower() or "repaired" in stderr_text.lower():
                logging.info(f"Ignoring recoverable warning: {stderr_text}")
            elif "warning" not in stderr_text.lower():
                logging.warning(f"Command stderr: {stderr_text}")
        return result
    except subprocess.CalledProcessError as e:
        logging.error(f"Command failed.\nSTDOUT: {e.stdout}\nSTDERR: {e.stderr}")
        raise ProcessingError(f"Tool failed: {e.stderr.strip() if e.stderr else 'Unknown Error'}")
    except FileNotFoundError as e:
        logging.error(f"Command not found: {command if use_shell else command[0]}")
        raise ProcessingError(f"Command not found: {e}.")
    except Exception as e:
        logging.error(f"Unexpected error: {e}")
        raise ProcessingError(str(e))