#region: backend.py
import os
import sys
import subprocess
import logging
import shutil
import tempfile
from pathlib import Path
import pikepdf
from tkinter import messagebox
import re
from PIL import Image

#region: Exceptions and Helpers
class GhostscriptNotFound(Exception): pass
class CpdfNotFound(Exception): pass
class ProcessingError(Exception): pass

def resource_path(relative_path):
    try:
        base_path = Path(sys._MEIPASS)
    except AttributeError:
        base_path = Path(__file__).parent
    return base_path / relative_path
#endregion

#region: Tool Finders
def find_ghostscript():
    exe_name = "gswin64c.exe" if sys.platform == "win32" else "gs"
    local_bin_path = resource_path('bin') / exe_name
    if local_bin_path.exists(): return str(local_bin_path)
    raise GhostscriptNotFound("Bundled Ghostscript executable not found in 'bin' folder.")

def find_cpdf():
    exe_name = "cpdf.exe" if sys.platform == "win32" else "cpdf"
    local_bin_path = resource_path('bin') / exe_name
    if local_bin_path.exists(): return str(local_bin_path)
    raise CpdfNotFound("Bundled cpdf executable not found in 'bin' folder.")

def find_srgb_profile():
    srgb_path = resource_path('lib/srgb.icc')
    if srgb_path.exists(): return str(srgb_path)
    raise FileNotFoundError("Could not find 'srgb.icc'. Please ensure it is in the 'lib' folder.")
#endregion

#region: Command Execution
def build_gs_command(gs_path, input_path, output_path, operation, options):
    cmd = [ gs_path, "-sDEVICE=pdfwrite", "-dCompatibilityLevel=1.4", "-dNOPAUSE", "-dBATCH", "-dQUIET", "-dSAFER",
            f"-dDownScaleFactor={options['downscale_factor']}", f"-dDownsampleType=/{options['downsample_type']}",
            f"-dSubsetFonts={'true' if options['subset_fonts'] else 'false'}", f"-dCompressFonts={'true' if options['compress_fonts'] else 'false'}" ]
    try:
        img_res = int(options['image_resolution'])
        cmd.extend([f"-dColorImageResolution={img_res}", f"-dGrayImageResolution={img_res}"])
    except (ValueError, TypeError): logging.warning(f"Invalid image resolution: {options['image_resolution']}.")
    if options['fast_web_view']: cmd.append("-dFastWebView=true")
    if options.get('remove_interactive', False): cmd.extend(["-dShowAnnots=false", "-dShowAcroForm=false"])
    
    preset_map = { "Compress (Screen - Smallest Size)": "/screen", "Compress (Ebook - Medium Size)": "/ebook",
                   "Compress (Printer - High Quality)": "/printer", "Compress (Prepress - Highest Quality)": "/prepress" }
    if operation in preset_map:
        cmd.extend([f"-dPDFSETTINGS={preset_map[operation]}", f"-sColorConversionStrategy={options['color_strategy']}"])
    
    cmd.extend([f"-sOutputFile={output_path}", input_path])
    return cmd

def build_pdfa_command(gs_path, input_path, output_path):
    pdfa_def_path_str = str(resource_path('lib/PDFA_def.ps'))
    if not Path(pdfa_def_path_str).exists():
        raise FileNotFoundError("Could not find 'PDFA_def.ps'. Please ensure it is in the 'lib' folder.")

    cmd = [
        gs_path,
        "-dPDFA=2",
        "-dBATCH",
        "-dNOPAUSE",
        "-dNOOUTERSAVE",
        "-sDEVICE=pdfwrite",
        "-dPDFACompatibilityPolicy=1",
        "-sColorConversionStrategy=UseDeviceIndependentColor",
        f"-sOutputFile={output_path}",
        pdfa_def_path_str,
        input_path
    ]
    return cmd

def build_pdf_to_image_command(gs_path, input_pdf, output_dir, options):
    fmt, dpi = options.get('format', 'png'), options.get('dpi', '300')
    out_path = Path(output_dir) / f"{Path(input_pdf).stem}_%d.{fmt}"
    device_map = {'png': 'png16m', 'jpeg': 'jpeg', 'tiff': 'tiffg4'}
    return [ gs_path, f"-sDEVICE={device_map.get(fmt, 'png16m')}", f"-r{dpi}", "-dNOPAUSE", "-dBATCH", "-dSAFER", f"-sOutputFile={str(out_path)}", input_pdf ]

def run_command(command, check=True):
    logging.info(f"Executing: {command}")
    try:
        kwargs = { 'stdin': subprocess.DEVNULL, 'check': check, 'capture_output': True, 'text': True, 'encoding': 'utf-8', 'errors': 'ignore' }
        if sys.platform == "win32": kwargs['creationflags'] = subprocess.CREATE_NO_WINDOW
        result = subprocess.run(command, **kwargs)
        if result.stderr and check and "warning" not in result.stderr.lower():
            logging.warning(f"Command produced stderr: {result.stderr[:200]}...")
    except subprocess.CalledProcessError as e: logging.error(f"Command failed: {e.stderr}"); raise ProcessingError(f"Tool failed: {e.stderr[:200]}...")
    except FileNotFoundError as e: raise ProcessingError(f"Command not found: {e}.")
    except Exception as e: logging.error(f"Unexpected error: {e}"); raise ProcessingError(str(e))
#endregion

#region: Final Processing & Page Logic
def apply_final_processing(file_path, options, cpdf_path):
    use_any_cpdf = (options.get('use_cpdf_squeeze', False) or options.get('darken_text', False) or
                    options.get('user_password') or options.get('owner_password'))
    if use_any_cpdf and cpdf_path:
        logging.info(f"Applying cpdf processing to {file_path}")
        with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as temp_cpdf_out:
            temp_path = str(temp_cpdf_out.name)
        try:
            cmd = [cpdf_path]
            if options.get('use_fast_processing', False): cmd.append("-fast")
            cmd.append(str(file_path))
            if options.get('use_cpdf_squeeze', False): cmd.extend(["AND", "-squeeze"])
            if options.get('darken_text', False): cmd.extend(["AND", "-blacktext"])

            user_pass, owner_pass = options.get('user_password'), options.get('owner_password')
            if user_pass or owner_pass:
                cmd.extend(["AND", "-encrypt", "AES", user_pass if user_pass else '""', owner_pass if owner_pass else '""'])

            cmd.extend(["-o", temp_path])
            run_command(cmd)
            shutil.move(temp_path, file_path)
        finally:
            if os.path.exists(temp_path): os.remove(temp_path)
    else:
        logging.info(f"Applying Pikepdf processing to {file_path}")
        try:
            pikepdf.settings.set_flate_compression_level(options.get('pikepdf_compression_level', 6))
            if (dp_str := options.get('decimal_precision', 'Default')).isdigit():
                pikepdf.settings.set_decimal_precision(int(dp_str))

            with pikepdf.open(file_path, allow_overwriting_input=True) as pdf:
                rotation_str = options.get('rotation', 'No Rotation')
                ROTATION_MAP = { "No Rotation": 0, "90° Right (Clockwise)": 90, "180°": 180, "90° Left (Counter-Clockwise)": 270 }
                angle = ROTATION_MAP.get(rotation_str, 0)
                
                if angle != 0:
                    for page in pdf.pages: page.rotate(angle, relative=True)
                if options.get('strip_metadata', False):
                    with pdf.open_metadata(set_pikepdf_as_editor=False) as meta:
                        for key in list(meta.keys()): del meta[key]

                pdf.save(file_path, object_stream_mode=pikepdf.ObjectStreamMode.generate, compress_streams=True)
        except Exception as e: logging.error(f"Pikepdf processing failed: {e}"); raise ProcessingError(f"Final processing step failed: {e}")

def parse_page_ranges(page_string, max_pages):
    indices = set()
    if not page_string.strip(): return []
    page_string = re.sub(r'\bend\b', str(max_pages), page_string, flags=re.IGNORECASE)
    for part in page_string.split(','):
        part = part.strip()
        if not part: continue
        if '-' in part:
            start, end = part.split('-', 1)
            start = 1 if start.strip() == '' else int(start)
            end = max_pages if end.strip() == '' else int(end)
            indices.update(range(start - 1, end))
        else: indices.add(int(part) - 1)
    return sorted(list(indices), reverse=True)
#endregion

#region: Main Task Runners
def run_conversion_task(params, is_folder, status_var, progress_var, cb):
    final_status = "Processing complete."
    try:
        def process_a_file(input_file, output_file_target):
            with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as gs_temp_out:
                gs_temp_path = str(gs_temp_out.name)
            try:
                cmd = build_gs_command(params['gs_path'], str(input_file), gs_temp_path, params['operation'], params['options'])
                run_command(cmd)
                apply_final_processing(gs_temp_path, params['options'], params.get('cpdf_path'))
                shutil.move(gs_temp_path, output_file_target)
            finally:
                if os.path.exists(gs_temp_path): os.remove(gs_temp_path)

        if is_folder:
            input_folder, output_folder = Path(params['input_path']), Path(params['output_path'])
            if not params['overwrite']: output_folder.mkdir(exist_ok=True)
            pdf_files = list(input_folder.rglob("*.pdf"))
            if not pdf_files: raise FileNotFoundError("No PDF files found in folder.")
            total = len(pdf_files)
            for i, pdf in enumerate(pdf_files):
                status_var.set(f"Processing ({i+1}/{total}): {pdf.name}")
                progress_var['value'] = (i / total) * 100
                if params['overwrite']:
                    with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as final_temp:
                        final_temp_path = Path(final_temp.name)
                    try: process_a_file(pdf, final_temp_path); shutil.move(str(final_temp_path), str(pdf))
                    finally:
                        if final_temp_path.exists(): final_temp_path.unlink()
                else:
                    out_name = f"{pdf.stem}_processed.pdf"; process_a_file(pdf, output_folder / out_name)
            final_status = f"Complete: Processed {total} files."
        else:
            status_var.set(f"Processing: {Path(params['input_path']).name}"); progress_var['value'] = 25
            if params['overwrite']:
                with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as final_temp:
                    final_temp_path = Path(final_temp.name)
                try: process_a_file(params['input_path'], final_temp_path); shutil.move(str(final_temp_path), params['input_path'])
                finally:
                    if final_temp_path.exists(): final_temp_path.unlink()
            else: process_a_file(params['input_path'], params['output_path'])
        progress_var['value'] = 100
    except Exception as e: final_status = f"Error: {e}"; messagebox.showerror("Error", str(e))
    finally: cb(final_status)

def run_pdfa_conversion_task(pdf_in, pdf_out, gs_path, status_var, progress_var, cb):
    final_status = "PDF/A conversion complete."
    temp_path = None
    try:
        status_var.set("Converting to PDF/A..."); progress_var['value'] = 25
        
        tf = tempfile.NamedTemporaryFile(suffix=".pdf", delete=False)
        temp_path = tf.name
        tf.close()

        cmd = build_pdfa_command(gs_path, pdf_in, temp_path)
        run_command(cmd)
        
        shutil.move(temp_path, pdf_out)
        temp_path = None
        
        progress_var['value'] = 100
    except Exception as e:
        final_status = f"Error: {e}"
        messagebox.showerror("Error", str(e))
    finally:
        if temp_path and os.path.exists(temp_path):
            os.remove(temp_path)
        cb(final_status)

def run_pdf_to_image_task(pdf_in, dir_out, options, gs_path, status_var, progress_var, cb):
    final_status = "Conversion to images completed."
    try:
        status_var.set("Preparing for conversion..."); progress_var['value'] = 10
        cmd = build_pdf_to_image_command(gs_path, pdf_in, dir_out, options)
        status_var.set("Converting PDF to images..."); run_command(cmd); progress_var['value'] = 100
    except Exception as e: final_status = f"Error: {e}"; messagebox.showerror("Error", str(e))
    finally: cb(final_status)

def run_merge_task(file_list, output_path, status_var, progress_var, cb):
    final_status = "Merge complete."
    try:
        status_var.set("Merging files..."); progress_var['value'] = 10
        with pikepdf.Pdf.new() as pdf:
            for i, file_path in enumerate(file_list):
                status_var.set(f"Adding {Path(file_path).name}..."); progress_var['value'] = 10 + (i / len(file_list) * 80)
                with pikepdf.open(file_path) as src: pdf.pages.extend(src.pages)
            pdf.save(output_path); progress_var['value'] = 100
    except Exception as e: final_status = f"Error: {e}"; messagebox.showerror("Error", str(e))
    finally: cb(final_status)

def run_split_task(input_path, output_dir, mode, value, status_var, progress_var, cb):
    final_status = "Splitting complete."
    try:
        status_var.set("Opening PDF..."); progress_var['value'] = 10; p_in = Path(input_path)
        with pikepdf.open(p_in) as pdf:
            total = len(pdf.pages)
            if mode == "Split to Single Pages":
                for i, page in enumerate(pdf.pages):
                    status_var.set(f"Saving page {i+1}/{total}"); progress_var['value'] = 10 + (i / total * 90)
                    with pikepdf.Pdf.new() as dst: dst.pages.append(page); dst.save(Path(output_dir) / f"{p_in.stem}_page_{i+1}.pdf")
            elif mode == "Split Every N Pages":
                n = int(value)
                for i in range(0, total, n):
                    status_var.set(f"Saving chunk starting at page {i+1}"); progress_var['value'] = 10 + (i / total * 90)
                    with pikepdf.Pdf.new() as dst: dst.pages.extend(pdf.pages[i:i+n]); dst.save(Path(output_dir) / f"{p_in.stem}_pages_{i+1}-{i+n}.pdf")
            elif mode == "Custom Range(s)":
                indices = parse_page_ranges(value, total)
                with pikepdf.Pdf.new() as dst:
                    for i in sorted([i for i in indices if i < total]):
                        dst.pages.append(pdf.pages[i])
                dst.save(Path(output_dir) / f"{p_in.stem}_custom_range.pdf")
            progress_var['value'] = 100
    except Exception as e: final_status = f"Error: {e}"; messagebox.showerror("Error", str(e))
    finally: cb(final_status)

def run_delete_pages_task(pdf_in, pdf_out, page_range, status_var, progress_var, cb):
    final_status = "Page deletion completed."
    try:
        status_var.set("Opening PDF..."); progress_var['value'] = 10
        with pikepdf.open(pdf_in) as pdf:
            indices_to_delete = parse_page_ranges(page_range, len(pdf.pages))
            if not indices_to_delete: raise ProcessingError("No valid pages to delete.")
            status_var.set(f"Deleting {len(indices_to_delete)} page(s)...")
            for i, idx in enumerate(indices_to_delete):
                del pdf.pages[idx]; progress_var['value'] = 10 + (i / len(indices_to_delete) * 80)
            pdf.save(pdf_out); progress_var['value'] = 100
    except Exception as e: final_status = f"Error: {e}"; messagebox.showerror("Error", str(e))
    finally: cb(final_status)

def run_rotate_task(pdf_in, pdf_out, angle, status_var, progress_var, cb):
    final_status = "Rotation complete."
    try:
        status_var.set("Opening PDF..."); progress_var['value'] = 20
        with pikepdf.open(pdf_in) as pdf:
            status_var.set(f"Rotating all pages by {angle} degrees..."); progress_var['value'] = 50
            for page in pdf.pages: page.rotate(angle, relative=True)
            pdf.save(pdf_out)
            progress_var['value'] = 100
    except Exception as e: final_status = f"Error: {e}"; messagebox.showerror("Error", str(e))
    finally: cb(final_status)

def run_stamp_task(pdf_in, pdf_out, stamp_opts, cpdf_path, status_var, progress_var, cb, mode, mode_opts):
    final_status = "Stamping complete."
    temp_stamp_pdf = None
    try:
        status_var.set("Stamping file..."); progress_var['value'] = 50

        pos_map = { "Center": ["-center"], "Bottom-Left": ["-bottomleft", "10"], "Bottom-Right": ["-bottomright", "10"] }
        pos_cmd = pos_map.get(stamp_opts['pos'], ["-center"])

        cmd = [cpdf_path]
        if mode == "Image":
            with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as temp_pdf_file:
                temp_stamp_pdf = temp_pdf_file.name

            with Image.open(mode_opts['image_path']) as img:
                width, height = img.size
                if new_w_str := mode_opts.get('width'):
                    try: width = int(new_w_str)
                    except ValueError: pass
                if new_h_str := mode_opts.get('height'):
                    try: height = int(new_h_str)
                    except ValueError: pass

                if (width, height) != img.size:
                    img = img.resize((width, height), Image.LANCZOS)

                if stamp_opts['opacity'] < 1.0:
                    if img.mode != 'RGBA':
                        img = img.convert('RGBA')
                    alpha = img.split()[3]
                    alpha = alpha.point(lambda p: p * stamp_opts['opacity'])
                    img.putalpha(alpha)

                img.save(temp_stamp_pdf, "PDF", resolution=100.0)

            cmd.extend([pdf_in, "-stamp-on" if stamp_opts['on_top'] else "-stamp-under", temp_stamp_pdf])
            cmd.extend(pos_cmd)
            cmd.extend(["-o", pdf_out])
        else: 
            text_to_stamp = mode_opts['text']
            cmd.extend([pdf_in, "-add-text", text_to_stamp.strip(), "-font", mode_opts['font'],
                        "-font-size", str(mode_opts['size']), "-color", mode_opts['color']])
            cmd.extend(pos_cmd)
            if (bates_start := mode_opts.get('bates_start')):
                cmd.extend(["-bates", bates_start])
            cmd.extend(["-opacity", str(stamp_opts['opacity'])])
            if not stamp_opts['on_top']: cmd.append("-underneath")
            cmd.extend(["-o", pdf_out])

        run_command(cmd)
        progress_var['value'] = 100
    except Exception as e: final_status = f"Error: {e}"; messagebox.showerror("Error", str(e))
    finally:
        if temp_stamp_pdf and os.path.exists(temp_stamp_pdf):
            os.remove(temp_stamp_pdf)
        cb(final_status)

def run_metadata_task(task_type, pdf_path, cpdf_path, metadata_dict=None):
    if task_type == 'load':
        cmd = [cpdf_path, "-info", pdf_path]
        proc = subprocess.run(cmd, capture_output=True, text=True, encoding='utf-8', errors='ignore')
        if proc.returncode != 0: raise ProcessingError(f"cpdf failed to read info: {proc.stderr[:200]}")
        info = {}
        for line in proc.stdout.splitlines():
            if ':' in line:
                key, val = line.split(':', 1)
                key = key.strip().lower()
                if key == 'title': info['title'] = val.strip()
                elif key == 'author': info['author'] = val.strip()
                elif key == 'subject': info['subject'] = val.strip()
                elif key == 'keywords': info['keywords'] = val.strip()
        return info
    elif task_type == 'save':
        with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as temp_out:
            temp_path = temp_out.name
        try:
            cmd = [cpdf_path, pdf_path]
            first_op = True
            for key, value in metadata_dict.items():
                if value:
                    if not first_op: cmd.append("AND")
                    cmd.extend([f"-set-{key}", value])
                    first_op = False

            if first_op: 
                return 
            
            cmd.extend(["-o", temp_path])
            run_command(cmd)
            shutil.move(temp_path, pdf_path)
        finally:
            if os.path.exists(temp_path): os.remove(temp_path)

def run_remove_open_action_task(pdf_in, cpdf_path, status_var, progress_var, cb):
    final_status = "Opening action removed."
    try:
        status_var.set("Removing opening action..."); progress_var['value'] = 50
        with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as temp_out:
            temp_path = temp_out.name
        cmd = [cpdf_path, pdf_in, "-remove-dict-entry", "/OpenAction", "-o", temp_path]
        run_command(cmd)
        shutil.move(temp_path, pdf_in) 
        progress_var['value'] = 100
    except Exception as e: final_status = f"Error: {e}"; messagebox.showerror("Error", str(e))
    finally: cb(final_status)
#endregion