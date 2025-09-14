# backend.py
import os
import sys
import subprocess
import logging
import shutil
import tempfile
from pathlib import Path
import pikepdf
import re
from PIL import Image
from io import BytesIO

from pdf_optimizer import PdfOptimizer
from constants import (SPLIT_SINGLE, SPLIT_EVERY_N, SPLIT_CUSTOM, STAMP_IMAGE,
                       POS_TOP_LEFT, POS_TOP_CENTER, POS_TOP_RIGHT,
                       POS_MIDDLE_LEFT, POS_CENTER, POS_MIDDLE_RIGHT,
                       POS_BOTTOM_LEFT, POS_BOTTOM_CENTER, POS_BOTTOM_RIGHT,
                       META_LOAD, META_SAVE)

class ToolNotFound(Exception): pass
class GhostscriptNotFound(ToolNotFound): pass
class CpdfNotFound(ToolNotFound): pass
class PngquantNotFound(ToolNotFound): pass
class JpegoptimNotFound(ToolNotFound): pass
class EctNotFound(ToolNotFound): pass
class OptipngNotFound(ToolNotFound): pass
class ProcessingError(Exception): pass

def resource_path(relative_path):
    try:
        base_path = Path(sys._MEIPASS)
    except AttributeError:
        base_path = Path(__file__).parent
    return base_path / relative_path

def find_executable(name, not_found_exception):
    exe_name = f"{name}.exe" if sys.platform == "win32" else name
    local_bin_path = resource_path('bin') / exe_name
    if local_bin_path.exists():
        return str(local_bin_path)
    if shutil.which(exe_name):
        return exe_name
    raise not_found_exception(f"Bundled {exe_name} not found and not in system PATH.")

def find_ghostscript(): return find_executable("gswin64c" if sys.platform == "win32" else "gs", GhostscriptNotFound)
def find_cpdf(): return find_executable("cpdf", CpdfNotFound)
def find_pngquant(): return find_executable("pngquant", PngquantNotFound)
def find_jpegoptim(): return find_executable("jpegoptim", JpegoptimNotFound)
def find_ect(): return find_executable("ect", EctNotFound)
def find_optipng(): return find_executable("optipng", OptipngNotFound)

def get_pdf_metadata(file_path):
    try:
        p = Path(file_path)
        size_bytes = p.stat().st_size
        if size_bytes > 1024 * 1024:
            size_str = f"{size_bytes / (1024*1024):.1f} MB"
        elif size_bytes > 1024:
            size_str = f"{size_bytes / 1024:.1f} KB"
        else:
            size_str = f"{size_bytes} B"

        with pikepdf.open(p) as pdf:
            page_count = len(pdf.pages)

        return {'name': p.name, 'pages': page_count, 'size': size_str}
    except FileNotFoundError:
        return {'name': Path(file_path).name, 'pages': 'N/A', 'size': 'N/A'}
    except Exception as e:
        logging.warning(f"Could not get metadata for {file_path}: {e}")
        return {'name': Path(file_path).name, 'pages': 'N/A', 'size': 'N/A'}

def run_command(command, check=True):
    logging.info(f"Executing command: {command}")
    try:
        kwargs = { 'stdin': subprocess.DEVNULL, 'check': check, 'capture_output': True, 'text': True, 'encoding': 'utf-8', 'errors': 'ignore' }
        if sys.platform == "win32": kwargs['creationflags'] = subprocess.CREATE_NO_WINDOW
        result = subprocess.run(command, **kwargs)
        if result.stderr:
            stderr_text = result.stderr.strip()
            if "wmic.exe" in stderr_text and "Failed to retrieve time" in stderr_text:
                pass
            elif "not permitted in PDF/A-2, overprint mode not set" in stderr_text:
                pass
            elif "warning" not in stderr_text.lower():
                logging.warning(f"Command stderr: {stderr_text}")
        return result
    except subprocess.CalledProcessError as e:
        logging.error(f"Command failed: {e.stderr}")
        raise ProcessingError(f"Tool failed: {e.stderr.strip()}")
    except FileNotFoundError as e:
        logging.error(f"Command not found: {e}")
        raise ProcessingError(f"Command not found: {e}.")
    except Exception as e:
        logging.error(f"Unexpected error: {e}")
        raise ProcessingError(str(e))

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

def get_total_size(path, is_folder):
    if is_folder:
        return sum(f.stat().st_size for f in Path(path).glob("*.pdf") if f.is_file())
    else:
        p = Path(path)
        return p.stat().st_size if p.is_file() else 0

def generate_preview(gs_path, cpdf_path, pdf_path, operation, options):
    if not pdf_path or not Path(pdf_path).exists():
        logging.warning("Input PDF for preview not found or path is empty.")
        return None

    with tempfile.TemporaryDirectory() as temp_dir:
        temp_dir_path = Path(temp_dir)
        first_page_pdf = temp_dir_path / "first_page.pdf"
        modified_pdf = temp_dir_path / "modified.pdf"
        preview_image_path = temp_dir_path / "preview.png"

        try:
            with pikepdf.open(pdf_path, allow_overwriting_input=False) as pdf:
                if not pdf.pages:
                    logging.warning("PDF has no pages to preview.")
                    return None
                with pikepdf.Pdf.new() as dst:
                    dst.pages.append(pdf.pages[0])
                    dst.save(first_page_pdf)
        except Exception as e:
            logging.error(f"Failed to extract first page for preview: {e}")
            return None

        if operation == 'rotate':
            try:
                with pikepdf.open(first_page_pdf) as pdf:
                    pdf.pages[0].rotate(options.get('angle', 0), relative=True)
                    pdf.save(modified_pdf)
            except Exception as e:
                logging.error(f"Failed to apply rotation for preview: {e}")

        elif operation == 'stamp':
            try:
                stamp_opts = options['stamp_opts']
                mode_opts = options['mode_opts']
                mode = options['mode']

                pos = stamp_opts['pos']
                margin = "20"
                pos_map = {
                    POS_TOP_LEFT: ["-topleft", margin], POS_TOP_CENTER: ["-top", margin], POS_TOP_RIGHT: ["-topright", margin],
                    POS_MIDDLE_LEFT: ["-left", margin], POS_CENTER: ["-center"], POS_MIDDLE_RIGHT: ["-right", margin],
                    POS_BOTTOM_LEFT: ["-bottomleft", margin], POS_BOTTOM_CENTER: ["-bottom", margin], POS_BOTTOM_RIGHT: ["-bottomright", margin],
                }
                pos_cmd = pos_map.get(pos, ["-center"])

                if mode == STAMP_IMAGE:
                    image_path = mode_opts.get('image_path')
                    if image_path and Path(image_path).exists():
                        stamp_to_apply_path = temp_dir_path / "image_stamp.pdf"

                        scale = mode_opts.get('image_scale', 1.0)

                        with Image.open(image_path) as img:
                            if scale != 1.0:
                                try:
                                    w = int(img.width * scale)
                                    h = int(img.height * scale)
                                    if w > 0 and h > 0:
                                        resample = Image.Resampling.LANCZOS if hasattr(Image, 'Resampling') else Image.ANTIALIAS
                                        img = img.resize((w, h), resample)
                                except Exception as e:
                                    logging.warning(f"Invalid image scale, using original. Error: {e}")

                            if stamp_opts['opacity'] < 1.0:
                                if img.mode != 'RGBA': img = img.convert('RGBA')
                                alpha = img.split()[3]; alpha = alpha.point(lambda p: p * stamp_opts['opacity']); img.putalpha(alpha)
                            img.save(stamp_to_apply_path, "PDF", resolution=100.0)

                        cmd = [cpdf_path, str(first_page_pdf), "-stamp-on" if stamp_opts['on_top'] else "-stamp-under", str(stamp_to_apply_path)]
                        cmd.extend(pos_cmd)
                        cmd.extend(["-o", str(modified_pdf)])
                        run_command(cmd)

                else:
                    final_text = mode_opts['text'].strip()
                    if mode_opts.get('bates_start') and "%Bates" in final_text:
                        final_text = final_text.replace("%Bates", str(mode_opts['bates_start']).zfill(6))

                    cmd = [cpdf_path, str(first_page_pdf)]
                    text_parts = ["-add-text", final_text, "-font", mode_opts['font'], "-font-size", str(mode_opts['size']), "-color", mode_opts['color'], "-opacity", str(stamp_opts['opacity'])]
                    cmd.extend(text_parts)
                    if not stamp_opts['on_top']: cmd.append("-underneath")
                    cmd.extend(pos_cmd)
                    cmd.extend(["-o", str(modified_pdf)])
                    run_command(cmd)

            except Exception as e:
                logging.error(f"Failed to apply stamp for preview: {e}", exc_info=True)

        if not modified_pdf.exists():
            shutil.copy(first_page_pdf, modified_pdf)

        try:
            render_cmd = [gs_path, "-sDEVICE=png16m", "-r96", "-dNOPAUSE", "-dBATCH", "-dSAFER", f"-sOutputFile={preview_image_path}", str(modified_pdf)]
            run_command(render_cmd)
        except Exception as e:
            logging.error(f"Failed to render preview image with Ghostscript: {e}")
            return None

        if not preview_image_path.exists() or preview_image_path.stat().st_size == 0:
            logging.error("Preview image was not generated or is empty.")
            return None

        with open(preview_image_path, 'rb') as f:
            image_data = f.read()
        return Image.open(BytesIO(image_data))

def run_compress_task(params, is_folder, q):
    try:
        optimizer = PdfOptimizer(
            gs_path=params['gs_path'], cpdf_path=params['cpdf_path'],
            pngquant_path=params['pngquant_path'],
            jpegoptim_path=params['jpegoptim_path'],
            ect_path=params['ect_path'],
            optipng_path=params['optipng_path'],
            q=q if not is_folder else None,
            user_password=params['user_password'], darken_text=params['darken_text'],
            remove_open_action=params.get('remove_open_action'),
            fast_web_view=params.get('fast_web_view'),
            fast_mode=params.get('fast_mode'),
            convert_to_grayscale=params.get('convert_to_grayscale', False),
            convert_to_cmyk=params.get('convert_to_cmyk', False),
            downsample_threshold_enabled=params.get('downsample_threshold_enabled', False),
            quantize_colors=params.get('quantize_colors', False),
            quantize_level=params.get('quantize_level', 4)
        )

        input_path = Path(params['input_path'])
        output_path = Path(params['output_path'])
        total_in_size = get_total_size(params['input_path'], is_folder)
        files_skipped = 0

        def process_a_file(input_file, output_file):
            nonlocal files_skipped
            mode = params.get('mode', 'Lossy')
            only_if_smaller = params.get('only_if_smaller', False)
            original_size = input_file.stat().st_size

            with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as temp_out:
                temp_output_path = Path(temp_out.name)

            try:
                if mode == 'Lossless':
                    if params.get('true_lossless', False):
                        optimizer.optimize_true_lossless(input_file, temp_output_path, strip_metadata=params['strip_metadata'])
                    else:
                        optimizer.optimize_lossless(input_file, temp_output_path, strip_metadata=params['strip_metadata'])
                elif mode == 'PDF/A':
                    optimizer.optimize_pdfa(input_file, temp_output_path)
                elif mode == 'Remove Images':
                    optimizer.optimize_text_only(input_file, temp_output_path, strip_metadata=params['strip_metadata'])
                else:
                    optimizer.optimize_lossy(
                        input_file, temp_output_path, params['dpi'],
                        strip_metadata=params['strip_metadata'],
                        remove_interactive=params['remove_interactive'],
                        use_bicubic=params['use_bicubic']
                    )

                if not temp_output_path.exists() or temp_output_path.stat().st_size == 0:
                    logging.warning(f"Processing failed for {input_file.name}, temp file is empty. Copying original.")
                    shutil.copy2(input_file, output_file)
                    return

                if mode == 'PDF/A':
                    shutil.move(str(temp_output_path), output_file)
                    return

                new_size = temp_output_path.stat().st_size
                if new_size < original_size:
                    shutil.move(str(temp_output_path), output_file)
                else:
                    if only_if_smaller:
                        files_skipped += 1
                        logging.info(f"Skipping save for {input_file.name}, new size {new_size} >= original size {original_size}")
                    else:
                        shutil.copy2(input_file, output_file)
                        logging.info(f"Saving original for {input_file.name}, new size {new_size} >= original size {original_size}")

            finally:
                if temp_output_path.exists():
                    os.remove(temp_output_path)

        if is_folder:
            output_path.mkdir(parents=True, exist_ok=True)
            pdf_files = sorted(list(input_path.glob("*.pdf")))
            total_files = len(pdf_files)
            if total_files == 0: raise ProcessingError(f"No PDF files found in '{input_path}'")
            for i, pdf_file in enumerate(pdf_files):
                q.put(('status', f"Processing {pdf_file.name} ({i+1}/{total_files})..."))
                process_a_file(pdf_file, output_path / pdf_file.name)
                q.put(('overall', ((i + 1) / total_files) * 100))
        else:
            q.put(('status', f"Processing {input_path.name}..."))
            process_a_file(input_path, output_path)

        total_out_size = get_total_size(output_path, is_folder)
        final_message = "Processing complete."

        if total_in_size > 0 and total_out_size > 0:
            saved_bytes = total_in_size - total_out_size
            if saved_bytes > 0:
                percent_saved = (saved_bytes / total_in_size) * 100
                if abs(saved_bytes) > 1024*1024: saved_str = f"{saved_bytes / (1024*1024):.2f} MB"
                elif abs(saved_bytes) > 1024: saved_str = f"{saved_bytes / 1024:.2f} KB"
                else: saved_str = f"{saved_bytes} bytes"
                final_message = f"Complete. Saved {saved_str} ({percent_saved:.1f}%)."
            else:
                final_message = "Complete. File size did not decrease."

        if files_skipped > 0:
            final_message += f" ({files_skipped} file(s) not saved)."

        q.put(('complete', final_message))

    except (ToolNotFound, ProcessingError, FileNotFoundError, ValueError, RuntimeError) as e:
        logging.error(f"Task failed: {e}"); q.put(('complete', f"Error: {e}"))
    except Exception as e:
        logging.error(f"An unexpected error occurred: {e}", exc_info=True); q.put(('complete', f"An unexpected error occurred: {e}"))

def run_merge_task(file_list, output_path, q):
    try:
        with pikepdf.Pdf.new() as pdf:
            total_files = len(file_list)
            if total_files == 0: raise ProcessingError("No files selected to merge.")
            for i, file_path in enumerate(file_list):
                q.put(('status', f"Adding {Path(file_path).name} ({i+1}/{total_files})"))
                with pikepdf.open(file_path) as src:
                    pdf.pages.extend(src.pages)
                q.put(('progress', ((i + 1) / total_files) * 100))
            q.put(('status', "Saving merged file...")); pdf.save(output_path); q.put(('overall', 100))
        q.put(('complete', "Merge complete."))
    except Exception as e: logging.error(f"Merge task failed: {e}"); q.put(('complete', f"Error: {e}"))

def run_split_task(input_path, output_dir, mode, value, q):
    try:
        q.put(('status', "Opening PDF...")); p_in = Path(input_path)
        output_dir_path = Path(output_dir)
        output_dir_path.mkdir(parents=True, exist_ok=True)
        with pikepdf.open(p_in) as pdf:
            total = len(pdf.pages)
            if mode == SPLIT_SINGLE:
                for i, page in enumerate(pdf.pages):
                    prog = ((i + 1) / total) * 100
                    q.put(('status', f"Saving page {i+1}/{total}")); q.put(('progress', prog)); q.put(('overall', prog))
                    with pikepdf.Pdf.new() as dst: dst.pages.append(page); dst.save(output_dir_path / f"{p_in.stem}_page_{i+1}.pdf")
            elif mode == SPLIT_EVERY_N:
                n = int(value);
                if n <= 0: raise ValueError("Number of pages must be positive.")
                for i in range(0, total, n):
                    prog = (min(i + n, total) / total) * 100
                    q.put(('status', f"Saving pages {i+1}-{min(i+n, total)}")); q.put(('progress', prog)); q.put(('overall', prog))
                    with pikepdf.Pdf.new() as dst: dst.pages.extend(pdf.pages[i:i+n]); dst.save(output_dir_path / f"{p_in.stem}_pages_{i+1}-{i+n}.pdf")
            elif mode == SPLIT_CUSTOM:
                indices = parse_page_ranges(value, total); q.put(('status', f"Extracting {len(indices)} pages..."))
                with pikepdf.Pdf.new() as dst:
                    for i, page_index in enumerate(sorted([i for i in indices if i < total])):
                        prog = ((i+1)/len(indices)) * 100
                        dst.pages.append(pdf.pages[page_index]); q.put(('progress', prog)); q.put(('overall', prog))
                dst.save(output_dir_path / f"{p_in.stem}_custom_range.pdf")
            else: raise ProcessingError(f"Unknown split mode: {mode}")
        q.put(('complete', "Splitting complete."))
    except Exception as e: logging.error(f"Split task failed: {e}"); q.put(('complete', f"Error: {e}"))

def run_delete_pages_task(pdf_in, pdf_out, page_range, q):
    try:
        q.put(('status', "Opening PDF..."))
        with pikepdf.open(pdf_in) as pdf:
            indices = parse_page_ranges(page_range, len(pdf.pages))
            if not indices: raise ProcessingError("No valid pages specified for deletion.")
            q.put(('status', f"Deleting {len(indices)} page(s)..."))
            for i in indices: del pdf.pages[i]
            pdf.save(pdf_out); q.put(('progress', 100)); q.put(('overall', 100))
        q.put(('complete', "Page deletion completed."))
    except Exception as e: logging.error(f"Delete pages task failed: {e}"); q.put(('complete', f"Error: {e}"))

def run_rotate_task(pdf_in, pdf_out, angle, q):
    try:
        q.put(('status', "Opening PDF..."))
        with pikepdf.open(pdf_in) as pdf:
            q.put(('status', f"Rotating all pages by {angle} degrees..."))
            for page in pdf.pages: page.rotate(angle, relative=True)
            pdf.save(pdf_out); q.put(('progress', 100)); q.put(('overall', 100))
        q.put(('complete', "Rotation complete."))
    except Exception as e: logging.error(f"Rotate task failed: {e}"); q.put(('complete', f"Error: {e}"))

def run_stamp_task(pdf_in, pdf_out, stamp_opts, cpdf_path, q, mode, mode_opts):
    with tempfile.TemporaryDirectory() as temp_dir_str:
        temp_dir = Path(temp_dir_str)
        try:
            q.put(('status', "Applying stamp..."))

            pos = stamp_opts['pos']
            margin = "20"
            pos_map = {
                POS_TOP_LEFT: ["-topleft", margin], POS_TOP_CENTER: ["-top", margin], POS_TOP_RIGHT: ["-topright", margin],
                POS_MIDDLE_LEFT: ["-left", margin], POS_CENTER: ["-center"], POS_MIDDLE_RIGHT: ["-right", margin],
                POS_BOTTOM_LEFT: ["-bottomleft", margin], POS_BOTTOM_CENTER: ["-bottom", margin], POS_BOTTOM_RIGHT: ["-bottomright", margin],
            }
            pos_cmd = pos_map.get(pos, ["-center"])

            if mode == STAMP_IMAGE:
                stamp_to_apply_path = temp_dir / "image_stamp.pdf"
                scale = mode_opts.get('image_scale', 1.0)

                with Image.open(mode_opts['image_path']) as img:
                    if scale != 1.0:
                        try:
                            w = int(img.width * scale)
                            h = int(img.height * scale)
                            if w > 0 and h > 0:
                                resample = Image.Resampling.LANCZOS if hasattr(Image, 'Resampling') else Image.ANTIALIAS
                                img = img.resize((w, h), resample)
                        except Exception as e:
                            logging.warning(f"Invalid image scale, using original. Error: {e}")

                    if stamp_opts['opacity'] < 1.0:
                        if img.mode != 'RGBA': img = img.convert('RGBA')
                        alpha = img.split()[3]; alpha = alpha.point(lambda p: p * stamp_opts['opacity']); img.putalpha(alpha)
                    img.save(stamp_to_apply_path, "PDF", resolution=100.0)

                cmd = [cpdf_path, pdf_in, "-stamp-on" if stamp_opts['on_top'] else "-stamp-under", str(stamp_to_apply_path)]
                cmd.extend(pos_cmd)
                cmd.extend(["-o", pdf_out])
                run_command(cmd)

            else:
                final_text = mode_opts['text'].strip()
                cmd = [cpdf_path, pdf_in]
                text_parts = ["-add-text", final_text, "-font", mode_opts['font'], "-font-size", str(mode_opts['size']), "-color", mode_opts['color'], "-opacity", str(stamp_opts['opacity'])]
                if mode_opts.get('bates_start'):
                    text_parts.extend(["-bates", mode_opts['bates_start']])
                cmd.extend(text_parts)
                if not stamp_opts['on_top']: cmd.append("-underneath")
                cmd.extend(pos_cmd)
                cmd.extend(["-o", pdf_out])
                run_command(cmd)

            q.put(('progress', 100)); q.put(('overall', 100)); q.put(('complete', "Stamping complete."))
        except Exception as e:
            logging.error(f"Stamp task failed: {e}", exc_info=True)
            q.put(('complete', f"Error: {e}"))

def run_page_number_task(pdf_in, pdf_out, cpdf_path, q, options):
    try:
        q.put(('status', "Adding page numbers/headers/footers..."))

        cmd = [cpdf_path, "-utf8"]

        cmd.extend(["-add-text", options['text']])
        cmd.extend(["-font", options['font']])
        cmd.extend(["-font-size", str(options['font_size'])])
        cmd.extend(["-color", options['color']])

        pos = options['pos']
        margin = "15"
        pos_map = {
            POS_TOP_LEFT: ["-topleft", margin], POS_TOP_CENTER: ["-top", margin], POS_TOP_RIGHT: ["-topright", margin],
            POS_BOTTOM_LEFT: ["-bottomleft", margin], POS_BOTTOM_CENTER: ["-bottom", margin], POS_BOTTOM_RIGHT: ["-bottomright", margin],
        }
        pos_cmd = pos_map.get(pos, ["-bottom", margin])
        cmd.extend(pos_cmd)

        cmd.append(pdf_in)
        page_range = options.get('page_range', '').strip()
        if page_range:
            cmd.append(page_range)

        cmd.extend(["-o", pdf_out])

        run_command(cmd)

        q.put(('progress', 100))
        q.put(('overall', 100))
        q.put(('complete', "Header/Footer task complete."))
    except Exception as e:
        logging.error(f"Page Number task failed: {e}", exc_info=True)
        q.put(('complete', f"Error: {e}"))

def run_metadata_task(task_type, pdf_path, cpdf_path, metadata_dict=None):
    if task_type == META_LOAD:
        proc = subprocess.run([cpdf_path, "-utf8", "-info", pdf_path], capture_output=True, text=True, encoding='utf-8', errors='ignore')
        if proc.returncode != 0: raise ProcessingError(f"cpdf failed to read info: {proc.stderr.strip()}")
        info = {}
        for line in proc.stdout.splitlines():
            if ':' in line:
                key, val = line.split(':', 1)
                if key.strip().lower() in ['title', 'author', 'subject', 'keywords']: info[key.strip().lower()] = val.strip()
        return info
    elif task_type == META_SAVE:
        with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as temp_out:
            temp_path = temp_out.name
        try:
            cmd = [cpdf_path, "-utf8", pdf_path]
            valid_metadata = {k: v for k, v in metadata_dict.items() if v}
            is_first_op = True
            for key, value in valid_metadata.items():
                if not is_first_op:
                    cmd.append("AND")
                cmd.extend([f"-set-{key}", value])
                is_first_op = False

            if not is_first_op:
                cmd.extend(["-o", temp_path])
                run_command(cmd)
                shutil.move(temp_path, pdf_path)
        finally:
            if os.path.exists(temp_path):
                os.remove(temp_path)

def run_pdf_to_image_task(gs_path, pdf_in, out_dir, options, q):
    try:
        q.put(('status', f"Converting PDF to {options.get('format')}..."))
        fmt, dpi = options.get('format', 'png'), options.get('dpi', '300')
        out_path = Path(out_dir) / f"{Path(pdf_in).stem}_%d.{fmt}"
        device_map = {'png': 'png16m', 'jpeg': 'jpeg', 'tiff': 'tiffg4'}
        cmd = [gs_path, f"-sDEVICE={device_map.get(fmt, 'png16m')}", f"-r{dpi}", "-dNOPAUSE", "-dBATCH", f"-sOutputFile={str(out_path)}", pdf_in]
        run_command(cmd)
        q.put(('progress', 100)); q.put(('overall', 100)); q.put(('complete', "Conversion to images complete."))
    except Exception as e: logging.error(f"PDF to image task failed: {e}"); q.put(('complete', f"Error: {e}"))

def run_repair_task(pdf_in, pdf_out, q):
    try:
        q.put(('status', "Attempting to repair PDF..."))
        with pikepdf.open(pdf_in, allow_overwriting_input=False) as pdf:
            pdf.save(pdf_out)
        q.put(('progress', 100))
        q.put(('overall', 100))
        q.put(('complete', "Repair attempt finished."))
    except Exception as e:
        logging.error(f"Repair task failed: {e}")
        q.put(('complete', f"Error: {e}"))

def run_toc_task(cpdf_path, pdf_in, pdf_out, options, q):
    try:
        q.put(('status', "Generating Table of Contents..."))
        with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as temp_out:
            temp_path = temp_out.name

        cmd = [cpdf_path, "-table-of-contents"]

        if options.get('title'):
            cmd.extend(["-toc-title", options.get('title')])
        if options.get('font'):
            cmd.extend(["-font", options.get('font')])
        if options.get('font_size'):
            cmd.extend(["-font-size", str(options.get('font_size'))])
        if options.get('dot_leaders'):
            cmd.append("-toc-dot-leaders")
        if options.get('no_bookmark'):
            cmd.append("-toc-no-bookmark")

        cmd.extend([pdf_in, "-o", temp_path])
        run_command(cmd)

        if not Path(temp_path).exists() or Path(temp_path).stat().st_size == 0:
            raise ProcessingError("cpdf failed to generate the table of contents. The output file is empty.")

        shutil.move(temp_path, pdf_out)
        q.put(('progress', 100))
        q.put(('overall', 100))
        q.put(('complete', "Table of Contents generation complete."))
    except Exception as e:
        logging.error(f"Table of Contents task failed: {e}", exc_info=True)
        q.put(('complete', f"Error: {e}"))