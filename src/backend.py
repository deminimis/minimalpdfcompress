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
                       META_LOAD, META_SAVE, ProcessingError)

from utils import (resource_path, find_ghostscript, find_cpdf, find_pngquant,
                   find_jpegoptim, find_ect, find_oxipng, format_size,
                   get_pdf_metadata, run_command)

from contextlib import contextmanager

@contextmanager
def task_context(q, success_msg="Task complete.", error_prefix="Task failed"):
    """Wraps background tasks to handle standard queue updates and exception logging."""
    try:
        yield
        if success_msg:
            q.put(('progress', 100))
            q.put(('overall', 100))
            q.put(('complete', success_msg))
    except Exception as e:
        logging.error(f"{error_prefix}: {e}", exc_info=True)
        q.put(('complete', f"Error: {e}"))

def _prepare_image_stamp(image_path, scale, opacity, output_path):
    """Helper to resize and apply opacity to image stamps."""
    with Image.open(image_path) as img:
        if scale != 1.0:
            try:
                w, h = int(img.width * scale), int(img.height * scale)
                if w > 0 and h > 0:
                    resample = Image.Resampling.LANCZOS if hasattr(Image, 'Resampling') else Image.ANTIALIAS
                    img = img.resize((w, h), resample)
            except Exception as e:
                logging.warning(f"Invalid image scale. Error: {e}")
        if opacity < 1.0:
            if img.mode != 'RGBA': img = img.convert('RGBA')
            alpha = img.split()[3].point(lambda p: p * opacity)
            img.putalpha(alpha)
        img.save(output_path, "PDF", resolution=100.0)

def get_cpdf_pos_cmd(pos, margin="20", default=None):
    """Centralized position mapping for cpdf text and stamp operations."""
    pos_map = {
        POS_TOP_LEFT: ["-topleft", margin], POS_TOP_CENTER: ["-top", margin], POS_TOP_RIGHT: ["-topright", margin],
        POS_MIDDLE_LEFT: ["-left", margin], POS_CENTER: ["-center"], POS_MIDDLE_RIGHT: ["-right", margin],
        POS_BOTTOM_LEFT: ["-bottomleft", margin], POS_BOTTOM_CENTER: ["-bottom", margin], POS_BOTTOM_RIGHT: ["-bottomright", margin],
    }
    return pos_map.get(pos, default or ["-center"])

def _update_progress(q, status_msg, current, total):
    """Helper to standardize progress queue updates."""
    prog = (current / total) * 100
    q.put(('status', status_msg))
    q.put(('progress', prog))
    q.put(('overall', prog))

def parse_page_ranges(page_string, max_pages):
    indices = set()
    if not page_string.strip(): return []
    for p in re.sub(r'\bend\b', str(max_pages), page_string, flags=re.IGNORECASE).split(','):
        if not (p := p.strip()): continue
        if '-' in p:
            s, e = [x.strip() for x in p.split('-', 1)]
            indices.update(range((int(s) if s else 1) - 1, int(e) if e else max_pages))
        else: indices.add(int(p) - 1)
    return sorted(list(indices), reverse=True)

def get_total_output_size(output_folder_path, processed_filenames):
    folder = Path(output_folder_path)
    if not folder.is_dir(): return 0
    return sum((folder / f).stat().st_size for f in processed_filenames if (folder / f).is_file())


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
            with pikepdf.open(str(pdf_path)) as pdf:
                with pikepdf.Pdf.new() as new_pdf:
                    new_pdf.pages.append(pdf.pages[0])
                    new_pdf.save(str(first_page_pdf))
            if not first_page_pdf.exists() or first_page_pdf.stat().st_size == 0:
                logging.warning("PDF has no pages or extraction failed.")
                return None
        except Exception as e:
            logging.error(f"Failed to extract first page for preview using pikepdf: {e}")
            return None

        if operation == 'rotate':
            try:
                angle = options.get('angle', 0)
                run_command([cpdf_path, str(first_page_pdf), "-rotate", str(angle), "-o", str(modified_pdf)])
            except Exception as e:
                logging.error(f"Failed to apply rotation for preview: {e}")

        elif operation == 'stamp':
            try:
                stamp_opts = options['stamp_opts']
                mode_opts = options['mode_opts']
                mode = options['mode']

                pos_cmd = get_cpdf_pos_cmd(stamp_opts['pos'], "20", ["-center"])    

                if mode == STAMP_IMAGE:
                    image_path = mode_opts.get('image_path')
                    if image_path and Path(image_path).exists():
                        stamp_to_apply_path = temp_dir_path / "image_stamp.pdf"

                        scale = mode_opts.get('image_scale', 1.0)

                        _prepare_image_stamp(image_path, scale, stamp_opts['opacity'], stamp_to_apply_path)

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

        elif operation == 'page_number':
            try:
                cmd = [cpdf_path, "-utf8"]
                
                preview_text = options.get('text', '').replace('%Page', '1').replace('%EndPage', '1') # Simulate first page of 1
                cmd.extend(["-add-text", preview_text])
                cmd.extend(["-font", options.get('font', 'Helvetica')])
                cmd.extend(["-font-size", str(options.get('font_size', '12'))])
                cmd.extend(["-color", options.get('color', '0 0 0')])

                pos_cmd = get_cpdf_pos_cmd(options.get('pos', POS_BOTTOM_CENTER), "15", ["-bottom", "15"])
                cmd.extend(pos_cmd)

                cmd.append(str(first_page_pdf))
                
                page_range = options.get('page_range', '').strip()
                should_apply = True
                if page_range:
                    try:
                        # Re-use existing parse_page_ranges logic (pass dummy max_pages=9999). 
                        # If index 0 (page 1) isn't in the parsed range, we skip applying the preview.
                        if 0 not in parse_page_ranges(page_range, 9999):
                            should_apply = False
                    except Exception:
                        logging.warning("Could not parse page range for preview, applying anyway.")
                
                if should_apply:
                    cmd.extend(["-o", str(modified_pdf)])
                    run_command(cmd)
                else:
                    logging.info("Page range doesn't include page 1, skipping preview modification.")


            except Exception as e:
                logging.error(f"Failed to apply page number/header/footer for preview: {e}", exc_info=True)


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


def run_compress_task(params, mode, q):
    with task_context(q, success_msg=None, error_prefix="Compress task failed"):
        optimizer = PdfOptimizer(
            gs_path=params['gs_path'],
            cpdf_path=params['cpdf_path'],
            pngquant_path=params['pngquant_path'],
            jpegoptim_path=params['jpegoptim_path'],
            ect_path=params['ect_path'],
            oxipng_path=params.get('oxipng_path'),
            q=q,
            darken_text=params['darken_text'],
            remove_open_action=params.get('remove_open_action'),
            fast_web_view=params.get('fast_web_view'),
            fast_mode=params.get('fast_mode'),
            safe_mode=params.get('safe_mode'),
            lossless_encoding=params.get('lossless_encoding', False),
            preserve_ocr=params.get('preserve_ocr', True),
            detect_duplicate_images=params.get('detect_duplicate_images', True),
            convert_to_grayscale=params.get('convert_to_grayscale', False),
            convert_to_cmyk=params.get('convert_to_cmyk', False),
            downsample_threshold_enabled=params.get('downsample_threshold_enabled', False),
            quantize_colors=params.get('quantize_colors', False),
            quantize_level=params.get('quantize_level', 4),
            pdfa_compression=params.get('pdfa_compression', False),
            pdfa_dpi=params.get('pdfa_dpi', 300)
        )

        output_path = Path(params['output_path'])
        total_in_size = 0
        files_skipped = 0
        processed_filenames = []

        def process_a_file(input_file, output_file):
            nonlocal files_skipped
            compression_mode = params.get('mode', 'Lossy')
            only_if_smaller = params.get('only_if_smaller', False)
            original_size = 0
            try:
                original_size = input_file.stat().st_size
            except FileNotFoundError:
                 logging.error(f"Input file not found: {input_file}")
                 raise ProcessingError(f"Input file not found: {input_file.name}")


            with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as temp_out:
                temp_output_path = Path(temp_out.name)

            try:
                if compression_mode == 'Lossless':
                    if params.get('true_lossless', False):
                        optimizer.optimize_true_lossless(input_file, temp_output_path, strip_metadata=params['strip_metadata'])
                    else:
                        optimizer.optimize_lossless(input_file, temp_output_path, strip_metadata=params['strip_metadata'])
                elif compression_mode == 'PDF/A':
                    optimizer.optimize_pdfa(input_file, temp_output_path)
                elif compression_mode == 'Remove Images':
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
                    processed_filenames.append(output_file.name)
                    return original_size

                new_size = temp_output_path.stat().st_size

                if compression_mode == 'PDF/A':
                    shutil.move(str(temp_output_path), output_file)
                    processed_filenames.append(output_file.name)
                    return new_size

                if new_size < original_size:
                    shutil.move(str(temp_output_path), output_file)
                    processed_filenames.append(output_file.name)

                    if params.get('delete_original') and input_file.resolve() != output_file.resolve():
                        try:
                            os.remove(input_file)
                            logging.info(f"Deleted original file: {input_file.name}")
                        except Exception as del_err:
                            logging.error(f"Failed to delete original file {input_file.name}: {del_err}")

                    return new_size
                else:
                    if only_if_smaller:
                        files_skipped += 1
                        logging.info(f"Skipping save for {input_file.name}, new size {new_size} >= original size {original_size}")
                        return 0 # Indicate no output size contribution
                    else:
                        shutil.copy2(input_file, output_file)
                        processed_filenames.append(output_file.name)
                        logging.info(f"Saving original for {input_file.name}, new size {new_size} >= original size {original_size}")
                        return original_size

            except Exception as proc_err:
                 logging.error(f"Error processing {input_file.name}: {proc_err}", exc_info=True)
                 try:
                     if not output_file.exists():
                          shutil.copy2(input_file, output_file)
                          processed_filenames.append(output_file.name)
                          logging.info(f"Copied original {input_file.name} due to processing error.")
                          return original_size # Count original size if copied
                 except Exception as copy_err:
                      logging.error(f"Could not copy original {input_file.name} after error: {copy_err}")
                 raise proc_err
            finally:
                if temp_output_path.exists():
                    os.remove(temp_output_path)
            return 0 # Return 0 size contribution if error wasn't handled by copying original

        pdf_files_paths = [Path(f) for f in params['input_files']]
        total_in_size = sum(f.stat().st_size for f in pdf_files_paths if f.is_file())
        output_path.mkdir(parents=True, exist_ok=True)
        total_files = len(pdf_files_paths)
        if total_files == 0: raise ProcessingError(f"No PDF files found in list.")

        errors_occurred = 0
        total_out_size = 0
        for i, pdf_file in enumerate(pdf_files_paths):
            output_file_path = output_path / pdf_file.name
            q.put(('status', f"Processing {pdf_file.name} ({i+1}/{total_files})..."))
            try:
                size = process_a_file(pdf_file, output_file_path)
                total_out_size += size
            except Exception:
                errors_occurred += 1
                q.put(('status', f"Error processing {pdf_file.name} ({i+1}/{total_files})..."))
            finally:
                q.put(('overall', ((i + 1) / total_files) * 100))

        final_message = "Processing complete."
        if errors_occurred > 0:
            final_message = f"Processing finished with {errors_occurred} error(s)."


        if total_in_size > 0:
            saved_bytes = total_in_size - total_out_size
            percent_saved = (saved_bytes / total_in_size) * 100
            
            saved_str = format_size(saved_bytes, decimals=2)

            if saved_bytes > 0:
                final_message = f"Complete. Saved {saved_str} ({percent_saved:.1f}%)."
            elif saved_bytes < 0:
                final_message = f"Complete. Size increased by {saved_str.lstrip('-')} ({-percent_saved:.1f}%)."
            else:
                 final_message = "Complete. No change in size."

            if errors_occurred > 0:
                 final_message += f" ({errors_occurred} error(s))"


        elif errors_occurred == 0:
            final_message = "Processing complete. Size comparison not available."

        if files_skipped > 0:
            final_message += f" ({files_skipped} file(s) not saved as output was larger)."

        q.put(('complete', final_message))


def run_merge_task(file_list, output_path, q):
    with task_context(q, "Merge complete.", "Merge task failed"):
        with pikepdf.Pdf.new() as pdf:
            total_files = len(file_list)
            if total_files == 0: raise ProcessingError("No files selected to merge.")
            for i, file_path in enumerate(file_list):
                q.put(('status', f"Adding {Path(file_path).name} ({i+1}/{total_files})"))
                with pikepdf.open(file_path) as src:
                    pdf.pages.extend(src.pages)
                q.put(('progress', ((i + 1) / total_files) * 100))
            q.put(('status', "Saving merged file..."))
            pdf.save(output_path)

def run_split_task(input_path, output_dir, mode, value, q):
    with task_context(q, "Splitting complete.", "Split task failed"):
        q.put(('status', "Opening PDF..."))
        p_in, output_dir_path = Path(input_path), Path(output_dir)
        output_dir_path.mkdir(parents=True, exist_ok=True)
        with pikepdf.open(p_in) as pdf:
            total = len(pdf.pages)
            if mode == SPLIT_SINGLE:
                for i, page in enumerate(pdf.pages):
                    _update_progress(q, f"Saving page {i+1}/{total}", i + 1, total)
                    with pikepdf.Pdf.new() as dst: 
                        dst.pages.append(page)
                        dst.save(output_dir_path / f"{p_in.stem}_page_{i+1}.pdf")
            elif mode == SPLIT_EVERY_N:
                n = int(value)
                if n <= 0: raise ValueError("Number of pages must be positive.")
                for i in range(0, total, n):
                    _update_progress(q, f"Saving pages {i+1}-{min(i+n, total)}", min(i + n, total), total)
                    with pikepdf.Pdf.new() as dst: 
                        dst.pages.extend(pdf.pages[i:i+n])
                        dst.save(output_dir_path / f"{p_in.stem}_pages_{i+1}-{i+n}.pdf")
            elif mode == SPLIT_CUSTOM:
                indices = parse_page_ranges(value, total)
                q.put(('status', f"Extracting {len(indices)} pages..."))
                with pikepdf.Pdf.new() as dst:
                    for i, page_index in enumerate(sorted([idx for idx in indices if idx < total])):
                        _update_progress(q, f"Extracting {len(indices)} pages...", i + 1, len(indices))
                        dst.pages.append(pdf.pages[page_index])
                dst.save(output_dir_path / f"{p_in.stem}_custom_range.pdf")
            else: raise ProcessingError(f"Unknown split mode: {mode}")

def run_delete_pages_task(pdf_in, pdf_out, page_range, q):
    with task_context(q, "Page deletion completed.", "Delete pages task failed"):
        q.put(('status', "Opening PDF..."))
        with pikepdf.open(pdf_in) as pdf:
            indices = parse_page_ranges(page_range, len(pdf.pages))
            if not indices: raise ProcessingError("No valid pages specified for deletion.")
            q.put(('status', f"Deleting {len(indices)} page(s)..."))
            for i in sorted(indices, reverse=True):
                 if 0 <= i < len(pdf.pages): del pdf.pages[i]
                 else: logging.warning(f"Page index {i+1} out of range, skipping deletion.")
            pdf.save(pdf_out)

def run_rotate_task(pdf_in, pdf_out, angle, q):
    with task_context(q, "Rotation complete.", "Rotate task failed"):
        q.put(('status', "Opening PDF..."))
        with pikepdf.open(pdf_in) as pdf:
            q.put(('status', f"Rotating all pages by {angle} degrees..."))
            for page in pdf.pages: page.rotate(angle, relative=True)
            pdf.save(pdf_out)

def run_stamp_task(pdf_in, pdf_out, stamp_opts, cpdf_path, q, mode, mode_opts):
    with task_context(q, "Stamping complete.", "Stamp task failed"):
        with tempfile.TemporaryDirectory() as temp_dir_str:
            temp_dir = Path(temp_dir_str)
            q.put(('status', "Applying stamp..."))
            pos_cmd = get_cpdf_pos_cmd(stamp_opts['pos'], "20", ["-center"])

            if mode == STAMP_IMAGE:
                image_file = Path(mode_opts['image_path'])
                if not image_file.exists(): raise ProcessingError(f"Stamp image not found: {image_file}")
                stamp_to_apply_path = temp_dir / "image_stamp.pdf"
                scale = mode_opts.get('image_scale', 1.0)

                _prepare_image_stamp(image_file, scale, stamp_opts['opacity'], stamp_to_apply_path)

                cmd = [cpdf_path, pdf_in, "-stamp-on" if stamp_opts['on_top'] else "-stamp-under", str(stamp_to_apply_path)]
                cmd.extend(pos_cmd)
                cmd.extend(["-o", pdf_out])
                run_command(cmd)

            else:
                final_text = mode_opts['text'].strip()
                if not final_text and not mode_opts.get('bates_start'): raise ProcessingError("Stamp text cannot be empty.")

                cmd = [cpdf_path, pdf_in]
                text_parts = ["-add-text", final_text, "-font", mode_opts['font'], "-font-size", str(mode_opts['size']), "-color", mode_opts['color'], "-opacity", str(stamp_opts['opacity'])]
                
                if mode_opts.get('bates_start'):
                    try:
                        bates_num = int(mode_opts['bates_start'])
                        if bates_num < 0: raise ValueError
                        text_parts.extend(["-bates", str(bates_num)])
                    except ValueError:
                         raise ProcessingError("Invalid Bates start number. Must be a non-negative integer.")

                cmd.extend(text_parts)
                if not stamp_opts['on_top']: cmd.append("-underneath")
                cmd.extend(pos_cmd)
                cmd.extend(["-o", pdf_out])
                run_command(cmd)

def run_page_number_task(pdf_in, pdf_out, cpdf_path, q, options):
    with task_context(q, "Header/Footer task complete.", "Page Number task failed"):
        q.put(('status', "Adding page numbers/headers/footers..."))
        cmd = [cpdf_path, "-utf8", "-add-text", options['text'], "-font", options['font'], "-font-size", str(options['font_size']), "-color", options['color']]
        cmd.extend(get_cpdf_pos_cmd(options['pos'], "15", ["-bottom", "15"]))
        cmd.append(pdf_in)
        
        page_range = options.get('page_range', '').strip()
        if page_range: cmd.append(page_range)
        cmd.extend(["-o", pdf_out])
        run_command(cmd)

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
        with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as temp_out: temp_path = temp_out.name
        try:
            cmd = [cpdf_path, "-utf8", pdf_path]
            valid_metadata = {k: v for k, v in metadata_dict.items() if v}
            is_first_op = True
            for key, value in valid_metadata.items():
                if not is_first_op: cmd.append("AND")
                cmd.extend([f"-set-{key}", value])
                is_first_op = False

            if not is_first_op:
                cmd.extend(["-o", temp_path])
                run_command(cmd)
                shutil.move(temp_path, pdf_path)
            else:
                 logging.info("No metadata changes specified, file not modified.")
        finally:
            if os.path.exists(temp_path): os.remove(temp_path)

def run_pdf_to_image_task(gs_path, pdf_in, out_dir, options, q):
    with task_context(q, "Conversion to images complete.", "PDF to image task failed"):
        q.put(('status', f"Converting PDF to {options.get('format')}..."))
        fmt, dpi = options.get('format', 'png'), options.get('dpi', '300')
        output_dir_path = Path(out_dir)
        output_dir_path.mkdir(parents=True, exist_ok=True)
        out_path = output_dir_path / f"{Path(pdf_in).stem}_%d.{fmt}"
        device_map = {'png': 'png16m', 'jpeg': 'jpeg', 'tiff': 'tiffg4'}
        cmd = [gs_path, f"-sDEVICE={device_map.get(fmt, 'png16m')}", f"-r{dpi}", "-dNOPAUSE", "-dBATCH", f"-sOutputFile={str(out_path)}", pdf_in]
        run_command(cmd)

def run_repair_task(pdf_in, pdf_out, q):
    with task_context(q, "Repair attempt finished.", "Repair task failed"):
        q.put(('status', "Attempting to repair PDF..."))
        with pikepdf.open(pdf_in, allow_overwriting_input=False, recover=True) as pdf:
            pdf.save(pdf_out)

def run_toc_task(cpdf_path, pdf_in, pdf_out, options, q):
    with task_context(q, "Table of Contents generation complete.", "Table of Contents task failed"):
        q.put(('status', "Generating Table of Contents..."))
        with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as temp_out: temp_path = temp_out.name
        cmd = [cpdf_path, "-table-of-contents"]
        if options.get('title'): cmd.extend(["-toc-title", options.get('title')])
        if options.get('font'): cmd.extend(["-font", options.get('font')])
        if options.get('font_size'): cmd.extend(["-font-size", str(options.get('font_size'))])
        if options.get('dot_leaders'): cmd.append("-toc-dot-leaders")
        if options.get('no_bookmark'): cmd.append("-toc-no-bookmark")
        cmd.extend([pdf_in, "-o", temp_path])
        run_command(cmd)

        if not Path(temp_path).exists() or Path(temp_path).stat().st_size == 0:
            raise ProcessingError("cpdf failed to generate the table of contents. The output file is empty.")
        shutil.move(temp_path, pdf_out)

def run_password_task(params, q):
    # Setting success_msg to None lets this task dictate its own completion texts
    with task_context(q, success_msg=None, error_prefix="Password task failed"): 
        input_path = params.get('input_path')
        output_path = params.get('output_path')
        mode = params.get('mode')

        if mode == 'add':
            q.put(('status', "Encrypting PDF..."))
            user_password, owner_password = params.get('user_password'), params.get('owner_password')

            if not user_password and not owner_password:
                raise ProcessingError("At least one password (user or owner) must be provided for encryption.")

            permissions = pikepdf.Permissions(print_highres=params.get('allow_printing'), print_lowres=params.get('allow_printing'), modify_other=params.get('allow_modification'), extract=params.get('allow_copy_and_extract'), modify_annotation=params.get('allow_annotations_and_forms'), modify_form=params.get('allow_annotations_and_forms'))

            with pikepdf.open(input_path) as pdf:
                pdf.save(output_path, encryption=pikepdf.Encryption(user=user_password, owner=owner_password, allow=permissions, R=6))
            
            q.put(('progress', 100)); q.put(('overall', 100)); q.put(('complete', "Encryption complete."))

        elif mode == 'remove':
            q.put(('status', "Decrypting PDF..."))
            password_provided = params.get('user_password')

            try:
                with pikepdf.open(input_path, allow_overwriting_input=True) as pdf:
                    if pdf.is_encrypted:
                        pdf.save(output_path)
                        q.put(('progress', 100)); q.put(('overall', 100)); q.put(('complete', "Decryption complete (owner password removed or none required)."))
                    else:
                        shutil.copy2(input_path, output_path)
                        q.put(('progress', 100)); q.put(('overall', 100)); q.put(('complete', "Info: This PDF is not encrypted."))

            except pikepdf.PasswordError:
                if not password_provided:
                    raise ProcessingError("This PDF requires a user password to open. Please provide it.")
                try:
                    with pikepdf.open(input_path, password=password_provided, allow_overwriting_input=True) as pdf:
                        pdf.save(output_path)
                        q.put(('progress', 100)); q.put(('overall', 100)); q.put(('complete', "Decryption complete."))
                except pikepdf.PasswordError:
                    raise ProcessingError("Wrong password provided.")
        else:
            raise ProcessingError(f"Unknown password mode: {mode}")