# pdf_optimizer.py
import subprocess
import logging
from pathlib import Path
import tempfile
import shutil
import pikepdf
import zlib
from PIL import Image
import sys

class PdfOptimizer:
    def __init__(self, gs_path, sam2p_path=None, cpdf_path=None, jbig2_path=None, pngout_path=None, compress_level=1, user_password=None, darken_text=None, remove_open_action=None, q=None):
        self.gs_path = gs_path
        self.sam2p_path = sam2p_path
        self.cpdf_path = cpdf_path
        self.jbig2_path = jbig2_path
        self.pngout_path = pngout_path
        self.compress_level = compress_level
        self.user_password = user_password
        self.darken_text = darken_text
        self.remove_open_action = remove_open_action
        self.q = q

    def _run_command(self, cmd, check=True):
        use_shell = isinstance(cmd, str)
        logging.info(f"Running command: {cmd}")
        try:
            kwargs = {
                'check': check,
                'capture_output': True,
                'text': True,
                'errors': 'ignore',
                'shell': use_shell
            }
            if sys.platform == "win32":
                kwargs['creationflags'] = subprocess.CREATE_NO_WINDOW
            
            result = subprocess.run(cmd, **kwargs)
            if result.stderr and "warning" not in result.stderr.lower():
                 logging.warning(f"Command produced stderr: {result.stderr.strip()}")
            return result
        except subprocess.CalledProcessError as e:
            logging.error(f"Command failed: {e.stderr}"); return None
        except FileNotFoundError:
            logging.error(f"Command not found: {cmd if use_shell else cmd[0]}"); return None

    def _optimize_images(self, pdf, temp_dir):
        logging.info("Starting image optimization phase.")
        images_optimized = 0
        for obj in pdf.objects:
            if isinstance(obj, pikepdf.Stream) and obj.get("/Subtype") == "/Image":
                if obj.get("/Filter") in ("/DCTDecode", "/JPXDecode"): continue
                try:
                    original_stream_data = obj.read_raw_bytes()
                    candidates = {'original': (original_stream_data, obj.Filter, obj.get("/DecodeParms"))}
                    image_obj = pikepdf.PdfImage(obj); pil_image = image_obj.as_pil_image()
                    
                    if pil_image.mode == 'CMYK':
                        pil_image = pil_image.convert('RGB')
                        
                    temp_png_path = temp_dir / f"img_{obj.objgen}.png"; pil_image.save(temp_png_path, "png")

                    if self.sam2p_path:
                        opt_path_recompress = temp_dir / f"img_{obj.objgen}.sam2p_recompress.png"
                        cmd_recompress = [self.sam2p_path, "-j:quiet", "-c", "zip:15:9", str(temp_png_path), str(opt_path_recompress)]
                        if self._run_command(cmd_recompress) and opt_path_recompress.exists() and opt_path_recompress.stat().st_size > 0:
                            candidates['sam2p_recompress'] = (opt_path_recompress.read_bytes(), pikepdf.Name("/FlateDecode"), None)

                        opt_path_png8 = temp_dir / f"img_{obj.objgen}.sam2p_png8.png"
                        cmd_png8 = [self.sam2p_path, "-j:quiet", str(temp_png_path), str(opt_path_png8)]
                        if self._run_command(cmd_png8) and opt_path_png8.exists() and opt_path_png8.stat().st_size > 0:
                            candidates['sam2p_png8'] = (opt_path_png8.read_bytes(), pikepdf.Name("/FlateDecode"), None)

                    if self.compress_level > 0:
                        if self.pngout_path:
                            opt_path_pngout = temp_dir / f"img_{obj.objgen}.pngout.png"
                            cmd_pngout = [self.pngout_path, "-force", "-q", str(temp_png_path), str(opt_path_pngout)]
                            if self._run_command(cmd_pngout) and opt_path_pngout.exists() and opt_path_pngout.stat().st_size > 0:
                                candidates['pngout'] = (opt_path_pngout.read_bytes(), pikepdf.Name("/FlateDecode"), None)

                        if pil_image.mode == '1' and self.jbig2_path:
                            pbm_path = temp_dir / f"img_{obj.objgen}.pbm"; pil_image.save(pbm_path)

                            opt_path_jbig2_p = temp_dir / f"img_{obj.objgen}.jbig2_p"
                            cmd_str_p = f'"{self.jbig2_path}" -p "{pbm_path}" > "{opt_path_jbig2_p}"'
                            if self._run_command(cmd_str_p, check=False) and opt_path_jbig2_p.exists() and opt_path_jbig2_p.stat().st_size > 0:
                                candidates['jbig2_p'] = (opt_path_jbig2_p.read_bytes(), pikepdf.Name("/JBIG2Decode"), None)

                            opt_path_jbig2_s = temp_dir / f"img_{obj.objgen}.jbig2_s"
                            cmd_str_s = f'"{self.jbig2_path}" -s -p "{pbm_path}" > "{opt_path_jbig2_s}"'
                            if self._run_command(cmd_str_s, check=False) and opt_path_jbig2_s.exists() and opt_path_jbig2_s.stat().st_size > 0:
                                candidates['jbig2_s'] = (opt_path_jbig2_s.read_bytes(), pikepdf.Name("/JBIG2Decode"), None)

                    best_method = min(candidates, key=lambda k: len(candidates[k][0]))
                    if best_method != 'original':
                        best_data, best_filter, best_params = candidates[best_method]
                        obj.write(best_data); obj.Filter = best_filter
                        if best_params: obj.DecodeParms = best_params
                        elif '/DecodeParms' in obj: del obj.DecodeParms
                        images_optimized += 1
                        logging.info(f"Optimized image {obj.objgen} using {best_method}.")
                except Exception as e:
                    logging.warning(f"Could not process image {obj.objgen}: {e}")
        logging.info(f"Finished image optimization. Optimized {images_optimized} images.")

    def _recompress_streams(self, pdf):
        logging.info("Recompressing non-image streams.")
        for obj in pdf.objects:
            if isinstance(obj, pikepdf.Stream) and obj.get("/Subtype") != "/Image":
                try:
                    uncompressed = obj.read_bytes(); compressed = zlib.compress(uncompressed, level=9)
                    if len(compressed) < len(obj.read_raw_bytes()):
                        obj.write(compressed); obj.Filter = pikepdf.Name("/FlateDecode")
                        if '/DecodeParms' in obj: del obj.DecodeParms
                except Exception as e:
                    logging.warning(f"Could not recompress stream {obj.objgen}: {e}")

    def _post_process_and_compare(self, candidate_path, temp_dir, current_best_size):
        if not candidate_path.exists() or candidate_path.stat().st_size == 0:
            return None, current_best_size

        try:
            with pikepdf.open(candidate_path, allow_overwriting_input=True) as pdf:
                self._optimize_images(pdf, temp_dir)
                self._recompress_streams(pdf)
                pdf.save(object_stream_mode=pikepdf.ObjectStreamMode.generate, recompress_flate=True, stream_decode_level=pikepdf.StreamDecodeLevel.all)
        except Exception as e:
            logging.warning(f"Post-processing failed for {candidate_path.name}: {e}")
            return None, current_best_size

        candidate_size = candidate_path.stat().st_size
        if candidate_size < current_best_size:
            logging.info(f"New best size found: {candidate_size} bytes with {candidate_path.name}")
            return candidate_path, candidate_size
        else:
            logging.info(f"Candidate {candidate_path.name} ({candidate_size} bytes) did not beat current best ({current_best_size} bytes).")
            return None, current_best_size

    def optimize_lossless(self, input_file, output_file, strip_metadata=False):
        logging.info(f"Starting SMART lossless optimization for {input_file}")
        with tempfile.TemporaryDirectory() as temp_dir_str:
            temp_dir = Path(temp_dir_str)
            original_size = Path(input_file).stat().st_size
            current_best_file = Path(input_file)
            current_best_size = original_size

            recipes = [
                {'type': 'pikepdf', 'name': 'Pikepdf Stream Recompression'},
            ]
            if self.cpdf_path:
                recipes.insert(0, {'type': 'cpdf', 'name': 'cpdf Squeeze'})
            if self.gs_path:
                recipes.append({'type': 'gs', 'name': 'Ghostscript Rebuild'})

            total_recipes = len(recipes)
            for i, recipe in enumerate(recipes):
                if self.q: self.q.put(('overall', ((i + 1) / total_recipes) * 100)); self.q.put(('status', f"Attempt {i+1}/{total_recipes}: {recipe['name']}..."))
                
                temp_candidate_path = temp_dir / f"lossless_attempt_{i+1}.pdf"
                cmd_success = False

                source_file = current_best_file if recipe['type'] != 'pikepdf' else input_file

                if recipe['type'] == 'cpdf':
                    cmd = [self.cpdf_path, "-squeeze", str(source_file), "-o", str(temp_candidate_path)]
                    if self._run_command(cmd): cmd_success = True
                elif recipe['type'] == 'gs':
                    cmd = [self.gs_path, "-sDEVICE=pdfwrite", "-dCompatibilityLevel=1.4", "-dNOPAUSE", "-dBATCH", "-dQUIET", f"-sOutputFile={temp_candidate_path}", str(source_file)]
                    if self._run_command(cmd): cmd_success = True
                elif recipe['type'] == 'pikepdf':
                    shutil.copy(source_file, temp_candidate_path)
                    try:
                        with pikepdf.open(temp_candidate_path, allow_overwriting_input=True) as pdf:
                            self._optimize_images(pdf, temp_dir)
                            self._recompress_streams(pdf)
                            pdf.save(object_stream_mode=pikepdf.ObjectStreamMode.generate, recompress_flate=True, stream_decode_level=pikepdf.StreamDecodeLevel.all)
                        cmd_success = True
                    except Exception as e:
                        logging.warning(f"Pikepdf optimization failed: {e}")

                if cmd_success and temp_candidate_path.exists():
                    candidate_size = temp_candidate_path.stat().st_size
                    if candidate_size < current_best_size:
                        logging.info(f"New best size found: {candidate_size} bytes with {recipe['name']}")
                        if current_best_file != Path(input_file):
                            current_best_file.unlink()
                        current_best_file = temp_candidate_path
                        current_best_size = candidate_size
                    else:
                        temp_candidate_path.unlink()

            if strip_metadata and current_best_file != Path(input_file):
                try:
                    with pikepdf.open(current_best_file, allow_overwriting_input=True) as pdf:
                        if '/Metadata' in pdf.Root:
                            del pdf.Root.Metadata
                            pdf.save()
                except Exception: pass
            
            self._finalize_pdf(current_best_file, output_file)

            if Path(output_file).resolve() != Path(input_file).resolve() and current_best_file == Path(input_file):
                shutil.copy(input_file, output_file)
            
            if self.q: self.q.put(('overall', 100))

    def optimize_lossy(self, input_file, output_file, dpi, strip_metadata=False, remove_interactive=False):
        logging.info(f"Starting SMART lossy optimization for {input_file} with DPI {dpi}")

        with tempfile.TemporaryDirectory() as temp_dir_str:
            temp_dir = Path(temp_dir_str)
            original_size = Path(input_file).stat().st_size
            current_best_file = Path(input_file)
            current_best_size = original_size
            
            dpi_val = int(dpi)
            
            recipes = []
            
            base_recipes = []
            if self.cpdf_path:
                base_recipes.append(
                    {'type': 'cpdf', 'name': 'cpdf Squeeze'}
                )
            base_recipes.extend([
                {'type': 'gs', 'name': 'Aggressive (/screen)', 'preset': '/screen', 'downsample': 'Average', 'dpi': str(dpi_val)},
                {'type': 'gs', 'name': 'Balanced (/ebook)', 'preset': '/ebook', 'downsample': 'Bicubic', 'dpi': str(dpi_val)},
                {'type': 'gs', 'name': 'High Quality (/printer)', 'preset': '/printer', 'downsample': 'Bicubic', 'dpi': str(dpi_val)}
            ])
            
            if self.compress_level == 0: # Fast
                recipes = [
                    {'type': 'cpdf', 'name': 'Quick Squeeze'},
                    {'type': 'gs', 'name': 'Balanced Downsampling', 'preset': '/ebook', 'downsample': 'Bicubic', 'dpi': str(dpi_val)}
                ] if self.cpdf_path else [
                    {'type': 'gs', 'name': 'Balanced Downsampling', 'preset': '/ebook', 'downsample': 'Bicubic', 'dpi': str(dpi_val)}
                ]
            elif self.compress_level == 1: # Normal
                recipes = base_recipes
            elif self.compress_level == 2: # Deep
                recipes = base_recipes + [{'type': 'lossless', 'name': 'Deep Clean Optimization'}]

            total_recipes = len(recipes)
            for i, recipe in enumerate(recipes):
                if self.q: self.q.put(('overall', (i / total_recipes) * 100)); self.q.put(('status', f"Attempt {i+1}/{total_recipes}: {recipe['name']}..."))

                temp_candidate_path = temp_dir / f"attempt_{i+1}.pdf"
                cmd_success = False

                if recipe['type'] == 'cpdf':
                    cmd = [self.cpdf_path, str(current_best_file), "-squeeze", "-o", str(temp_candidate_path)]
                    if self._run_command(cmd): cmd_success = True
                
                elif recipe['type'] == 'gs':
                    cmd = [
                        self.gs_path, "-sDEVICE=pdfwrite", "-dCompatibilityLevel=1.4",
                        "-dNOPAUSE", "-dBATCH", "-dQUIET", "-dSubsetFonts=true", "-dCompressFonts=true",
                        f"-dPDFSETTINGS={recipe['preset']}",
                        f"-dColorImageDownsampleType=/{recipe['downsample']}", f"-dGrayImageDownsampleType=/{recipe['downsample']}",
                        f"-dColorImageResolution={recipe['dpi']}", f"-dGrayImageResolution={recipe['dpi']}",
                        f"-sOutputFile={temp_candidate_path}", str(input_file)
                    ]
                    if remove_interactive: cmd.extend(["-dShowAnnots=false", "-dShowAcroForm=false"])
                    if self._run_command(cmd): cmd_success = True
                
                elif recipe['type'] == 'lossless':
                    shutil.copy(input_file, temp_candidate_path)
                    try:
                        with pikepdf.open(temp_candidate_path, allow_overwriting_input=True) as pdf:
                            self._optimize_images(pdf, temp_dir)
                            self._recompress_streams(pdf)
                            pdf.save(object_stream_mode=pikepdf.ObjectStreamMode.generate, recompress_flate=True, stream_decode_level=pikepdf.StreamDecodeLevel.all)
                        cmd_success = True
                    except Exception as e:
                        logging.warning(f"Lossless recipe step failed: {e}")

                if cmd_success and temp_candidate_path.exists() and temp_candidate_path.stat().st_size > 0:
                    new_best_path, new_best_size = self._post_process_and_compare(temp_candidate_path, temp_dir, current_best_size)
                    if new_best_path:
                        if current_best_file != Path(input_file) and current_best_file.exists():
                           current_best_file.unlink()
                        current_best_file = new_best_path
                        current_best_size = new_best_size
            
            if strip_metadata and current_best_file != Path(input_file):
                try:
                    with pikepdf.open(current_best_file, allow_overwriting_input=True) as pdf:
                        if '/Metadata' in pdf.Root:
                            del pdf.Root.Metadata
                            pdf.save()
                except Exception as e:
                    logging.warning(f"Could not strip metadata from best candidate: {e}")

            self._finalize_pdf(current_best_file, output_file)
            
            if Path(output_file).resolve() != Path(input_file).resolve() and current_best_file == Path(input_file):
                 shutil.copy(input_file, output_file)
            
            if self.q: self.q.put(('overall', 100))

    def _finalize_pdf(self, optimized_candidate_path, final_output_path):
        current_best_file = optimized_candidate_path
        
        temp_files_to_clean = []
        def get_temp_path():
            f = tempfile.NamedTemporaryFile(suffix=".pdf", delete=False)
            f.close()
            temp_files_to_clean.append(f.name)
            return f.name

        if self.cpdf_path and optimized_candidate_path.exists():
            if self.darken_text:
                dark_path = get_temp_path()
                cmd = [self.cpdf_path, str(current_best_file), "-blacktext", "-o", dark_path]
                if self._run_command(cmd) and Path(dark_path).exists() and Path(dark_path).stat().st_size > 0: current_best_file = dark_path
                else: logging.warning("cpdf blacktext failed.")
            if self.remove_open_action:
                open_action_path = get_temp_path()
                cmd = [self.cpdf_path, str(current_best_file), "-remove-dict-entry", "/OpenAction", "-o", open_action_path]
                if self._run_command(cmd) and Path(open_action_path).exists() and Path(open_action_path).stat().st_size > 0: current_best_file = open_action_path
                else: logging.warning("cpdf remove open action failed.")
            
            squeeze_path = get_temp_path()
            cmd = [self.cpdf_path, str(current_best_file), "-squeeze", "-o", squeeze_path]
            if self._run_command(cmd) and Path(squeeze_path).exists() and Path(squeeze_path).stat().st_size > 0:
                current_best_file = squeeze_path
            else:
                logging.warning("cpdf squeeze failed.")

        if self.user_password:
            logging.info(f"Applying final AES-256 encryption.")
            try:
                with pikepdf.open(current_best_file) as pdf:
                    pdf.save(final_output_path, encryption=pikepdf.Encryption(user=self.user_password, owner=self.user_password, R=6), object_stream_mode=pikepdf.ObjectStreamMode.generate, recompress_flate=True, stream_decode_level=pikepdf.StreamDecodeLevel.all)
            except Exception as e: logging.error(f"Pikepdf encryption failed: {e}"); raise RuntimeError(f"Failed to apply password: {e}")
        else:
            if Path(current_best_file).exists():
                shutil.move(str(current_best_file), final_output_path)
            
        for f in temp_files_to_clean:
            try:
                Path(f).unlink(missing_ok=True)
            except Exception:
                pass