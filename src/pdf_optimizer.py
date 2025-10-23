# pdf_optimizer.py
import subprocess
import logging
from pathlib import Path
import tempfile
import shutil
import pikepdf
from PIL import Image
from io import BytesIO
import sys
import os
import oxipng
import zlib

from constants import ProcessingError

def resource_path(relative_path):
    try:
        base_path = Path(sys._MEIPASS)
    except AttributeError:
        base_path = Path(__file__).parent
    return base_path / relative_path

class PdfOptimizer:
    _blank_image_data = b'\xff\xff\xff'

    def __init__(self, gs_path, cpdf_path, pngquant_path, q=None, **kwargs):
        self.gs_path = gs_path
        self.cpdf_path = cpdf_path
        self.pngquant_path = pngquant_path
        self.jpegoptim_path = kwargs.get('jpegoptim_path')
        self.ect_path = kwargs.get('ect_path')
        self.q = q
        self.darken_text = kwargs.get('darken_text')
        self.remove_open_action = kwargs.get('remove_open_action')
        self.linearize = kwargs.get('fast_web_view', False)
        self.fast_mode = kwargs.get('fast_mode', False)
        self.safe_mode = kwargs.get('safe_mode', False)
        self.convert_to_grayscale = kwargs.get('convert_to_grayscale', False)
        self.convert_to_cmyk = kwargs.get('convert_to_cmyk', False)
        self.downsample_threshold_enabled = kwargs.get('downsample_threshold_enabled', False)
        self.quantize_colors = kwargs.get('quantize_colors', False)
        self.quantize_level = kwargs.get('quantize_level', 4)
        self.pdfa_compression = kwargs.get('pdfa_compression', False)
        self.pdfa_dpi = kwargs.get('pdfa_dpi', 300)

    def _run_command(self, cmd, check=True):
        use_shell = isinstance(cmd, str)
        logging.info(f"Running command: {cmd}")
        try:
            kwargs = {
                'check': check, 'stdout': subprocess.PIPE, 'stderr': subprocess.PIPE,
                'text': True, 'errors': 'ignore', 'shell': use_shell
            }
            if sys.platform == "win32":
                kwargs['creationflags'] = subprocess.CREATE_NO_WINDOW

            result = subprocess.run(cmd, **kwargs)
            if result.stderr and "warning" not in result.stderr.lower():
                # Ignore specific recoverable warnings if needed
                stderr_lower = result.stderr.lower()
                if "invalid xref" in stderr_lower or "repaired" in stderr_lower:
                    logging.info(f"Ignoring recoverable GS warning: {result.stderr.strip()}")
                else:
                    logging.warning(f"Command produced stderr: {result.stderr.strip()}")
            return result
        except subprocess.CalledProcessError as e:
            full_error = f"Command failed.\nSTDOUT: {e.stdout.strip()}\nSTDERR: {e.stderr.strip()}"
            logging.error(full_error)
            # Raise exception to be caught by the calling optimize function
            raise ProcessingError(f"Tool failed: {e.stderr.strip()}")
        except FileNotFoundError as e:
            logging.error(f"Command not found: {cmd if use_shell else cmd[0]}")
            raise ProcessingError(f"Command not found: {e}")
        except Exception as e:
            logging.error(f"Unexpected error running command: {e}")
            raise ProcessingError(str(e))


    def _replace_image_stream(self, obj):
        try:
            if not isinstance(obj, pikepdf.Stream) or obj.get("/Subtype") != "/Image":
                return

            # Keep essential keys, remove others that might prevent blanking
            essential_keys = {pikepdf.Name.Type, pikepdf.Name.Subtype, pikepdf.Name.Width,
                            pikepdf.Name.Height, pikepdf.Name.ColorSpace, pikepdf.Name.BitsPerComponent}
            for key in list(obj.keys()):
                if key not in essential_keys:
                    del obj[key]

            obj.Width = 1
            obj.Height = 1
            obj.ColorSpace = pikepdf.Name.DeviceRGB
            obj.BitsPerComponent = 8
            # Explicitly remove filter/decode params if they existed
            if pikepdf.Name.Filter in obj: del obj[pikepdf.Name.Filter]
            if pikepdf.Name.DecodeParms in obj: del obj[pikepdf.Name.DecodeParms]

            obj.write(self._blank_image_data)
            logging.info(f"Replaced image {obj.objgen} with a blank pixel.")
        except Exception as e:
            logging.warning(f"Could not replace image {obj.objgen}: {e}")

    def _lossless_optimize_jpeg_stream(self, obj, temp_dir):
        if not (self.jpegoptim_path or self.ect_path):
            return 0

        try:
            original_data = obj.read_raw_bytes()
            original_size = len(original_data)
            if original_size == 0:
                return 0

            tmp = temp_dir / f"img_{obj.objgen}.jpg"
            tmp.write_bytes(original_data)

            optimized = False
            if self.jpegoptim_path:
                cmd = [self.jpegoptim_path, "--strip-all", "-q", str(tmp)]
                self._run_command(cmd) # Ignore errors, proceed if possible
                optimized = True

            if self.ect_path and tmp.exists() and not self.fast_mode:
                cmd = [self.ect_path, "-quiet", "-strip", "-progressive", "-3", str(tmp)]
                self._run_command(cmd) # Ignore errors
                optimized = True

            if optimized and tmp.exists():
                new_bytes = tmp.read_bytes()
                new_size = len(new_bytes)
                if 0 < new_size < original_size:
                    obj.write(new_bytes)
                    # Ensure filter is correct, remove params
                    obj.Filter = pikepdf.Name.DCTDecode
                    if '/DecodeParms' in obj:
                        del obj['/DecodeParms']
                    logging.info(f"Losslessly optimized JPEG {obj.objgen}, saved {original_size - new_size} bytes.")
                    return original_size - new_size
        except Exception as e:
             logging.warning(f"Failed to optimize JPEG stream {obj.objgen}: {e}")
        return 0


    def _optimize_image_stream(self, pdf, obj, temp_dir, mode='lossless', dpi=150):
        # (This function remains largely the same as before, handling PNG/Flate optimization)
        # It correctly modifies the 'obj' in place, which is saved later.
        try:
            if not isinstance(obj, pikepdf.Stream) or obj.get("/Subtype") != "/Image":
                return 0

            if mode == 'lossless':
                try:
                    color_space = obj.get('/ColorSpace')
                    if color_space == '/DeviceCMYK' or (isinstance(color_space, pikepdf.Array) and color_space[0] == '/ICCBased'):
                        logging.info(f"Skipping CMYK image {obj.objgen} in lossless mode to preserve it.")
                        return 0
                except Exception:
                    pass

            original_size = len(obj.read_raw_bytes())
            if original_size == 0:
                return 0

            try:
                pdf_image = pikepdf.PdfImage(obj)
                pil_image = pdf_image.as_pil_image()
            except Exception as e:
                # Common issue: Masked images, unsupported color spaces
                logging.warning(f"Could not extract image {obj.objgen} for optimization (possibly masked or unsupported format): {e}")
                return 0 # Skip optimization for this image

            temp_img_path = temp_dir / f"img_{obj.objgen}.png"
            pil_image.save(temp_img_path, "png")

            optimized_path = None
            if mode == 'lossy' and self.pngquant_path:
                quality_str = "80-95"
                if dpi <= 100:
                    quality_str = "40-60"
                elif dpi <= 200:
                    quality_str = "65-80"
                quant_path = temp_dir / f"img_{obj.objgen}.quant.png"
                cmd = [self.pngquant_path, "--force", "--skip-if-larger", f"--quality={quality_str}", "--output", str(quant_path), "256", str(temp_img_path)]
                # Run command, check return code and if output exists
                result = self._run_command(cmd, check=False) # Don't raise error on failure
                if result and result.returncode == 0 and quant_path.exists() and quant_path.stat().st_size > 0:
                    optimized_path = quant_path
                elif result and result.returncode != 0 and result.returncode != 99: # 99 = skipped
                     logging.warning(f"Pngquant failed for image {obj.objgen}: {result.stderr.strip()}")


            final_optimized_path = optimized_path if optimized_path else temp_img_path

            try:
                oxipng_out_path = temp_dir / f"img_{obj.objgen}.oxipng.png"
                options = {"level": 2 if self.fast_mode else 6, "strip": oxipng.StripChunks.all()}
                if mode == 'lossy':
                    options["optimize_alpha"] = True
                    options["scale_16"] = True
                oxipng.optimize(final_optimized_path, oxipng_out_path, **options)
                if oxipng_out_path.exists() and oxipng_out_path.stat().st_size > 0:
                    final_optimized_path = oxipng_out_path
            except Exception as e:
                logging.warning(f"Could not process PNG with pyoxipng for {obj.objgen}: {e}")

            if self.ect_path and not self.fast_mode and final_optimized_path.exists():
                ect_target_path = temp_dir / f"img_{obj.objgen}.ect.png"
                shutil.copy(final_optimized_path, ect_target_path)
                cmd_ect = [self.ect_path, "-S2", "-strip", "-quiet", str(ect_target_path)]
                # Run ECT, check size comparison
                result_ect = self._run_command(cmd_ect, check=False)
                if result_ect and result_ect.returncode == 0 and ect_target_path.exists() and ect_target_path.stat().st_size < final_optimized_path.stat().st_size:
                    final_optimized_path = ect_target_path
                elif result_ect and result_ect.returncode != 0:
                    logging.warning(f"ECT failed for image {obj.objgen}: {result_ect.stderr.strip()}")


            final_pil_image = Image.open(final_optimized_path)
            # Ensure we handle palette images correctly by converting before compression
            if final_pil_image.mode == 'P':
                final_pil_image = final_pil_image.convert('RGBA' if 'A' in final_pil_image.info.get('transparency', '') else 'RGB')

            has_transparency = 'A' in final_pil_image.mode
            total_new_size = float('inf')

            if has_transparency:
                if final_pil_image.mode != 'RGBA':
                    final_pil_image = final_pil_image.convert('RGBA')
                # Create RGB and Alpha separately
                rgb_image = Image.new("RGB", final_pil_image.size, (255, 255, 255)) # White background
                rgb_image.paste(final_pil_image, mask=final_pil_image.split()[3]) # Paste using alpha mask
                alpha_image = final_pil_image.split()[3]

                compressed_rgb = zlib.compress(rgb_image.tobytes())
                compressed_alpha = zlib.compress(alpha_image.tobytes())
                total_new_size = len(compressed_rgb) + len(compressed_alpha)

                if total_new_size < original_size:
                    # Clear existing stream dictionary before writing new data
                    for key in list(obj.keys()): del obj[key]

                    obj.write(compressed_rgb)
                    obj.Type = pikepdf.Name.XObject
                    obj.Subtype = pikepdf.Name.Image
                    obj.Filter = pikepdf.Name.FlateDecode
                    obj.Width = final_pil_image.width
                    obj.Height = final_pil_image.height
                    obj.ColorSpace = pikepdf.Name.DeviceRGB
                    obj.BitsPerComponent = 8

                    smask_stream = pdf.new_stream(compressed_alpha)
                    smask_stream.Type = pikepdf.Name.XObject
                    smask_stream.Subtype = pikepdf.Name.Image
                    smask_stream.Filter = pikepdf.Name.FlateDecode
                    smask_stream.Width = final_pil_image.width
                    smask_stream.Height = final_pil_image.height
                    smask_stream.ColorSpace = pikepdf.Name.DeviceGray
                    smask_stream.BitsPerComponent = 8
                    obj.SMask = smask_stream # Assign the new stream object
            else:
                if final_pil_image.mode != 'RGB':
                    final_pil_image = final_pil_image.convert('RGB')
                compressed_rgb = zlib.compress(final_pil_image.tobytes())
                total_new_size = len(compressed_rgb)

                if total_new_size < original_size:
                     # Clear existing stream dictionary
                    for key in list(obj.keys()): del obj[key]

                    obj.write(compressed_rgb)
                    obj.Type = pikepdf.Name.XObject
                    obj.Subtype = pikepdf.Name.Image
                    obj.Filter = pikepdf.Name.FlateDecode
                    obj.Width = final_pil_image.width
                    obj.Height = final_pil_image.height
                    obj.ColorSpace = pikepdf.Name.DeviceRGB
                    obj.BitsPerComponent = 8
                    # Ensure SMask is removed if transparency was lost
                    if pikepdf.Name.SMask in obj: del obj[pikepdf.Name.SMask]


            if total_new_size < original_size:
                saved = original_size - total_new_size
                logging.info(f"Optimized Flate image {obj.objgen}, saved {saved} bytes.")
                return saved

        except Exception as e:
            logging.warning(f"Could not process image stream {obj.objgen}: {e}", exc_info=True)
        return 0


    def _post_process_pdf(self, pdf_path_in, pdf_path_out, strip_metadata=False):
        # Writes from pdf_path_in to pdf_path_out
        if not self.cpdf_path:
            logging.warning("cpdf not found, skipping post-processing.")
            if pdf_path_in != pdf_path_out:
                 shutil.copy2(pdf_path_in, pdf_path_out)
            return

        final_in_path = Path(pdf_path_in)
        final_out_path = Path(pdf_path_out)

        cmd = [self.cpdf_path]
        if self.linearize:
            cmd.append("-l") # Linearize should be first operation on input

        cmd.append(str(final_in_path))

        operations = []
        if self.darken_text:
            operations.append(["-blacktext"])
        if self.remove_open_action:
            operations.append(["-remove-dict-entry", "/OpenAction"])
        if strip_metadata:
            operations.append(["-remove-metadata"])

        # Squeeze and object streams last for best compression
        operations.append(["-squeeze", "-create-objstm"])

        is_first_op = not self.linearize # Linearize counts as first op if present
        for op in operations:
            if not is_first_op:
                 cmd.append("AND")
            cmd.extend(op)
            is_first_op = False

        cmd.extend(["-o", str(final_out_path)])

        try:
            result = self._run_command(cmd) # run_command now raises ProcessingError on failure
            if not final_out_path.exists() or final_out_path.stat().st_size == 0:
                 logging.error(f"cpdf processing failed to create output file or file is empty.")
                 # Copy original if output failed
                 if pdf_path_in != pdf_path_out: shutil.copy2(pdf_path_in, pdf_path_out)
                 raise ProcessingError("cpdf post-processing failed.")
        except ProcessingError as e:
             logging.error(f"cpdf processing step failed: {e}")
             # Ensure original is copied to output path if cpdf fails
             if pdf_path_in != pdf_path_out:
                 try:
                     shutil.copy2(pdf_path_in, pdf_path_out)
                     logging.info(f"Copied original file to {pdf_path_out} due to cpdf failure.")
                 except Exception as copy_err:
                     logging.error(f"Failed to copy original file after cpdf failure: {copy_err}")
             # Re-raise the error so the main task knows it failed
             raise e
        except Exception as e:
            logging.error(f"Unexpected error during cpdf post-processing: {e}")
            if pdf_path_in != pdf_path_out:
                 try: shutil.copy2(pdf_path_in, pdf_path_out)
                 except Exception as copy_err: logging.error(f"Failed to copy original file: {copy_err}")
            raise ProcessingError(f"Unexpected cpdf error: {e}")

    # --- MODIFIED optimize functions to write to temp_output_path ---

    def optimize_lossless(self, input_file, temp_output_path, strip_metadata=False):
        internal_temp_pdf = None
        try:
            with tempfile.TemporaryDirectory() as temp_dir_str:
                temp_dir = Path(temp_dir_str)
                # Use a named temporary file within the context for pikepdf processing
                with tempfile.NamedTemporaryFile(suffix=".pdf", dir=temp_dir, delete=False) as tf:
                    internal_temp_pdf = Path(tf.name)
                shutil.copy2(input_file, internal_temp_pdf) # Copy input to processable temp

                if self.q: self.q.put(('status', f"Opening PDF for lossless..."))
                with pikepdf.open(internal_temp_pdf, allow_overwriting_input=True) as pdf:
                    if self.q: self.q.put(('status', f"Optimizing images losslessly..."))
                    total_optimized = 0
                    # Iterate safely through objects
                    obj_list = list(pdf.objects)
                    for obj in obj_list:
                         if isinstance(obj, pikepdf.Stream) and obj.get("/Subtype") == "/Image":
                            filt = obj.get("/Filter")
                            if isinstance(filt, pikepdf.Array) and len(filt) > 0: filt = filt[0]

                            if filt == "/DCTDecode":
                                total_optimized += self._lossless_optimize_jpeg_stream(obj, temp_dir)
                            elif filt != "/JPXDecode": # Skip JPX
                                total_optimized += self._optimize_image_stream(pdf, obj, temp_dir, mode='lossless')

                    if self.q: self.q.put(('status', f"Recompressing streams..."))
                    # Save changes back to the internal temp file
                    pdf.save(internal_temp_pdf, object_stream_mode=pikepdf.ObjectStreamMode.generate, recompress_flate=True)

            # Post-process from internal temp to the final temp path provided
            if self.q: self.q.put(('status', f"Finalizing with cpdf..."))
            self._post_process_pdf(internal_temp_pdf, temp_output_path, strip_metadata)

        except Exception as e:
            logging.error(f"Lossless optimization failed: {e}", exc_info=True)
            # Ensure temp_output_path gets the original if anything fails
            try: shutil.copy2(input_file, temp_output_path)
            except Exception as copy_err: logging.error(f"Failed to copy original on error: {copy_err}")
            raise # Re-raise exception to be caught by run_compress_task
        finally:
            # Clean up the internal temporary file if it exists
            if internal_temp_pdf and internal_temp_pdf.exists():
                os.remove(internal_temp_pdf)

    def optimize_true_lossless(self, input_file, temp_output_path, strip_metadata=False):
        # Similar structure to optimize_lossless, writing to temp_output_path via _post_process_pdf
        internal_temp_pdf = None
        try:
            with tempfile.TemporaryDirectory() as temp_dir_str:
                temp_dir = Path(temp_dir_str)
                with tempfile.NamedTemporaryFile(suffix=".pdf", dir=temp_dir, delete=False) as tf:
                    internal_temp_pdf = Path(tf.name)
                shutil.copy2(input_file, internal_temp_pdf)

                if self.q: self.q.put(('status', f"Opening PDF for true lossless..."))
                with pikepdf.open(internal_temp_pdf, allow_overwriting_input=True) as pdf:
                    if self.q: self.q.put(('status', f"Optimizing non-JPEG images losslessly..."))
                    obj_list = list(pdf.objects)
                    for obj in obj_list:
                         if isinstance(obj, pikepdf.Stream) and obj.get("/Subtype") == "/Image":
                            # Only optimize non-DCT images
                            filt = obj.get("/Filter")
                            if isinstance(filt, pikepdf.Array) and len(filt) > 0: filt = filt[0]
                            if filt != "/DCTDecode" and filt != "/JPXDecode":
                                self._optimize_image_stream(pdf, obj, temp_dir, mode='lossless')

                    if self.q: self.q.put(('status', f"Recompressing streams..."))
                    pdf.save(internal_temp_pdf, object_stream_mode=pikepdf.ObjectStreamMode.generate, recompress_flate=True)

            if self.q: self.q.put(('status', f"Finalizing with cpdf..."))
            self._post_process_pdf(internal_temp_pdf, temp_output_path, strip_metadata)

        except Exception as e:
            logging.error(f"True Lossless optimization failed: {e}", exc_info=True)
            try: shutil.copy2(input_file, temp_output_path)
            except Exception as copy_err: logging.error(f"Failed to copy original on error: {copy_err}")
            raise
        finally:
            if internal_temp_pdf and internal_temp_pdf.exists():
                os.remove(internal_temp_pdf)

    def optimize_text_only(self, input_file, temp_output_path, strip_metadata=False):
        # Writes to temp_output_path via _post_process_pdf
        internal_temp_pdf = None
        try:
            with tempfile.TemporaryDirectory() as temp_dir_str:
                temp_dir = Path(temp_dir_str) # Needed if _replace_image needs it later? No.
                with tempfile.NamedTemporaryFile(suffix=".pdf", dir=temp_dir_str, delete=False) as tf:
                     internal_temp_pdf = Path(tf.name)
                shutil.copy2(input_file, internal_temp_pdf)

                if self.q: self.q.put(('status', f"Opening PDF to remove images..."))
                with pikepdf.open(internal_temp_pdf, allow_overwriting_input=True) as pdf:
                    if self.q: self.q.put(('status', f"Finding images to replace..."))
                    image_objects = [obj for obj in pdf.objects if isinstance(obj, pikepdf.Stream) and obj.get("/Subtype") == "/Image"]

                    if self.q: self.q.put(('status', f"Replacing {len(image_objects)} images..."))
                    for obj in image_objects:
                        self._replace_image_stream(obj)

                    if self.q: self.q.put(('status', f"Recompressing streams..."))
                    pdf.save(internal_temp_pdf, object_stream_mode=pikepdf.ObjectStreamMode.generate, recompress_flate=True)

            if self.q: self.q.put(('status', f"Finalizing with cpdf..."))
            self._post_process_pdf(internal_temp_pdf, temp_output_path, strip_metadata)

        except Exception as e:
            logging.error(f"Remove Images optimization failed: {e}", exc_info=True)
            try: shutil.copy2(input_file, temp_output_path)
            except Exception as copy_err: logging.error(f"Failed to copy original on error: {copy_err}")
            raise
        finally:
            if internal_temp_pdf and internal_temp_pdf.exists():
                os.remove(internal_temp_pdf)

    def _optimize_lossy_gs_preset_fallback(self, input_file, temp_output_path, dpi, strip_metadata, remove_interactive, use_bicubic):
        # Writes to temp_output_path
        if self.q: self.q.put(('status', "Fallback: Using GS preset mode..."))
        gs_output_temp_pdf = None # Use a separate temp file for GS output

        try:
            with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as temp_out_gs:
                gs_output_temp_pdf = Path(temp_out_gs.name)

            # Determine preset based on DPI
            if dpi <= 72: preset = '/screen'
            elif dpi <= 150: preset = '/ebook'
            elif dpi <= 300: preset = '/printer'
            else: preset = '/prepress'

            cmd = [
                self.gs_path, '-sDEVICE=pdfwrite', '-dCompatibilityLevel=1.4',
                '-dNOPAUSE', '-dBATCH', '-dQUIET', '-dSAFER', f'-dPDFSETTINGS={preset}',
                '-dColorImageDownsampleThreshold=1.0', '-dGrayImageDownsampleThreshold=1.0',
                f'-dColorImageResolution={dpi}', f'-dGrayImageResolution={dpi}',
                f'-dDownsampleType=/{"Bicubic" if use_bicubic else "Average"}',
                '-dSubsetFonts=true', '-dCompressFonts=true'
            ]
            if self.linearize: cmd.append("-dFastWebView=true")
            if remove_interactive: cmd.extend(["-dShowAnnots=false", "-dShowAcroForm=false"])
            if self.convert_to_grayscale: cmd.append('-sColorConversionStrategy=Gray')
            else: cmd.append('-sColorConversionStrategy=LeaveColorUnchanged')
            cmd.extend([f'-sOutputFile={gs_output_temp_pdf}', str(input_file)])

            # Run Ghostscript command
            self._run_command(cmd) # Will raise ProcessingError on failure

            if not gs_output_temp_pdf.exists() or gs_output_temp_pdf.stat().st_size == 0:
                raise ProcessingError("Ghostscript preset fallback failed: Output file empty or not created.")

            # Post-process with Pikepdf (minimal) + cpdf
            if self.q: self.q.put(('status', "Fallback: Finalizing..."))
            try:
                # Minimal pikepdf pass primarily for potential structure cleanup
                 with pikepdf.open(gs_output_temp_pdf, allow_overwriting_input=True) as pdf:
                     # Only remove metadata if requested, avoid extra recompression here
                     if strip_metadata and pdf.docinfo:
                         for key in list(pdf.docinfo.keys()): del pdf.docinfo[key]
                     if strip_metadata and pdf.Root.Metadata: del pdf.Root.Metadata
                     pdf.save(gs_output_temp_pdf) # Save potentially cleaned structure
            except Exception as e:
                logging.warning(f"Pikepdf finalization in fallback failed: {e}")
                # Continue to cpdf even if pikepdf fails

            # Final step: cpdf processing from gs_output_temp_pdf to temp_output_path
            self._post_process_pdf(gs_output_temp_pdf, temp_output_path, strip_metadata)

        except Exception as e:
            logging.error(f"GS preset fallback mode failed: {e}", exc_info=True)
            # Copy original to ensure *something* is in temp_output_path for the check
            try: shutil.copy2(input_file, temp_output_path)
            except Exception as copy_err: logging.error(f"Failed to copy original on fallback error: {copy_err}")
            raise # Re-raise the original error
        finally:
            # Clean up the intermediate GS output file
            if gs_output_temp_pdf and gs_output_temp_pdf.exists():
                os.remove(gs_output_temp_pdf)


    def optimize_lossy(self, input_file, temp_output_path, dpi, strip_metadata=False, remove_interactive=False, use_bicubic=False):
        # Writes to temp_output_path
        gs_output_temp_pdf = None # Intermediate file for GS output

        try:
            if self.q: self.q.put(('status', "Attempting high-compression mode..."))
            with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as temp_out_gs:
                gs_output_temp_pdf = Path(temp_out_gs.name)

            # Build Ghostscript Command
            cmd = [
                self.gs_path, '-dNOSAFER', f'-sOutputFile={gs_output_temp_pdf}',
                '-sDEVICE=pdfwrite', '-dCompatibilityLevel=1.7', '-dNOPAUSE', '-dQUIET', '-dBATCH',
                '-dDetectDuplicateImages=true', '-dDownsampleGrayImages=true',
                '-dDownsampleMonoImages=true', '-dDownsampleColorImages=true',
                f'-dColorImageResolution={dpi}', f'-dGrayImageResolution={dpi}', f'-dMonoImageResolution={dpi}',
                '-dMonoImageFilter=/CCITTFaxEncode',
            ]
            if self.downsample_threshold_enabled:
                cmd.extend(['-dColorImageDownsampleThreshold=1.0', '-dGrayImageDownsampleThreshold=1.0', '-dMonoImageDownsampleThreshold=1.0'])

            jpeg_quality = "70" if self.fast_mode else ("70" if dpi <= 100 else "85")

            if self.convert_to_grayscale:
                cmd.extend(['-sColorConversionStrategy=Gray', '-dProcessColorModel=/DeviceGray', '-dOverrideICC', f'-dJPEGQ={jpeg_quality}'])
            else:
                cmd.extend(['-sColorConversionStrategy=sRGB', '-dProcessColorModel=/DeviceRGB', f'-dJPEGQ={jpeg_quality}'])

            if self.safe_mode:
                cmd.extend(['-dAutoFilterColorImages=true', '-dAutoFilterGrayImages=true'])
            else:
                cmd.extend(['-dColorImageFilter=/DCTEncode', '-dGrayImageFilter=/DCTEncode'])

            cmd.append('-dApplyTrCF=true') 

            if use_bicubic:
                cmd.extend(['-dColorImageDownsampleType=/Bicubic', '-dGrayImageDownsampleType=/Bicubic'])
            if remove_interactive:
                cmd.extend(["-dShowAnnots=false", "-dShowAcroForm=false"])

            cmd.append(str(input_file))

            # --- Run Ghostscript ---
            self._run_command(cmd) 

            if not gs_output_temp_pdf.exists() or gs_output_temp_pdf.stat().st_size == 0:
                raise ProcessingError("Ghostscript high-compression failed: Output file empty or not created.")

            # --- Post-GS Lossless Optimization ---
            with tempfile.TemporaryDirectory() as final_opt_dir_str:
                final_opt_dir = Path(final_opt_dir_str)
                try:
                    if self.q: self.q.put(('status', "Optimizing images losslessly post-GS..."))
                    with pikepdf.open(gs_output_temp_pdf, allow_overwriting_input=True) as pdf:
                        optimized_bytes = 0
                        obj_list = list(pdf.objects) # Iterate over a copy
                        for obj in obj_list:
                            if isinstance(obj, pikepdf.Stream) and obj.get("/Subtype") == "/Image":
                                filt = obj.get("/Filter")
                                if isinstance(filt, pikepdf.Array) and len(filt) > 0: filt = filt[0]

                                if filt == pikepdf.Name.DCTDecode:
                                    optimized_bytes += self._lossless_optimize_jpeg_stream(obj, final_opt_dir)
                                elif filt == pikepdf.Name.FlateDecode:
                                     optimized_bytes += self._optimize_image_stream(pdf, obj, final_opt_dir, mode='lossless')
                        if optimized_bytes > 0:
                            logging.info(f"Post-GS optimization saved an additional {optimized_bytes} bytes.")
                            pdf.save(gs_output_temp_pdf, recompress_flate=True)
                except Exception as e:
                    logging.warning(f"Post-Ghostscript optimization step failed: {e}")
                    

            # --- Final cpdf Processing ---
            if self.q: self.q.put(('status', "Finalizing with cpdf..."))
            # Process from the (potentially pikepdf-optimized) GS output to the final temp path
            self._post_process_pdf(gs_output_temp_pdf, temp_output_path, strip_metadata)

        except Exception as e: # Catch GS failure or cpdf failure
            logging.warning(f"High-compression mode failed: {e}. Switching to preset fallback mode.")
            if self.q: self.q.put(('status', "High-compression failed, switching to preset mode..."))
            # Fallback writes directly to temp_output_path
            try:
                self._optimize_lossy_gs_preset_fallback(input_file, temp_output_path, dpi, strip_metadata, remove_interactive, use_bicubic)
            except Exception as fallback_e:
                logging.error(f"Fallback optimization also failed: {fallback_e}")
                # Ensure original is copied if fallback fails too
                try: shutil.copy2(input_file, temp_output_path)
                except Exception as copy_err: logging.error(f"Failed to copy original on fallback error: {copy_err}")
                raise fallback_e # Re-raise the fallback error
        finally:
            # Clean up the intermediate GS output file
            if gs_output_temp_pdf and gs_output_temp_pdf.exists():
                os.remove(gs_output_temp_pdf)

    def optimize_pdfa(self, input_file, temp_output_path):
        # Writes to temp_output_path
        if self.q: self.q.put(('status', "Converting to PDF/A..."))
        gs_output_temp_pdf = None # Intermediate file

        try:
            gs_path_obj = Path(self.gs_path)
            # Determine srgb.icc path
            if sys.platform == "win32":
                srgb_profile = gs_path_obj.parent.parent / "lib" / "srgb.icc"
            else: # Linux/macOS might be in share
                srgb_profile = gs_path_obj.parent.parent / "share" / "ghostscript" / "iccprofiles" / "srgb.icc"

            
            if not srgb_profile.exists(): srgb_profile = resource_path('lib/srgb.icc')
            if not srgb_profile.exists(): raise FileNotFoundError("Required 'srgb.icc' profile not found.")

            with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as temp_out_gs:
                 gs_output_temp_pdf = Path(temp_out_gs.name)

            cmd = [
                self.gs_path, '-dPDFA=1', '-dBATCH', '-dNOPAUSE', '-dQUIET', '-dNOSAFER',
                '-sDEVICE=pdfwrite', '-dPDFACompatibilityPolicy=1', '-sColorConversionStrategy=RGB',
                f'-sOutputICCProfile={str(srgb_profile)}',
            ]
            if self.pdfa_compression:
                 cmd.extend([
                     f'-dColorImageResolution={self.pdfa_dpi}', f'-dGrayImageResolution={self.pdfa_dpi}', f'-dMonoImageResolution={self.pdfa_dpi}',
                     '-dPDFSETTINGS=/prepress', 
                 ])
                 if self.downsample_threshold_enabled: cmd.extend(['-dColorImageDownsampleThreshold=1.0', '-dGrayImageDownsampleThreshold=1.0', '-dMonoImageDownsampleThreshold=1.0'])
                 jpeg_quality = "70" if self.fast_mode else ("70" if self.pdfa_dpi <= 100 else "85")
                 
                 cmd.extend(['-dColorImageFilter=/DCTEncode', '-dGrayImageFilter=/DCTEncode', f'-dJPEGQ={jpeg_quality}'])
                 cmd.extend(['-dColorImageDownsampleType=/Bicubic', '-dGrayImageDownsampleType=/Bicubic'])

            cmd.extend([f'-sOutputFile={gs_output_temp_pdf}', str(input_file)])

            # --- Run Ghostscript ---
            self._run_command(cmd) 

            if not gs_output_temp_pdf.exists() or gs_output_temp_pdf.stat().st_size == 0:
                 raise ProcessingError("Ghostscript PDF/A conversion failed: Output file empty or not created.")

            
            try:
                 with pikepdf.open(gs_output_temp_pdf, allow_overwriting_input=True) as pdf:
                    
                     changed_meta = False
                     with pdf.open_metadata(set_pikepdf_as_editor=False) as meta:
                         if 'pdf:Producer' not in meta: meta['pdf:Producer'] = 'Unknown' # Required for PDF/A
                         meta['pdf:Producer'] += '; MinimalPDF Optimizer'
                         meta['xmp:CreatorTool'] = 'MinimalPDF Optimizer'
                         changed_meta = True
                     if changed_meta:
                         
                         pdf.save(gs_output_temp_pdf, object_stream_mode=pikepdf.ObjectStreamMode.generate, recompress_flate=True)
            except Exception as e:
                logging.warning(f"Pikepdf metadata step for PDF/A failed: {e}")
                

           
            self._post_process_pdf(gs_output_temp_pdf, temp_output_path, strip_metadata=False) 

        except Exception as e:
            logging.error(f"PDF/A optimization failed: {e}", exc_info=True)
            try: shutil.copy2(input_file, temp_output_path)
            except Exception as copy_err: logging.error(f"Failed to copy original on PDF/A error: {copy_err}")
            raise
        finally:
             if gs_output_temp_pdf and gs_output_temp_pdf.exists():
                 os.remove(gs_output_temp_pdf)