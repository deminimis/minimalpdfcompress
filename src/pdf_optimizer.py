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
                logging.warning(f"Command produced stderr: {result.stderr.strip()}")
            return result
        except subprocess.CalledProcessError as e:
            full_error = f"Command failed.\nSTDOUT: {e.stdout.strip()}\nSTDERR: {e.stderr.strip()}"
            logging.error(full_error)
            return None
        except FileNotFoundError:
            logging.error(f"Command not found: {cmd if use_shell else cmd[0]}")
            return None

    def _replace_image_stream(self, obj):
        try:
            if not isinstance(obj, pikepdf.Stream) or obj.get("/Subtype") != "/Image":
                return

            for key in ('/Filter', '/DecodeParms', '/SMask', '/Intent', '/ImageMask', '/Mask'):
                if key in obj:
                    del obj[key]

            obj.Width = 1
            obj.Height = 1
            obj.ColorSpace = pikepdf.Name.DeviceRGB
            obj.BitsPerComponent = 8

            obj.write(self._blank_image_data)
            logging.info(f"Replaced image {obj.objgen} with a blank pixel.")
        except Exception as e:
            logging.warning(f"Could not replace image {obj.objgen}: {e}")

    def _lossless_optimize_jpeg_stream(self, obj, temp_dir):
        if not (self.jpegoptim_path or self.ect_path):
            return 0

        original_data = obj.read_raw_bytes()
        original_size = len(original_data)
        if original_size == 0:
            return 0

        tmp = temp_dir / f"img_{obj.objgen}.jpg"
        tmp.write_bytes(original_data)

        if self.jpegoptim_path:
            cmd = [self.jpegoptim_path, "--strip-all", "-q", str(tmp)]
            self._run_command(cmd)

        if self.ect_path and tmp.exists() and not self.fast_mode:
            cmd = [self.ect_path, "-quiet", "-strip", "-progressive", "-3", str(tmp)]
            self._run_command(cmd)

        if tmp.exists():
            new_bytes = tmp.read_bytes()
            new_size = len(new_bytes)
            if 0 < new_size < original_size:
                obj.write(new_bytes)
                obj.Filter = pikepdf.Name.DCTDecode
                if '/DecodeParms' in obj:
                    del obj['/DecodeParms']
                logging.info(f"Losslessly optimized JPEG {obj.objgen}, saved {original_size - new_size} bytes.")
                return original_size - new_size
        return 0

    def _optimize_image_stream(self, pdf, obj, temp_dir, mode='lossless', dpi=150):
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
                logging.warning(f"Could not extract image {obj.objgen}: {e}")
                return 0

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
                if self._run_command(cmd) and quant_path.exists() and quant_path.stat().st_size > 0:
                    optimized_path = quant_path

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
                logging.warning(f"Could not process PNG with pyoxipng: {e}")

            if self.ect_path and not self.fast_mode and final_optimized_path.exists():
                ect_target_path = temp_dir / f"img_{obj.objgen}.ect.png"
                shutil.copy(final_optimized_path, ect_target_path)
                cmd_ect = [self.ect_path, "-S2", "-strip", "-quiet", str(ect_target_path)]
                if self._run_command(cmd_ect) and ect_target_path.exists() and ect_target_path.stat().st_size < final_optimized_path.stat().st_size:
                    final_optimized_path = ect_target_path

            final_pil_image = Image.open(final_optimized_path)
            has_transparency = 'A' in final_pil_image.mode
            total_new_size = float('inf')

            if has_transparency:
                if final_pil_image.mode != 'RGBA':
                    final_pil_image = final_pil_image.convert('RGBA')
                rgb_image = Image.new("RGB", final_pil_image.size)
                rgb_image.paste(final_pil_image)
                alpha_image = final_pil_image.split()[3]
                compressed_rgb = zlib.compress(rgb_image.tobytes())
                compressed_alpha = zlib.compress(alpha_image.tobytes())
                total_new_size = len(compressed_rgb) + len(compressed_alpha)

                if total_new_size < original_size:
                    for key in list(obj.keys()):
                        del obj[key]
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
                    obj.SMask = smask_stream
            else:
                if final_pil_image.mode != 'RGB':
                    final_pil_image = final_pil_image.convert('RGB')
                compressed_rgb = zlib.compress(final_pil_image.tobytes())
                total_new_size = len(compressed_rgb)

                if total_new_size < original_size:
                    for key in list(obj.keys()):
                        del obj[key]
                    obj.write(compressed_rgb)
                    obj.Type = pikepdf.Name.XObject
                    obj.Subtype = pikepdf.Name.Image
                    obj.Filter = pikepdf.Name.FlateDecode
                    obj.Width = final_pil_image.width
                    obj.Height = final_pil_image.height
                    obj.ColorSpace = pikepdf.Name.DeviceRGB
                    obj.BitsPerComponent = 8

            if total_new_size < original_size:
                saved = original_size - total_new_size
                logging.info(f"Optimized Flate image {obj.objgen}, saved {saved} bytes.")
                return saved

        except Exception as e:
            logging.warning(f"Could not process image {obj.objgen}: {e}", exc_info=True)
        return 0

    def _post_process_pdf(self, pdf_path, strip_metadata=False):
        if not self.cpdf_path:
            return

        final_path = Path(pdf_path)
        temp_output_path = final_path.with_name(f"cpdf_processed_{final_path.name}")

        cmd = [self.cpdf_path]
        if self.linearize:
            cmd.append("-l")

        cmd.append(str(final_path))

        operations = []
        if self.darken_text:
            operations.append(["-blacktext"])
        if self.remove_open_action:
            operations.append(["-remove-dict-entry", "/OpenAction"])
        if strip_metadata:
            operations.append(["-remove-metadata"])

        operations.append(["-squeeze", "-create-objstm"])

        for i, op in enumerate(operations):
            cmd.extend(op)
            if i < len(operations) - 1:
                cmd.append("AND")

        cmd.extend(["-o", str(temp_output_path)])

        if self._run_command(cmd) and temp_output_path.exists() and temp_output_path.stat().st_size > 0:
            shutil.move(str(temp_output_path), str(final_path))
        else:
            logging.warning("cpdf processing step failed.")

    def optimize_lossless(self, input_file, output_file, strip_metadata=False):
        with tempfile.TemporaryDirectory() as temp_dir_str:
            temp_dir = Path(temp_dir_str)
            temp_pdf_path = temp_dir / "processed.pdf"
            shutil.copy(input_file, temp_pdf_path)

            with pikepdf.open(temp_pdf_path, allow_overwriting_input=True) as pdf:
                if self.q:
                    self.q.put(('status', f"Optimizing images (true lossless)..."))
                for obj in pdf.objects:
                    if isinstance(obj, pikepdf.Stream) and obj.get("/Subtype") == "/Image":
                        filt = obj.get("/Filter")
                        if isinstance(filt, pikepdf.Array) and len(filt) > 0:
                            filt = filt[0]

                        if filt == "/DCTDecode":
                            self._lossless_optimize_jpeg_stream(obj, temp_dir)
                        elif filt != "/JPXDecode":
                            self._optimize_image_stream(pdf, obj, temp_dir, mode='lossless')

                if self.q:
                    self.q.put(('status', f"Recompressing streams..."))
                pdf.save(object_stream_mode=pikepdf.ObjectStreamMode.generate, recompress_flate=True)

            if self.q:
                self.q.put(('status', f"Finalizing with cpdf..."))
            self._post_process_pdf(temp_pdf_path, strip_metadata)
            shutil.move(str(temp_pdf_path), output_file)

    def optimize_true_lossless(self, input_file, output_file, strip_metadata=False):
        with tempfile.TemporaryDirectory() as temp_dir_str:
            temp_dir = Path(temp_dir_str)
            temp_pdf_path = temp_dir / "processed.pdf"
            shutil.copy(input_file, temp_pdf_path)

            with pikepdf.open(temp_pdf_path, allow_overwriting_input=True) as pdf:
                if self.q:
                    self.q.put(('status', f"Optimizing images losslessly..."))
                for obj in pdf.objects:
                    self._optimize_image_stream(pdf, obj, temp_dir, mode='lossless')

                if self.q:
                    self.q.put(('status', f"Recompressing streams..."))
                pdf.save(object_stream_mode=pikepdf.ObjectStreamMode.generate, recompress_flate=True)

            if self.q:
                self.q.put(('status', f"Finalizing with cpdf..."))
            self._post_process_pdf(temp_pdf_path, strip_metadata)
            shutil.move(str(temp_pdf_path), output_file)

    def optimize_text_only(self, input_file, output_file, strip_metadata=False):
        with tempfile.TemporaryDirectory() as temp_dir_str:
            temp_dir = Path(temp_dir_str)
            temp_pdf_path = temp_dir / "processed.pdf"
            shutil.copy(input_file, temp_pdf_path)

            with pikepdf.open(temp_pdf_path, allow_overwriting_input=True) as pdf:
                if self.q:
                    self.q.put(('status', f"Finding images to replace..."))
                image_objects = [obj for obj in pdf.objects if isinstance(obj, pikepdf.Stream) and obj.get("/Subtype") == "/Image"]

                if self.q:
                    self.q.put(('status', f"Replacing {len(image_objects)} images with blanks..."))
                for obj in image_objects:
                    self._replace_image_stream(obj)

                if self.q:
                    self.q.put(('status', f"Recompressing streams..."))
                pdf.save(object_stream_mode=pikepdf.ObjectStreamMode.generate, recompress_flate=True)

            if self.q:
                self.q.put(('status', f"Finalizing with cpdf..."))
            self._post_process_pdf(temp_pdf_path, strip_metadata)
            shutil.move(str(temp_pdf_path), output_file)

    def _optimize_lossy_gs_preset_fallback(self, input_file, output_file, dpi, strip_metadata, remove_interactive, use_bicubic):
        if self.q:
            self.q.put(('status', "Fallback: Using GS preset mode..."))

        with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as temp_out:
            gs_output_path = Path(temp_out.name)

        try:
            if dpi <= 72:
                preset = '/screen'
            elif dpi <= 150:
                preset = '/ebook'
            elif dpi <= 300:
                preset = '/printer'
            else:
                preset = '/prepress'

            cmd = [
                self.gs_path,
                '-sDEVICE=pdfwrite',
                '-dCompatibilityLevel=1.4',
                '-dNOPAUSE', '-dBATCH', '-dQUIET', '-dSAFER',
                f'-dPDFSETTINGS={preset}',
                '-dColorImageDownsampleThreshold=1.0',
                '-dGrayImageDownsampleThreshold=1.0',
                f'-dColorImageResolution={dpi}',
                f'-dGrayImageResolution={dpi}',
                f'-dDownsampleType=/{"Bicubic" if use_bicubic else "Average"}',
                '-dSubsetFonts=true',
                '-dCompressFonts=true'
            ]

            if self.linearize:
                cmd.append("-dFastWebView=true")

            if remove_interactive:
                cmd.extend(["-dShowAnnots=false", "-dShowAcroForm=false"])

            if self.convert_to_grayscale:
                cmd.append('-sColorConversionStrategy=Gray')
            else:
                cmd.append('-sColorConversionStrategy=LeaveColorUnchanged')

            cmd.extend([f'-sOutputFile={gs_output_path}', str(input_file)])

            if not self._run_command(cmd) or not gs_output_path.exists() or gs_output_path.stat().st_size == 0:
                logging.error(f"Ghostscript preset-based fallback failed. Command was: {' '.join(cmd)}")
                raise Exception("Ghostscript preset-based fallback failed.")

            try:
                if self.q:
                    self.q.put(('status', "Fallback: Finalizing with Pikepdf..."))
                with pikepdf.open(gs_output_path, allow_overwriting_input=True) as pdf:
                    if strip_metadata:
                        with pdf.open_metadata(set_pikepdf_as_editor=False) as meta:
                            for key in list(meta.keys()):
                                del meta[key]
                    pdf.save(
                        gs_output_path,
                        object_stream_mode=pikepdf.ObjectStreamMode.generate,
                        recompress_flate=True
                    )
            except Exception as e:
                logging.warning(f"Pikepdf finalization in fallback failed: {e}")

            shutil.move(str(gs_output_path), output_file)

        except Exception as e:
            logging.error(f"GS preset fallback mode failed entirely: {e}", exc_info=True)
            if input_file != output_file:
                shutil.copy2(input_file, output_file)
            raise
        finally:
            if gs_output_path.exists():
                os.remove(gs_output_path)

    def optimize_lossy(self, input_file, output_file, dpi, strip_metadata=False, remove_interactive=False, use_bicubic=False):
        try:
            if self.q:
                self.q.put(('status', "Attempting high-compression mode..."))
            with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as temp_out:
                gs_output_path = Path(temp_out.name)

            try:
                cmd = [
                    self.gs_path, '-dNOSAFER', f'-sOutputFile={gs_output_path}',
                    '-sDEVICE=pdfwrite', '-dCompatibilityLevel=1.7', '-dNOPAUSE', '-dQUIET', '-dBATCH',
                    '-dDetectDuplicateImages=true', '-dDownsampleGrayImages=true',
                    '-dDownsampleMonoImages=true', '-dDownsampleColorImages=true',
                    f'-dColorImageResolution={dpi}', f'-dGrayImageResolution={dpi}', f'-dMonoImageResolution={dpi}',
                    '-dMonoImageFilter=/CCITTFaxEncode',
                ]
                if self.downsample_threshold_enabled:
                    cmd.extend(['-dColorImageDownsampleThreshold=1.0', '-dGrayImageDownsampleThreshold=1.0', '-dMonoImageDownsampleThreshold=1.0'])
                jpeg_quality = "70" if self.fast_mode else ("70" if dpi <= 100 else "85")
                gray_filter_cmd = ['-dGrayImageFilter=/FlateEncode', '-dApplyTrCF=true']
                if self.convert_to_grayscale:
                    cmd.extend(['-sColorConversionStrategy=Gray', '-dProcessColorModel=/DeviceGray', '-dOverrideICC', f'-dJPEGQ={jpeg_quality}'])
                    cmd.extend(gray_filter_cmd)
                else:
                    cmd.extend(['-sColorConversionStrategy=sRGB', '-dProcessColorModel=/DeviceRGB', '-dColorImageFilter=/DCTEncode', f'-dJPEGQ={jpeg_quality}'])
                    cmd.extend(gray_filter_cmd)
                if use_bicubic:
                    cmd.extend(['-dColorImageDownsampleType=/Bicubic', '-dGrayImageDownsampleType=/Bicubic'])
                if remove_interactive:
                    cmd.extend(["-dShowAnnots=false", "-dShowAcroForm=false"])
                cmd.append(str(input_file))

                if not self._run_command(cmd) or not gs_output_path.exists() or gs_output_path.stat().st_size == 0:
                    raise Exception("Ghostscript high-compression failed.")

                with tempfile.TemporaryDirectory() as final_opt_dir_str:
                    final_opt_dir = Path(final_opt_dir_str)
                    try:
                        with pikepdf.open(gs_output_path, allow_overwriting_input=True) as pdf:
                            for obj in pdf.objects:
                                if isinstance(obj, pikepdf.Stream) and obj.get("/Subtype") == "/Image":
                                    filt = obj.get("/Filter")
                                    if isinstance(filt, pikepdf.Array) and len(filt) > 0:
                                        filt = filt[0]

                                    if filt == pikepdf.Name.DCTDecode:
                                        self._lossless_optimize_jpeg_stream(obj, final_opt_dir)
                                    elif filt == pikepdf.Name.FlateDecode:
                                        self._optimize_image_stream(pdf, obj, final_opt_dir, mode='lossless')
                            pdf.save(recompress_flate=True)
                    except Exception as e:
                        logging.warning(f"Post-Ghostscript optimization failed: {e}")

                self._post_process_pdf(gs_output_path, strip_metadata)
                shutil.move(str(gs_output_path), output_file)

            finally:
                if gs_output_path.exists():
                    os.remove(gs_output_path)

        except Exception as e:
            logging.warning(f"High-compression mode failed: {e}. Switching to preset fallback mode.")
            if self.q:
                self.q.put(('status', "High-compression failed, switching to preset mode..."))
            self._optimize_lossy_gs_preset_fallback(input_file, output_file, dpi, strip_metadata, remove_interactive, use_bicubic)

    def optimize_pdfa(self, input_file, output_file):
        if self.q:
            self.q.put(('status', "Converting to PDF/A..."))

        gs_path_obj = Path(self.gs_path)
        if sys.platform == "win32":
            srgb_profile = gs_path_obj.parent.parent / "lib" / "srgb.icc"
        else:
            srgb_profile = gs_path_obj.parent.parent / "share" / "ghostscript" / "iccprofiles" / "srgb.icc"

        if not srgb_profile.exists():
             srgb_profile = resource_path('lib/srgb.icc')

        if not srgb_profile.exists():
            raise FileNotFoundError("The required 'srgb.icc' color profile was not found.")

        with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as temp_out:
            temp_pdf_path = Path(temp_out.name)

        try:
            cmd = [
                self.gs_path,
                '-dPDFA=1',
                '-dBATCH',
                '-dNOPAUSE',
                '-dQUIET',
                '-dNOSAFER',
                '-sDEVICE=pdfwrite',
                '-dPDFACompatibilityPolicy=1',
                '-sColorConversionStrategy=RGB',
                f'-sOutputICCProfile={str(srgb_profile)}',
            ]
            
            if self.pdfa_compression:
                cmd.extend([
                    f'-dColorImageResolution={self.pdfa_dpi}',
                    f'-dGrayImageResolution={self.pdfa_dpi}',
                    f'-dMonoImageResolution={self.pdfa_dpi}',
                    '-dPDFSETTINGS=/prepress',
                ])
                if self.downsample_threshold_enabled:
                    cmd.extend(['-dColorImageDownsampleThreshold=1.0', '-dGrayImageDownsampleThreshold=1.0', '-dMonoImageDownsampleThreshold=1.0'])
                jpeg_quality = "70" if self.fast_mode else ("70" if self.pdfa_dpi <= 100 else "85")
                cmd.extend(['-dColorImageFilter=/DCTEncode', f'-dJPEGQ={jpeg_quality}'])
                cmd.extend(['-dColorImageDownsampleType=/Bicubic', '-dGrayImageDownsampleType=/Bicubic'])


            cmd.extend([
                f'-sOutputFile={temp_pdf_path}',
                str(input_file)
            ])

            result = self._run_command(cmd)
            if not result or not temp_pdf_path.exists() or temp_pdf_path.stat().st_size == 0:
                error_details = result.stderr.strip() if result and result.stderr else "Unknown error"
                logging.error(f"Ghostscript PDF/A command failed: {' '.join(cmd)}")
                logging.error(f"Ghostscript STDERR: {error_details}")
                raise ProcessingError(f"Ghostscript PDF/A conversion failed: {error_details}")

            with pikepdf.open(temp_pdf_path, allow_overwriting_input=True) as pdf:
                with pdf.open_metadata(set_pikepdf_as_editor=False) as meta:
                    meta['pdf:Producer'] = 'MinimalPDF Optimizer'
                    meta['xmp:CreatorTool'] = 'MinimalPDF Optimizer'
                pdf.save(temp_pdf_path, object_stream_mode=pikepdf.ObjectStreamMode.generate, recompress_flate=True)

            self._post_process_pdf(temp_pdf_path, strip_metadata=False)
            shutil.move(str(temp_pdf_path), output_file)

        finally:
            if temp_pdf_path.exists():
                os.remove(temp_pdf_path)