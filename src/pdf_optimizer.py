# pdf_optimizer.py
import subprocess
import logging
from pathlib import Path
import tempfile
import shutil
import pikepdf
from PIL import Image
import sys
import oxipng

def resource_path(relative_path):
    try:
        base_path = Path(sys._MEIPASS)
    except AttributeError:
        base_path = Path(__file__).parent
    return base_path / relative_path

class PdfOptimizer:
    _blank_image_data = b'\xff\xff\xff'

    def __init__(self, gs_path, cpdf_path, pngquant_path, jbig2_path, zopfli_path, q=None, **kwargs):
        self.gs_path = gs_path
        self.cpdf_path = cpdf_path
        self.pngquant_path = pngquant_path
        self.jpegoptim_path = kwargs.get('jpegoptim_path')
        self.jbig2_path = jbig2_path
        self.ect_path = kwargs.get('ect_path')
        self.zopfli_path = zopfli_path
        self.q = q
        self.user_password = kwargs.get('user_password')
        self.darken_text = kwargs.get('darken_text')
        self.remove_open_action = kwargs.get('remove_open_action')
        self.deep_png_compress = kwargs.get('deep_png_compress', False)
        self.linearize = kwargs.get('fast_web_view', False)
        self.fast_mode = kwargs.get('fast_mode', False)
        self.convert_to_grayscale = kwargs.get('convert_to_grayscale', False)

    def _run_command(self, cmd, check=True):
        use_shell = isinstance(cmd, str)
        logging.info(f"Running command: {cmd}")
        try:
            kwargs = {
                'check': check, 'capture_output': True, 'text': True,
                'errors': 'ignore', 'shell': use_shell
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

    def _optimize_image_stream(self, pdf, obj, temp_dir, mode='lossless', dpi=150):
        try:
            if not isinstance(obj, pikepdf.Stream) or obj.get("/Subtype") != "/Image":
                return 0

            if obj.get("/Filter") in ("/DCTDecode", "/JPXDecode"):
                if not self.ect_path and not self.jpegoptim_path:
                    return 0
                original_size = len(obj.read_raw_bytes())
                if original_size == 0:
                    return 0

                temp_jpg_path = temp_dir / f"img_{obj.objgen}.jpg"
                try:
                    pdf_image = pikepdf.PdfImage(obj)
                    pil_image = pdf_image.as_pil_image()

                    if pil_image.mode != 'RGB':
                        pil_image = pil_image.convert('RGB')

                    pil_image.save(temp_jpg_path, "jpeg", quality=95, optimize=True)

                    if self.jpegoptim_path and temp_jpg_path.exists():
                        cmd_jpegoptim = [self.jpegoptim_path, "--strip-all", "-q", str(temp_jpg_path)]
                        if mode == 'lossy':
                            quality = 85
                            if dpi <= 100: quality = 70
                            elif dpi <= 200: quality = 80
                            cmd_jpegoptim.insert(1, f"-m{quality}")
                        self._run_command(cmd_jpegoptim)

                    if self.ect_path and temp_jpg_path.exists() and not self.fast_mode:
                        cmd_ect = [self.ect_path, "-quiet", "-strip", "-progressive", "-3", str(temp_jpg_path)]
                        self._run_command(cmd_ect)

                    if temp_jpg_path.exists():
                        optimized_data = temp_jpg_path.read_bytes()
                        new_size = len(optimized_data)
                        if 0 < new_size < original_size:
                            obj.write(optimized_data)
                            obj.Filter = pikepdf.Name.DCTDecode
                            obj.ColorSpace = pikepdf.Name.DeviceRGB
                            obj.BitsPerComponent = 8
                            if '/Decode' in obj:
                                del obj['/Decode']
                            logging.info(f"Optimized JPEG {obj.objgen} with jpegoptim/ECT, saved {original_size - new_size} bytes.")
                            return original_size - new_size
                except Exception as e:
                    logging.warning(f"Could not process JPEG {obj.objgen}: {e}")
                return 0

            original_size = len(obj.read_raw_bytes())
            temp_img_path = temp_dir / f"img_{obj.objgen}.png"

            try:
                pdf_image = pikepdf.PdfImage(obj)
                pil_image = pdf_image.as_pil_image()
            except Exception as e:
                logging.warning(f"Could not extract image {obj.objgen}: {e}")
                return 0

            pil_image.save(temp_img_path, "png")
            optimized_path = None
            best_data = None

            if pil_image.mode == '1' and self.jbig2_path:
                pbm_path = temp_dir / f"img_{obj.objgen}.pbm"
                pil_image.save(pbm_path)
                jbig2_out_path = temp_dir / f"img_{obj.objgen}.jbig2"

                cmd_parts = [self.jbig2_path, "-p", str(pbm_path)]
                if mode == 'lossy':
                    cmd_parts.append("-s")

                cmd_str = f'"{cmd_parts[0]}" {" ".join(cmd_parts[1:])} > "{jbig2_out_path}"'

                if self._run_command(cmd_str, check=False) and jbig2_out_path.exists() and jbig2_out_path.stat().st_size > 0:
                    best_data = jbig2_out_path.read_bytes()
                    obj.Filter = pikepdf.Name("/JBIG2Decode")
                    if '/DecodeParms' in obj: del obj.DecodeParms

            elif mode == 'lossy' and self.pngquant_path:
                quality_str = "80-95"
                if dpi <= 100:
                    quality_str = "40-60"
                elif dpi <= 200:
                    quality_str = "65-80"

                quant_path = temp_dir / f"img_{obj.objgen}.quant.png"
                cmd = [self.pngquant_path, "--force", "--skip-if-larger", f"--quality={quality_str}", "--output", str(quant_path), "256", str(temp_img_path)]
                if self._run_command(cmd) and quant_path.exists() and quant_path.stat().st_size > 0:
                    optimized_path = quant_path

            try:
                source_for_oxipng = optimized_path if optimized_path else temp_img_path
                if source_for_oxipng.exists() and source_for_oxipng.stat().st_size > 0:
                    oxipng_out_path = temp_dir / f"img_{obj.objgen}.oxipng.png"
                    
                    options = {
                        "level": 2 if self.fast_mode else 6,
                        "strip": oxipng.StripChunks.safe()
                    }
                    
                    if mode == 'lossy':
                        options["optimize_alpha"] = True
                        options["scale_16"] = True
                        
                    oxipng.optimize(source_for_oxipng, oxipng_out_path, **options)

                    if oxipng_out_path.exists() and oxipng_out_path.stat().st_size > 0:
                        use_zopfli = self.deep_png_compress and self.zopfli_path and not self.fast_mode
                        if use_zopfli:
                            zopfli_out_path = temp_dir / f"img_{obj.objgen}.zopfli.png"
                            cmd_zopfli = [self.zopfli_path, str(oxipng_out_path), str(zopfli_out_path)]
                            if self._run_command(cmd_zopfli) and zopfli_out_path.exists() and zopfli_out_path.stat().st_size > 0:
                                current_best_data = zopfli_out_path.read_bytes()
                            else:
                                current_best_data = oxipng_out_path.read_bytes()
                        else:
                            current_best_data = oxipng_out_path.read_bytes()
                        
                        if best_data is None or len(current_best_data) < len(best_data):
                            best_data = current_best_data
                            obj.Filter = pikepdf.Name("/FlateDecode")

            except oxipng.PngError as e:
                logging.warning(f"Could not process PNG with pyoxipng: {e}")
            except Exception as e:
                logging.warning(f"An unexpected error occurred during pyoxipng processing: {e}")

            if best_data and len(best_data) < original_size:
                obj.write(best_data)
                logging.info(f"Optimized image {obj.objgen}, saved {original_size - len(best_data)} bytes.")
                return original_size - len(best_data)

        except Exception as e:
            logging.warning(f"Could not process image {obj.objgen}: {e}")

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

        if self.user_password:
            encrypted_path = final_path.with_name(f"encrypted_{final_path.name}")
            try:
                with pikepdf.open(final_path) as pdf:
                    pdf.save(encrypted_path, encryption=pikepdf.Encryption(user=self.user_password, owner=self.user_password, R=6))
                shutil.move(str(encrypted_path), str(final_path))
            except Exception as e:
                logging.error(f"Pikepdf encryption failed: {e}")

    def optimize_lossless(self, input_file, output_file, strip_metadata=False):
        with tempfile.TemporaryDirectory() as temp_dir_str:
            temp_dir = Path(temp_dir_str)
            temp_pdf_path = temp_dir / "processed.pdf"
            shutil.copy(input_file, temp_pdf_path)

            with pikepdf.open(temp_pdf_path, allow_overwriting_input=True) as pdf:
                if self.q: self.q.put(('status', f"Optimizing images losslessly..."))
                for obj in pdf.objects:
                    self._optimize_image_stream(pdf, obj, temp_dir, mode='lossless')

                if self.q: self.q.put(('status', f"Recompressing streams..."))
                pdf.save(object_stream_mode=pikepdf.ObjectStreamMode.generate, recompress_flate=True)

            if self.q: self.q.put(('status', f"Finalizing with cpdf..."))
            self._post_process_pdf(temp_pdf_path, strip_metadata)
            shutil.move(str(temp_pdf_path), output_file)

    def optimize_text_only(self, input_file, output_file, strip_metadata=False):
        with tempfile.TemporaryDirectory() as temp_dir_str:
            temp_dir = Path(temp_dir_str)
            temp_pdf_path = temp_dir / "processed.pdf"
            shutil.copy(input_file, temp_pdf_path)

            with pikepdf.open(temp_pdf_path, allow_overwriting_input=True) as pdf:
                if self.q: self.q.put(('status', f"Finding images to replace..."))
                image_objects = [obj for obj in pdf.objects if isinstance(obj, pikepdf.Stream) and obj.get("/Subtype") == "/Image"]
                
                if self.q: self.q.put(('status', f"Replacing {len(image_objects)} images with blanks..."))
                for obj in image_objects:
                    self._replace_image_stream(obj)

                if self.q: self.q.put(('status', f"Recompressing streams..."))
                pdf.save(object_stream_mode=pikepdf.ObjectStreamMode.generate, recompress_flate=True)

            if self.q: self.q.put(('status', f"Finalizing with cpdf..."))
            self._post_process_pdf(temp_pdf_path, strip_metadata)
            shutil.move(str(temp_pdf_path), output_file)

    def optimize_lossy(self, input_file, output_file, dpi, strip_metadata=False, remove_interactive=False, use_bicubic=False):
        with tempfile.TemporaryDirectory() as temp_dir_str:
            temp_dir = Path(temp_dir_str)
            gs_output_path = temp_dir / "gs_downsampled.pdf"

            if self.q: self.q.put(('status', f"Processing with Ghostscript..."))

            pdf_settings = "/screen" if self.fast_mode else "/ebook"

            cmd = [
                self.gs_path,
                "-sDEVICE=pdfwrite",
                "-dCompatibilityLevel=1.7",
                "-dNOPAUSE",
                "-dBATCH",
                "-dQUIET",
                f"-dPDFSETTINGS={pdf_settings}",
                f"-dColorImageResolution={dpi}",
                f"-dGrayImageResolution={dpi}",
                f"-dMonoImageResolution={dpi}",
                f"-sOutputFile={gs_output_path}"
            ]

            if self.convert_to_grayscale:
                cmd.extend([
                    "-sColorConversionStrategy=Gray",
                    "-dProcessColorModel=/DeviceGray",
                    "-dOverrideICC",
                ])

            if use_bicubic:
                cmd.extend([
                    '-dColorImageDownsampleType=/Bicubic',
                    '-dGrayImageDownsampleType=/Bicubic'
                ])

            if remove_interactive:
                cmd.extend(["-dShowAnnots=false", "-dShowAcroForm=false"])
            
            cmd.append(str(input_file))

            if not self._run_command(cmd) or not gs_output_path.exists() or gs_output_path.stat().st_size == 0:
                logging.warning("Ghostscript downsampling failed, proceeding with original file.")
                shutil.copy(input_file, gs_output_path)

            with pikepdf.open(gs_output_path, allow_overwriting_input=True) as pdf:
                if self.q: self.q.put(('status', f"Optimizing images with ECT/pngquant/pyoxipng..."))
                for obj in pdf.objects:
                    self._optimize_image_stream(pdf, obj, temp_dir, mode='lossy', dpi=dpi)

                if self.q: self.q.put(('status', f"Recompressing streams..."))
                pdf.save(object_stream_mode=pikepdf.ObjectStreamMode.generate, recompress_flate=True)

            if self.q: self.q.put(('status', f"Finalizing with cpdf..."))
            self._post_process_pdf(gs_output_path, strip_metadata)
            shutil.move(str(gs_output_path), output_file)

    def optimize_pdfa(self, input_file, output_file):
        if self.q: self.q.put(('status', "Converting to PDF/A..."))
        pdfa_def = resource_path('lib/PDFA_def.ps')
        if not pdfa_def.exists():
            raise FileNotFoundError("PDFA_def.ps not found in lib folder.")
        cmd = [
            self.gs_path,
            "-dPDFA=2",
            "-dBATCH",
            "-dNOPAUSE",
            "-sDEVICE=pdfwrite",
            "-dPDFACompatibilityPolicy=1",
            "-sColorConversionStrategy=sRGB",
            f"-sOutputFile={output_file}",
            str(pdfa_def),
            str(input_file)
        ]
        if not self._run_command(cmd):
            raise ProcessingError("Ghostscript PDF/A conversion failed.")