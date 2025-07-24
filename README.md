# MinimalPDF Compress

<img src="https://github.com/deminimis/minimalpdfcompress/blob/main/assets/pdf.png?raw=true" alt="Project Logo" width="250">


## Overview
Minimal PDF Compress is a user-friendly, cross-platform application designed to simplify PDF compression and conversion tasks using [Ghostscript](https://www.ghostscript.com/), [Pikepdf](https://github.com/pikepdf/pikepdf), and [cpdf](https://github.com/coherentgraphics/cpdf-binaries). It allows users compress files, edit pages, apply watermarks, and more. 

<img src="https://github.com/deminimis/minimalpdfcompress/blob/main/assets/pic1.png?raw=true" alt="Dark Picture" style="max-width: 50%;">


Note: Ghostscript's pdfwrite device doesn't technically "compress" PDFs in the traditional sense. Instead, it recreates a new PDF that is generally smaller due to several optimizations. I added more traditional "compression" with Pikepdf and cpdf. 



## How to Use

### Installation

1. 💾 Download the latest version [here](https://github.com/deminimis/minimalpdfcompress/releases).
2. Double-click `.exe` to launch the app. 


<img src="https://github.com/deminimis/minimalpdfcompress/blob/main/assets/pic2.png?raw=true" alt="Light Picture" width="70%">
<img src="https://github.com/deminimis/minimalpdfcompress/blob/main/assets/pic3.png?raw=true" alt="Dark Picture" width="70%">
<img src="https://github.com/deminimis/minimalpdfcompress/blob/main/assets/new tabs.png?raw=true" alt="PDF Tool tab" width="70%">






### Usage Guide

#### Compress & Optimize Tab
_Tooltips have been added to the options inside the app to refer to._
1. **Launch the App**:
   - Open the app via the `.exe` or by running the Python script.
2. **Select Input**:
   - Click the "Input File or Folder" button.
   - Choose whether to process a single PDF file or a folder:
     - **File**: Select a `.pdf` file to process.
     - **Folder**: Select a folder to process all `.pdf` files in it (and its subfolders).
3. **Set Output Location**:
   - For a single file, the output path is automatically set to the input file’s directory with a suffix (e.g., `input_out.pdf`).
   - For a folder, the output path defaults to the input folder but can be changed.
   - Click "Browse" to manually choose the output location.
   - You can also drag and drop a folder or file into the window in the compression tab (Windows only at the moment)
4. **Choose Operation**:
   - Select the operation from the dropdown:
     - **Compress (Screen - Smallest Size)**: Best for emailing or on-screen viewing.
     - **Compress (Ebook - Medium Size)**: Good quality for tablets and e-readers.
     - **Compress (Printer - High Quality)**: Optimized for printing on standard printers.
     - **Compress (Prepress - Highest Quality)**: Best quality for professional printing; results in the largest file size.
5. **Advanced Options**:
   - Check the "Advanced Options" box to reveal additional settings:
* **Resolution (DPI)**: Set a custom DPI for images within the PDF.

   * **Downscaling & Downsampling**: Control how images are resized and re-sampled.
   * **Color Conversion**: Convert the entire document's color space (e.g., to Grayscale).
   * **Web Optimization**: Enable Fast Web View for quicker online loading.
   * **Font Control**: Subset and compress embedded fonts.
   * **Pikepdf Compression**: Apply a final compression pass with a level from 0-9.
   * **Decimal Precision**: Reduce file size for vector-heavy documents.
   * **Strip Content**: Remove metadata, annotations, and forms.
   * **Password Protection**: Secure your PDF with user and owner passwords.
   * **cpdf Optimizations**:
     * **Squeeze with cpdf**: Restructures the PDF to remove redundant objects, often significantly reducing file size.
     * **Darken Text (cpdf)**: Converts all text to pure black, perfect for improving the readability of scanned documents.
     * **Fast Processing (cpdf)**: Speeds up `cpdf` operations for modern, well-formed PDFs.

* **Overwrite Originals**: Option to overwrite input files instead of creating new ones. 
   - *Note*: Some options (e.g., downscaling factor, color strategy) are constrained if you are using PDF/A, for PDF/A to ensure compliance.
6. **Process the PDF(s)**:
   - Click the "Process" button.
   - For folder inputs, confirm batch processing in the popup.
   - A UAC prompt may appear due to Ghostscript execution.
   - The status bar at the bottom will update (e.g., "Processing...", "Complete", or error messages).
7. **Check Results**:
   - Output files are saved to the specified location.
   - For batch processing, output files are named with prefixes like `out_filename_operation.pdf`.

#### Page Tools Tab
   * **Merge PDFs**: Combine multiple files into one, with controls to reorder them first.
   * **Split PDF**: Split a document by single pages, every 'N' pages, or by a custom range (`1-5, 8, 12`).
   * **Delete Pages**: Remove specific pages or page ranges.
   * **Rotate Pages**: Rotate all pages in a document.

#### PDF Tools Tab
   * **Stamp & Watermark**:
     * **Image Stamping**: Stamp any JPG or PNG image onto a PDF, with full control over position, size, and opacity.
     * **Text Stamping**: Add custom text with controls for font, size, color, position, and opacity. Includes dynamic fields for **filename**, **date/time**, and sequential **Bates numbering**.
   * **Metadata Editor**: View, edit, and save the Title, Author, Subject, and Keywords fields.
   * **Convert PDF to Image**: Convert a PDF's pages to PNG, JPEG, or TIFF images at a specified DPI.
   * **Remove Opening Action**: Strip commands that force a PDF to open at a specific page or zoom level.
  

### FAQ

1. Isn't Python slow? Why not write in something like c++?  
    * Yes, and no. This app my open a bit slower, but the processing is primarily done with backends like Ghostscript written in C. So there will be little to no difference. 




### Building the Executable Yourself
To create a portable `.exe`:
1. Install PyInstaller:
   ```bash
   pip install pyinstaller
   ```
2. Run the following command in the directory containing the `.py` files and `pdf.ico`:
   ```bash
   pyinstaller --noconsole --icon="pdf.ico" --add-data="pdf.ico:." main.py
   ```
   * Alternatively, use the .spec files in the `src` folder and run: `pyinstaller main.spec` (changing the name of the spec for each type. To use the standalone versions you will need to include the ghostscript .dll and .exe in their lib and bin folders in the same directory. 
3. Find the `.exe` in the `dist` folder and distribute/run it.

   
## Code Structure
   
   * **Language**: Python 3.10+
   
   * **GUI Framework**: Tkinter with `ttk` widgets.
   
   * **Backends**: Ghostscript, Pikepdf, cpdf.
   
   * **Code Structure**:
   
     * `main.py`: Entry point. Handles DPI awareness and launches the GUI.
   
     * `gui.py`: Manages the entire UI, event handling, and threading.
   
     * `backend.py`: Contains all non-GUI logic. Constructs commands, runs subprocesses, and handles all PDF processing.
   
     * `styles.py`: Defines the application's visual theme for light and dark modes.

### Ghostscript Integration
- **Command Construction**: Builds Ghostscript commands dynamically based on user inputs. Example command for compression:
  ```bash
  gswin64c.exe -sDEVICE=pdfwrite -dCompatibilityLevel=1.4 -dNOPAUSE -dBATCH -dQUIET -dSAFER -r150 -dDownScaleFactor=1 -dFastWebView=true -dSubsetFonts=true -dCompressFonts=true -dPDFSETTINGS=/screen -sColorConversionStrategy=LeaveColorUnchanged -sOutputFile=output.pdf input.pdf
  ```







## Contributing
Contributions are welcome! Please fork the repository, create a branch, and submit a pull request with your changes. 



