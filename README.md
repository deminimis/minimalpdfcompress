# Minimal PDF Compress

<img src="https://github.com/deminimis/minimalpdfcompress/blob/main/assets/pdf.png?raw=true" alt="Project Logo" width="250">


## Overview
Minimal PDF Compress is a user-friendly, graphical desktop application designed to simplify PDF compression and conversion tasks using [Ghostscript](https://www.ghostscript.com/) and [Pikepdf](https://github.com/pikepdf/pikepdf). It allows users to compress PDF files or convert them to PDF/A format with customizable options, all through an intuitive GUI (Graphical User Interface). 


Note: Ghostscript's pdfwrite device doesn't technically "compress" PDFs in the traditional sense. Instead, it recreates a new PDF that is generally smaller due to several optimizations. I added more traditional "compression" with Pikepdf. 



## How to Use

### Installation

1. Download either the `.zip` or `.exe` from the [Releases page](https://github.com/deminimis/minimalpdfcompress/releases).
2. Double-click `.exe` to launch the app. 
3. If you are not using the standalone version, it will prompt you to download Ghostscript if you haven't already.


<img src="https://github.com/deminimis/minimalpdfcompress/blob/main/assets/darkpic.png?raw=true" alt="Dark Picture" style="max-width: 50%;">
<img src="https://github.com/deminimis/minimalpdfcompress/blob/main/assets/lightpic.png?raw=true" alt="Dark Picture" style="max-width: 50%;">

### Prerequisites
- **Windows OS**: The app is currently optimized for Windows, as it uses `gswin64c.exe` and Windows-specific paths. It should be very simple to use on Linux if you have Ghostscript installed. If the popularity gets high I will make a Linux standalone version. 






### Usage Guide
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
4. **Choose Operation**:
   - Select the operation from the dropdown:
     - **Compress (Screen - Smallest Size)**: Best for emailing or on-screen viewing.
     - **Compress (Ebook - Medium Size)**: Good quality for tablets and e-readers.
     - **Compress (Printer - High Quality)**: Optimized for printing on standard printers.
     - **Compress (Prepress - Highest Quality)**: Best quality for professional printing; results in the largest file size.
     - **Convert to PDF/A**: Converts the PDF to PDF/A-1b format for archival purposes.
5. **Advanced Options**:
   - Check the "Advanced Options" box to reveal additional settings:
     - **Resolution (dpi)**: Choose 72, 150, or 300 dpi.
     - **Downscaling Factor**: Set to 1, 2, or 3 to reduce image resolution.
     - **Color Conversion Strategy**: Options include `LeaveColorUnchanged`, `Gray`, `RGB`, or `CMYK`.
     - **Downsample Method**: Choose `Subsample`, `Average`, or `Bicubic`.
     - **Enable Fast Web View**: Optimizes the PDF for faster online viewing.
     - **Subset Fonts**: Embeds only used font subsets.
     - **Compress Fonts**: Compresses embedded fonts.
     - **Compress PDF/A Output**: (Available for PDF/A conversion) Applies compression to PDF/A output.
     - **Remove metadata**: Removes any metadata it can. There might not be much it can remove.
     - **Traditional compression**: Uses Pikepdf after Ghostscript optimizations. 
   - *Note*: Some options (e.g., downscaling factor, color strategy) are constrained if you are using PDF/A, for PDF/A to ensure compliance.
6. **Process the PDF(s)**:
   - Click the "Process" button.
   - For folder inputs, confirm batch processing in the popup.
   - A UAC prompt may appear due to Ghostscript execution.
   - The status bar at the bottom will update (e.g., "Processing...", "Complete", or error messages).
7. **Check Results**:
   - Output files are saved to the specified location.
   - For batch processing, output files are named with prefixes like `out_filename_operation.pdf`.


### Building the Executable Yourself
To create a portable `.exe`:
1. Install PyInstaller:
   ```bash
   pip install pyinstaller
   ```
2. Run the following command in the directory containing `ghostscript_gui.py` and `pdf.ico`:
   ```bash
   pyinstaller --name "Minimal PDF Compress" --onedir --windowed --icon="pdf.ico" --add-data "bin;bin" --add-data "lib;lib" main.py
   ```
3. Find the `.exe` in the `dist` folder and distribute/run it.

   
## Technical Specifications

### Architecture and Design
- **Language**: Python 3.10+ (compatible with later versions).
- **GUI Framework**: Tkinter (standard library) for the interface, with `ttk` widgets for a native look.
- **External Dependency**: Ghostscript; Pikepdf.
- **File I/O**: Uses Python’s `pathlib` for cross-platform path handling and `subprocess` for running Ghostscript commands.

### Code Structure
- **Main Class**: `GhostscriptGUI`
  - Initializes the Tkinter window, sets up variables, and builds the GUI.
  - Key methods:
    - `build_gui()`: Constructs the interface with sections for Input, Output, Operation, Advanced Options, and Status.
    - `run_ghostscript()`: Constructs and executes Ghostscript commands for single-file processing.
    - `process_single()`: Handles single-file processing.
    - `process_batch()`: Manages batch processing via a temporary `.bat` file to minimize UAC prompts.
    - `find_ghostscript()`: Locates the Ghostscript executable via predefined paths and Windows registry.
- **Dynamic Sizing**: The window uses a grid layout with `sticky="nsew"` and `columnconfigure`/`rowconfigure` weights to adapt to different screen sizes. Minimum size is set to 600x500 pixels to ensure content visibility.
- **Error Handling**:
  - Comprehensive exception handling for Ghostscript execution (e.g., `subprocess.TimeoutExpired`, `subprocess.CalledProcessError`).
  - User-friendly error messages via Tkinter `messagebox`.
  - Logging to `ghostscript_gui.log` for debugging.

### Ghostscript Integration
- **Command Construction**: Builds Ghostscript commands dynamically based on user inputs. Example command for compression:
  ```bash
  gswin64c.exe -sDEVICE=pdfwrite -dCompatibilityLevel=1.4 -dNOPAUSE -dBATCH -dQUIET -dSAFER -r150 -dDownScaleFactor=1 -dFastWebView=true -dSubsetFonts=true -dCompressFonts=true -dPDFSETTINGS=/screen -sColorConversionStrategy=LeaveColorUnchanged -sOutputFile=output.pdf input.pdf
  ```




### Limitations and Future Improvements
- **Platform**: Currently Windows-only due to Ghostscript executable naming (`gswin64c.exe`) and registry checks. Future versions could add macOS/Linux support by detecting `gs` and adjusting paths.
- **Ghostscript Dependency**: Requires separate installation of Ghostscript. This could be done easily with something like [Inno Setup](https://jrsoftware.org/isinfo.php). A future enhancement could bundle a Ghostscript binary, but I prefer this way as you don't have to put your trust into more entities.



## Contributing
Contributions are welcome! Please fork the repository, create a branch, and submit a pull request with your changes. 



