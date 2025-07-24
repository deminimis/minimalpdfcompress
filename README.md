# MinimalPDF Compress

<img src="https://github.com/deminimis/minimalpdfcompress/blob/main/assets/pdf.png?raw=true" alt="Project Logo" width="250">


## Overview
Minimal PDF Compress is a user-friendly, cross-platform application designed to simplify PDF compression and conversion tasks using [Ghostscript](https://www.ghostscript.com/), [Pikepdf](https://github.com/pikepdf/pikepdf), and [cpdf](https://github.com/coherentgraphics/cpdf-binaries). It allows users to compress PDF files or convert them to PDF/A format with customizable options, all through an intuitive GUI (Graphical User Interface). 

<img src="https://github.com/deminimis/minimalpdfcompress/blob/main/assets/pic1.png?raw=true" alt="Dark Picture" style="max-width: 50%;">


Note: Ghostscript's pdfwrite device doesn't technically "compress" PDFs in the traditional sense. Instead, it recreates a new PDF that is generally smaller due to several optimizations. I added more traditional "compression" with Pikepdf. 



## How to Use

### Installation

1. 💾 Download the latest version [here](https://github.com/deminimis/minimalpdfcompress/releases).
2. Double-click `.exe` to launch the app. 
3. If you are not using the standalone version, it will prompt you to download Ghostscript if you haven't already.


<img src="https://github.com/deminimis/minimalpdfcompress/blob/main/assets/pic2.png?raw=true" alt="Dark Picture" style="max-width: 50%;">
<img src="https://github.com/deminimis/minimalpdfcompress/blob/main/assets/pic3.png?raw=true" alt="Dark Picture" style="max-width: 50%;">

### Prerequisites
- **Windows OS**: The app is currently optimized for Windows, as it uses `gswin64c.exe` and Windows-specific paths. It should be very simple to use on Linux if you have Ghostscript installed. If the popularity gets high I will make a Linux standalone version. 






### Usage Guide
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
   - You can also drag and drop a folder or file into the window (Windows only at the moment)
4. **Choose Operation**:
   - Select the operation from the dropdown:
     - **Compress (Screen - Smallest Size)**: Best for emailing or on-screen viewing.
     - **Compress (Ebook - Medium Size)**: Good quality for tablets and e-readers.
     - **Compress (Printer - High Quality)**: Optimized for printing on standard printers.
     - **Compress (Prepress - Highest Quality)**: Best quality for professional printing; results in the largest file size.
     - **Convert to PDF/A**: Converts the PDF to PDF/A-1b format for archival purposes.
5. **Advanced Options**:
   - Check the "Advanced Options" box to reveal additional settings:
     - **Resolution (dpi)**: Choose 72, 150, or 300 dpi as defaults. To the right you can customize the dpi. 
     - **Downscaling Factor**: Set to 1, 2, or 3 to reduce image resolution.
     - **Color Conversion Strategy**: Options include `LeaveColorUnchanged`, `Gray`, `RGB`, or `CMYK`.
     - **Downsample Method**: Choose `Subsample`, `Average`, or `Bicubic`.
     - **Enable Fast Web View**: Optimizes the PDF for faster online viewing.
     - **Subset Fonts**: Embeds only used font subsets.
     - **Compress Fonts**: Compresses embedded fonts.
     - **Compress PDF/A Output**: (Available for PDF/A conversion) Applies compression to PDF/A output.
     - **Remove metadata**: Removes any metadata it can. There might not be much it can remove.
     - **Pikepdf Compression Level**: A slider (0-9) to control the final compression strength applied by Pikepdf, after processing by Ghostscript.
     - **Decimal Precision**: Reduces file size for documents with vector art by limiting the number of decimal places in coordinates.
     - **Remove Annotations & Forms**: Strips all comments, highlights, and interactive form fields from the PDF.
     - **Overwrite original file**: Overwrites the original file rather than created a new processed file. 
   - *Note*: Some options (e.g., downscaling factor, color strategy) are constrained if you are using PDF/A, for PDF/A to ensure compliance.
6. **Process the PDF(s)**:
   - Click the "Process" button.
   - For folder inputs, confirm batch processing in the popup.
   - A UAC prompt may appear due to Ghostscript execution.
   - The status bar at the bottom will update (e.g., "Processing...", "Complete", or error messages).
7. **Check Results**:
   - Output files are saved to the specified location.
   - For batch processing, output files are named with prefixes like `out_filename_operation.pdf`.
  

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

   
## Technical Specifications

### Architecture and Design
- **Language**: Python 3.10+ (compatible with later versions).
- **GUI Framework**: Tkinter (standard library) for the interface, with `ttk` widgets for a native look.
- **External Dependency**: Ghostscript; Pikepdf.
- **File I/O**: Uses Python’s `pathlib` for cross-platform path handling and `subprocess` for running Ghostscript commands.

### Code Structure

* **`main.py`**: The main entry point of the application. Its sole responsibilities are to handle Windows DPI awareness and to initialize and run the main GUI window.

* **`gui.py`**: Contains the `GhostscriptGUI` class, which manages the entire user interface.

  * **Responsibilities**: Builds all widgets, manages Tkinter variables, handles user events (button clicks, drag-and-drop), saves/loads settings, and initiates the processing task in a separate thread to keep the UI responsive.

  * **Key Methods**: `build_gui()`, `process()`, `load_settings()`, `save_settings()`.

  * **Helper Class**: Includes a `Tooltip` class for creating hover-text explanations for UI elements.

* **`backend.py`**: Contains all the non-GUI logic for processing PDFs. It is called by `gui.py`.

  * **Responsibilities**: Locates the system's Ghostscript installation, constructs the command-line arguments, executes the Ghostscript subprocess, and applies final processing steps with Pikepdf.

  * **Key Methods**: `run_processing_task()`, `build_gs_command()`, `apply_final_processing()`, `find_ghostscript()`.

* **`styles.py`**: A dedicated module for styling the application.

  * **Responsibilities**: Defines the color palettes for light and dark modes and contains the `apply_theme()` function, which uses `ttk.Style` to configure the appearance of all widgets.

* **Threading**: The core processing logic (`run_processing_task`) is executed in a separate thread to prevent the GUI from freezing during potentially long operations. Callbacks are used to update the UI (status bar, buttons) when the task is complete.

* **Error Handling**:

  * Custom exceptions are defined in `backend.py` (e.g., `GhostscriptNotFound`, `ProcessingError`).

  * The backend's `run_processing_task` function uses a `try...except` block to catch errors and displays them to the user via a Tkinter `messagebox`.

  * Detailed logs are written to `ghostscript_gui.log` for debugging purposes.

### Ghostscript Integration
- **Command Construction**: Builds Ghostscript commands dynamically based on user inputs. Example command for compression:
  ```bash
  gswin64c.exe -sDEVICE=pdfwrite -dCompatibilityLevel=1.4 -dNOPAUSE -dBATCH -dQUIET -dSAFER -r150 -dDownScaleFactor=1 -dFastWebView=true -dSubsetFonts=true -dCompressFonts=true -dPDFSETTINGS=/screen -sColorConversionStrategy=LeaveColorUnchanged -sOutputFile=output.pdf input.pdf
  ```







## Contributing
Contributions are welcome! Please fork the repository, create a branch, and submit a pull request with your changes. 



