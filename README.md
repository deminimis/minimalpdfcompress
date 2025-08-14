# MinimalPDF Compress

<div style="overflow: auto;">
<img src="https://github.com/deminimis/minimalpdfcompress/blob/main/assets/pdf.png?raw=true" alt="Project Logo" width="20%" align="left" style="margin-right: 20px;">
<strong>MinimalPDF Compress</strong> is a user-friendly, cross-platform application designed to simplify PDF compression and utility tasks. It combines the power of <a href="https://www.ghostscript.com/">Ghostscript</a>, <a href="https://github.com/pikepdf/pikepdf">Pikepdf</a>, and <a href="https://github.com/coherentgraphics/cpdf-binaries">cpdf</a> into a single intuitive interface, allowing users to compress files, merge and split pages, apply watermarks, and much more.
</div> <br>

<i><strong>Note:</strong> Ghostscript's `pdfwrite` device doesn't "compress" PDFs in the traditional sense. Instead, it recreates a new, optimized PDF. True stream compression has been added via Pikepdf and cpdf to further reduce file size.</i> <br><br>

<p align="center">
<img src="https://github.com/deminimis/minimalpdfcompress/blob/main/assets/pic1.png?raw=true" alt="Dark Picture" style="max-width: 50%;">
</p>

## Features

<img src="https://github.com/deminimis/minimalpdfcompress/blob/main/assets/pic2.png?raw=true" alt="Light Picture" width="70%">
<img src="https://github.com/deminimis/minimalpdfcompress/blob/main/assets/pic3.png?raw=true" alt="Dark Picture" width="70%">
<img src="https://github.com/deminimis/minimalpdfcompress/blob/main/assets/new tabs.png?raw=true" alt="PDF Tool tab" width="70%">

The application is organized into task-based tabs for a clear workflow:

-   **Compress**: Powerful PDF optimization with presets for screen, ebook, and print. Includes advanced controls for image resolution, color conversion, font subsetting, and password protection.
    
-   **Merge**: Combine multiple PDF files into a single document with tools to reorder files before merging.
    
-   **Split**: Break apart a PDF into separate files. Modes include splitting into single pages, splitting every 'N' pages, or extracting custom page ranges (e.g., `1-5, 8, 12-end`).
    
-   **Rotate**: Rotate all pages of a document by 90°, 180°, or 270°.
    
-   **Delete Pages**: Remove specific pages or ranges from a PDF.
    
-   **PDF to Images**: Convert PDF pages into high-quality PNG, JPEG, or TIFF images at a custom DPI.
    
-   **PDF/A**: Convert standard PDFs into the PDF/A-2b format for long-term archival.
    
-   **Stamp/Watermark**:
    
    -   **Image Stamping**: Apply any JPG or PNG as a watermark, with full control over position, size, and opacity.
        
    -   **Text Stamping**: Add custom text watermarks with controls for font, size, color, and opacity. Includes dynamic fields for **filename**, **date/time**, and sequential **Bates numbering**.
        
-   **Metadata**: View, edit, and save core metadata fields like Title, Author, Subject, and Keywords.
    
-   **Utilities**:
    
    -   **Remove Open Action**: Strip commands that force a PDF to open at a specific page or zoom level.
        

## How to Use

### Installation

1.  💾 **Download** the latest version from the [Releases page](https://github.com/deminimis/minimalpdfcompress/releases).
    
2.  Launch the application by double-clicking the **`.exe`** file.
    

### General Usage

1.  **Select a Tab**: Choose the task you want to perform (e.g., Compress, Merge, Split).
    
2.  **Select Input File(s)**: Use the "Browse" button to select your source PDF or folder. **Drag-and-drop** is also supported for all tabs on Windows.
    
3.  **Configure Options**: Adjust the settings available in the tab. Nearly every option has a **tooltip** explaining what it does—just hover over it.
    
4.  **Set Output Location**: Choose where to save the processed file.
    
5.  **Process**: Click the main button (e.g., "Process," "Merge," "Split") to start the task. A progress window will appear for longer operations.
    

## Building From Source

To create the executable yourself:

1.  **Install Python** (3.10+ recommended) and `pip`.
    
2.  **Clone the repository** and navigate into the project directory.
    
3.  **Install dependencies**:
    
    Bash
    
        pip install pyinstaller pillow pikepdf windnd
    
4.  **Run the PyInstaller command** using the provided spec file. This correctly bundles all required assets.
    
    Bash
    
        pyinstaller main.spec
    
5.  Find your completed application inside the `dist/MinimalPDF Compress v1.6` folder.
    

## Code Structure

The application is modularized to separate concerns between the UI, backend logic, and styling.

-   **`main.py`**: The main entry point. Handles high-DPI awareness on Windows and launches the GUI.
    
-   **`gui.py`**: Manages the entire UI, including window layout, widgets, event handling, and threading for backend tasks.
    
-   **`backend.py`**: Contains all non-GUI logic. It finds tools, builds shell commands, runs subprocesses, and handles all PDF manipulation tasks.
    
-   **`styles.py`**: Defines the application's visual theme and color palettes for light and dark modes.
    
-   **`constants.py`**: Stores shared constants like operation names and rotation maps to prevent duplication.
    
-   **`ui_components.py`**: Contains reusable custom UI widgets, such as the `FileSelector` and `Tooltip` classes.
    
-   **`tooltips.py`**: A dedicated file storing all the tooltip text strings for easy management.
    

## FAQ

**Q: Isn't Python slow? Why not write this in C++?** A: While Python's startup time can be slower, this application's heavy lifting (PDF processing) is handled by powerful backends like Ghostscript (written in C) and cpdf. The performance difference in processing is therefore negligible, while the speed of development and ease of maintenance in Python are significant advantages.

## Contributing

Contributions are welcome! Please fork the repository, create a feature branch, and submit a pull request with your changes.

