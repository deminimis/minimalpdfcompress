# MinimalPDF Compress

<div style="overflow: auto;">
<img src="https://github.com/deminimis/minimalpdfcompress/blob/main/assets/pdf.png?raw=true" alt="Project Logo" width="20%" align="left" style="margin-right: 20px;">
<strong>MinimalPDF Compress</strong> is a user-friendly, Windows application designed to simplify PDF compression and utility tasks. It leverages the power of multiple backends, including Ghostscript, Pikepdf, cpdf, sam2p, jbig2enc, and PNGOUT, and wraps them in a single, intuitive interface. <br><br><br><br>


This application employs a multi-stage compression pipeline to achieve maximum file size reduction. For lossy operations, the process begins with Ghostscript, which rebuilds the entire PDF while downsampling images to the user-selected DPI. After this initial step, the application performs a granular optimization of the file's internal components. Each non-jpeg image is extracted and subjected to a competitive "bakeoff" using specialized tools such as sam2p, pngout, and jbig2, to find and re-insert the absolute smallest version. Concurrently, all non-image data streams, such as text and fonts, are aggressively re-compressed using zlib. The newly assembled file is then saved using Pikepdf, and as a final, mandatory step, the cpdf tool's -squeeze command is applied to the entire file, performing a final structural analysis to wring out any remaining redundancies.

## Installation

1.  💾 **Download** the latest version from the [Releases page](https://github.com/deminimis/minimalpdfcompress/releases).
    
2.  Launch the application by double-clicking the **`.exe`** file.

## Features

<img src="https://github.com/deminimis/minimalpdfcompress/blob/main/assets/Compress.png?raw=true" alt="Light Picture" width="70%">
<img src="https://github.com/deminimis/minimalpdfcompress/blob/main/assets/live preview.png?raw=true" alt="Dark Picture" width="70%">
<img src="https://github.com/deminimis/minimalpdfcompress/blob/main/assets/output.png?raw=true" alt="PDF Tool tab" width="70%">

The application is organized into logical tabs for a clear and efficient workflow.

### Main Tab: Compress

The compression tab has been rebuilt to offer two distinct optimization modes:

- **Lossy Compression (High-Quality Downsampling)**: Recreates the PDF using Ghostscript to intelligently downsample images. An interactive **Quality Dial** makes it easy to set the desired image resolution (DPI), with presets for `/screen` (72 DPI), `/ebook` (150 DPI), and `/printer` (300+ DPI).
    
- **Lossless Compression (Advanced Optimization)**: A powerful new pipeline that reduces file size without any quality degradation. It works by:
    
    - Re-compressing all non-image data streams with maximum zlib compression.
        
    - Intelligently optimizing embedded images by testing multiple formats and choosing the smallest one. This includes using **sam2p** and **PNGOUT** for PNGs and **jbig2enc** for monochrome images, often resulting in dramatic size reductions.
        
- **Common Options**: Both modes support additional options like password protection, stripping metadata, darkening scanned text, and removing interactive elements.


### Utilities

A suite of tools for common PDF management tasks:

- **Merge**: Combine multiple PDF files into a single document, with controls to reorder files before merging.
    
- **Split/Extract**: Break apart a PDF. Modes include splitting into single pages, splitting every 'N' pages, or extracting custom page ranges (e.g., `1-5, 8, 12-end`).
    
- **Rotate**: Rotate all pages of a document by 90°, 180°, or 270°. Includes a live preview of the rotation.
    
- **Delete Pages**: Remove specific pages or page ranges from a PDF.
    
- **Stamp/Watermark**: Apply image or text-based watermarks. Features a live preview and full control over position, opacity, font, color, and size.
    
- **Metadata**: View and edit core metadata fields like Title, Author, Subject, and Keywords.
    
- **PDF to Image**: Convert PDF pages into high-quality PNG, JPEG, or TIFF images at a custom DPI.
    
- **PDF/A**: Convert standard PDFs into the PDF/A format for long-term archival.
    
- **PDF Repair**: Attempts to repair corrupted or damaged PDF files by rebuilding their structure.

## Building From Source

To create the executable from the source code:

1. **Install Python** (3.10+ recommended).
    
2. **Clone the repository** and navigate into the project directory.
    
3. **Install dependencies**:
    
    Bash
    
    ```
    pip install pyinstaller pillow pikepdf windnd
    ```
    
4. **Run the PyInstaller command** using the provided spec file. This correctly bundles all required binary tools and assets.
    
    Bash
    
    ```
    pyinstaller main.spec
    ```
    
5. Find your completed application inside the
    
    `dist/MinimalPDF Compress v1.6` folder.
    

## Code Structure

The application is modularized to separate concerns between the UI, backend logic, and styling.

- `main.py`: The application's main entry point. Handles setup and launches the GUI.
    
- `gui.py`: Manages the entire Tkinter UI, including window layout, widgets, live previews, event handling, and threading for backend tasks.
    
- `backend.py`: Contains all non-GUI logic for the utility tabs (Merge, Split, etc.). It finds the required tools, builds commands, runs subprocesses, and handles PDF manipulation.
    
- `pdf_optimizer.py`: Houses the core compression logic. The `PdfOptimizer` class manages the complex pipelines for both lossless and lossy optimization.
    
- `styles.py`: Defines the application's visual theme and color palettes for modern light and dark modes.
    
- `constants.py`: Stores shared constants like operation names, UI text, and rotation maps to ensure consistency.
    
- `ui_components.py`: Contains reusable, custom-built Tkinter widgets, such as the `QualityDial`, `ModernToggle`, and `DropZone`.
    

## FAQ

**Q: What's the difference between Lossy and Lossless compression?**

**A:** **Lossy** reduces file size by removing data. This leads to much smaller files but also a reduction in quality. It's best for documents where perfect image fidelity isn't critical. **Lossless** compression reduces file size _without_ any loss of quality. It works by finding more efficient ways to store the exact same data, such as re-compressing data streams and optimizing image formats without changing the pixels.

**Q: Do I need to install Ghostscript or other tools separately?**

**A:** No. All required command-line tools (Ghostscript, cpdf, sam2p, etc.) are bundled with the application for your convenience. The application is self-contained and ready to run after unzipping.

**Q: Why did my file size _increase_ after "lossless" compression?**

**A:** This can happen if the original PDF was already highly optimized. 

**Q: Isn't Python slow? Why not use C++?**

**A:** While Python itself can be slower for CPU-intensive tasks, this application's heavy lifting is handled by powerful external programs written in C/C++ (like Ghostscript, cpdf, and jbig2enc). Python acts as the front-end user interface. The performance difference in the actual PDF processing is therefore negligible.

**Q: Is my data safe? Does this app connect to the internet?**

**A:** All PDF processing happens locally on your computer. The application does not send your files or any data over the internet. In fact, this program never connects to the internet (unless you click the two links in the settings tab to open the webpages). 

## Contributing

Contributions are welcome! Feel free to post problems in the Issues tab or suggestions in the Discussion tab. 

If you want to implement code yourself, please fork the repository, create a feature branch, and submit a pull request with your changes. 

## Special Thanks

This code relies on:

- [GhostScript](https://www.ghostscript.com/)
- [Pdfsizeopt](https://github.com/pts/pdfsizeopt-jbig2)
- [CoherentPDF](https://github.com/coherentgraphics/cpdf-binaries)
- [Sam2p](https://github.com/pts/sam2p)
- 

