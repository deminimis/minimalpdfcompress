# MinimalPDF Compress

<div style="overflow: auto;">
<img src="https://github.com/deminimis/minimalpdfcompress/blob/main/assets/pdf1.png?raw=true" alt="Project Logo" width="20%" align="left" style="margin-right: 20px;">
<strong>MinimalPDF Compress</strong>  is a user-friendly, cross-platform application designed to simplify PDF compression and utility tasks. It leverages the power of multiple best-in-class backends, including Ghostscript, cpdf, oxipng, pngquant, jbig2, zopfli, and ECT, wrapping them in a single, intuitive interface. <br><br><br><br>


This application employs a highly-refined compression pipeline tailored to the user's selected mode. For lossy Compression, the process begins with Ghostscript, which rebuilds the PDF while downsampling images to the user-selected DPI. The file is then passed through a granular optimization stage where each image is individually optimized using specialized tools. For Lossless mode, the application uses pikepdf to iterate through the document, optimizing each image and data stream without reducing quality. All modes conclude with a final processing step using cpdf to linearize the file for fast web viewing, apply security settings, and squeeze out any final structural redundancies.

## Installation

1.  💾 **Download** the latest version from the [Releases page](https://github.com/deminimis/minimalpdfcompress/releases).
    
2.  Launch the application by double-clicking the **`.exe`** file.

## Features

<img src="https://github.com/deminimis/minimalpdfcompress/blob/main/assets/Compress.png?raw=true" alt="Light Picture" width="70%">
<img src="https://github.com/deminimis/minimalpdfcompress/blob/main/assets/live preview.png?raw=true" alt="Dark Picture" width="70%">
<img src="https://github.com/deminimis/minimalpdfcompress/blob/main/assets/output.png?raw=true" alt="PDF Tool tab" width="70%">

The application is organized into logical tabs for a clear and efficient workflow.

### Primary Tab: Smart Compression

The main tab offers a streamlined approach with four distinct optimization modes:

- **Compression**: Recreates the PDF using Ghostscript to intelligently downsample images. An interactive **custom slider** makes it easy to set the desired image resolution (DPI).
    
- **Lossless**: A powerful pipeline that reduces file size without any quality degradation. It works by re-compressing data streams and intelligently optimizing embedded images (PNG, JPG, etc.) using `oxipng`, `ECT`, and `jbig2`.
     - **Note**: actively being worked on, it's not truely lossless yet. 
    
- **PDF/A**: Converts the document to the PDF/A-2b archival format. This ensures long-term viewability but may increase file size.
    
- **Remove Images**: Strips all image data from the document, replacing them with blank placeholders. This is useful for text-only archival.
    
- **Common Options**: All modes support a rich set of options, including password protection, metadata stripping, darkening scanned text, converting to grayscale, and linearizing for fast web view.


### Utilities

A comprehensive suite of tools for common PDF management tasks:

- **Merge**: Combine multiple PDF files into a single document, with controls to reorder files before merging.
    
- **Split/Extract**: Break apart a PDF. Modes include splitting into single pages, splitting every 'N' pages, or extracting custom page ranges (e.g., `1-5, 8, 12-end`).
    
- **Rotate**: Rotate all pages of a document by 90°, 180°, or 270°. Includes a live preview of the rotation.
    
- **Delete Pages**: Remove specific pages or page ranges from a PDF.

- **Password**: Encrypt or decrypt PDF password. Must know the password for user-encrypted pdf, but can decrypt owner passwords without knowing the password. 
    
- **Stamp/Watermark**: Apply image or text-based watermarks. Features a live preview and full control over position, opacity, font, color, and size. Now includes image scaling.
    
- **Header/Footer**: Add dynamic page numbers, dates, filenames, or custom text to the top or bottom of pages.
    
- **Table of Contents**: Automatically generate a new, bookmarked Table of Contents page based on the document's existing bookmarks.
    
- **Metadata**: View and edit core metadata fields like Title, Author, Subject, and Keywords.
    
- **PDF to Image**: Convert PDF pages into high-quality PNG, JPEG, or TIFF images at a custom DPI.
    
- **PDF Repair**: Attempts to repair corrupted or damaged PDF files by rebuilding their structure.

## Building From Source

To create the executable from the source code:

1. **Install Python** (3.10+ recommended).
    
2. **Clone the repository** and navigate into the project directory.
    
3. **Install dependencies**:
    
    Bash
    
    ```
    pip install pyinstaller pillow pikepdf windnd oxipng
    ```
    
4. **Run the PyInstaller command** using the provided spec file. This correctly bundles all required binary tools and assets.
    
    Bash
    
    ```
    pyinstaller main.spec
    ```
    
5. Find your completed application inside the `dist/MinimalPDF v1.8` folder.
    

### Third party binaries
See: [Windows PDF Compression Binaries](https://github.com/deminimis/minimalpdfcompress/releases/tag/Win_Binaries)    
    

## Code Structure

The application is modularized to separate concerns between the UI, backend logic, and styling.

- `main.py`: The application's main entry point. Handles setup and launches the GUI.
    
- `gui.py`: Manages the entire Tkinter UI, including window layout, widgets, live previews, event handling, and threading for backend tasks.
    
- `backend.py`: Contains all non-GUI logic. It finds the required tools, builds commands, runs subprocesses, and handles PDF manipulation for all utility tabs.
    
- `pdf_optimizer.py`: Houses the core compression logic. The `PdfOptimizer` class manages the distinct pipelines for each compression mode.
    
- `styles.py`: Defines the application's visual theme and color palettes for modern light and dark modes.
    
- `constants.py`: Stores shared constants like operation names, UI text, and font lists to ensure consistency.
    
- `ui_components.py`: Contains reusable, custom-built Tkinter widgets, such as the `CompressionGauge`, `ModernToggle`, `CustomSlider`, and `PositionSelector`.
    
- `tooltips.py`: A centralized dictionary containing all tooltip text for the UI elements, making it easy to manage and update help text.
    



## FAQ

**Q: What's the difference between "Compression" and "Lossless" modes?**

**A:** **Compression** (lossy) reduces file size by removing data, primarily by downsampling images to a lower resolution (DPI). This leads to much smaller files but also a reduction in image quality. It's best for documents where perfect image fidelity isn't critical. **Lossless** reduces file size _without_ any loss of quality. It works by finding more efficient ways to store the exact same data, such as re-compressing data streams and optimizing image formats without changing the pixels.

**Q: Do I need to install Ghostscript or other tools separately?**

**A:** No. All required command-line tools (Ghostscript, cpdf, pngquant, etc.) are bundled with the application for your convenience. The application is self-contained and ready to run after unzipping.

**Q: Why did my file size _increase_ after compression?**

**A:** This can happen if the original PDF was already highly optimized or if you are converting to PDF/A, which can add structural data. You can use the "Don't save if larger than original" option in the Output Settings to prevent this.

**Q: Isn't Python slow? Why not use C++?**

**A:** While Python itself can be slower for CPU-intensive tasks, this application's heavy lifting is handled by powerful external programs written in C/C++ (like Ghostscript, cpdf, and oxipng). Python acts as the front-end user interface. The performance difference in the actual PDF processing is therefore negligible.

**Q: Is my data safe? Does this app connect to the internet?**

**A:** All PDF processing happens locally on your computer. The application does not send your files or any data over the internet. The only network activity occurs if you explicitly click the "Check for Updates" or "Buy me a coffee" links in the settings tab.

## Contributing

Contributions are welcome! Feel free to post problems in the Issues tab or suggestions in the Discussion tab.

If you want to implement code yourself, please fork the repository, create a feature branch, and submit a pull request with your changes.

## Special Thanks

This code relies on:

- [GhostScript](https://www.ghostscript.com/)
    
- [CoherentPDF](https://github.com/coherentgraphics/cpdf-binaries)
    
- [oxipng](https://github.com/shssoichiro/oxipng)
    
- [pngquant](https://pngquant.org/)
    
- [jpegoptim](https://github.com/tjko/jpegoptim)
    
- [zopfli](https://github.com/google/zopfli)
    
- [Efficient Compression Tool (ECT)](https://github.com/fhanau/Efficient-Compression-Tool)
    
- [jbig2enc](https://github.com/agl/jbig2enc)

PDF Testing with:

- [Sample-Files](https://sample-files.com/documents/pdf/)
