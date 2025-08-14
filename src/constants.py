# constants.py
APP_VERSION = "1.6"

OP_COMPRESS_SCREEN = "Compress (Screen - Smallest Size)"
OP_COMPRESS_EBOOK = "Compress (Ebook - Medium Size)"
OP_COMPRESS_PRINTER = "Compress (Printer - High Quality)"
OP_COMPRESS_PREPRESS = "Compress (Prepress - Highest Quality)"
OPERATIONS = [ OP_COMPRESS_SCREEN, OP_COMPRESS_EBOOK, OP_COMPRESS_PRINTER, OP_COMPRESS_PREPRESS ]

ROTATION_MAP = {
    "No Rotation": 0,
    "90° Right (Clockwise)": 90,
    "180°": 180,
    "90° Left (Counter-Clockwise)": 270
}

PDF_FONTS = [
    "Times-Roman", "Times-Bold", "Times-Italic", "Times-BoldItalic",
    "Helvetica", "Helvetica-Bold", "Helvetica-Oblique", "Helvetica-BoldOblique",
    "Courier", "Courier-Bold", "Courier-Oblique", "Courier-BoldOblique",
    "Symbol", "ZapfDingbats"
]

DPI_PRESETS = {
    OP_COMPRESS_SCREEN: "72",
    OP_COMPRESS_EBOOK: "150",
    OP_COMPRESS_PRINTER: "300",
    OP_COMPRESS_PREPRESS: "300"
}