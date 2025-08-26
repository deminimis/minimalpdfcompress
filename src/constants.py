# constants.py
APP_VERSION = "1.7"

OP_LOSSLESS = "Lossless"

SPLIT_SINGLE = "Split to Single Pages"
SPLIT_EVERY_N = "Split Every N Pages"
SPLIT_CUSTOM = "Custom Range(s)"
SPLIT_MODES = [SPLIT_SINGLE, SPLIT_EVERY_N, SPLIT_CUSTOM]

STAMP_IMAGE = "Image"
STAMP_TEXT = "Text"

POS_CENTER = "Center"
POS_BOTTOM_LEFT = "Bottom-Left"
POS_BOTTOM_RIGHT = "Bottom-Right"
STAMP_POSITIONS = [POS_CENTER, POS_BOTTOM_LEFT, POS_BOTTOM_RIGHT]

IMG_FORMAT_PNG = "png"
IMG_FORMAT_JPEG = "jpeg"
IMG_FORMAT_TIFF = "tiff"
IMAGE_FORMATS = [IMG_FORMAT_PNG, IMG_FORMAT_JPEG, IMG_FORMAT_TIFF]

META_LOAD = 'load'
META_SAVE = 'save'

PRESET_TICKS = {
    "Monitor": 72,
    "E-book": 150,
    "Printer": 300,
}

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