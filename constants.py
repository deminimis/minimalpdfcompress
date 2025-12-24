# constants.py
APP_VERSION = "1.9.5"

class ToolNotFound(Exception): pass
class GhostscriptNotFound(ToolNotFound): pass
class CpdfNotFound(ToolNotFound): pass
class PngquantNotFound(ToolNotFound): pass
class JpegoptimNotFound(ToolNotFound): pass
class EctNotFound(ToolNotFound): pass
class OptipngNotFound(ToolNotFound): pass
class OxipngNotFound(ToolNotFound): pass
class ProcessingError(Exception): pass

SPLIT_SINGLE = "Split to Single Pages"
SPLIT_EVERY_N = "Split Every N Pages"
SPLIT_CUSTOM = "Custom Range(s)"
SPLIT_MODES = [SPLIT_SINGLE, SPLIT_EVERY_N, SPLIT_CUSTOM]

STAMP_IMAGE = "Image"
STAMP_TEXT = "Text"

POS_TOP_LEFT = "Top Left"
POS_TOP_CENTER = "Top Center"
POS_TOP_RIGHT = "Top Right"
POS_MIDDLE_LEFT = "Middle Left"
POS_CENTER = "Center"
POS_MIDDLE_RIGHT = "Middle Right"
POS_BOTTOM_LEFT = "Bottom Left"
POS_BOTTOM_CENTER = "Bottom Center"
POS_BOTTOM_RIGHT = "Bottom Right"
STAMP_POSITIONS = [
    POS_TOP_LEFT, POS_TOP_CENTER, POS_TOP_RIGHT,
    POS_MIDDLE_LEFT, POS_CENTER, POS_MIDDLE_RIGHT,
    POS_BOTTOM_LEFT, POS_BOTTOM_CENTER, POS_BOTTOM_RIGHT,
]
PAGE_NUMBER_POSITIONS = [
    POS_TOP_LEFT, POS_TOP_CENTER, POS_TOP_RIGHT,
    POS_BOTTOM_LEFT, POS_BOTTOM_CENTER, POS_BOTTOM_RIGHT,
]

IMG_FORMAT_PNG = "png"
IMG_FORMAT_JPEG = "jpeg"
IMG_FORMAT_TIFF = "tiff"
IMAGE_FORMATS = [IMG_FORMAT_PNG, IMG_FORMAT_JPEG, IMG_FORMAT_TIFF]

META_LOAD = 'load'
META_SAVE = 'save'

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