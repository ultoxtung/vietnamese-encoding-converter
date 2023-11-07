import colorama
from enum import Enum


# Cmd arguments
class InputType(Enum):
    DIRECTORY = 0
    FILE      = 1
    STRING    = 2


class Font:
    def __init__(self, typeface, panose=None):
        self.typeface = typeface
        self.panose = panose


# OpenOfficeXML namespace
NSMAP = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
    "cp": "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
    "dc": "http://purl.org/dc/elements/1.1/",
    "dcmitype": "http://purl.org/dc/dcmitype/",
    "dcterms": "http://purl.org/dc/terms/",
    "dgm": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
    "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
    "pic": "http://schemas.openxmlformats.org/drawingml/2006/picture",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "sl": "http://schemas.openxmlformats.org/schemaLibrary/2006/main",
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    'w14': "http://schemas.microsoft.com/office/word/2010/wordml",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "xml": "http://www.w3.org/XML/1998/namespace",
    "xsi": "http://www.w3.org/2001/XMLSchema-instance",
}

# Config file
CONFIG_FILE = './config.yaml'
SOFFICE_PATH_KEY = 'SOFFICE_PATH'
SOFFICE_CMD = 'soffice'

# Console logging
ERROR_COLOR = colorama.Fore.RED
PASS_COLOR = colorama.Fore.GREEN

# Font
# Leave default fonts on top
FONTS = [
    Font('Times New Roman', '02020603050405020304'),
    Font('Arial', '02110604020202020204'),
    Font('Tahoma', '02110604030504040204'),
    Font('Calibri', '02150502020204030204'),
]
