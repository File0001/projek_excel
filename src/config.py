"""Configuration and constants for Excel to PDF converter.

Adheres to camelCase variable naming and modular design.
"""

from enum import Enum
from typing import Dict, Any


class PaperSizeEnum(str, Enum):
    A4 = "A4"
    F4 = "F4 (Folio)"


class PageOrientationEnum(str, Enum):
    PORTRAIT = "Portrait"
    LANDSCAPE = "Landscape"


class PageScalingEnum(str, Enum):
    NORMAL = "Normal (100% / No Scaling)"
    FIT_SHEET = "Fit Sheet on One Page"
    FIT_ALL_COLUMNS = "Fit All Columns on One Page"
    FIT_ALL_ROWS = "Fit All Rows on One Page"


class OutputModeEnum(str, Enum):
    SINGLE_MERGED = "Gabung Jadi 1 File PDF"
    SEPARATE_FILES = "Pisah Per Sheet (Banyak File PDF)"


# Excel COM specific constants mapping
# Paper sizes: xlPaperA4 = 9, xlPaperFolio = 14 (Folio 8.5 x 13 in / 215 x 330 mm Indonesian F4 standard)
paperSizeComMapping: Dict[PaperSizeEnum, int] = {
    PaperSizeEnum.A4: 9,
    PaperSizeEnum.F4: 14,
}

# Fallback sequence for F4 if printer driver does not support 14
f4FallbackPaperSizes = [14, 41, 5]  # Folio (14) -> FanfoldGermanLegal (41) -> Legal (5)

# Orientations: xlPortrait = 1, xlLandscape = 2
orientationComMapping: Dict[PageOrientationEnum, int] = {
    PageOrientationEnum.PORTRAIT: 1,
    PageOrientationEnum.LANDSCAPE: 2,
}

# Scaling configurations: (fitToPagesWide, fitToPagesTall, zoom)
scalingComMapping: Dict[PageScalingEnum, Dict[str, Any]] = {
    PageScalingEnum.NORMAL: {
        "fitToPagesWide": False,
        "fitToPagesTall": False,
        "zoom": 100,
    },
    PageScalingEnum.FIT_SHEET: {
        "fitToPagesWide": 1,
        "fitToPagesTall": 1,
        "zoom": False,
    },
    PageScalingEnum.FIT_ALL_COLUMNS: {
        "fitToPagesWide": 1,
        "fitToPagesTall": False,
        "zoom": False,
    },
    PageScalingEnum.FIT_ALL_ROWS: {
        "fitToPagesWide": False,
        "fitToPagesTall": 1,
        "zoom": False,
    },
}

# Supported file extensions
supportedExtensions = (".xlsx", ".xls", ".xlsm")

# Application metadata
appTitle = "Excel to PDF Converter"
appVersion = "1.0.0"
defaultWindowWidth = 860
defaultWindowHeight = 720
