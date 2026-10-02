"""UI Theme and styling definitions aligned with CheapNFriendly brand identity.

Eye-friendly Soft Light Mode:
- Calming soft light grayish-blue background (#F0F3F8) eliminating harsh glare
- Pure crisp white surfaces (#FFFFFF) with gentle borders (#D6DFEC)
- Authentic brand harmony directly derived from the CheapNFriendly logo:
  * Brand Royal Navy (#1C2456) for primary headings, identity, and key buttons
  * Brand Crimson Red (#D81E5B) for the signature 'N' and accents
  * Brand Warm Amber Gold (#DF9E1B / #A6710A) for primary action and subtitle
- Deep navy-slate typography (#161D3B) for crystal-clear, strain-free readability
"""

from typing import Dict, Any


themeColors = {
    # Soft Light Mode Backgrounds
    "windowBg": "#F0F3F8",
    "cardBg": "#FFFFFF",
    "cardSubBg": "#EEF2F8",
    "cardBorder": "#D6DFEC",

    # Brand Colors (Aligned with Logo)
    "brandNavy": "#1C2456",
    "brandNavyHover": "#2B367A",
    "brandCrimson": "#D81E5B",
    "brandCrimsonHover": "#BA144A",
    "brandGold": "#DF9E1B",
    "brandGoldHover": "#C5880E",
    "brandGoldPressed": "#A87208",
    "brandGoldText": "#141A3B",

    # Typography (Eye-Friendly Soft Light)
    "textPrimary": "#161D3B",
    "textSecondary": "#4D5975",
    "textMuted": "#7E8BA4",
    "textCrimson": "#D81E5B",
    "textGold": "#A6710A",

    # Status & Utility
    "success": "#15803D",
    "warning": "#B45309",
    "error": "#DC2626",
    "info": "#1D4ED8",

    # Controls & Inputs
    "inputBg": "#F7FAFD",
    "inputBorder": "#CFD8E8",
    "checkboxActive": "#1C2456",
    "checkboxCheckmark": "#FFFFFF",
    "segmentedSelected": "#DF9E1B",
    "segmentedSelectedText": "#141A3B",
    "segmentedUnselected": "#EEF2F8",
    "segmentedUnselectedHover": "#E1E8F2",
}

# Typography configurations
fontFamily = "Segoe UI"

themeFonts = {
    "brandTitle": (fontFamily, 20, "bold"),
    "brandSubtitle": (fontFamily, 11, "bold"),
    "sectionHeader": (fontFamily, 13, "bold"),
    "labelBold": (fontFamily, 11, "bold"),
    "body": (fontFamily, 11),
    "bodySmall": (fontFamily, 10),
    "badge": (fontFamily, 9, "bold"),
    "buttonLarge": (fontFamily, 13, "bold"),
}


