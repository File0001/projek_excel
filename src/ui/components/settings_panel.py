"""Settings panel component for print layout and PDF output configurations."""

from typing import Dict, Any
import customtkinter as ctk

from src.config import (
    PaperSizeEnum,
    PageOrientationEnum,
    PageScalingEnum,
    OutputModeEnum,
)
from src.ui.theme import themeColors, themeFonts


class SettingsPanel(ctk.CTkFrame):
    def __init__(self, master: Any, **kwargs):
        super().__init__(
            master,
            fg_color=themeColors["cardBg"],
            corner_radius=12,
            border_width=1,
            border_color=themeColors["cardBorder"],
            **kwargs
        )
        self._buildUi()

    def _buildUi(self) -> None:
        self.grid_columnconfigure(0, weight=1)

        # Header
        ctk.CTkLabel(
            self,
            text="⚙️ Pengaturan Halaman & Cetak",
            font=themeFonts["sectionHeader"],
            text_color=themeColors["textPrimary"],
            anchor="w",
        ).grid(row=0, column=0, padx=16, pady=(14, 10), sticky="w")

        content = ctk.CTkFrame(self, fg_color="transparent")
        content.grid(row=1, column=0, padx=16, pady=(0, 14), sticky="nsew")
        content.grid_columnconfigure(0, weight=1)

        # 1. Paper Size
        ctk.CTkLabel(
            content,
            text="Ukuran Kertas",
            font=themeFonts["labelBold"],
            text_color=themeColors["textSecondary"],
            anchor="w",
        ).grid(row=0, column=0, pady=(4, 2), sticky="w")

        self.paperSizeVar = ctk.StringVar(value=PaperSizeEnum.A4.value)
        self.paperSizeSeg = ctk.CTkSegmentedButton(
            content,
            values=[PaperSizeEnum.A4.value, PaperSizeEnum.F4.value],
            variable=self.paperSizeVar,
            selected_color=themeColors["segmentedSelected"],
            selected_hover_color=themeColors["brandGoldHover"],
            unselected_color=themeColors["segmentedUnselected"],
            unselected_hover_color=themeColors["segmentedUnselectedHover"],
            text_color=themeColors["segmentedSelectedText"],
            height=32,
        )
        self.paperSizeSeg.grid(row=1, column=0, pady=(0, 12), sticky="ew")

        # 2. Orientation
        ctk.CTkLabel(
            content,
            text="Orientasi Halaman",
            font=themeFonts["labelBold"],
            text_color=themeColors["textSecondary"],
            anchor="w",
        ).grid(row=2, column=0, pady=(4, 2), sticky="w")

        self.orientationVar = ctk.StringVar(value=PageOrientationEnum.PORTRAIT.value)
        self.orientationSeg = ctk.CTkSegmentedButton(
            content,
            values=[PageOrientationEnum.PORTRAIT.value, PageOrientationEnum.LANDSCAPE.value],
            variable=self.orientationVar,
            selected_color=themeColors["segmentedSelected"],
            selected_hover_color=themeColors["brandGoldHover"],
            unselected_color=themeColors["segmentedUnselected"],
            unselected_hover_color=themeColors["segmentedUnselectedHover"],
            text_color=themeColors["segmentedSelectedText"],
            height=32,
        )
        self.orientationSeg.grid(row=3, column=0, pady=(0, 12), sticky="ew")

        # 3. Scaling
        ctk.CTkLabel(
            content,
            text="Penskalaan (Scaling)",
            font=themeFonts["labelBold"],
            text_color=themeColors["textSecondary"],
            anchor="w",
        ).grid(row=4, column=0, pady=(4, 2), sticky="w")

        scalingValues = [s.value for s in PageScalingEnum]
        self.scalingVar = ctk.StringVar(value=PageScalingEnum.NORMAL.value)
        self.scalingOption = ctk.CTkOptionMenu(
            content,
            values=scalingValues,
            variable=self.scalingVar,
            fg_color=themeColors["inputBg"],
            button_color=themeColors["brandNavy"],
            button_hover_color=themeColors["brandNavyHover"],
            dropdown_fg_color=themeColors["cardBg"],
            dropdown_hover_color=themeColors["cardSubBg"],
            text_color=themeColors["textPrimary"],
            dropdown_text_color=themeColors["textPrimary"],
            height=32,
        )
        self.scalingOption.grid(row=5, column=0, pady=(0, 12), sticky="ew")

        # 4. Output Mode
        ctk.CTkLabel(
            content,
            text="Mode Output PDF",
            font=themeFonts["labelBold"],
            text_color=themeColors["textSecondary"],
            anchor="w",
        ).grid(row=6, column=0, pady=(4, 2), sticky="w")

        self.outputModeVar = ctk.StringVar(value=OutputModeEnum.SINGLE_MERGED.value)
        self.outputModeSeg = ctk.CTkSegmentedButton(
            content,
            values=[OutputModeEnum.SINGLE_MERGED.value, OutputModeEnum.SEPARATE_FILES.value],
            variable=self.outputModeVar,
            selected_color=themeColors["segmentedSelected"],
            selected_hover_color=themeColors["brandGoldHover"],
            unselected_color=themeColors["segmentedUnselected"],
            unselected_hover_color=themeColors["segmentedUnselectedHover"],
            text_color=themeColors["segmentedSelectedText"],
            height=32,
        )
        self.outputModeSeg.grid(row=7, column=0, pady=(0, 4), sticky="ew")

    def getSettings(self) -> Dict[str, Any]:
        """Returns the user configuration mapped to config enums."""
        return {
            "paperSize": PaperSizeEnum(self.paperSizeVar.get()),
            "orientation": PageOrientationEnum(self.orientationVar.get()),
            "scaling": PageScalingEnum(self.scalingVar.get()),
            "outputMode": OutputModeEnum(self.outputModeVar.get()),
        }
