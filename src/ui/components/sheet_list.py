"""Sheet list component allowing multi-sheet selection."""

from typing import List, Dict, Any, Callable, Optional
import customtkinter as ctk
from src.ui.theme import themeColors, themeFonts


class SheetList(ctk.CTkFrame):
    def __init__(self, master: Any, onSelectionChanged: Optional[Callable[[], None]] = None, **kwargs):
        super().__init__(
            master,
            fg_color=themeColors["cardBg"],
            corner_radius=12,
            border_width=1,
            border_color=themeColors["cardBorder"],
            **kwargs
        )
        self.onSelectionChanged = onSelectionChanged
        self.sheetItems: List[Dict[str, Any]] = []
        self._buildUi()

    def _buildUi(self) -> None:
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(2, weight=1)

        # Header
        header = ctk.CTkFrame(self, fg_color="transparent")
        header.grid(row=0, column=0, padx=16, pady=(12, 6), sticky="ew")
        header.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            header,
            text="📑 Daftar Sheet",
            font=themeFonts["sectionHeader"],
            text_color=themeColors["textPrimary"],
        ).grid(row=0, column=0, sticky="w")

        self.counterLabel = ctk.CTkLabel(
            header,
            text="0 sheet",
            font=themeFonts["bodySmall"],
            text_color=themeColors["brandNavy"],
        )
        self.counterLabel.grid(row=0, column=1, sticky="e")

        # Action buttons
        actionBar = ctk.CTkFrame(self, fg_color="transparent")
        actionBar.grid(row=1, column=0, padx=16, pady=(0, 8), sticky="ew")

        ctk.CTkButton(
            actionBar,
            text="Pilih Semua",
            font=themeFonts["bodySmall"],
            fg_color=themeColors["cardSubBg"],
            hover_color=themeColors["cardBorder"],
            text_color=themeColors["textPrimary"],
            corner_radius=6,
            height=26,
            width=80,
            command=self.selectAll,
        ).pack(side="left", padx=(0, 6))

        ctk.CTkButton(
            actionBar,
            text="Batal Pilih",
            font=themeFonts["bodySmall"],
            fg_color=themeColors["cardSubBg"],
            hover_color=themeColors["cardBorder"],
            text_color=themeColors["textMuted"],
            corner_radius=6,
            height=26,
            width=80,
            command=self.deselectAll,
        ).pack(side="left")

        # Scrollable container
        self.scrollFrame = ctk.CTkScrollableFrame(
            self,
            fg_color=themeColors["inputBg"],
            corner_radius=8,
            border_width=1,
            border_color=themeColors["inputBorder"],
        )
        self.scrollFrame.grid(row=2, column=0, padx=16, pady=(0, 12), sticky="nsew")
        self.scrollFrame.grid_columnconfigure(0, weight=1)

        self.emptyLabel = ctk.CTkLabel(
            self.scrollFrame,
            text="Pilih file Excel terlebih dahulu.",
            font=themeFonts["bodySmall"],
            text_color=themeColors["textMuted"],
        )
        self.emptyLabel.grid(row=0, column=0, pady=30, padx=12)

    def populateSheets(self, sheets: List[Dict[str, Any]]) -> None:
        for item in self.sheetItems:
            item["row"].destroy()
        self.sheetItems.clear()

        if not sheets:
            self.emptyLabel.configure(text="Tidak ada sheet ditemukan.")
            self.emptyLabel.grid()
            self._updateCounter()
            return

        self.emptyLabel.grid_remove()
        for idx, sheet in enumerate(sheets):
            if isinstance(sheet, dict):
                name = sheet.get("name", f"Sheet{idx+1}")
                isVisible = sheet.get("isVisible", sheet.get("visible", True))
            else:
                name = getattr(sheet, "name", f"Sheet{idx+1}")
                isVisible = getattr(sheet, "isVisible", getattr(sheet, "visible", True))

            row = ctk.CTkFrame(self.scrollFrame, fg_color="transparent")
            row.grid(row=idx, column=0, sticky="ew", padx=4, pady=2)
            row.grid_columnconfigure(0, weight=1)

            var = ctk.BooleanVar(value=isVisible)
            chk = ctk.CTkCheckBox(
                row,
                text=name,
                variable=var,
                font=themeFonts["body"],
                text_color=themeColors["textPrimary"] if isVisible else themeColors["textMuted"],
                fg_color=themeColors["checkboxActive"],
                hover_color=themeColors["brandNavyHover"],
                checkmark_color=themeColors["checkboxCheckmark"],
                border_color=themeColors["cardBorder"],
                corner_radius=4,
                command=self._onCheckChanged,
            )
            chk.grid(row=0, column=0, sticky="w")

            if not isVisible:
                badge = ctk.CTkLabel(
                    row,
                    text="Hidden",
                    font=themeFonts["badge"],
                    text_color=themeColors["textMuted"],
                    fg_color=themeColors["cardSubBg"],
                    corner_radius=4,
                    padx=4,
                    pady=1,
                )
                badge.grid(row=0, column=1, sticky="e")

            self.sheetItems.append({"name": name, "variable": var, "row": row})

        self._updateCounter()

    def selectAll(self) -> None:
        for item in self.sheetItems:
            item["variable"].set(True)
        self._updateCounter()

    def deselectAll(self) -> None:
        for item in self.sheetItems:
            item["variable"].set(False)
        self._updateCounter()

    def _onCheckChanged(self) -> None:
        self._updateCounter()
        if self.onSelectionChanged:
            self.onSelectionChanged()

    def _updateCounter(self) -> None:
        total = len(self.sheetItems)
        selected = len(self.getSelectedSheetNames())
        self.counterLabel.configure(text=f"{selected} / {total} dipilih")

    def getSelectedSheetNames(self) -> List[str]:
        return [it["name"] for it in self.sheetItems if it["variable"].get()]
