"""File selector component for choosing Excel workbooks."""

import os
from typing import Callable, Optional, Any
import customtkinter as ctk

from src.ui.theme import themeColors, themeFonts
from src.config import supportedExtensions


class FileSelector(ctk.CTkFrame):
    """Card component providing file selection for Excel workbooks."""

    def __init__(
        self,
        master: Any,
        onFileSelectedCallback: Optional[Callable[[str], None]] = None,
        **kwargs
    ):
        super().__init__(
            master,
            fg_color=themeColors["cardBg"],
            corner_radius=12,
            border_width=1,
            border_color=themeColors["cardBorder"],
            **kwargs
        )
        self.onFileSelectedCallback = onFileSelectedCallback
        self.selectedFilePath: Optional[str] = None

        self._buildUi()

    def _buildUi(self) -> None:
        self.grid_columnconfigure(0, weight=1)

        # Header section inside card
        headerFrame = ctk.CTkFrame(self, fg_color="transparent")
        headerFrame.grid(row=0, column=0, padx=16, pady=(14, 8), sticky="ew")
        headerFrame.grid_columnconfigure(0, weight=1)

        titleLabel = ctk.CTkLabel(
            headerFrame,
            text="📁 Pilih File Excel",
            font=themeFonts["sectionHeader"],
            text_color=themeColors["textPrimary"],
            anchor="w",
        )
        titleLabel.grid(row=0, column=0, sticky="w")

        hintLabel = ctk.CTkLabel(
            headerFrame,
            text=".xlsx, .xls, .xlsm",
            font=themeFonts["bodySmall"],
            text_color=themeColors["textMuted"],
            anchor="e",
        )
        hintLabel.grid(row=0, column=1, sticky="e")

        # File display container
        self.displayFrame = ctk.CTkFrame(
            self,
            fg_color=themeColors["inputBg"],
            corner_radius=8,
            border_width=1,
            border_color=themeColors["inputBorder"],
        )
        self.displayFrame.grid(row=1, column=0, padx=16, pady=(0, 14), sticky="ew")
        self.displayFrame.grid_columnconfigure(0, weight=1)

        self.fileNameLabel = ctk.CTkLabel(
            self.displayFrame,
            text="Belum ada file yang dipilih",
            font=themeFonts["body"],
            text_color=themeColors["textMuted"],
            anchor="w",
        )
        self.fileNameLabel.grid(row=0, column=0, padx=12, pady=10, sticky="ew")

        # Action button
        self.browseButton = ctk.CTkButton(
            self.displayFrame,
            text="Jelajahi...",
            font=themeFonts["labelBold"],
            fg_color=themeColors["brandNavy"],
            hover_color=themeColors["brandNavyHover"],
            text_color="#FFFFFF",
            corner_radius=6,
            width=90,
            height=32,
            command=self._handleBrowse,
        )
        self.browseButton.grid(row=0, column=1, padx=(6, 8), pady=6, sticky="e")

        # Metadata info row (size / path)
        self.metaLabel = ctk.CTkLabel(
            self,
            text="",
            font=themeFonts["bodySmall"],
            text_color=themeColors["textSecondary"],
            anchor="w",
        )
        self.metaLabel.grid(row=2, column=0, padx=16, pady=(0, 10), sticky="w")
        self.metaLabel.grid_remove()  # Hidden until file selected

    def _handleBrowse(self) -> None:
        fileFilter = [
            ("Excel Workbooks", "*.xlsx *.xls *.xlsm"),
            ("All Files", "*.*"),
        ]
        chosenFile = ctk.filedialog.askopenfilename(
            title="Pilih Dokumen Excel",
            filetypes=fileFilter,
        )
        if chosenFile:
            self.setFile(chosenFile)

    def setFile(self, filePath: str) -> None:
        """Sets the selected file and triggers callback."""
        if not os.path.exists(filePath):
            return

        self.selectedFilePath = filePath
        fileName = os.path.basename(filePath)
        fileSizeBytes = os.path.getsize(filePath)
        fileSizeFormatted = self._formatFileSize(fileSizeBytes)

        self.fileNameLabel.configure(
            text=f"📄 {fileName}",
            text_color=themeColors["textPrimary"],
        )

        self.metaLabel.configure(
            text=f"Ukuran: {fileSizeFormatted}  •  Lokasi: {filePath}"
        )
        self.metaLabel.grid()

        if self.onFileSelectedCallback:
            self.onFileSelectedCallback(filePath)

    def getFilePath(self) -> Optional[str]:
        return self.selectedFilePath

    @staticmethod
    def _formatFileSize(sizeInBytes: int) -> str:
        if sizeInBytes < 1024:
            return f"{sizeInBytes} B"
        elif sizeInBytes < 1024 * 1024:
            return f"{sizeInBytes / 1024:.1f} KB"
        else:
            return f"{sizeInBytes / (1024 * 1024):.2f} MB"
