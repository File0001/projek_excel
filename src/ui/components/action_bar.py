"""Action bar component containing convert button, progress bar, and status display."""

import os
import subprocess
from typing import Callable, Optional, Any
import customtkinter as ctk

from src.ui.theme import themeColors, themeFonts


class ActionBar(ctk.CTkFrame):
    def __init__(
        self,
        master: Any,
        onConvertCallback: Optional[Callable[[], None]] = None,
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
        self.onConvertCallback = onConvertCallback
        self.lastOutputFolder: Optional[str] = None
        self._buildUi()

    def _buildUi(self) -> None:
        self.grid_columnconfigure(0, weight=1)

        # Status text
        self.statusLabel = ctk.CTkLabel(
            self,
            text="Pilih file Excel dan sheet untuk memulai konversi.",
            font=themeFonts["body"],
            text_color=themeColors["textMuted"],
            anchor="w",
        )
        self.statusLabel.grid(row=0, column=0, padx=16, pady=(14, 6), sticky="w")

        # Progress bar
        self.progressBar = ctk.CTkProgressBar(
            self,
            progress_color=themeColors["brandGold"],
            fg_color=themeColors["cardSubBg"],
            height=8,
            corner_radius=4,
        )
        self.progressBar.grid(row=1, column=0, padx=16, pady=(0, 12), sticky="ew")
        self.progressBar.set(0.0)

        # Buttons row
        buttonRow = ctk.CTkFrame(self, fg_color="transparent")
        buttonRow.grid(row=2, column=0, padx=16, pady=(0, 14), sticky="ew")
        buttonRow.grid_columnconfigure(0, weight=1)

        self.convertButton = ctk.CTkButton(
            buttonRow,
            text="⚡ Konversi ke PDF",
            font=themeFonts["buttonLarge"],
            fg_color=themeColors["brandGold"],
            hover_color=themeColors["brandGoldHover"],
            text_color=themeColors["brandGoldText"],
            height=42,
            corner_radius=8,
            command=self._handleConvert,
        )
        self.convertButton.grid(row=0, column=0, sticky="ew")

        self.openFolderBtn = ctk.CTkButton(
            buttonRow,
            text="📂 Buka Folder",
            font=themeFonts["labelBold"],
            fg_color=themeColors["brandNavy"],
            hover_color=themeColors["brandNavyHover"],
            text_color="#FFFFFF",
            height=42,
            width=120,
            corner_radius=8,
            command=self._handleOpenFolder,
        )
        self.openFolderBtn.grid(row=0, column=1, padx=(10, 0), sticky="e")
        self.openFolderBtn.grid_remove()

    def _handleConvert(self) -> None:
        if self.onConvertCallback:
            self.onConvertCallback()

    def _handleOpenFolder(self) -> None:
        if self.lastOutputFolder and os.path.exists(self.lastOutputFolder):
            try:
                os.startfile(self.lastOutputFolder)
            except Exception:
                subprocess.Popen(["explorer", self.lastOutputFolder])

    def setRunning(self, isRunning: bool, message: str = "") -> None:
        if isRunning:
            self.convertButton.configure(state="disabled", text="⏳ Memproses Konversi...")
            self.statusLabel.configure(
                text=message or "Sedang memproses dokumen Excel...",
                text_color=themeColors["brandGold"],
            )
            self.progressBar.configure(mode="indeterminate")
            self.progressBar.start()
            self.openFolderBtn.grid_remove()
        else:
            self.convertButton.configure(state="normal", text="⚡ Konversi ke PDF")
            self.progressBar.stop()
            self.progressBar.configure(mode="determinate")

    def setSuccess(self, message: str, outputFolder: Optional[str] = None) -> None:
        self.setRunning(False)
        self.progressBar.set(1.0)
        self.statusLabel.configure(
            text=f"✓ {message}",
            text_color=themeColors["success"],
        )
        if outputFolder:
            self.lastOutputFolder = outputFolder
            self.openFolderBtn.grid()

    def setError(self, errorMessage: str) -> None:
        self.setRunning(False)
        self.progressBar.set(0.0)
        self.statusLabel.configure(
            text=f"✕ Error: {errorMessage}",
            text_color=themeColors["error"],
        )
