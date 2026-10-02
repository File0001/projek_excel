"""Main application window for CheapNFriendly Excel to PDF Converter."""

import os
import queue
import threading
from typing import Optional, List, Dict, Any, Callable
import tkinter as tk
from tkinter import filedialog
import customtkinter as ctk

from src.config import appTitle, appVersion, OutputModeEnum
from src.core.excel_inspector import inspectWorkbook
from src.core.excel_converter import convertExcelToPdf, ConversionOptions
from src.ui.theme import themeColors, themeFonts
from src.ui.components import FileSelector, SheetList, SettingsPanel, ActionBar
from src.utils.logger import appLogger


class MainWindow(ctk.CTk):
    """Primary application window with 2-column layout."""

    def __init__(self) -> None:
        ctk.set_appearance_mode("light")
        super().__init__()
        self.title(f"{appTitle} v{appVersion} - CheapNFriendly")
        self.geometry("980x680")
        self.minsize(860, 600)
        self.configure(fg_color=themeColors["windowBg"])

        self.currentFilePath: Optional[str] = None
        self.isConverting: bool = False
        self.isDestroyed: bool = False
        self.taskQueue: queue.Queue = queue.Queue()

        self._buildUi()
        self._pollTaskQueue()

    def runOnMainThread(self, callback: Callable[[], None]) -> None:
        """Thread-safe dispatch of UI updates to the main Tkinter thread."""
        self.taskQueue.put(callback)

    def _pollTaskQueue(self) -> None:
        """Periodically flushes queued tasks on the main Tkinter thread."""
        if self.isDestroyed:
            return
        try:
            while not self.taskQueue.empty():
                callback = self.taskQueue.get_nowait()
                try:
                    callback()
                except Exception as taskErr:
                    appLogger.error(f"Error dalam task main thread: {taskErr}", exc_info=True)
        finally:
            if not self.isDestroyed:
                self.after(35, self._pollTaskQueue)

    def destroy(self) -> None:
        self.isDestroyed = True
        super().destroy()

    def _buildUi(self) -> None:
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(2, weight=1)

        # 1. Top Crimson Brand Stripe
        stripe = ctk.CTkFrame(self, fg_color=themeColors["brandCrimson"], height=3, corner_radius=0)
        stripe.grid(row=0, column=0, sticky="ew")

        # 2. Modern Header
        self._buildHeader()

        # 3. Two-Column Dashboard Content
        content = ctk.CTkFrame(self, fg_color="transparent")
        content.grid(row=2, column=0, padx=20, pady=(0, 10), sticky="nsew")
        content.grid_columnconfigure(0, weight=5)
        content.grid_columnconfigure(1, weight=5)
        content.grid_rowconfigure(0, weight=1)

        # Left Column
        leftCol = ctk.CTkFrame(content, fg_color="transparent")
        leftCol.grid(row=0, column=0, padx=(0, 8), sticky="nsew")
        leftCol.grid_columnconfigure(0, weight=1)
        leftCol.grid_rowconfigure(1, weight=1)

        self.fileSelector = FileSelector(leftCol, onFileSelectedCallback=self._onFileSelected)
        self.fileSelector.grid(row=0, column=0, pady=(0, 10), sticky="ew")

        self.sheetList = SheetList(leftCol, onSelectionChanged=self._onSheetSelectionChanged)
        self.sheetList.grid(row=1, column=0, sticky="nsew")

        # Right Column
        rightCol = ctk.CTkFrame(content, fg_color="transparent")
        rightCol.grid(row=0, column=1, padx=(8, 0), sticky="nsew")
        rightCol.grid_columnconfigure(0, weight=1)
        rightCol.grid_rowconfigure(0, weight=1)

        self.settingsPanel = SettingsPanel(rightCol)
        self.settingsPanel.grid(row=0, column=0, pady=(0, 10), sticky="nsew")

        self.actionBar = ActionBar(rightCol, onConvertCallback=self._onConvertRequested)
        self.actionBar.grid(row=1, column=0, sticky="ew")

        # 4. Footer
        footer = ctk.CTkFrame(self, fg_color="transparent", height=20)
        footer.grid(row=3, column=0, padx=20, pady=(0, 6), sticky="ew")
        ctk.CTkLabel(
            footer,
            text="CHEAP N FRIENDLY DigitAllPrinting  •  Excel to PDF Native Engine",
            font=themeFonts["bodySmall"],
            text_color=themeColors["textMuted"],
        ).pack(side="left")

    def _buildHeader(self) -> None:
        header = ctk.CTkFrame(
            self,
            fg_color=themeColors["cardBg"],
            corner_radius=0,
            height=64,
            border_width=1,
            border_color=themeColors["cardBorder"],
        )
        header.grid(row=1, column=0, pady=(0, 12), sticky="ew")
        header.grid_columnconfigure(1, weight=1)

        # Canvas with brand monogram emblem (Navy + Gold)
        icon = tk.Canvas(header, width=38, height=38, bg=themeColors["cardBg"], highlightthickness=0)
        icon.grid(row=0, column=0, padx=(18, 10), pady=10)
        icon.create_polygon([4, 6, 34, 6, 28, 12, 10, 12], fill=themeColors["brandNavy"], outline=themeColors["brandNavyHover"])
        icon.create_polygon([8, 14, 30, 14, 20, 22, 14, 22, 19, 34, 12, 34], fill=themeColors["brandGold"], outline=themeColors["brandGoldHover"])

        # Brand title matching logo typography
        box = ctk.CTkFrame(header, fg_color="transparent")
        box.grid(row=0, column=1, sticky="w")
        row = ctk.CTkFrame(box, fg_color="transparent")
        row.pack(anchor="w")

        ctk.CTkLabel(row, text="CHEAP", font=themeFonts["brandTitle"], text_color=themeColors["brandNavy"]).pack(side="left")
        ctk.CTkLabel(row, text="N", font=themeFonts["brandTitle"], text_color=themeColors["brandCrimson"]).pack(side="left")
        ctk.CTkLabel(row, text="FRIENDLY", font=themeFonts["brandTitle"], text_color=themeColors["brandNavy"]).pack(side="left")
        ctk.CTkLabel(box, text="DigitAllPrinting  •  Excel to PDF", font=themeFonts["brandSubtitle"], text_color=themeColors["textGold"]).pack(anchor="w")

        ctk.CTkLabel(
            header,
            text="PRO",
            font=themeFonts["badge"],
            text_color="#FFFFFF",
            fg_color=themeColors["brandNavy"],
            corner_radius=4,
            padx=8,
            pady=2,
        ).grid(row=0, column=2, padx=18, sticky="e")

    def _onFileSelected(self, filePath: str) -> None:
        self.currentFilePath = filePath
        self.actionBar.statusLabel.configure(
            text="Membaca struktur sheet...",
            text_color=themeColors["textGold"],
        )

        def inspectTask():
            try:
                sheets = inspectWorkbook(filePath)
                self.runOnMainThread(lambda: self._onInspectSuccess(sheets))
            except Exception as inspectErr:
                appLogger.error(f"Gagal inspect file: {inspectErr}")
                self.runOnMainThread(lambda err=inspectErr: self.actionBar.setError(str(err)))

        threading.Thread(target=inspectTask, daemon=True).start()

    def _onInspectSuccess(self, sheets: List[Any]) -> None:
        try:
            self.sheetList.populateSheets(sheets)
            totalSheets = len(sheets)
            self.actionBar.statusLabel.configure(
                text=f"Siap. {totalSheets} sheet ditemukan. Silakan konfigurasi dan konversi.",
                text_color=themeColors["textSecondary"],
            )
        except Exception as populateErr:
            appLogger.error(f"Gagal menampilkan daftar sheet: {populateErr}", exc_info=True)
            self.actionBar.setError(f"Gagal memproses daftar sheet: {populateErr}")

    def _onSheetSelectionChanged(self) -> None:
        selected = self.sheetList.getSelectedSheetNames()
        if not selected:
            self.actionBar.statusLabel.configure(
                text="Peringatan: Tidak ada sheet yang dipilih.",
                text_color=themeColors["warning"],
            )

    def _onConvertRequested(self, customOutputPath: Optional[str] = None) -> None:
        if self.isConverting:
            return

        if not self.currentFilePath or not os.path.exists(self.currentFilePath):
            self.actionBar.setError("Pilih file Excel terlebih dahulu!")
            return

        selectedSheets = self.sheetList.getSelectedSheetNames()
        if not selectedSheets:
            self.actionBar.setError("Pilih minimal satu sheet untuk dikonversi!")
            return

        settings = self.settingsPanel.getSettings()
        isMerge = settings["outputMode"] == OutputModeEnum.SINGLE_MERGED
        sourceDir = os.path.dirname(self.currentFilePath)
        baseName = os.path.splitext(os.path.basename(self.currentFilePath))[0]

        if customOutputPath:
            outputPath = customOutputPath
        elif isMerge:
            defaultFileName = f"{baseName}_converted.pdf"
            selectedPath = filedialog.asksaveasfilename(
                parent=self,
                title="Simpan File PDF Sebagai",
                initialdir=sourceDir,
                initialfile=defaultFileName,
                defaultextension=".pdf",
                filetypes=[("PDF Documents (*.pdf)", "*.pdf"), ("Semua File (*.*)", "*.*")],
            )
            if not selectedPath:
                self.actionBar.statusLabel.configure(
                    text="Konversi dibatalkan oleh pengguna.",
                    text_color=themeColors["textSecondary"],
                )
                return
            outputPath = selectedPath
        else:
            selectedFolder = filedialog.askdirectory(
                parent=self,
                title="Pilih Folder Penyimpanan Hasil PDF",
                initialdir=sourceDir,
            )
            if not selectedFolder:
                self.actionBar.statusLabel.configure(
                    text="Konversi dibatalkan oleh pengguna.",
                    text_color=themeColors["textSecondary"],
                )
                return
            outputPath = selectedFolder

        self.isConverting = True
        self.actionBar.setRunning(True, "Memulai konversi dokumen Excel...")

        def runConversion():
            try:
                def progressCallback(percent: int, message: str):
                    self.runOnMainThread(lambda p=percent, m=message: self._updateProgress(p, m))

                options = ConversionOptions(
                    inputExcelPath=self.currentFilePath,
                    outputPdfPath=outputPath,
                    targetSheetNames=selectedSheets,
                    paperSize=settings["paperSize"],
                    orientation=settings["orientation"],
                    scaling=settings["scaling"],
                    mergeToSinglePdf=isMerge,
                    progressCallback=progressCallback,
                )

                resultPath = convertExcelToPdf(options)
                outFolder = resultPath if os.path.isdir(resultPath) else os.path.dirname(resultPath)
                self.runOnMainThread(lambda: self._onConversionSuccess(resultPath, outFolder))

            except Exception as convErr:
                appLogger.error(f"Gagal saat konversi: {convErr}", exc_info=True)
                self.runOnMainThread(lambda err=convErr: self._onConversionError(str(err)))

        threading.Thread(target=runConversion, daemon=True).start()

    def _updateProgress(self, percent: int, message: str) -> None:
        self.actionBar.progressBar.set(percent / 100.0)
        self.actionBar.statusLabel.configure(
            text=f"({percent}%) {message}",
            text_color=themeColors["brandGold"],
        )

    def _onConversionSuccess(self, resultPath: str, outputFolder: str) -> None:
        self.isConverting = False
        fileName = os.path.basename(resultPath)
        self.actionBar.setSuccess(f"Berhasil: {fileName}", outputFolder=outputFolder)

    def _onConversionError(self, errorMessage: str) -> None:
        self.isConverting = False
        self.actionBar.setError(errorMessage)

