"""Main application window for CheapNFriendly Excel to PDF Converter.

Unified queue interface with per-file configuration and batch conversion.
Drag-and-drop supported across the entire window.
"""

import os
import queue
import threading
import shutil
import zipfile
import tempfile
from typing import Optional, List, Dict, Any, Callable
import tkinter as tk
from tkinterdnd2 import DND_FILES
from tkinter import filedialog
import customtkinter as ctk

from src.config import (
    appTitle,
    appVersion,
    PaperSizeEnum,
    PageOrientationEnum,
    PageScalingEnum,
    supportedExtensions,
)
from src.core.excel_inspector import inspectWorkbook
from src.core.excel_converter import convertExcelToPdf, ConversionOptions
from src.ui.theme import themeColors, themeFonts
from src.ui.components import SheetList, SettingsPanel
from src.ui.file_task import FileTask
from src.ui.queue_manager import QueueManager
from src.utils.logger import appLogger


# Status visual configuration (soft & light palette)
statusConfig: Dict[str, Dict[str, Any]] = {
    "pending": {"label": "Menunggu", "icon": "●", "fg": "#EEF2F8", "text": "#7E8BA4"},
    "processing": {"label": "Memproses", "icon": "●", "fg": "#FEF3C7", "text": "#B45309"},
    "completed": {"label": "Selesai", "icon": "●", "fg": "#D1FAE5", "text": "#15803D"},
    "error": {"label": "Error", "icon": "●", "fg": "#FEE2E2", "text": "#DC2626"},
}


class MainWindow(ctk.CTkFrame):
    """Primary application window with unified queue + configuration layout."""

    def __init__(self, master) -> None:
        super().__init__(master, fg_color=themeColors["windowBg"])
        self.master = master

        # Configure root window (TkinterDnD.Tk)
        master.title(f"{appTitle} v{appVersion} - CheapNFriendly")
        master.geometry("1100x720")
        master.minsize(980, 640)
        master.configure(bg=themeColors["windowBg"])

        # Fill entire window
        self.pack(fill="both", expand=True)

        # State
        self.isDestroyed: bool = False
        self.taskQueue: queue.Queue = queue.Queue()
        self.queueManager: QueueManager = QueueManager()
        self.isProcessingQueue: bool = False
        self.selectedTaskId: Optional[str] = None
        self.fileListItems: Dict[str, Dict[str, Any]] = {}
        self.tempExtractionDirs: List[str] = []
        self._suppressConfigSave: bool = False

        self._buildUi()
        self._registerDragAndDrop()
        self._refreshFileList()
        self._showConfigPlaceholder()
        self._pollTaskQueue()

    # ──────────────────────────────────────────────────────────────────────
    #  Thread-safe main-thread dispatch
    # ──────────────────────────────────────────────────────────────────────

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
                self.master.after(35, self._pollTaskQueue)

    def destroy(self) -> None:
        self.isDestroyed = True
        # Cleanup temporary extraction dirs
        for tmpDir in self.tempExtractionDirs:
            try:
                shutil.rmtree(tmpDir, ignore_errors=True)
            except Exception:
                pass
        self.tempExtractionDirs.clear()
        try:
            super().destroy()
        except Exception:
            pass

    # ──────────────────────────────────────────────────────────────────────
    #  UI Construction
    # ──────────────────────────────────────────────────────────────────────

    def _buildUi(self) -> None:
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(2, weight=1)

        # 1. Top Crimson Brand Stripe
        stripe = ctk.CTkFrame(self, fg_color=themeColors["brandCrimson"], height=3, corner_radius=0)
        stripe.grid(row=0, column=0, sticky="ew")

        # 2. Modern Header
        self._buildHeader()

        # 3. Content Area (2 columns)
        content = ctk.CTkFrame(self, fg_color="transparent")
        content.grid(row=2, column=0, padx=20, pady=(4, 10), sticky="nsew")
        content.grid_columnconfigure(0, weight=4, uniform="cols")
        content.grid_columnconfigure(1, weight=5, uniform="cols")
        content.grid_rowconfigure(0, weight=1)

        # Left: File Queue Panel
        self._buildQueuePanel(content)

        # Right: Configuration Panel
        self._buildConfigPanel(content)

        # 4. Footer
        footer = ctk.CTkFrame(self, fg_color="transparent", height=20)
        footer.grid(row=3, column=0, padx=20, pady=(0, 6), sticky="ew")
        ctk.CTkLabel(
            footer,
            text="CHEAP N FRIENDLY DigitAllPrinting  •  Excel to PDF Native Engine",
            font=themeFonts["bodySmall"],
            text_color=themeColors["textMuted"],
        ).pack(side="left")

        # 5. Full-window drop overlay (hidden by default)
        self._buildDropOverlay()

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

    def _buildDropOverlay(self) -> None:
        """Build the full-window drag-and-drop overlay (hidden by default)."""
        self.dropOverlay = ctk.CTkFrame(
            self.master,
            fg_color=themeColors["brandNavy"],
            corner_radius=0,
            border_width=3,
            border_color=themeColors["brandGold"],
        )
        inner = ctk.CTkFrame(self.dropOverlay, fg_color="transparent")
        inner.place(relx=0.5, rely=0.5, anchor="center")
        ctk.CTkLabel(
            inner,
            text="📂",
            font=(themeFonts["brandTitle"][0], 56),
            text_color=themeColors["brandGold"],
        ).pack(pady=(0, 10))
        ctk.CTkLabel(
            inner,
            text="Lepaskan file di sini",
            font=themeFonts["buttonLarge"],
            text_color="#FFFFFF",
        ).pack()
        ctk.CTkLabel(
            inner,
            text="File Excel, folder, atau arsip ZIP akan ditambahkan ke antrian",
            font=themeFonts["body"],
            text_color="#C7CDE4",
        ).pack(pady=(4, 0))
        self.dropOverlayPlaceholder = None

    def _buildQueuePanel(self, parent) -> None:
        """Left column: file queue list with action buttons."""
        panel = ctk.CTkFrame(
            parent,
            fg_color=themeColors["cardBg"],
            corner_radius=12,
            border_width=1,
            border_color=themeColors["cardBorder"],
        )
        panel.grid(row=0, column=0, padx=(0, 8), sticky="nsew")
        panel.grid_columnconfigure(0, weight=1)
        panel.grid_rowconfigure(2, weight=1)

        # Header
        header = ctk.CTkFrame(panel, fg_color="transparent")
        header.grid(row=0, column=0, padx=16, pady=(14, 8), sticky="ew")
        header.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            header,
            text="📋 Antrian File",
            font=themeFonts["sectionHeader"],
            text_color=themeColors["textPrimary"],
        ).grid(row=0, column=0, sticky="w")

        self.queueCountLabel = ctk.CTkLabel(
            header,
            text="0 file",
            font=themeFonts["bodySmall"],
            text_color=themeColors["brandNavy"],
        )
        self.queueCountLabel.grid(row=0, column=1, sticky="e")

        # Toolbar
        toolbar = ctk.CTkFrame(panel, fg_color="transparent")
        toolbar.grid(row=1, column=0, padx=16, pady=(0, 8), sticky="ew")

        ctk.CTkButton(
            toolbar,
            text="+ Tambah File",
            font=themeFonts["bodySmall"],
            fg_color=themeColors["brandNavy"],
            hover_color=themeColors["brandNavyHover"],
            text_color="#FFFFFF",
            corner_radius=6,
            height=30,
            command=self._browseFiles,
        ).pack(side="left", padx=(0, 6))

        ctk.CTkButton(
            toolbar,
            text="Hapus Terpilih",
            font=themeFonts["bodySmall"],
            fg_color=themeColors["cardSubBg"],
            hover_color=themeColors["cardBorder"],
            text_color=themeColors["textPrimary"],
            corner_radius=6,
            height=30,
            command=self._removeSelectedTask,
        ).pack(side="left", padx=(0, 6))

        ctk.CTkButton(
            toolbar,
            text="Bersihkan",
            font=themeFonts["bodySmall"],
            fg_color=themeColors["cardSubBg"],
            hover_color=themeColors["cardBorder"],
            text_color=themeColors["textMuted"],
            corner_radius=6,
            height=30,
            command=self._clearQueue,
        ).pack(side="left")

        # Scrollable file list
        self.fileListFrame = ctk.CTkScrollableFrame(
            panel,
            fg_color=themeColors["inputBg"],
            corner_radius=8,
            border_width=1,
            border_color=themeColors["inputBorder"],
        )
        self.fileListFrame.grid(row=2, column=0, padx=16, pady=(0, 10), sticky="nsew")
        self.fileListFrame.grid_columnconfigure(0, weight=1)

        self.emptyListLabel = ctk.CTkLabel(
            self.fileListFrame,
            text="Belum ada file.\nDrag & drop ke mana saja di jendela ini,\natau klik Tambah File.",
            font=themeFonts["bodySmall"],
            text_color=themeColors["textMuted"],
            justify="center",
        )
        self.emptyListLabel.grid(row=0, column=0, pady=40, padx=12)

        # Bottom action bar
        actionRow = ctk.CTkFrame(panel, fg_color="transparent")
        actionRow.grid(row=3, column=0, padx=16, pady=(0, 14), sticky="ew")
        actionRow.grid_columnconfigure(0, weight=1)

        self.convertAllBtn = ctk.CTkButton(
            actionRow,
            text="▶  Konversi Semua",
            font=themeFonts["buttonLarge"],
            fg_color=themeColors["brandGold"],
            hover_color=themeColors["brandGoldHover"],
            text_color=themeColors["brandGoldText"],
            corner_radius=8,
            height=42,
            command=self._convertAllPendingFiles,
        )
        self.convertAllBtn.grid(row=0, column=0, sticky="ew")

        self.queueStatusLabel = ctk.CTkLabel(
            actionRow,
            text="",
            font=themeFonts["bodySmall"],
            text_color=themeColors["textMuted"],
        )
        self.queueStatusLabel.grid(row=1, column=0, pady=(6, 0), sticky="w")

    def _buildConfigPanel(self, parent) -> None:
        """Right column: per-file configuration panel."""
        self.configPanel = ctk.CTkFrame(parent, fg_color="transparent")
        self.configPanel.grid(row=0, column=1, padx=(8, 0), sticky="nsew")
        self.configPanel.grid_columnconfigure(0, weight=1)
        self.configPanel.grid_rowconfigure(0, weight=1)

        # --- Empty placeholder (shown when no file selected) ---
        self.configPlaceholder = ctk.CTkFrame(
            self.configPanel,
            fg_color=themeColors["cardBg"],
            corner_radius=12,
            border_width=1,
            border_color=themeColors["cardBorder"],
        )
        self.configPlaceholder.grid(row=0, column=0, sticky="nsew")
        phInner = ctk.CTkFrame(self.configPlaceholder, fg_color="transparent")
        phInner.place(relx=0.5, rely=0.5, anchor="center")
        ctk.CTkLabel(phInner, text="⚙️", font=(themeFonts["brandTitle"][0], 44), text_color=themeColors["textMuted"]).pack(pady=(0, 10))
        ctk.CTkLabel(
            phInner,
            text="Pilih file dari antrian",
            font=themeFonts["sectionHeader"],
            text_color=themeColors["textSecondary"],
        ).pack()
        ctk.CTkLabel(
            phInner,
            text="Klik salah satu file di antrian untuk\nmengatur sheet dan pengaturan cetaknya.",
            font=themeFonts["bodySmall"],
            text_color=themeColors["textMuted"],
            justify="center",
        ).pack(pady=(4, 0))

        # --- Active config view (shown when a file is selected) ---
        self.configView = ctk.CTkFrame(self.configPanel, fg_color="transparent")
        self.configView.grid(row=0, column=0, sticky="nsew")
        self.configView.grid_columnconfigure(0, weight=1)
        self.configView.grid_rowconfigure(1, weight=1)

        # File info header card
        infoCard = ctk.CTkFrame(
            self.configView,
            fg_color=themeColors["cardBg"],
            corner_radius=12,
            border_width=1,
            border_color=themeColors["cardBorder"],
        )
        infoCard.grid(row=0, column=0, pady=(0, 10), sticky="ew")
        infoCard.grid_columnconfigure(0, weight=1)

        self.selectedFileNameLabel = ctk.CTkLabel(
            infoCard,
            text="",
            font=themeFonts["sectionHeader"],
            text_color=themeColors["textPrimary"],
            anchor="w",
        )
        self.selectedFileNameLabel.grid(row=0, column=0, padx=16, pady=(14, 2), sticky="ew")

        self.selectedFilePathLabel = ctk.CTkLabel(
            infoCard,
            text="",
            font=themeFonts["bodySmall"],
            text_color=themeColors["textMuted"],
            anchor="w",
        )
        self.selectedFilePathLabel.grid(row=1, column=0, padx=16, pady=(0, 6), sticky="ew")

        self.selectedFileStatusLabel = ctk.CTkLabel(
            infoCard,
            text="",
            font=themeFonts["badge"],
            text_color=themeColors["textPrimary"],
            corner_radius=4,
            padx=8,
            pady=2,
        )
        self.selectedFileStatusLabel.grid(row=2, column=0, padx=16, pady=(0, 14), sticky="w")

        # Config content container (scrollable so the sheet list always has room)
        configInner = ctk.CTkScrollableFrame(
            self.configView,
            fg_color="transparent",
            scrollbar_button_color=themeColors["cardBorder"],
            scrollbar_button_hover_color=themeColors["brandNavy"],
        )
        configInner.grid(row=1, column=0, sticky="nsew")
        configInner.grid_columnconfigure(0, weight=1)

        self.sheetList = SheetList(configInner, onSelectionChanged=self._onConfigChanged)
        self.sheetList.grid(row=0, column=0, pady=(0, 10), sticky="ew")
        # Guarantee a comfortable minimum height so sheets are always visible
        self.sheetList.configure(height=300, width=400)

        self.settingsPanel = SettingsPanel(configInner)
        self.settingsPanel.grid(row=1, column=0, sticky="ew")
        self._bindSettingsEvents()

    def _bindSettingsEvents(self) -> None:
        """Attach change listeners to settings controls for auto-save."""
        try:
            self.settingsPanel.paperSizeVar.trace_add("write", lambda *_: self._onConfigChanged())
            self.settingsPanel.orientationVar.trace_add("write", lambda *_: self._onConfigChanged())
            self.settingsPanel.outputModeVar.trace_add("write", lambda *_: self._onConfigChanged())
            self.settingsPanel.scalingVar.trace_add("write", lambda *_: self._onConfigChanged())
        except Exception as bindErr:
            appLogger.warning(f"Gagal memasang listener settings: {bindErr}")

    # ──────────────────────────────────────────────────────────────────────
    #  Drag & Drop (entire window)
    # ──────────────────────────────────────────────────────────────────────

    def _registerDragAndDrop(self) -> None:
        """Register the entire window as drop target for files, folders, ZIPs."""
        target = self.master
        try:
            target.drop_target_register(DND_FILES)
            target.dnd_bind("<<DragEnter>>", self._onDragEnter)
            target.dnd_bind("<<DragLeave>>", self._onDragLeave)
            target.dnd_bind("<<Drop>>", self._onDrop)
            appLogger.info("Drag-and-drop registered on entire window")
        except Exception as dndErr:
            appLogger.error(f"Gagal mendaftarkan drag-and-drop: {dndErr}", exc_info=True)

    def _showDropOverlay(self) -> None:
        self.dropOverlay.place(relx=0, rely=0, relwidth=1, relheight=1)
        self.dropOverlay.lift()

    def _hideDropOverlay(self) -> None:
        try:
            self.dropOverlay.place_forget()
        except Exception:
            pass

    def _onDragEnter(self, event) -> None:
        self._showDropOverlay()
        return event.action

    def _onDragLeave(self, event) -> None:
        self._hideDropOverlay()
        return event.action

    def _onDrop(self, event) -> None:
        self._hideDropOverlay()
        appLogger.info(f"Drop event received: {repr(event.data)}")
        try:
            paths = self.master.tk.splitlist(event.data)
        except Exception:
            paths = event.data.split()

        addedCount = 0
        for rawPath in paths:
            path = rawPath.strip("{}").strip('"').strip("'")
            if not path:
                continue
            if os.path.isdir(path):
                addedCount += self._addFolderToQueue(path)
            elif path.lower().endswith(".zip"):
                addedCount += self._addZipToQueue(path)
            elif os.path.isfile(path) and path.lower().endswith(supportedExtensions):
                self._queueManagerAdd(path, "file")
                addedCount += 1
            elif os.path.isfile(path):
                appLogger.warning(f"Format file tidak didukung: {path}")

        self._refreshFileList()
        busy = f"{addedCount} item ditambahkan ke antrian."
        self.queueStatusLabel.configure(
            text=busy if addedCount else "Tidak ada file Excel/folder/ZIP yang dikenali.",
            text_color=themeColors["success"] if addedCount else themeColors["warning"],
        )
        return event.action

    # ──────────────────────────────────────────────────────────────────────
    #  Queue management
    #──────────────────────────────────────────────────────────────────────

    def _browseFiles(self) -> None:
        """Open a native file dialog to add Excel/ZIP files to the queue."""
        selected = filedialog.askopenfilenames(
            parent=self,
            title="Pilih File Excel atau ZIP",
            filetypes=[
                ("Excel & ZIP", "*.xlsx *.xls *.xlsm *.zip"),
                ("Excel Files", "*.xlsx *.xls *.xlsm"),
                ("ZIP Archive", "*.zip"),
                ("Semua File", "*.*"),
            ],
        )
        if not selected:
            return
        addedCount = 0
        for path in selected:
            if path.lower().endswith(".zip"):
                addedCount += self._addZipToQueue(path)
            else:
                self._queueManagerAdd(path, "file")
                addedCount += 1
        self._refreshFileList()
        self.queueStatusLabel.configure(
            text=f"{addedCount} item ditambahkan ke antrian.",
            text_color=themeColors["success"],
        )

    def _queueManagerAdd(self, path: str, taskType: str) -> FileTask:
        """Small wrapper adding a task to the queue manager."""
        return self.queueManager.add_file_task(path, taskType)

    def _addFolderToQueue(self, folderPath: str) -> int:
        """Recursively add all Excel files inside a folder. Returns count added."""
        count = 0
        try:
            for root, _dirs, files in os.walk(folderPath):
                for name in files:
                    if name.lower().endswith(supportedExtensions):
                        fullPath = os.path.join(root, name)
                        self._queueManagerAdd(fullPath, "file")
                        count += 1
        except Exception as scanErr:
            appLogger.error(f"Gagal memindai folder {folderPath}: {scanErr}")
        return count

    def _addZipToQueue(self, zipPath: str) -> int:
        """Extract Excel files from a ZIP into a temp folder and add them."""
        count = 0
        try:
            tmpDir = tempfile.mkdtemp(prefix="cnf_excel_zip_")
            self.tempExtractionDirs.append(tmpDir)
            with zipfile.ZipFile(zipPath, "r") as zf:
                zf.extractall(tmpDir)
            for root, _dirs, files in os.walk(tmpDir):
                for name in files:
                    if name.lower().endswith(supportedExtensions):
                        fullPath = os.path.join(root, name)
                        self._queueManagerAdd(fullPath, "file")
                        count += 1
        except Exception as zipErr:
            appLogger.error(f"Gagal mengekstrak ZIP {zipPath}: {zipErr}")
        return count

    def _removeSelectedTask(self) -> None:
        if not self.selectedTaskId:
            self.queueStatusLabel.configure(text="Pilih file terlebih dahulu.", text_color=themeColors["warning"])
            return
        self.queueManager.remove_task(self.selectedTaskId)
        self.selectedTaskId = None
        self._refreshFileList()
        self._showConfigPlaceholder()

    def _clearQueue(self) -> None:
        self.queueManager.clear_all()
        self.selectedTaskId = None
        self._refreshFileList()
        self._showConfigPlaceholder()
        self.queueStatusLabel.configure(text="Antrian dibersihkan.", text_color=themeColors["textMuted"])

    # ──────────────────────────────────────────────────────────────────────
    #  File list rendering
    # ──────────────────────────────────────────────────────────────────────

    def _refreshFileList(self) -> None:
        """Rebuild the left file list from the queue manager."""
        for item in self.fileListItems.values():
            try:
                item["row"].destroy()
            except Exception:
                pass
        self.fileListItems.clear()

        tasks = self.queueManager.get_all_tasks()

        if not tasks:
            self.emptyListLabel.grid()
            self.queueCountLabel.configure(text="0 file")
            self._updateConvertButtonState()
            return

        self.emptyListLabel.grid_remove()

        for idx, task in enumerate(tasks):
            self._renderFileRow(idx, task)

        pending = self.queueManager.get_pending_count()
        self.queueCountLabel.configure(text=f"{len(tasks)} file  •  {pending} menunggu")
        self._updateConvertButtonState()

    def _renderFileRow(self, idx: int, task: FileTask) -> None:
        """Render a single clickable row for a file task."""
        status = statusConfig.get(task.status, statusConfig["pending"])
        isSelected = task.id == self.selectedTaskId

        row = ctk.CTkFrame(
            self.fileListFrame,
            fg_color=themeColors["cardSubBg"] if isSelected else "transparent",
            corner_radius=8,
            border_width=1 if isSelected else 0,
            border_color=themeColors["brandNavy"] if isSelected else themeColors["inputBorder"],
        )
        row.grid(row=idx, column=0, sticky="ew", padx=4, pady=3)
        row.grid_columnconfigure(1, weight=1)

        # Status dot
        dot = ctk.CTkLabel(
            row,
            text=status["icon"],
            font=(themeFonts["body"][0], 16),
            text_color=status["text"],
            width=18,
        )
        dot.grid(row=0, column=0, rowspan=2, padx=(10, 4), pady=8)

        # File name
        nameLabel = ctk.CTkLabel(
            row,
            text=self._truncateMiddle(os.path.basename(task.path), 34),
            font=themeFonts["body"],
            text_color=themeColors["textPrimary"],
            anchor="w",
        )
        nameLabel.grid(row=0, column=1, sticky="ew", pady=(8, 0))

        # Folder path
        folderLabel = ctk.CTkLabel(
            row,
            text=self._truncateMiddle(os.path.dirname(task.path), 40),
            font=themeFonts["bodySmall"],
            text_color=themeColors["textMuted"],
            anchor="w",
        )
        folderLabel.grid(row=1, column=1, sticky="ew", pady=(0, 8))

        # Status badge
        badge = ctk.CTkLabel(
            row,
            text=status["label"],
            font=themeFonts["badge"],
            text_color=status["text"],
            fg_color=status["fg"],
            corner_radius=4,
            padx=8,
            pady=2,
        )
        badge.grid(row=0, column=2, rowspan=1, padx=10, pady=(8, 0), sticky="e")

        # Per-file config hint (paper + orientation)
        hint = ctk.CTkLabel(
            row,
            text=f"{task.paperSize} • {task.orientation}",
            font=themeFonts["bodySmall"],
            text_color=themeColors["textMuted"],
        )
        hint.grid(row=1, column=2, padx=10, pady=(0, 8), sticky="e")

        # Click bindings (on row + children)
        for widget in (row, dot, nameLabel, folderLabel, badge, hint):
            widget.bind("<Button-1>", lambda _e, tid=task.id: self._selectTask(tid))
            try:
                widget.configure(cursor="hand2")
            except Exception:
                pass

        self.fileListItems[task.id] = {"row": row, "task": task}

    def _truncateMiddle(self, text: str, maxLen: int) -> str:
        if len(text) <= maxLen:
            return text
        keep = (maxLen - 3) // 2
        return text[:keep] + "..." + text[-keep:]

    def _updateConvertButtonState(self) -> None:
        pending = self.queueManager.get_pending_count()
        if self.isProcessingQueue:
            self.convertAllBtn.configure(state="disabled", text="⏳  Memproses...")
        elif pending == 0:
            self.convertAllBtn.configure(state="disabled", text="▶  Konversi Semua")
        else:
            self.convertAllBtn.configure(state="normal", text=f"▶  Konversi Semua ({pending})")

    # ──────────────────────────────────────────────────────────────────────
    #  Selection & per-file configuration
    # ──────────────────────────────────────────────────────────────────────

    def _selectTask(self, taskId: str) -> None:
        """Select a task, auto-saving the previous one's config first."""
        if taskId == self.selectedTaskId:
            return

        self._autoSaveCurrentConfig()
        self.selectedTaskId = taskId
        self._refreshFileList()
        self._loadTaskIntoConfig(taskId)

    def _showConfigPlaceholder(self) -> None:
        self.configView.grid_remove()
        self.configPlaceholder.grid()
        self.selectedTaskId = None

    def _loadTaskIntoConfig(self, taskId: str) -> None:
        """Populate right-side config panel from a task."""
        task = self.queueManager.get_task(taskId)
        if not task:
            self._showConfigPlaceholder()
            return

        self.configPlaceholder.grid_remove()
        self.configView.grid()

        self.selectedFileNameLabel.configure(text=os.path.basename(task.path))
        self.selectedFilePathLabel.configure(text=self._truncateMiddle(task.path, 70))

        status = statusConfig.get(task.status, statusConfig["pending"])
        self.selectedFileStatusLabel.configure(
            text=f"  {status['label']}  ",
            text_color=status["text"],
            fg_color=status["fg"],
        )

        # Settings widgets
        self._suppressConfigSave = True
        try:
            self.settingsPanel.setValues(
                paperSize=task.paperSize,
                orientation=task.orientation,
                scaling=task.scaling,
            )
        finally:
            self._suppressConfigSave = False

        # Load sheet list in background thread
        self.sheetList.populateSheets([])
        threading.Thread(
            target=self._loadSheetsForTask,
            args=(task,),
            daemon=True,
        ).start()

    def _loadSheetsForTask(self, task: FileTask) -> None:
        """Inspect the workbook in a background thread and populate sheet list."""
        try:
            sheets = inspectWorkbook(task.path)
            sheetDicts = []
            for sheet in sheets:
                if isinstance(sheet, dict):
                    name = sheet.get("name", "")
                    visible = sheet.get("isVisible", sheet.get("visible", True))
                else:
                    name = getattr(sheet, "name", str(sheet))
                    visible = getattr(sheet, "isVisible", getattr(sheet, "visible", True))
                sheetDicts.append({"name": name, "isVisible": visible})

            def applySheets():
                # Guard: still the same selected task
                if self.selectedTaskId != task.id or self.isDestroyed:
                    return
                self.sheetList.populateSheets(sheetDicts)
                # Restore previously selected sheets
                if task.selectedSheets:
                    for item in self.sheetList.sheetItems:
                        item["variable"].set(item["name"] in task.selectedSheets)
                    self.sheetList._updateCounter()

            self.runOnMainThread(applySheets)
        except Exception as inspectErr:
            appLogger.warning(f"Gagal membaca sheet untuk {task.path}: {inspectErr}")
            self.runOnMainThread(lambda: self.sheetList.populateSheets([]))

    def _onConfigChanged(self) -> None:
        """Handler invoked whenever any config control changes."""
        if self._suppressConfigSave:
            return
        self._autoSaveCurrentConfig()

    def _autoSaveCurrentConfig(self) -> None:
        """Persist the current config panel values to the selected task."""
        if self._suppressConfigSave:
            return
        if not self.selectedTaskId:
            return
        task = self.queueManager.get_task(self.selectedTaskId)
        if not task:
            return
        try:
            self.queueManager.update_task_config(task.id, self._gatherConfigFromUi(task))
            # Refresh hint text on the queue row without full rebuild flicker
            self._updateRowHint(task.id)
        except Exception as saveErr:
            appLogger.warning(f"Gagal auto-save config: {saveErr}")

    def _gatherConfigFromUi(self, task: FileTask) -> Dict[str, Any]:
        """Read current values from config widgets into a dict."""
        config: Dict[str, Any] = {}
        if self.sheetList.sheetItems:
            selected = self.sheetList.getSelectedSheetNames()
            allNames = [it["name"] for it in self.sheetList.sheetItems]
            # Store empty list when all sheets selected (means "all")
            config["selectedSheets"] = [] if set(selected) == set(allNames) else selected
        try:
            config["paperSize"] = self.settingsPanel.paperSizeVar.get()
            config["orientation"] = self.settingsPanel.orientationVar.get()
            config["scaling"] = self.settingsPanel.scalingVar.get()
        except Exception:
            pass
        return config

    def _updateRowHint(self, taskId: str) -> None:
        """Refresh queue list so rows reflect latest persisted config."""
        item = self.fileListItems.get(taskId)
        if not item:
            return
        try:
            self._refreshFileList()
        except Exception:
            pass

    # ──────────────────────────────────────────────────────────────────────
    #  Batch conversion
    # ──────────────────────────────────────────────────────────────────────

    def _convertAllPendingFiles(self) -> None:
        """Kick off sequential conversion of all pending tasks."""
        if self.isProcessingQueue:
            return
        self._autoSaveCurrentConfig()

        pendingTasks = self.queueManager.get_pending_tasks()
        if not pendingTasks:
            self.queueStatusLabel.configure(
                text="Tidak ada file menunggu untuk dikonversi.",
                text_color=themeColors["warning"],
            )
            return

        self.isProcessingQueue = True
        self._updateConvertButtonState()
        self.queueStatusLabel.configure(
            text=f"Memulai konversi {len(pendingTasks)} file...",
            text_color=themeColors["brandNavy"],
        )
        threading.Thread(target=self._runBatchConversion, args=(pendingTasks,), daemon=True).start()

    def _runBatchConversion(self, tasks: List[FileTask]) -> None:
        """Worker thread: convert each task sequentially."""
        successCount = 0
        errorCount = 0

        for task in tasks:
            if self.isDestroyed:
                break

            self.runOnMainThread(lambda t=task: self._setTaskStatus(t.id, "processing", 0))

            try:
                outputPath = self._resolveOutputPath(task)
                options = self._buildConversionOptions(task, outputPath)
                convertExcelToPdf(options)

                self.runOnMainThread(lambda t=task: self._setTaskStatus(t.id, "completed", 100))
                successCount += 1
            except Exception as convErr:
                appLogger.error(f"Konversi gagal untuk {task.path}: {convErr}", exc_info=True)
                self.runOnMainThread(lambda t=task: self._setTaskStatus(t.id, "error", 0))
                errorCount += 1

        self.runOnMainThread(lambda: self._onBatchFinished(successCount, errorCount))

    def _buildConversionOptions(self, task: FileTask, outputPath: str) -> ConversionOptions:
        """Map a FileTask's config into ConversionOptions."""
        if task.selectedSheets:
            targetSheets = list(task.selectedSheets)
        else:
            targetSheets = self._listAllSheets(task.path)

        paperEnum = PaperSizeEnum.A4
        for candidate in PaperSizeEnum:
            if candidate.value == task.paperSize:
                paperEnum = candidate
                break

        orientationEnum = PageOrientationEnum.PORTRAIT
        for candidate in PageOrientationEnum:
            if candidate.value == task.orientation:
                orientationEnum = candidate
                break

        scalingEnum = PageScalingEnum.NORMAL
        for candidate in PageScalingEnum:
            if candidate.value == task.scaling:
                scalingEnum = candidate
                break

        return ConversionOptions(
            inputExcelPath=task.path,
            outputPdfPath=outputPath,
            targetSheetNames=targetSheets,
            paperSize=paperEnum,
            orientation=orientationEnum,
            scaling=scalingEnum,
            mergeToSinglePdf=True,
            progressCallback=lambda pct, msg, t=task: self.runOnMainThread(
                lambda: self._setTaskStatus(t.id, "processing", pct)
            ),
        )

    def _listAllSheets(self, path: str) -> List[str]:
        """Return every sheet name in the workbook."""
        try:
            sheets = inspectWorkbook(path)
        except Exception as inspectErr:
            appLogger.warning(f"Gagal membaca sheet {path}: {inspectErr}")
            return []
        names = []
        for sheet in sheets:
            names.append(sheet.name if hasattr(sheet, "name") else str(sheet))
        return names

    def _resolveOutputPath(self, task: FileTask) -> str:
        """Resolve output PDF path: source folder with auto-numbering."""
        if task.outputPath:
            return task.outputPath

        sourceDir = os.path.dirname(os.path.abspath(task.path))
        baseName = os.path.splitext(os.path.basename(task.path))[0]
        candidate = os.path.join(sourceDir, f"{baseName}_converted.pdf")

        if not os.path.exists(candidate):
            return candidate

        counter = 1
        while True:
            candidate = os.path.join(sourceDir, f"{baseName}_converted({counter}).pdf")
            if not os.path.exists(candidate):
                return candidate
            counter += 1

    def _setTaskStatus(self, taskId: str, status: str, progress: int) -> None:
        """Update a task's status/progress and refresh its UI."""
        self.queueManager.update_task_status(taskId, status, progress)
        self._refreshFileList()
        if taskId == self.selectedTaskId:
            task = self.queueManager.get_task(taskId)
            if task:
                cfg = statusConfig.get(task.status, statusConfig["pending"])
                self.selectedFileStatusLabel.configure(
                    text=f"  {cfg['label']}  ",
                    text_color=cfg["text"],
                    fg_color=cfg["fg"],
                )

    def _onBatchFinished(self, successCount: int, errorCount: int) -> None:
        self.isProcessingQueue = False
        self._updateConvertButtonState()
        if errorCount == 0:
            self.queueStatusLabel.configure(
                text=f"✓ Selesai. {successCount} file berhasil dikonversi.",
                text_color=themeColors["success"],
            )
        else:
            self.queueStatusLabel.configure(
                text=f"Selesai. {successCount} berhasil, {errorCount} gagal.",
                text_color=themeColors["warning"] if successCount else themeColors["error"],
            )








