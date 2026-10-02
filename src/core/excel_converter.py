"""Excel to PDF conversion engine utilizing Microsoft Excel COM automation.

Handles page setup (A4/F4, orientation, scaling), selective sheet exporting,
and seamless PDF merging.
Adheres to camelCase variable naming and modular design.
"""

from dataclasses import dataclass
import os
import shutil
import tempfile
from typing import Callable, List, Optional
from pypdf import PdfWriter

from src.config import (
    PaperSizeEnum,
    PageOrientationEnum,
    PageScalingEnum,
    paperSizeComMapping,
    orientationComMapping,
    scalingComMapping,
    f4FallbackPaperSizes,
)
from src.core.process_guard import ExcelSession
from src.utils.logger import appLogger


@dataclass
class ConversionOptions:
    """Configuration options for the Excel to PDF conversion job."""
    inputExcelPath: str
    outputPdfPath: str
    targetSheetNames: List[str]
    paperSize: PaperSizeEnum = PaperSizeEnum.A4
    orientation: PageOrientationEnum = PageOrientationEnum.PORTRAIT
    scaling: PageScalingEnum = PageScalingEnum.NORMAL
    mergeToSinglePdf: bool = True
    progressCallback: Optional[Callable[[int, str], None]] = None


def reportProgress(
    progressCallback: Optional[Callable[[int, str], None]],
    percent: int,
    message: str,
) -> None:
    """Helper to report progress safely to caller or GUI."""
    appLogger.info(f"Progress ({percent}%): {message}")
    if progressCallback:
        try:
            progressCallback(percent, message)
        except Exception as callbackErr:
            appLogger.warning(f"Error pada progressCallback: {callbackErr}")


def applyPaperSize(worksheetCom, paperSize: PaperSizeEnum) -> None:
    """Applies paper size with fallback support for custom sizes like F4/Folio."""
    if paperSize == PaperSizeEnum.A4:
        try:
            worksheetCom.PageSetup.PaperSize = paperSizeComMapping[PaperSizeEnum.A4]
        except Exception as a4Err:
            appLogger.warning(f"Gagal mengatur ukuran kertas A4: {a4Err}")
    elif paperSize == PaperSizeEnum.F4:
        success = False
        for fallbackId in f4FallbackPaperSizes:
            try:
                worksheetCom.PageSetup.PaperSize = fallbackId
                appLogger.info(f"Ukuran kertas F4 berhasil diatur dengan ID: {fallbackId}")
                success = True
                break
            except Exception:
                continue
        if not success:
            appLogger.warning("Driver printer tidak mendukung ID kertas Folio/Legal bawaan.")


def applyPageSetup(
    worksheetCom,
    paperSize: PaperSizeEnum,
    orientation: PageOrientationEnum,
    scaling: PageScalingEnum,
) -> None:
    """Applies comprehensive print page setup to a worksheet."""
    # 1. Orientation
    worksheetCom.PageSetup.Orientation = orientationComMapping[orientation]

    # 2. Paper Size
    applyPaperSize(worksheetCom, paperSize)

    # 3. Scaling Options
    scalingConfig = scalingComMapping[scaling]
    if scalingConfig["zoom"] is not False:
        worksheetCom.PageSetup.Zoom = scalingConfig["zoom"]
    else:
        # In Excel COM, FitToPagesWide/Tall only takes effect when Zoom is False
        worksheetCom.PageSetup.Zoom = False
        worksheetCom.PageSetup.FitToPagesWide = scalingConfig["fitToPagesWide"]
        worksheetCom.PageSetup.FitToPagesTall = scalingConfig["fitToPagesTall"]


def mergePdfFiles(sourcePdfPaths: List[str], destinationPdfPath: str) -> None:
    """Merges multiple PDF files into a single output PDF using modern pypdf."""
    if not sourcePdfPaths:
        raise ValueError("Daftar file PDF untuk digabungkan kosong.")

    pdfWriter = PdfWriter()
    try:
        for pdfPath in sourcePdfPaths:
            pdfWriter.append(pdfPath)

        parentDirectory = os.path.dirname(os.path.abspath(destinationPdfPath))
        os.makedirs(parentDirectory, exist_ok=True)

        with open(destinationPdfPath, "wb") as outputStream:
            pdfWriter.write(outputStream)
    finally:
        pdfWriter.close()


def convertExcelToPdf(options: ConversionOptions) -> str:
    """Converts selected sheets from an Excel workbook into PDF.

    Returns the path to the generated PDF (or output directory if separated).
    """
    absInputPath = os.path.abspath(options.inputExcelPath)
    absOutputPath = os.path.abspath(options.outputPdfPath)

    if not os.path.exists(absInputPath):
        raise FileNotFoundError(f"File input Excel tidak ditemukan: {absInputPath}")

    if not options.targetSheetNames:
        raise ValueError("Tidak ada sheet yang dipilih untuk dikonversi.")

    reportProgress(options.progressCallback, 5, "Memulai sesi Microsoft Excel COM...")

    tempDirectory = None
    tempPdfFiles: List[str] = []

    try:
        with ExcelSession() as excelApp:
            reportProgress(options.progressCallback, 15, "Membuka file Excel...")

            workbookCom = excelApp.Workbooks.Open(
                absInputPath,
                ReadOnly=True,
                UpdateLinks=0,
                IgnoreReadOnlyRecommended=True,
            )

            try:
                availableSheetNames = [sheet.Name for sheet in workbookCom.Worksheets]
                totalSheetsToConvert = len(options.targetSheetNames)

                for targetName in options.targetSheetNames:
                    if targetName not in availableSheetNames:
                        raise ValueError(f"Sheet '{targetName}' tidak ditemukan di workbook.")

                # If only 1 sheet and single PDF is requested, export directly
                if totalSheetsToConvert == 1 and options.mergeToSinglePdf:
                    targetSheetName = options.targetSheetNames[0]
                    reportProgress(options.progressCallback, 40, f"Mengatur halaman {targetSheetName}...")
                    worksheetCom = workbookCom.Worksheets(targetSheetName)
                    applyPageSetup(worksheetCom, options.paperSize, options.orientation, options.scaling)

                    reportProgress(options.progressCallback, 70, f"Mengekspor {targetSheetName} ke PDF...")
                    parentDir = os.path.dirname(absOutputPath)
                    os.makedirs(parentDir, exist_ok=True)

                    worksheetCom.ExportAsFixedFormat(0, absOutputPath)
                    reportProgress(options.progressCallback, 100, "Konversi selesai.")
                    return absOutputPath

                # Multiple sheets or separate files requested
                tempDirectory = tempfile.mkdtemp(prefix="excel_to_pdf_")
                if options.mergeToSinglePdf:
                    outputDir = os.path.dirname(absOutputPath)
                else:
                    outputDir = absOutputPath if not absOutputPath.lower().endswith(".pdf") else os.path.dirname(absOutputPath)
                os.makedirs(outputDir, exist_ok=True)

                for index, sheetName in enumerate(options.targetSheetNames):
                    percent = 20 + int((index / totalSheetsToConvert) * 60)
                    reportProgress(options.progressCallback, percent, f"Memproses sheet: {sheetName}...")

                    worksheetCom = workbookCom.Worksheets(sheetName)
                    applyPageSetup(worksheetCom, options.paperSize, options.orientation, options.scaling)

                    if options.mergeToSinglePdf:
                        tempPdfPath = os.path.join(tempDirectory, f"sheet_{index}.pdf")
                        worksheetCom.ExportAsFixedFormat(0, tempPdfPath)
                        tempPdfFiles.append(tempPdfPath)
                    else:
                        cleanSheetName = "".join(c for c in sheetName if c.isalnum() or c in (" ", "_", "-")).rstrip()
                        separatePdfPath = os.path.join(outputDir, f"{cleanSheetName}.pdf")
                        worksheetCom.ExportAsFixedFormat(0, separatePdfPath)
                        tempPdfFiles.append(separatePdfPath)

                # If merging is requested, combine using pypdf
                if options.mergeToSinglePdf:
                    reportProgress(options.progressCallback, 85, "Menggabungkan semua sheet ke satu PDF...")
                    mergePdfFiles(tempPdfFiles, absOutputPath)
                    reportProgress(options.progressCallback, 100, "Konversi dan penggabungan selesai.")
                    return absOutputPath
                else:
                    reportProgress(options.progressCallback, 100, "Seluruh sheet berhasil diekspor per file.")
                    return outputDir

            finally:
                if workbookCom is not None:
                    try:
                        workbookCom.Close(SaveChanges=False)
                    except Exception:
                        pass

    finally:
        if tempDirectory and os.path.exists(tempDirectory):
            try:
                shutil.rmtree(tempDirectory, ignore_errors=True)
            except Exception as cleanupErr:
                appLogger.warning(f"Gagal menghapus direktori temp: {cleanupErr}")

