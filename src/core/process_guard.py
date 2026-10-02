"""Process guard and COM lifecycle manager for Microsoft Excel.

Ensures that Excel instances are safely launched, manipulated, and terminated
without leaving zombie processes or leaking memory.
Adheres to camelCase variable naming and modular architecture.
"""

import gc
import os
import subprocess
from typing import Optional
import pythoncom
import win32com.client
import win32process

from src.utils.logger import appLogger


def killProcessByPid(targetPid: int) -> bool:
    """Terminates a process by its PID using Windows taskkill command."""
    if not targetPid or targetPid <= 0:
        return False
    try:
        subprocess.run(
            ["taskkill", "/F", "/T", "/PID", str(targetPid)],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            check=False,
        )
        return True
    except Exception as killError:
        appLogger.warning(f"Gagal menghentikan PID {targetPid}: {killError}")
        return False


def isProcessRunning(targetPid: int) -> bool:
    """Checks if a process with targetPid is currently running."""
    if not targetPid or targetPid <= 0:
        return False
    try:
        checkResult = subprocess.run(
            ["tasklist", "/FI", f"PID eq {targetPid}"],
            capture_output=True,
            text=True,
            check=False,
        )
        return str(targetPid) in checkResult.stdout
    except Exception:
        return False


class ExcelSession:
    """Context manager for managing an isolated, headless Microsoft Excel COM instance."""

    def __init__(self) -> None:
        self.excelApp: Optional[win32com.client.CDispatch] = None
        self.excelPid: Optional[int] = None
        self.comInitialized: bool = False

    def __enter__(self) -> win32com.client.CDispatch:
        """Initializes COM apartment and starts an isolated Excel application."""
        try:
            pythoncom.CoInitialize()
            self.comInitialized = True

            # Use DispatchEx to ensure an isolated process instance, not attaching to user's open Excel
            self.excelApp = win32com.client.DispatchEx("Excel.Application")
            self.excelApp.Visible = False
            self.excelApp.DisplayAlerts = False
            self.excelApp.ScreenUpdating = False
            self.excelApp.EnableEvents = False
            self.excelApp.Interactive = False

            # Capture the exact OS PID of this Excel instance
            try:
                windowHandle = self.excelApp.Hwnd
                _, processId = win32process.GetWindowThreadProcessId(windowHandle)
                self.excelPid = processId
                appLogger.info(f"Excel COM instance aktif dengan PID: {self.excelPid}")
            except Exception as pidError:
                appLogger.warning(f"Tidak dapat membaca PID Excel window: {pidError}")
                self.excelPid = None

            return self.excelApp

        except Exception as initError:
            appLogger.error(f"Gagal menginisialisasi sesi Excel COM: {initError}")
            self.cleanup()
            raise

    def __exit__(self, excType, excVal, excTb) -> None:
        """Safely terminates the Excel instance and cleans up COM apartment."""
        self.cleanup()

    def cleanup(self) -> None:
        """Safely shuts down Excel COM instance and prevents zombie processes."""
        if self.excelApp is not None:
            try:
                # Close any remaining open workbooks without saving
                workbooks = getattr(self.excelApp, "Workbooks", None)
                if workbooks is not None:
                    for wbIndex in range(workbooks.Count, 0, -1):
                        try:
                            workbooks(wbIndex).Close(SaveChanges=False)
                        except Exception:
                            pass
                self.excelApp.Quit()
            except Exception as quitError:
                appLogger.warning(f"Peringatan saat excelApp.Quit(): {quitError}")
            finally:
                self.excelApp = None

        # Force garbage collection to release OLE COM pointers
        gc.collect()

        # Check if process is still lingering and forcefully kill if needed
        if self.excelPid and isProcessRunning(self.excelPid):
            appLogger.info(f"Membersihkan sisa proses Excel PID {self.excelPid}...")
            killProcessByPid(self.excelPid)

        # Uninitialize COM for this thread
        if self.comInitialized:
            try:
                pythoncom.CoUninitialize()
            except Exception as comError:
                appLogger.warning(f"Peringatan CoUninitialize: {comError}")
            self.comInitialized = False
