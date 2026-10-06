"""Main entrypoint for CheapNFriendly Excel to PDF Converter application."""

import os
import sys
import ctypes
import customtkinter as ctk
import tkinterdnd2

from src.ui.main_window import MainWindow
from src.utils.logger import appLogger


def resolveResourcePath(relativePath: str) -> str:
    """Resolve a bundled resource path, supporting both dev and PyInstaller frozen modes."""
    basePath = getattr(sys, "_MEIPASS", os.path.abspath(os.path.dirname(__file__)))
    return os.path.join(basePath, relativePath)


def applyWindowIcon(root) -> None:
    """Apply the application icon to the window title bar and taskbar."""
    candidates = [
        resolveResourcePath("icon.ico"),
        resolveResourcePath(os.path.join("assets", "icon.ico")),
        os.path.join(os.path.dirname(os.path.abspath(__file__)), "icon.ico"),
    ]
    for iconPath in candidates:
        if os.path.exists(iconPath):
            try:
                root.iconbitmap(iconPath)
                return
            except Exception as iconErr:
                appLogger.warning(f"Gagal memasang ikon window dari {iconPath}: {iconErr}")
    appLogger.warning("File ikon window tidak ditemukan, memakai ikon default.")


def setupWindowsDpi() -> None:
    """Enables Windows Per-Monitor DPI awareness for crisp fonts and graphics."""
    if sys.platform == "win32":
        try:
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            try:
                ctypes.windll.user32.SetProcessDPIAware()
            except Exception:
                pass


def main() -> None:
    setupWindowsDpi()

    # Configure CustomTkinter defaults
    ctk.set_appearance_mode("light")
    ctk.set_default_color_theme("blue")

    appLogger.info("Memulai aplikasi CheapNFriendly Excel to PDF Converter...")

    # Use TkinterDnD.Tk() as root for drag-and-drop support
    root = tkinterdnd2.Tk()
    applyWindowIcon(root)
    app = MainWindow(root)
    root.mainloop()
    appLogger.info("Aplikasi ditutup dengan bersih.")


if __name__ == "__main__":
    main()
