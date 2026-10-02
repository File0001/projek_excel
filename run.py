"""Main entrypoint for CheapNFriendly Excel to PDF Converter application."""

import sys
import ctypes
import customtkinter as ctk

from src.ui.main_window import MainWindow
from src.utils.logger import appLogger


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
    app = MainWindow()
    app.mainloop()
    appLogger.info("Aplikasi ditutup dengan bersih.")


if __name__ == "__main__":
    main()
