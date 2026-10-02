"""Structured logger utility for Excel to PDF converter."""

import logging
import sys


def setupLogger(loggerName: str = "ExcelToPdfApp", logLevel: int = logging.INFO) -> logging.Logger:
    """Creates or retrieves a configured logger instance with formatted console output."""
    appLogger = logging.getLogger(loggerName)
    if not appLogger.handlers:
        appLogger.setLevel(logLevel)
        logHandler = logging.StreamHandler(sys.stdout)
        logFormatter = logging.Formatter(
            fmt="[%(asctime)s] [%(levelname)s] [%(name)s]: %(message)s",
            datefmt="%Y-%m-%d %H:%M:%S",
        )
        logHandler.setFormatter(logFormatter)
        appLogger.addHandler(logHandler)
    return appLogger


# Global app logger instance
appLogger = setupLogger()
