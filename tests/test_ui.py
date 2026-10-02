"""Automated verification test for UI components and MainWindow."""

import os
import sys

# Ensure project root is in sys.path
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

import customtkinter as ctk
from src.config import PaperSizeEnum, PageOrientationEnum, PageScalingEnum, OutputModeEnum
from src.core.excel_inspector import SheetMetadata
from src.ui.main_window import MainWindow
from src.ui.components import FileSelector, SheetList, SettingsPanel, ActionBar



def testUiComponents() -> None:
    print("Testing UI components initialization...")
    ctk.set_appearance_mode("dark")
    root = ctk.CTk()
    root.withdraw()  # Hide root window during headless test

    try:
        # 1. Test FileSelector
        fileSelected = []
        selector = FileSelector(root, onFileSelectedCallback=lambda p: fileSelected.append(p))
        selector.setFile(__file__)
        assert selector.getFilePath() == __file__, "FileSelector getFilePath mismatch"
        assert len(fileSelected) == 1, "FileSelector callback not called"
        print("[PASS] FileSelector verified.")

        # 2. Test SheetList (with dicts and SheetMetadata objects)
        changedCount = [0]
        sheetList = SheetList(root, onSelectionChanged=lambda: changedCount.append(1))
        sampleSheets = [
            {"name": "Summary", "visible": True},
            {"name": "Details", "visible": True},
            {"name": "Audit", "visible": False},
        ]
        sheetList.populateSheets(sampleSheets)
        selected = sheetList.getSelectedSheetNames()
        assert "Summary" in selected and "Details" in selected, "Visible sheets should be selected by default"
        assert "Audit" not in selected, "Hidden sheets should not be selected by default"

        sheetList.selectAll()
        assert len(sheetList.getSelectedSheetNames()) == 3, "Select all should select all 3 sheets"

        sheetList.deselectAll()
        assert len(sheetList.getSelectedSheetNames()) == 0, "Deselect all should select 0 sheets"

        # Test with SheetMetadata dataclass (real engine output)
        metadataSheets = [
            SheetMetadata(name="DataSheet", index=1, isVisible=True, state="visible"),
            SheetMetadata(name="HiddenConfig", index=2, isVisible=False, state="hidden"),
        ]
        sheetList.populateSheets(metadataSheets)
        metaSelected = sheetList.getSelectedSheetNames()
        assert metaSelected == ["DataSheet"], f"Expected ['DataSheet'], got {metaSelected}"
        print("[PASS] SheetList verified (both dict and SheetMetadata supported).")

        # 3. Test SettingsPanel
        panel = SettingsPanel(root)
        settings = panel.getSettings()
        assert settings["paperSize"] == PaperSizeEnum.A4
        assert settings["orientation"] == PageOrientationEnum.PORTRAIT
        assert settings["scaling"] == PageScalingEnum.NORMAL
        assert settings["outputMode"] == OutputModeEnum.SINGLE_MERGED

        panel.paperSizeVar.set(PaperSizeEnum.F4.value)
        panel.orientationVar.set(PageOrientationEnum.LANDSCAPE.value)
        updated = panel.getSettings()
        assert updated["paperSize"] == PaperSizeEnum.F4
        assert updated["orientation"] == PageOrientationEnum.LANDSCAPE
        print("[PASS] SettingsPanel verified.")

        # 4. Test ActionBar
        clicked = []
        bar = ActionBar(root, onConvertCallback=lambda: clicked.append(True))
        bar._handleConvert()
        assert len(clicked) == 1, "ActionBar convert callback not fired"

        bar.setRunning(True, "Processing...")
        bar.setSuccess("Completed!", outputFolder=os.path.dirname(__file__))
        print("[PASS] ActionBar verified.")

    finally:
        root.destroy()

    # 5. Test MainWindow instantiation
    print("Testing MainWindow instantiation...")
    app = MainWindow()
    app.withdraw()  # Don't show window
    try:
        app.update()
        assert app.fileSelector is not None
        assert app.sheetList is not None
        assert app.settingsPanel is not None
        assert app.actionBar is not None
        print("[PASS] MainWindow verified successfully.")
    finally:
        app.update()
        app.destroy()

    print("=== ALL UI TESTS PASSED SUCCESSFULLY ===")


if __name__ == "__main__":
    testUiComponents()
