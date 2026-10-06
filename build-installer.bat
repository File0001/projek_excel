@echo off
REM =====================================================================
REM  Build CheapNFriendly Excel to PDF Converter installer
REM  Steps:
REM    1. Generate icon.ico from icon.png (if needed)
REM    2. Create the portable EXE with PyInstaller
REM    3. Wrap it in a single-file installer with Inno Setup
REM =====================================================================
setlocal enabledelayedexpansion

cd /d "%~dp0"

echo.
echo [1/3] Menyiapkan ikon...
if not exist icon.ico (
    python _make_icon.py
    if errorlevel 1 (
        echo GAGAL: tidak bisa membuat icon.ico. Pastikan Pillow terpasang ^(pip install pillow^).
        exit /b 1
    )
) else (
    echo icon.ico sudah ada, dilewati.
)

echo.
echo [2/3] Membangun EXE dengan PyInstaller...
python -m PyInstaller --noconfirm --clean ExcelToPdfConverter.spec
if errorlevel 1 (
    echo GAGAL: PyInstaller gagal membangun EXE.
    exit /b 1
)

echo.
echo [3/3] Membangun installer dengan Inno Setup...
set "ISCC=%ProgramFiles(x86)%\Inno Setup 6\ISCC.exe"
if not exist "!ISCC!" set "ISCC=%ProgramFiles%\Inno Setup 6\ISCC.exe"
if not exist "!ISCC!" (
    echo GAGAL: ISCC.exe ^(Inno Setup 6^) tidak ditemukan.
    echo Silakan pasang Inno Setup 6 dari https://jrsoftware.org/isdl.php
    exit /b 1
)
"!ISCC!" "installer\ExcelToPdfInstaller.iss"
if errorlevel 1 (
    echo GAGAL: Inno Setup gagal membangun installer.
    exit /b 1
)

echo.
echo =====================================================================
echo  SELESAI!
echo  Executable : dist\ExcelToPdfConverter.exe
echo  Installer  : installer\CheapNFriendly-ExcelToPdf-Setup-1.0.0.exe
echo =====================================================================
endlocal
