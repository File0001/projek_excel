# Build & Installer Guide

Panduan membuat executable dan installer untuk **CheapNFriendly Excel to PDF Converter**.

## Prasyarat

| Tool | Versi | Keterangan |
|------|-------|------------|
| Python | 3.11+ | Runtime aplikasi |
| PyInstaller | 6.x | Membuat EXE (`pip install pyinstaller`) |
| Pillow | 12.x | Membuat `icon.ico` dari PNG (`pip install pillow`) |
| Inno Setup 6 | 6.x | Membuat installer ([unduh](https://jrsoftware.org/isdl.php)) |

Install semua dependensi Python:

```bat
pip install -r requirements.txt
pip install pyinstaller pillow
```

## Cara Cepat (otomatis)

Jalankan skrip build dari root proyek:

```bat
build-installer.bat
```

Skrip ini akan:
1. Membuat `icon.ico` dari `icon.png` (jika belum ada).
2. Membangun `dist\ExcelToPdfConverter.exe` (satu file, ikon tertanam).
3. Membungkusnya menjadi installer `installer\CheapNFriendly-ExcelToPdf-Setup-1.0.0.exe`.

## Cara Manual

### 1. Membuat ikon

```bat
python -c "from PIL import Image; Image.open('icon.png').convert('RGBA').save('icon.ico', format='ICO', sizes=[(s,s) for s in (16,24,32,48,64,128,256)])"
```

### 2. Membangun EXE (PyInstaller)

```bat
python -m PyInstaller --noconfirm --clean ExcelToPdfConverter.spec
```

Hasil: `dist\ExcelToPdfConverter.exe` (sekitar 20 MB, tanpa konsol, ikon tertanam).

### 3. Membangun installer (Inno Setup)

```bat
"C:\Program Files (x86)\Inno Setup 6\ISCC.exe" installer\ExcelToPdfInstaller.iss
```

Hasil: `installer\CheapNFriendly-ExcelToPdf-Setup-1.0.0.exe`.

## Detail Teknis

- **Ikon aplikasi**: `run.py` memasang ikon window lewat `applyWindowIcon()` yang mencari
  `icon.ico` (mendukung mode dev maupun PyInstaller via `sys._MEIPASS`).
- **Data queue**: saat diinstal, `queue_data.json` disimpan di
  `%LOCALAPPDATA%\CheapNFriendly\queue_data.json` (folder aplikasi read-only).
  Saat dev, tetap di `src/ui/queue_data.json`.
- **Bundling tkdnd**: binari drag-and-drop `tkinterdnd2` dan tema `customtkinter`
  dikumpulkan otomatis melalui `collect_data_files` di dalam `.spec`.

## Struktur File Build

```
ExcelToPdfConverter.spec          # Konfigurasi PyInstaller
build-installer.bat               # Skrip build otomatis
icon.png                          # Ikon sumber (1254x1254)
icon.ico                          # Ikon multi-resolusi (dibuat otomatis)
installer/
  ExcelToPdfInstaller.iss         # Skrip Inno Setup
  CheapNFriendly-...-Setup.exe    # Output installer (tidak di-commit)
dist/
  ExcelToPdfConverter.exe         # Output EXE (tidak di-commit)
```

> **Catatan**: `build/`, `dist/`, dan `installer/*.exe` tidak di-commit ke Git.
> Sumber build (`*.spec`, `installer/*.iss`, `build-installer.bat`, `icon.ico`) di-commit.
