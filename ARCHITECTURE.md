# ARCHITECTURE.md - Excel to PDF Converter Architecture Specification

## 1. High-Level Architecture Flowchart

```
+-------------------------------------------------------------+
|                      User Interface                         |
|             (CustomTkinter Native Desktop App)              |
+------------------------------+------------------------------+
                               |
               1. Select File  |  2. Load Sheets (Metadata)
                               v
+-------------------------------------------------------------+
|                 Core: Excel Inspector                       |
|         (openpyxl - Fast Read-Only Introspection)           |
+------------------------------+------------------------------+
                               |
                               v
                     Populate Sheet List
                     (Multi-select Checkboxes)
                               |
               3. Configure Settings (A4/F4, Orient, Scale)
               4. Click "Buat PDF"
                               v
+-------------------------------------------------------------+
|               Core: Background Worker Thread                |
|       (Non-blocking UI, Thread-safe Progress Updates)       |
+------------------------------+------------------------------+
                               |
                               v
+-------------------------------------------------------------+
|                 Core: Excel Converter COM                   |
|              (win32com.client - Excel Engine)               |
|                                                             |
|  - ProcessGuard: Spawn Excel Headless                       |
|  - PageSetup: PaperSize (A4 / F4), Orientation, Zoom, Fit   |
|  - Export: ExportAsFixedFormat(Type=0 [PDF])                |
|  - Safe Teardown: Close, Quit, Cleanup Orphan EXCEL.EXE     |
+------------------------------+------------------------------+
                               |
                               v
+-------------------------------------------------------------+
|                        Output PDF                           |
|        (Saved to Destination + Open Notification)           |
+-------------------------------------------------------------+
```

## 2. Module Responsibilities

1. **`src/config.py`**:
   - Berisi konstanta enum untuk Ukuran Kertas (A4, F4/Folio), Orientasi (Portrait, Landscape), Scaling Options, default margins, dan palet warna UI.
2. **`src/core/excel_inspector.py`**:
   - Bertanggung jawab membaca metadata workbook (`.xlsx`, `.xls`, `.xlsm`) secara kilat tanpa memicu startup COM Excel yang berat.
   - Mengembalikan list: `[{"name": "Sheet1", "visible": True}, ...]`.
3. **`src/core/excel_converter.py`**:
   - Membuka workbook via COM automation.
   - Mengatur parameter `PageSetup` per sheet terpilih sesuai opsi konfigurasi pengguna.
   - Melakukan ekspor selektif ke file PDF menggunakan `ExportAsFixedFormat`.
4. **`src/core/process_guard.py`**:
   - Mengelola siklus hidup COM Excel. Menjamin instance Excel ditutup bersih dalam skenario normal maupun crash.
5. **`src/ui/main_window.py` & components**:
   - Komponen visual: File Selector bar, Sheet list viewer dengan "Pilih Semua", Settings card (Kertas, Orientasi, Skala), Tombol Eksekusi, dan Status/Progress bar.

## 3. Reliability & Maintenance Principles
- **KISS & YAGNI:** Tidak menggunakan server web atau container jika kebutuhan utama adalah desktop utility.
- **Fail-Safe Lifecycle:** Menghindari memory leak dan zombie process COM dengan context manager (`with ExcelSession() as excel:`).

