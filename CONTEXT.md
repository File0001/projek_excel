# CONTEXT.md - Excel to PDF Converter Desktop Application

## 1. Project Overview & Vision
Aplikasi desktop berbasis Python (Windows Native) yang dirancang untuk mengonversi dokumen Microsoft Excel (`.xlsx`, `.xls`, `.xlsm`) menjadi dokumen PDF berkualitas tinggi (pixel-perfect) dengan kontrol penuh terhadap pemilihan sheet, ukuran kertas (A4 & F4/Folio), orientasi, dan opsi penskalaan (scaling).

---

## 2. Core User Journey & Workflow
1. **Pilih / Upload File Excel:**
   - Pengguna memilih file Excel via dialog file (`.xlsx`, `.xls`, `.xlsm`).
2. **Inspeksi Workbook:**
   - Sistem membaca struktur workbook dan menampilkan seluruh daftar sheet beserta statusnya.
3. **Konfigurasi Konversi:**
   - **Pemilihan Sheet:** Pilih satu sheet, beberapa sheet spesifik, atau seluruh sheet (*Select All*).
   - **Ukuran Kertas:** A4 (210 x 297 mm) atau F4 / Folio (215 x 330 mm / 8.5 x 13 in).
   - **Orientasi:** Portrait atau Landscape.
   - **Penskalaan (Scaling):**
     - *Normal (100% / No Scaling)*
     - *Fit Sheet on One Page* (Muat 1 halaman)
     - *Fit All Columns on One Page* (Muat semua kolom dalam 1 halaman lebar)
     - *Fit All Rows on One Page* (Muat semua baris dalam 1 halaman tinggi)
   - **Opsi Output:** Gabung semua sheet terpilih menjadi 1 PDF atau pisah per-sheet.
4. **Eksekusi Konversi:**
   - Pengguna menekan tombol "Buat PDF".
   - Sistem mengeksekusi konversi di latar belakang (*worker thread*) agar UI tidak freeze.
   - Progress bar / status indicator memberikan umpan balik langsung.
5. **Output & Notifikasi:**
   - File PDF disimpan di lokasi tujuan.
   - Notifikasi sukses dengan opsi langsung membuka file / folder tujuan.

---

## 3. System Architecture & Tech Stack

### Tech Stack Decisions
- **Bahasa:** Python 3.11+
- **GUI Framework:** `CustomTkinter` (Modern, clean, native look-and-feel, dark/light theme support).
- **Workbook Inspector:** `openpyxl` (Membaca daftar sheet secara instan tanpa overhead peluncuran COM process).
- **PDF Conversion Engine:** `pywin32` (`win32com.client` untuk Microsoft Excel COM Automation).
  - *Justifikasi:* Excel terpasang di sistem (`C:\Program Files\Microsoft Office\Root\Office16\EXCEL.EXE`). COM Automation menjamin rendering 100% identik dengan hasil cetak asli Microsoft Excel, mendukung margin, font custom, conditional formatting, rumus, dan bagan/chart tanpa cacat visual.
- **PDF Merging / Utility:** `pypdf` (Untuk penggabungan modular non-kontigu jika dibutuhkan).

### Component Design (Modular Architecture)
```
projek_excel/
├── CONTEXT.md                     # Source of truth projek & dokumentasi arsitektur
├── ARCHITECTURE.md                # Detail arsitektur teknik & diagram alur
├── requirements.txt               # Daftar dependensi Python
├── run.py                         # Entry point aplikasi
├── src/
│   ├── __init__.py
│   ├── config.py                  # Konstanta aplikasi (paper sizes, scaling enums, default paths)
│   ├── core/
│   │   ├── __init__.py
│   │   ├── excel_inspector.py     # Pembaca metadata file & sheet (cepat, non-blocking)
│   │   ├── excel_converter.py     # Engine konversi COM (PageSetup & ExportAsFixedFormat)
│   │   └── process_guard.py       # Pengaman proses (COM lifecycle, safe release, prevent zombie EXCEL.EXE)
│   ├── ui/
│   │   ├── __init__.py
│   │   ├── main_window.py         # Container utama antarmuka CustomTkinter
│   │   ├── components/            # Komponen modular (FileSelector, SheetList, PrintSettings, StatusBar)
│   │   └── theme.py               # Palet warna, tipografi, dan gaya visual
│   └── utils/
│       ├── __init__.py
│       └── logger.py              # Logging terstruktur untuk audit & troubleshooting
└── tests/
    ├── __init__.py
    └── test_conversion_engine.py  # Uji otomatis konversi dan edge case
```

---

## 4. Technical Specifications & Engine Mapping

### Excel COM PageSetup Matrix
| Parameter UI | Excel COM Property | Value / Mapping |
| :--- | :--- | :--- |
| **Kertas A4** | `PageSetup.PaperSize` | `xlPaperA4` (Value: `9`) |
| **Kertas F4 / Folio** | `PageSetup.PaperSize` | `xlPaperFanfoldLegalGerman` (Value: `41` / 8.5 x 13 in) |
| **Orientasi Portrait** | `PageSetup.Orientation` | `xlPortrait` (Value: `1`) |
| **Orientasi Landscape** | `PageSetup.Orientation` | `xlLandscape` (Value: `2`) |
| **No Scaling** | `PageSetup.Zoom` | `100` (atau `False` jika custom) |
| **Fit Sheet on 1 Page** | `FitToPagesWide`, `FitToPagesTall`, `Zoom` | `Wide = 1`, `Tall = 1`, `Zoom = False` |
| **Fit Columns on 1 Page**| `FitToPagesWide`, `FitToPagesTall`, `Zoom` | `Wide = 1`, `Tall = 0 / False`, `Zoom = False` |
| **Fit Rows on 1 Page**   | `FitToPagesWide`, `FitToPagesTall`, `Zoom` | `Wide = 0 / False`, `Tall = 1`, `Zoom = False` |

---

## 5. Security, Edge Cases & Error Handling

1. **COM Process Hanging & Zombie Prevention:**
   - Setiap pemanggilan COM diisolasi dalam blok `try ... finally`.
   - Mengatur `excel.Visible = False`, `excel.DisplayAlerts = False`, `excel.ScreenUpdating = False`.
   - Wajib memanggil `workbook.Close(SaveChanges=False)`, `excel.Quit()`, `del workbook`, `del excel`, dan `pythoncom.CoUninitialize()`.
   - `process_guard` memverifikasi tidak ada proses `EXCEL.EXE` tertinggal (orphan).
2. **File Locked / In-Use:**
   - Deteksi jika file Excel sedang terbuka di aplikasi lain; berikan peringatan informatif tanpa crash.
3. **Password-Protected / Encrypted Workbooks:**
   - Tangkap exception enkripsi/kata sandi secara spesifik dan tampilkan pesan kesalahan ramah pengguna.
4. **Hidden / Empty Sheets:**
   - Berikan penanda visual pada sheet yang kosong atau tersembunyi (*xlSheetHidden* / *xlSheetVeryHidden*).
5. **Non-Blocking UI (Concurrency):**
   - Operasi I/O dan konversi COM dijalankan pada worker thread terpisah (`threading.Thread`).
   - Pembaruan UI dilakukan secara thread-safe menggunakan event queue atau `widget.after()`.

---

## 6. Current Progress Status
- [x] Diskusi dan persetujuan arsitektur awal (Desktop GUI via CustomTkinter).
- [x] Inisialisasi dokumentasi `CONTEXT.md` dan spesifikasi teknis `ARCHITECTURE.md`.
- [x] Setup dependensi (`requirements.txt`) & verifikasi instalasi paket (`customtkinter`, `openpyxl`, `pywin32`, `pypdf`).
- [x] Pembentukan struktur modular project (`src/config.py`, `src/core/`, `src/ui/`, `src/utils/`, `tests/`).
- [x] Implementasi modul `core`:
  - `src/core/process_guard.py`: Manajemen siklus hidup Excel COM, Apartment threading initialization, isolasi PID, dan pembersihan zombie process tanpa kebocoran memori.
  - `src/core/excel_inspector.py`: Fast introspection workbook dan deteksi sheet (visible & hidden) via `openpyxl` dengan fallback otomatis ke COM untuk format `.xls`.
  - `src/core/excel_converter.py`: Konversi Excel ke PDF dengan PageSetup presisi (A4 = 9, F4/Folio = 14, Orientasi, Scaling: Fit Sheet/Cols/Rows) dan penggabungan via modern `pypdf.PdfWriter`.
  - `tests/test_conversion_engine.py`: Uji otomatis end-to-end terverifikasi 100% lolos.
- [x] Implementasi modul `ui` (Modern Desktop Interface CustomTkinter):
  - `src/ui/theme.py`: Definisi color palette brand yang telah disempurnakan menjadi **Eye-Friendly Soft Light Mode** (Latar `#F0F3F8`, Kartu Putih Bersih `#FFFFFF`, Royal Navy `#1C2456`, Warm Amber Gold `#DF9E1B`, Crimson Red `#D81E5B`, Teks Slate Navy `#161D3B`) yang selaras 100% dengan identitas logo CHEAP N FRIENDLY DigitAllPrinting, ramah di mata, dan bebas silau.
  - `src/ui/components/file_selector.py`: Kartu interaktif pemilihan file Excel dengan format size & validasi ekstensi.
  - `src/ui/components/sheet_list.py`: Daftar checklist scrollable dengan opsi "Pilih Semua", badge "Hidden", dan filter otomatis.
  - `src/ui/components/settings_panel.py`: Segmented controls untuk Kertas (A4 / F4 Folio), Orientasi (Portrait / Landscape), Skala halaman, dan Mode PDF (Single Merged / Pisah Per-Sheet).
  - `src/ui/components/action_bar.py`: Indikator status real-time, progress bar, tombol konversi emas mencolok, serta shortcut langsung buka folder hasil konversi.
  - `src/ui/main_window.py`: Dashboard 2-kolom responsif, ikon geometric emas pada canvas, terintegrasi mulus dengan background worker thread & error handling.
  - `run.py`: Entrypoint root dengan dukungan Windows High-DPI Per-Monitor scaling.
  - `tests/test_ui.py`: Smoke test otomatis antarmuka dan komponen UI terverifikasi 100% lolos.
- [x] Integrasi End-to-End & Background Worker (Non-blocking UI teruji stabil).
- [x] Bug Fix: Perbaikan Masalah Stuck di "Membaca struktur sheet...":
  - **Akar Masalah:** Objek hasil `inspectWorkbook` adalah dataclass `SheetMetadata`, namun `SheetList.populateSheets` memanggil `.get()`, memicu `AttributeError` di event callback. Selain itu, pemanggilan `self.after()` langsung dari worker thread rentan memicu `RuntimeError: main thread is not in main loop`.
  - **Solusi:**
    1. `SheetList.populateSheets`: Dibuat polimorfik, mendukung baik dataclass `SheetMetadata` (`getattr`) maupun dictionary (`.get()`).
    2. Thread-Safety Queue: Mengimplementasikan `queue.Queue` (`runOnMainThread` & `_pollTaskQueue`) di `MainWindow` sehingga seluruh komunikasi worker thread ke UI 100% aman dan dieksekusi di main thread Tkinter.
    3. Error Boundary: Menambahkan `try-except` di `_onInspectSuccess` agar kegagalan tidak membekukan UI dan langsung ditampilkan di status bar.
    4. Seluruh automated test (`test_ui.py`, `test_conversion_engine.py`, dan live end-to-end test) lolos 100%.
- [x] Fitur Dialog Penyimpanan Dinamis (Save As & Folder Picker):
  - **Mode Single Merged PDF:** Menampilkan dialog native Windows `asksaveasfilename` dengan nama default terisi (`<baseName>_converted.pdf`), pengguna bebas mengganti nama file & menentukan folder simpan.
  - **Mode Pisah Per-Sheet:** Menampilkan dialog native Windows `askdirectory` untuk memilih folder tujuan seluruh file PDF sheet.
  - **Pembatalan Aman:** Jika pengguna menekan tombol Cancel pada dialog, konversi dibatalkan secara elegan tanpa error.
- [x] **Unified Single-View Redesign (Batch Queue + Per-File Config):**
  - **Status:** **DONE** - 2026-10-02
  - **Motivasi:** Menghapus dua-tab (Single / Batch) yang membingungkan dan menggantinya dengan satu tampilan terpadu yang lebih profesional dan ringkas.
  - **Fitur:**
    - **Layout 1 tampilan, 2 kolom:** Kiri = *Antrian File* (daftar file yang bisa diklik), Kanan = *Panel Konfigurasi* per file.
    - **Per-file config:** Setiap file di antrian punya konfigurasinya sendiri (sheet terpilih, ukuran kertas, orientasi, scaling). Tidak lagi memakai setting global.
    - **Auto-save on switch:** Saat berpindah file, konfigurasi file sebelumnya otomatis tersimpan ke disk (`queue_data.json`).
    - **Whole-window drag & drop:** Drop file/folder/ZIP bisa di seluruh area jendela (bukan drop zone terpisah), dengan overlay feedback saat drag-enter.
    - **Sequential batch convert:** Tombol "Konversi Semua" memproses semua task `pending` satu per satu secara berurutan di worker thread.
    - **Output ke folder sumber + auto-numbering:** Hasil disimpan di folder asal file (`report_converted.pdf` → `report_converted(1).pdf` → `report_converted(2).pdf`, dst).
    - **Status warna soft/light:** pending `#EEF2F8`/`#7E8BA4`, processing `#FEF3C7`/`#B45309`, completed `#D1FAE5`/`#15803D`, error `#FEE2E2`/`#DC2626`.
    - **Styling brand:** Mengikuti `themeColors`/`themeFonts` (Navy/Gold/Crimson), gap konsisten 12–16px, tinggi tombol 32–42px, radius sudut 6–12px.
  - **Perubahan Model Data (`FileTask`):**
    - Field baru: `selectedSheets: List[str]`, `paperSize: str`, `orientation: str`, `scaling: str`, `outputPath: Optional[str]`.
  - **Perubahan `QueueManager`:**
    - `update_task_config(task_id, config)`: menyimpan config per-file.
    - Serialisasi/deserialisasi `_save_to_disk` & `_load_from_disk` diperluas untuk field konfigurasi baru.
  - **Perubahan `SettingsPanel`:**
    - Method baru `setValues(paperSize, orientation, scaling, outputMode)` untuk memuat nilai saat berpindah file.
  - **Files baru/update:**
    - `src/ui/file_task.py`: Field konfigurasi per-file ditambahkan.
    - `src/ui/queue_manager.py`: `update_task_config()` + serialization diperluas.
    - `src/ui/main_window.py`: **Ditulis ulang total** — hapus tab & drop zone, daftar file + panel config 2 kolom, overlay drag full-window, auto-save, batch convert, auto-numbering.
    - `src/ui/components/settings_panel.py`: Tambah `setValues()`.
    - `tests/test_ui.py`: Disesuaikan dengan API `MainWindow(master)` yang baru (frame-based).
  - **Testing:** `py_compile` bersih pada semua file; `tests/test_ui.py` **ALL UI TESTS PASSED**; verifikasi live: instantiasi, seleksi file, load sheet, auto-save config, dan auto-numbering output path semuanya OK.

- [ ] Panduan Pengujian Manual & Packaging Executable (Opsional/Next Step).


