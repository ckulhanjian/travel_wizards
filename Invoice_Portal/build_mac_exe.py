"""
build_exe.py — Run this to build the app.

Works unchanged on both platforms:
  - Windows (e.g. your GitHub Actions runner): produces InvoicePortal.exe
  - macOS (local builds): produces InvoicePortal.app

The only thing that differs between platforms is the separator PyInstaller
expects in --add-data ("SRC;DEST" on Windows, "SRC:DEST" on Mac/Linux) —
handled automatically below via os.pathsep, so this one script is the only
one you need to maintain.
"""
import os
import sys
import PyInstaller.__main__

SEP = os.pathsep  # ';' on Windows, ':' on macOS/Linux

# Every data file the app needs bundled alongside the entry point.
DATA_FILES = [
    "overlay.pdf",
    "backside.pdf",
    "invoice_processor.py",
    "invoice_generator.py",
    "invoice_pdf.py",
    "airport_lookup.py",
    "airport_resolver.py",
    "airport_manager.py",
    "state_parser.py",
    "updater.py",
    "compare_view.py",
    "hotel_invoice_editor.py",
    "hotel_invoice_processor.py",
]

# Only bundle files that actually exist here — keeps this script safe to
# run even if one of these hasn't been created yet on a given checkout,
# instead of failing the whole build over a single missing optional file.
add_data_args = []
for fname in DATA_FILES:
    if os.path.exists(fname):
        add_data_args += ["--add-data", f"{fname}{SEP}."]
    else:
        print(f"[build_exe] Skipping {fname} — not found in this folder.")

HIDDEN_IMPORTS = [
    "invoice_processor",
    "invoice_generator",
    "invoice_pdf",
    "airport_lookup",
    "airport_resolver",
    "airport_manager",
    "state_parser",
    "updater",
    "compare_view",
    "hotel_invoice_editor",
    "hotel_invoice_processor",
    "reportlab",
    "reportlab.lib",
    "reportlab.platypus",
]
hidden_import_args = []
for mod in HIDDEN_IMPORTS:
    hidden_import_args += ["--hidden-import", mod]

PyInstaller.__main__.run([
    "portal.py",
    "--onefile",
    "--windowed",
    "--name", "InvoicePortal",
    *add_data_args,
    *hidden_import_args,
])