"""build_exe.py — Run this to build the Windows executable."""
import PyInstaller.__main__
import sys

PyInstaller.__main__.run([
    "portal.py",
    "--onefile",
    "--windowed",
    "--name", "InvoicePortal",
    "--add-data", "overlay.pdf;.",
    "--add-data", "backside.pdf;.",
    "--add-data", "invoice_processor.py;.",
    "--add-data", "invoice_generator.py;.",
    "--add-data", "invoice_pdf.py;.",
    "--add-data", "airport_lookup.py;.",
    "--add-data", "airport_resolver.py;.",
    "--add-data", "state_parser.py;.",
    "--hidden-import", "invoice_processor",
    "--hidden-import", "invoice_generator",
    "--hidden-import", "invoice_pdf",
    "--hidden-import", "airport_lookup",
    "--hidden-import", "airport_resolver",
    "--hidden-import", "state_parser",
    "--hidden-import", "reportlab",
    "--hidden-import", "reportlab.lib",
    "--hidden-import", "reportlab.platypus",
])