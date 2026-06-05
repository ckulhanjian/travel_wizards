"""build_exe.py — Run this to build the Windows executable."""
import PyInstaller.__main__
import sys

PyInstaller.__main__.run([
    "portal.py",
    "--onefile",
    "--windowed",
    "--name", "InvoicePortal",
    "--add-data", "overlay.pdf" + (";" if sys.platform == "win32" else ":") + ".",
    "--add-data", "backside.pdf" + (";" if sys.platform == "win32" else ":") + ".",
    "--add-data", "invoice_processor.py" + (";" if sys.platform == "win32" else ":") + ".",
    "--add-data", "invoice_parser.py" + (";" if sys.platform == "win32" else ":") + ".",
    "--add-data", "itin_parser.py" + (";" if sys.platform == "win32" else ":") + ".",
    "--add-data", "invoice_generator.py" + (";" if sys.platform == "win32" else ":") + ".",
    "--add-data", "airport_lookup.py" + (";" if sys.platform == "win32" else ":") + ".",
    "--add-data", "airport_resolver.py" + (";" if sys.platform == "win32" else ":") + ".",
    "--hidden-import", "invoice_processor",
    "--hidden-import", "invoice_parser",
    "--hidden-import", "itin_parser",
    "--hidden-import", "invoice_generator",
    "--hidden-import", "airport_lookup",
    "--hidden-import", "airport_resolver",
    "--hidden-import", "reportlab",
    "--hidden-import", "reportlab.lib",
    "--hidden-import", "reportlab.platypus",
])