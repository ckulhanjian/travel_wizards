# portal.spec  —  builds TravelWizards.exe (single entry point)
#
# Required files in this folder before building:
#   overlay.pdf, backside.pdf, Hotel_Invoice-Template.xlsx

import os
from PyInstaller.utils.hooks import collect_all, collect_submodules

block_cipher = None

# Verify required assets exist before wasting build time
required = ["overlay.pdf", "backside.pdf", "Hotel_Invoice-Template.xlsx"]
missing  = [f for f in required if not os.path.exists(f)]
if missing:
    raise SystemExit(f"\n\nBUILD ERROR — missing required asset(s):\n" +
                     "\n".join(f"  {f}" for f in missing) +
                     "\n\nAdd these files to hotel_invoice_processor/ then retry.\n")

fitz_datas,  fitz_binaries,  fitz_hiddenimports  = collect_all("fitz")
rl_datas,    rl_binaries,    rl_hiddenimports    = collect_all("reportlab")
pil_datas,   pil_binaries,   pil_hiddenimports   = collect_all("PIL")

a = Analysis(
    ["portal.py"],
    pathex=["."],
    binaries=fitz_binaries + rl_binaries + pil_binaries,
    datas=[
        ("overlay.pdf",                 "."),
        ("backside.pdf",                "."),
        ("Hotel_Invoice-Template.xlsx", "."),
        ("customer-review.png",         "."),
        ("hotel.png",                   "."),
        ("invoice_processor.py",        "."),
        ("hotel_invoice_processor.py",  "."),
        ("hotel_invoice_editor.py",     "."),
        ("invoice_pdf.py",              "."),
        *fitz_datas,
        *rl_datas,
        *pil_datas,
    ],
    hiddenimports=[
        "invoice_processor",
        "hotel_invoice_processor",
        "hotel_invoice_editor",
        "invoice_pdf",
        "pandas",
        "openpyxl",
        "openpyxl.styles",
        "openpyxl.utils",
        "et_xmlfile",
        "PIL", "PIL.Image", "PIL.ImageTk", "PIL.ImageDraw",
        *fitz_hiddenimports,
        *rl_hiddenimports,
        *pil_hiddenimports,
    ],
    hookspath=[],
    runtime_hooks=[],
    excludes=["pandas.tests"],   # skip test suite — shaves ~50MB
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts, a.binaries, a.zipfiles, a.datas, [],
    name="TravelWizards",
    debug=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    console=False,
    # icon="icon.ico",
)
