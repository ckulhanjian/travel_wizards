# portal.spec
# Builds a single TravelWizards.exe that includes:
#   - portal.py (launcher)
#   - invoice_processor.py (PDF renamer)
#   - hotel_invoice_processor.py
#   - hotel_invoice_editor.py
#   - invoice_pdf.py
# Bundled assets: overlay.pdf, backside.pdf, Hotel_Invoice-Template.xlsx,
#                 customer-review.png, hotel.png

from PyInstaller.utils.hooks import collect_all, collect_submodules

block_cipher = None

fitz_datas,    fitz_binaries,    fitz_hiddenimports    = collect_all("fitz")
rl_datas,      rl_binaries,      rl_hiddenimports      = collect_all("reportlab")
pil_datas,     pil_binaries,     pil_hiddenimports     = collect_all("PIL")
pandas_hidden  = collect_submodules("pandas")
openpyxl_hidden = collect_submodules("openpyxl")

a = Analysis(
    ["portal.py"],
    pathex=["."],
    binaries=fitz_binaries + rl_binaries + pil_binaries,
    datas=[
        # Assets
        ("overlay.pdf",                  "."),
        ("backside.pdf",                 "."),
        ("Hotel_Invoice-Template.xlsx",  "."),
        ("customer-review.png",          "."),
        ("hotel.png",                    "."),
        # All source modules (portal imports them dynamically)
        ("invoice_processor.py",         "."),
        ("hotel_invoice_processor.py",   "."),
        ("hotel_invoice_editor.py",      "."),
        ("invoice_pdf.py",               "."),
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
        "PIL",
        "PIL.Image",
        "PIL.ImageTk",
        "PIL.ImageDraw",
        *fitz_hiddenimports,
        *rl_hiddenimports,
        *pil_hiddenimports,
        *pandas_hidden,
        *openpyxl_hidden,
    ],
    hookspath=[],
    runtime_hooks=[],
    excludes=[],
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name="TravelWizards",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,          # no terminal window
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,              # add icon="icon.ico" here when ready
)
