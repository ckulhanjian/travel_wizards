# hotel_invoice_processor.spec
# PyInstaller spec for the Hotel Invoice Processor (Travel Wizards)
#
# Bundled assets (must exist in this folder before building):
#   Hotel_Invoice-Template.xlsx
#
# To add more templates later:
#   1. Place the .xlsx file in this folder
#   2. Add ("YourTemplate.xlsx", ".") to the datas list below
#   3. Register it in the TEMPLATES dict in hotel_invoice_processor.py

block_cipher = None

a = Analysis(
    ["hotel_invoice_processor.py"],
    pathex=["."],
    binaries=[],
    datas=[
        ("Hotel_Invoice-Template.xlsx", "."),
        # ("Hotel_Invoice-Template-Intl.xlsx",  "."),   # uncomment when ready
        # ("Hotel_Invoice-Template-Group.xlsx", "."),   # uncomment when ready
    ],
    hiddenimports=[
        "pandas",
        "openpyxl",
        "openpyxl.styles",
        "openpyxl.utils",
        "et_xmlfile",
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
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
    name="HotelInvoiceProcessor",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,          # no terminal window — GUI only
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    # icon="icon.ico",      # uncomment and add icon.ico to use a custom icon
)
