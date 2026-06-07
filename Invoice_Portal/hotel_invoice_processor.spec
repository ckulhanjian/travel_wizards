# hotel_invoice_processor.spec
from PyInstaller.utils.hooks import collect_all

block_cipher = None

fitz_datas, fitz_binaries, fitz_hiddenimports = collect_all("fitz")
rl_datas, rl_binaries, rl_hiddenimports = collect_all("reportlab")

a = Analysis(
    ["hotel_invoice_processor.py", "hotel_invoice_editor.py", "invoice_pdf.py"],
    pathex=["."],
    binaries=fitz_binaries + rl_binaries,
    datas=[
        ("Hotel_Invoice-Template.xlsx", "."),
        ("overlay.pdf",                 "."),
        *fitz_datas,
        *rl_datas,
    ],
    hiddenimports=[
        "pandas", "openpyxl", "openpyxl.styles", "openpyxl.utils",
        "et_xmlfile", "hotel_invoice_editor", "invoice_pdf",
        *fitz_hiddenimports,
        *rl_hiddenimports,
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
    a.scripts, a.binaries, a.zipfiles, a.datas, [],
    name="HotelInvoiceProcessor",
    debug=False, strip=False, upx=True,
    console=False,
)
