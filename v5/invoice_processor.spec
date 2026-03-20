# invoice_processor.spec
# PyInstaller spec for the PDF Invoice Processor (Travel Wizards)
#
# Bundled assets (must exist in this folder before building):
#   overlay.pdf
#   backside.pdf

from PyInstaller.utils.hooks import collect_all

block_cipher = None

# Collect all PyMuPDF / fitz internals automatically
fitz_datas, fitz_binaries, fitz_hiddenimports = collect_all("fitz")

a = Analysis(
    ["invoice_processor.py"],
    pathex=["."],
    binaries=fitz_binaries,
    datas=[
        ("overlay.pdf",   "."),
        ("backside.pdf",  "."),
        *fitz_datas,
    ],
    hiddenimports=[
        *fitz_hiddenimports,
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
    name="InvoiceProcessor",
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
