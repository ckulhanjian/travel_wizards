#!/usr/bin/env python3
"""
build.py  –  Compile pdf_invoice_renamer.py into a standalone executable.

Requirements:
    pip install pyinstaller PyMuPDF

Usage:
    python build.py          # auto-detects platform
    python build.py --mac    # force macOS (.app)
    python build.py --win    # force Windows (.exe)
"""

import subprocess
import sys
import os
import platform

# ---------------------------------------------------------------------------
# Paths (adjust if your files live elsewhere)
# ---------------------------------------------------------------------------
SCRIPT      = "pdf_invoice_renamer.py"   # main script
OVERLAY     = "overlay.pdf"              # bundled overlay
BACKSIDE    = "backside.pdf"             # bundled back page
APP_NAME    = "InvoiceProcessor"         # output executable name
ICON_WIN    = "icon.ico"                 # optional – Windows icon
ICON_MAC    = "icon.icns"                # optional – macOS icon

# ---------------------------------------------------------------------------
def ensure_pyinstaller():
    try:
        import PyInstaller  # noqa: F401
    except ImportError:
        print("PyInstaller not found – installing...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "pyinstaller"])

def check_assets():
    for f in [SCRIPT, OVERLAY, BACKSIDE]:
        if not os.path.exists(f):
            print(f"ERROR: Required file not found: {f}")
            sys.exit(1)

def build(target: str):
    """
    target: 'win' or 'mac'
    """
    # Separator differs by platform (even when cross-compiling the spec must
    # use the *host* separator because PyInstaller runs on the host).
    sep = ";" if sys.platform.startswith("win") else ":"

    # --add-data adds the PDFs so _asset() can find them at runtime.
    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--onefile",                          # single executable
        "--windowed",                         # no console window
        f"--name={APP_NAME}",
        f"--add-data={OVERLAY}{sep}.",        # copy overlay.pdf next to the exe
        f"--add-data={BACKSIDE}{sep}.",       # copy backside.pdf next to the exe
    ]

    # Optional icon
    if target == "win" and os.path.exists(ICON_WIN):
        cmd.append(f"--icon={ICON_WIN}")
    elif target == "mac" and os.path.exists(ICON_MAC):
        cmd.append(f"--icon={ICON_MAC}")

    if target == "mac":
        cmd.append("--osx-bundle-identifier=com.travelwizards.invoiceprocessor")

    cmd.append(SCRIPT)

    print("\n" + "="*60)
    print(f"Building for: {target.upper()}")
    print("Command:", " ".join(cmd))
    print("="*60 + "\n")

    result = subprocess.run(cmd)

    if result.returncode == 0:
        out_dir = os.path.join("dist")
        print("\n✓ Build succeeded!")
        if target == "win":
            print(f"  Executable: {out_dir}\\{APP_NAME}.exe")
        else:
            print(f"  App bundle: {out_dir}/{APP_NAME}.app  (or {APP_NAME} binary)")
        print("\nNOTE: The overlay.pdf and backside.pdf are embedded inside the")
        print("executable – users do NOT need those files separately.\n")
    else:
        print("\n✗ Build failed. See output above for details.")
        sys.exit(1)


# ---------------------------------------------------------------------------
if __name__ == "__main__":
    args = sys.argv[1:]

    ensure_pyinstaller()
    check_assets()

    if "--win" in args:
        target = "win"
    elif "--mac" in args:
        target = "mac"
    else:
        # Auto-detect
        target = "win" if platform.system() == "Windows" else "mac"

    build(target)