"""
invoice_pdf.py — Travel Wizards
Renders a hotel invoice directly to PDF using ReportLab, then stamps
the overlay.pdf on top with PyMuPDF.  No LibreOffice required.
"""

import os
import fitz  # PyMuPDF

from reportlab.pdfgen import canvas as rl_canvas
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.colors import black


# ---------------------------------------------------------------------------
# Page constants  (all in points, letter = 612 x 792)
# ---------------------------------------------------------------------------
W, H   = letter          # 612 x 792
ML     = 51.4            # left margin  (matches measured x)
MR     = 561.0           # right edge of right column
COL_C  = 369.7           # left edge of label column (C)
COL_D  = 560.0           # right edge of value column (D) — right-aligned to here
FS     = 11              # base font size


# ---------------------------------------------------------------------------
# Font registration  (Times New Roman via system or reportlab built-in)
# ---------------------------------------------------------------------------
def _register_fonts():
    """Register Times New Roman if available, else fall back to Times-Roman."""
    try:
        # Mac / Windows system paths
        import sys
        candidates = []
        if sys.platform == "darwin":
            candidates = [
                "/Library/Fonts/Times New Roman.ttf",
                "/System/Library/Fonts/Times New Roman.ttf",
                "/Library/Fonts/Microsoft/Times New Roman.ttf",
            ]
        elif sys.platform == "win32":
            candidates = [
                r"C:\Windows\Fonts\times.ttf",
                r"C:\Windows\Fonts\timesnewroman.ttf",
            ]
        for path in candidates:
            if os.path.exists(path):
                pdfmetrics.registerFont(TTFont("TNR", path))
                bold_path = path.replace(".ttf", " Bold.ttf").replace(
                    "times.ttf", "timesbd.ttf")
                if os.path.exists(bold_path):
                    pdfmetrics.registerFont(TTFont("TNR-Bold", bold_path))
                else:
                    pdfmetrics.registerFont(TTFont("TNR-Bold", path))
                return "TNR", "TNR-Bold"
    except Exception:
        pass
    return "Times-Roman", "Times-Bold"


FONT_REG, FONT_BOLD = _register_fonts()


# ---------------------------------------------------------------------------
# Drawing helpers
# ---------------------------------------------------------------------------
def _y(pdf_y: float) -> float:
    """Convert PDF coordinate (top-down) to ReportLab (bottom-up)."""
    return H - pdf_y


def _text(c, x, y_pdf, text, bold=False, size=FS, align="left"):
    font = FONT_BOLD if bold else FONT_REG
    c.setFont(font, size)
    c.setFillColor(black)
    rl_y = _y(y_pdf)
    if align == "right":
        c.drawRightString(x, rl_y, str(text))
    elif align == "center":
        c.drawCentredString(x, rl_y, str(text))
    else:
        c.drawString(x, rl_y, str(text))


def _money(value) -> str:
    try:
        v = float(str(value).replace("$", "").replace(",", "")) if value else 0.0
        return f"${v:,.2f}"
    except (ValueError, TypeError):
        return "$0.00"


def _hline(c, x1, x2, y_pdf, width=0.5):
    c.setLineWidth(width)
    c.setStrokeColor(black)
    rl_y = _y(y_pdf)
    c.line(x1, rl_y, x2, rl_y)


# ---------------------------------------------------------------------------
# Main renderer
# ---------------------------------------------------------------------------
def render_invoice_pdf(fields: dict, out_path: str):
    """
    fields keys match the FIELDS cell map in hotel_invoice_editor.py:
      D3  Invoice #     D4  Date        D5  Account     D7  PNR Locator
      A17 Hotel Name    A18 Hotel Addr  A19 Hotel City  A20 Hotel Phone
      D23 Confirm #     A24 Guest(s)    D25 Arrive      D26 Depart
      D27 Rate          D28 Total       D29 Commission  D30 Pd To Date
    Subtotal is computed here as Total - Commission - Pd.To.Date.
    """
    import io
    buf = io.BytesIO()
    c = rl_canvas.Canvas(buf, pagesize=letter)
    c.setTitle("Travel Wizards Invoice")

    f = fields  # shorthand

    # ── Static header (left column) ──────────────────────────────────
    _text(c, ML,    134.6, "TRAVEL WIZARDS, INC.", bold=True)
    _text(c, ML,    148.1, "P. O. BOX 711")
    _text(c, ML,    161.5, "BURLINGAME, CA 94011")
    _text(c, ML,    175.0, "IATAN: 05 90893 2")
    _text(c, ML,    188.9, "VAT/TIN: 94-2713343")

    # ── Static header (right labels) ────────────────────────────────
    _text(c, COL_C, 134.9, "Invoice #:")
    _text(c, COL_C, 148.1, "Date:")
    _text(c, COL_C, 161.5, "Account:")
    _text(c, COL_C, 175.0, "Page #:")
    _text(c, COL_C, 188.5, "PNR Locator:")

    # ── Dynamic header values (right-aligned) ───────────────────────
    _text(c, COL_D, 134.6, f.get("D3", ""),  align="right")
    _text(c, COL_D, 148.1, f.get("D4", ""),  align="right")
    _text(c, COL_D, 161.5, f.get("D5", ""),  align="right")
    _text(c, COL_D, 175.0, "1",               align="right")
    _text(c, COL_D, 188.5, f.get("D7", ""),  align="right")

    # ── Banking block ────────────────────────────────────────────────
    _text(c, ML, 211.5, "JPMORGAN CHASE BANK")
    _text(c, ML, 225.1, "270 PARK AVENUE")
    _text(c, ML, 239.0, "NEW YORK, NY 10017")
    _text(c, ML, 252.8, "ACCT #: 80010264218")
    _text(c, ML, 266.6, "SWIFT CODE: CHASUS33")
    _text(c, ML, 280.4, "ABA: 322271627")

    # ── Hotel block ──────────────────────────────────────────────────
    _text(c, ML, 307.7, "HOTEL:", bold=True)
    _text(c, ML, 321.5, f.get("A17", ""))
    _text(c, ML, 335.3, f.get("A18", ""))
    _text(c, ML, 349.1, f.get("A19", ""))
    _text(c, ML, 362.8, f.get("A20", ""))

    # ── Guest block ──────────────────────────────────────────────────
    _hline(c, 50.4, 548.3, 379.98, width=1.75)

    _text(c, ML,    403.2, "Guest(s):",      bold=True)
    _text(c, COL_C, 403.2, "Confirmation #:", bold=True)
    _text(c, COL_D, 402.8, f.get("D23", ""), align="right")

    # Guest names — one per line starting at y=415.8
    guests = [g.strip() for g in f.get("A24", "").splitlines() if g.strip()]
    for idx, guest in enumerate(guests):
        _text(c, ML, 415.8 + idx * 13.4, guest)

    # ── Dates / financials block ─────────────────────────────────────
    _text(c, COL_C, 501.7, "Arrive: ")
    _text(c, COL_D, 501.4, f.get("D25", ""), align="right")

    _text(c, COL_C, 515.2, "Depart: ")
    _text(c, COL_D, 514.9, f.get("D26", ""), align="right")

    _text(c, COL_C, 528.7, "Rate: ")
    _text(c, COL_D, 528.3, _money(f.get("D27", "")), align="right")

    _text(c, COL_C, 542.2, "Total: ")
    _text(c, COL_D, 541.8, _money(f.get("D28", "")), align="right")

    _text(c, COL_C, 555.7, "Commission:")
    _text(c, COL_D, 555.7, _money(f.get("D29", "")), align="right")

    _text(c, COL_C, 569.2, "Pd. To Date")
    _text(c, COL_D, 568.8, _money(f.get("D30", "")), align="right")

    # Sub-total line
    _hline(c, COL_C, MR, 579.0)

    try:
        total    = float(str(f.get("D28","0")).replace("$","").replace(",","") or 0)
        comm     = float(str(f.get("D29","0")).replace("$","").replace(",","") or 0)
        pd_date  = float(str(f.get("D30","0")).replace("$","").replace(",","") or 0)
        subtotal = total - comm - pd_date
    except (ValueError, TypeError):
        subtotal = 0.0

    # ── Total Due ────────────────────────────────────────────────────
    _text(c, COL_C + 19.2, 595.8, "TOTAL DUE:", bold=True, align="left")
    _text(c, COL_D,        595.8, _money(subtotal), bold=True, align="right")
    _hline(c, COL_C, MR, 600.5, width=1.0)

    c.save()

    # ── Stamp overlay ────────────────────────────────────────────────
    buf.seek(0)
    content = fitz.open("pdf", buf.read())
    overlay = fitz.open(out_path.replace(out_path, _overlay_path_for(out_path)))
    content[0].show_pdf_page(content[0].rect, overlay, 0)
    content.save(out_path)
    content.close()
    overlay.close()


def _overlay_path_for(out_path: str) -> str:
    """Resolved at call time so bundled and dev paths both work."""
    import sys
    base = sys._MEIPASS if getattr(sys, "frozen", False) \
           else os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base, "overlay.pdf")


# ---------------------------------------------------------------------------
# Public entry point used by the editor
# ---------------------------------------------------------------------------
def build_pdf(fields: dict, overlay_path: str, out_pdf: str):
    """
    Build the invoice PDF and stamp the overlay.
    overlay_path is passed explicitly so the editor stays in control.
    """
    import io
    buf = io.BytesIO()
    c = rl_canvas.Canvas(buf, pagesize=letter)
    c.setTitle("Travel Wizards Invoice")

    f = fields

    _text(c, ML,    134.6, "TRAVEL WIZARDS, INC.", bold=True)
    _text(c, ML,    148.1, "P. O. BOX 711")
    _text(c, ML,    161.5, "BURLINGAME, CA 94011")
    _text(c, ML,    175.0, "IATAN: 05 90893 2")
    _text(c, ML,    188.9, "VAT/TIN: 94-2713343")

    _text(c, COL_C, 134.9, "Invoice #:")
    _text(c, COL_C, 148.1, "Date:")
    _text(c, COL_C, 161.5, "Account:")
    _text(c, COL_C, 175.0, "Page #:")
    _text(c, COL_C, 188.5, "PNR Locator:")

    _text(c, COL_D, 134.6, f.get("D3", ""),  align="right")
    _text(c, COL_D, 148.1, f.get("D4", ""),  align="right")
    _text(c, COL_D, 161.5, f.get("D5", ""),  align="right")
    _text(c, COL_D, 175.0, "1",               align="right")
    _text(c, COL_D, 188.5, f.get("D7", ""),  align="right")

    _text(c, ML, 211.5, "JPMORGAN CHASE BANK")
    _text(c, ML, 225.1, "270 PARK AVENUE")
    _text(c, ML, 239.0, "NEW YORK, NY 10017")
    _text(c, ML, 252.8, "ACCT #: 80010264218")
    _text(c, ML, 266.6, "SWIFT CODE: CHASUS33")
    _text(c, ML, 280.4, "ABA: 322271627")

    _text(c, ML, 307.7, "HOTEL:", bold=True)
    _text(c, ML, 321.5, f.get("A17", ""))
    _text(c, ML, 335.3, f.get("A18", ""))
    _text(c, ML, 349.1, f.get("A19", ""))
    _text(c, ML, 362.8, f.get("A20", ""))

    _hline(c, 50.4, 548.3, 379.98, width=1.75)

    _text(c, ML,    403.2, "Guest(s):",       bold=True)
    _text(c, COL_C, 403.2, "Confirmation #:", bold=True)
    _text(c, COL_D, 402.8, f.get("D23", ""),  align="right")

    guests = [g.strip() for g in f.get("A24", "").splitlines() if g.strip()]
    for idx, guest in enumerate(guests):
        _text(c, ML, 415.8 + idx * 13.4, guest)

    _text(c, COL_C, 501.7, "Arrive: ")
    _text(c, COL_D, 501.4, f.get("D25", ""), align="right")

    _text(c, COL_C, 515.2, "Depart: ")
    _text(c, COL_D, 514.9, f.get("D26", ""), align="right")

    _text(c, COL_C, 528.7, "Rate: ")
    _text(c, COL_D, 528.3, _money(f.get("D27", "")), align="right")

    _text(c, COL_C, 542.2, "Total: ")
    _text(c, COL_D, 541.8, _money(f.get("D28", "")), align="right")

    _text(c, COL_C, 555.7, "Commission:")
    _text(c, COL_D, 555.7, _money(f.get("D29", "")), align="right")

    _text(c, COL_C, 569.2, "Pd. To Date")
    _text(c, COL_D, 568.8, _money(f.get("D30", "")), align="right")

    _hline(c, COL_C, MR, 579.0)

    try:
        total    = float(str(f.get("D28","0")).replace("$","").replace(",","") or 0)
        comm     = float(str(f.get("D29","0")).replace("$","").replace(",","") or 0)
        pd_date  = float(str(f.get("D30","0")).replace("$","").replace(",","") or 0)
        subtotal = total - comm - pd_date
    except (ValueError, TypeError):
        subtotal = 0.0

    _text(c, COL_C + 19.2, 595.8, "TOTAL DUE:", bold=True)
    _text(c, COL_D,        595.8, _money(subtotal), bold=True, align="right")
    _hline(c, COL_C, MR, 600.5, width=1.0)

    c.save()

    buf.seek(0)
    content = fitz.open("pdf", buf.read())
    overlay = fitz.open(overlay_path)
    content[0].show_pdf_page(content[0].rect, overlay, 0)
    content.save(out_pdf)
    content.close()
    overlay.close()