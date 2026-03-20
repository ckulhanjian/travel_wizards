#!/usr/bin/env python3
"""
PDF Invoice Renamer - Travel Wizards

Two folder layouts under globalware/:
  itin/    -> ITIN format  (has SALES PERSON, full overlay on every page)
  tipitin/ -> TIPITIN format (no SALES PERSON, footer-only overlay on pages 2+)

Overlay and back page are bundled with the application.
"""

import os
import sys
import shutil
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import threading
import re

try:
    import fitz  # PyMuPDF
except ImportError as e:
    root = tk.Tk()
    root.withdraw()
    messagebox.showerror("Missing Dependencies",
                         f"Missing required library: {e}\n\n"
                         "Please install:\n"
                         "pip install PyMuPDF")
    sys.exit(1)


# ---------------------------------------------------------------------------
# Resolve bundled asset paths (works both as .py and as a PyInstaller bundle)
# ---------------------------------------------------------------------------
def _asset(filename):
    if getattr(sys, "frozen", False):
        base = sys._MEIPASS
    else:
        base = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base, filename)


OVERLAY_PATH  = _asset("overlay.pdf")
BACKSIDE_PATH = _asset("backside.pdf")

FOLDER_ITIN    = "itin"
FOLDER_TIPITIN = "tipitin"


# ---------------------------------------------------------------------------
# Detect format from file text content
# ---------------------------------------------------------------------------
def detect_format(text: str):
    """
    Detect invoice format from the text of page 1.
    itin    -> has 'SALES PERSON:' line
    tipitin -> has 'ITIN NO:' but no 'SALES PERSON:'
    """
    if re.search(r'SALES PERSON:', text):
        return FOLDER_ITIN
    if re.search(r'ITIN NO:', text):
        return FOLDER_TIPITIN
    return None


# ---------------------------------------------------------------------------
# Field extraction
# ---------------------------------------------------------------------------
def extract_fields(text, fmt: str):
    """
    Extract (agent_initials, invoice_no, last_name) based on known format.

    itin:
        SALES PERSON: 89SSV   ITIN/INVOICE NO.   328016
        FOR: POOLE/EDWARD GRAY

    tipitin:
        ITIN NO:   360085
        FAHRNEY/DONALD WILLIAM  (first NAME/FIRSTNAME line)
        No SALES PERSON line.
    """
    if fmt == FOLDER_ITIN:
        inv   = re.search(r'ITIN/INVOICE NO\.\s+(\d+)', text)
        agent = re.search(r'SALES PERSON:\s*(\S+)', text)
        last  = re.search(r'FOR:\s+([A-Z]+)/', text)
        invoice_no     = inv.group(1)        if inv   else None
        agent_initials = agent.group(1)[-2:] if agent else None
        last_name      = last.group(1)       if last  else None
    else:  # tipitin
        inv   = re.search(r'ITIN NO:\s*(\d+)', text)
        last  = re.search(r'^\s{0,10}([A-Z]{3,})/', text, re.MULTILINE)
        invoice_no     = inv.group(1)  if inv  else None
        agent_initials = None
        last_name      = last.group(1) if last else None

    return agent_initials, invoice_no, last_name


def build_filename(agent_initials, invoice_no, last_name):
    if agent_initials:
        return f"{agent_initials}_{invoice_no}_{last_name}.pdf"
    return f"{invoice_no}_{last_name}.pdf"


# ---------------------------------------------------------------------------
# GUI
# ---------------------------------------------------------------------------
class PDFRenamerGUI:
    # Colour palette — black and white
    CLR_BG        = "#ffffff"
    CLR_PANEL     = "#f5f5f5"
    CLR_ACCENT    = "#000000"
    CLR_ACCENT2   = "#333333"
    CLR_TEXT      = "#000000"
    CLR_MUTED     = "#555555"
    CLR_ENTRY_BG  = "#f5f5f5"
    CLR_BORDER    = "#cccccc"
    CLR_BTN_BG    = "#000000"
    CLR_BTN_FG    = "#000000"
    CLR_LOG_BG    = "#f9f9f9"
    CLR_LOG_FG    = "#111111"

    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Travel Wizards — Invoice Processor")
        self.root.geometry("860x320")
        self.root.resizable(True, True)
        self.root.configure(bg=self.CLR_BG)
        # Force Windows to respect button colours instead of using the system theme
        self.root.option_add("*Button.Background", "#e0e0e0")
        self.root.option_add("*Button.Foreground", "#000000")
        self.root.option_add("*Button.activeBackground", "#cccccc")
        self.root.option_add("*Button.activeForeground", "#000000")
        self.root.option_add("*Button.relief", "flat")

        self.source_folder = tk.StringVar()
        self.detected_fmt  = tk.StringVar(value="—")
        self._draw_logo()
        self.setup_ui()

    def _draw_logo(self):
        """Draw the Travel Wizards wordmark using a Canvas — no image file needed."""
        c = tk.Canvas(self.root, height=74, bg=self.CLR_BG,
                      highlightthickness=0, bd=0)
        c.pack(fill="x", padx=0, pady=0)

        # Horizontal rule at bottom of header
        c.create_line(0, 73, 2000, 73, fill=self.CLR_BORDER, width=1)

        # Wordmark — spaced caps, gold
        c.create_text(430, 26, text="TRAVEL  WIZARDS",
                      font=("Georgia", 26, "bold"),
                      fill=self.CLR_ACCENT, anchor="center")

        # Tagline
        c.create_text(430, 52, text="M A G I C A L   J O U R N E Y S",
                      font=("Arial", 10),
                      fill="#cc0000", anchor="center")

    def setup_ui(self):
        # ── Folder selection row ──────────────────────────────────────────
        folder_outer = tk.Frame(self.root, bg=self.CLR_BG)
        folder_outer.pack(fill="x", padx=24, pady=(18, 4))

        tk.Label(folder_outer, text="INVOICE FOLDER",
                 font=("Arial", 10, "bold"),
                 bg=self.CLR_BG, fg=self.CLR_MUTED).pack(anchor="w")

        folder_row = tk.Frame(folder_outer, bg=self.CLR_PANEL,
                              highlightbackground=self.CLR_BORDER,
                              highlightthickness=1)
        folder_row.pack(fill="x", pady=(4, 0))

        tk.Entry(folder_row, textvariable=self.source_folder,
                 relief="flat",
                 bg=self.CLR_PANEL, fg=self.CLR_TEXT,
                 readonlybackground=self.CLR_PANEL,
                 insertbackground=self.CLR_TEXT,
                 font=("Consolas", 11),
                 bd=0).pack(side="left", padx=10, pady=8, fill="x", expand=True)

        tk.Button(folder_row, text="Browse",
                  command=self.browse_folder,
                  relief="flat", cursor="hand2",
                  bg="#e0e0e0", fg="#000000",
                  activebackground="#cccccc",
                  activeforeground="#000000",
                  font=("Arial", 10, "bold"),
                  padx=14, pady=6, bd=0).pack(side="right", padx=6, pady=4)

        # ── Format indicator ──────────────────────────────────────────────
        fmt_row = tk.Frame(self.root, bg=self.CLR_BG)
        fmt_row.pack(fill="x", padx=24, pady=(2, 4))
        tk.Label(fmt_row, text="Format:",
                 font=("Arial", 8), bg=self.CLR_BG, fg=self.CLR_MUTED).pack(side="left")
        tk.Label(fmt_row, textvariable=self.detected_fmt,
                 font=("Arial", 10, "bold"),
                 bg=self.CLR_BG, fg=self.CLR_ACCENT).pack(side="left", padx=6)

        # ── Process button ────────────────────────────────────────────────
        btn_frame = tk.Frame(self.root, bg=self.CLR_BG)
        btn_frame.pack(fill="x", padx=24, pady=(8, 8))

        self.process_btn = tk.Button(
            btn_frame, text="▶  PROCESS INVOICES",
            command=self.start_processing,
            relief="flat", cursor="hand2",
            bg="#e0e0e0", fg="#000000",
            activebackground="#cccccc",
            activeforeground="#000000",
            font=("Arial", 13, "bold"),
            pady=10, bd=0,
        )
        self.process_btn.pack(fill="x")

        # ── Log area ──────────────────────────────────────────────────────
        log_label_row = tk.Frame(self.root, bg=self.CLR_BG)
        log_label_row.pack(fill="x", padx=24, pady=(6, 2))
        tk.Label(log_label_row, text="PROGRESS LOG",
                 font=("Arial", 10, "bold"),
                 bg=self.CLR_BG, fg=self.CLR_MUTED).pack(side="left")

        log_outer = tk.Frame(self.root,
                             highlightbackground=self.CLR_BORDER,
                             highlightthickness=1, bg=self.CLR_BORDER)
        log_outer.pack(fill="both", expand=True, padx=24, pady=(0, 16))

        self.log_text = scrolledtext.ScrolledText(
            log_outer,
            relief="flat", bd=0,
            bg=self.CLR_LOG_BG, fg=self.CLR_LOG_FG,
            insertbackground=self.CLR_LOG_FG,
            font=("Consolas", 11),
            wrap="word",
        )
        self.log_text.pack(fill="both", expand=True, padx=1, pady=1)

    def browse_folder(self):
        folder = filedialog.askdirectory(title="Select folder containing PDF invoices")
        if folder:
            self.source_folder.set(folder)
            self.detected_fmt.set("auto-detected per file")
            self.log(f"Selected folder: {folder}")

    def log(self, message):
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.root.update_idletasks()

    def start_processing(self):
        if not self.source_folder.get():
            messagebox.showerror("Error", "Please select a folder first!")
            return

        for path, label in [(OVERLAY_PATH, "overlay.pdf"), (BACKSIDE_PATH, "backside.pdf")]:
            if not os.path.exists(path):
                messagebox.showerror(
                    "Missing Asset",
                    f"Bundled asset not found: {label}\n"
                    f"Expected at: {path}\n\n"
                    "Please ensure the file is in the same directory as this script.",
                )
                return

        self.process_btn.config(state="disabled", text="⏳  Processing...", bg="#aaaaaa")
        self.log_text.delete(1.0, tk.END)

        thread = threading.Thread(target=self.process_pdfs)
        thread.daemon = True
        thread.start()

    # ------------------------------------------------------------------
    @staticmethod
    def _make_bottom_overlay(overlay_path: str) -> fitz.Document:
        """
        Footer-only overlay: copies just the footer strip (y=730+) of the
        overlay into a blank same-size page. The top is left empty so no
        existing page text is obscured.
        """
        src = fitz.open(overlay_path)
        page_rect  = src[0].rect
        footer_rect = fitz.Rect(0, 730, page_rect.width, page_rect.height)
        out = fitz.open()
        new_page = out.new_page(width=page_rect.width, height=page_rect.height)
        new_page.show_pdf_page(footer_rect, src, 0, clip=footer_rect)
        src.close()
        return out  # caller must close

    def apply_overlay_and_backside(self, pdf_path: str, fmt: str) -> bool:
        """
        Stamp overlay on every page then append the back page.

        itin    : full overlay (logo + footer) on ALL pages
        tipitin : full overlay on page 1, footer-only on pages 2+
        """
        temp_path = pdf_path.replace(".pdf", "_temp.pdf")
        try:
            background   = fitz.open(pdf_path)
            overlay_full = fitz.open(OVERLAY_PATH)
            backside     = fitz.open(BACKSIDE_PATH)

            if len(background) < 1 or len(overlay_full) < 1:
                self.log("    Warning: background or overlay has no pages")
                background.close(); overlay_full.close(); backside.close()
                return False

            overlay_bottom = None
            if fmt == FOLDER_TIPITIN and len(background) > 1:
                overlay_bottom = self._make_bottom_overlay(OVERLAY_PATH)

            for i, page in enumerate(background):
                if i == 0 or fmt == FOLDER_ITIN:
                    page.show_pdf_page(page.rect, overlay_full, 0)
                else:
                    page.show_pdf_page(page.rect, overlay_bottom, 0)

            background.insert_pdf(backside, from_page=0, to_page=0)
            background.save(temp_path)
            background.close()
            overlay_full.close()
            backside.close()
            if overlay_bottom:
                overlay_bottom.close()

            os.remove(pdf_path)
            os.rename(temp_path, pdf_path)
            return True

        except Exception as e:
            self.log(f"    ✗ Overlay error: {e}")
            if os.path.exists(temp_path):
                try:
                    os.remove(temp_path)
                except Exception:
                    pass
            return False

    # ------------------------------------------------------------------
    def process_pdfs(self):
        try:
            source_path = self.source_folder.get()
            target_path = os.path.join(source_path, "processed_invoices")

            # Create target directory
            if not os.path.exists(target_path):
                os.makedirs(target_path)
                self.log(f"Created directory: {target_path}")

            # Get PDF files — flat list, same folder only (matches original behaviour)
            pdf_files = [f for f in os.listdir(source_path)
                         if f.lower().endswith('.pdf')
                         and os.path.isfile(os.path.join(source_path, f))]

            if not pdf_files:
                self.log("No PDF files found in the selected folder!")
                self.process_btn.config(state="normal", text="▶  PROCESS INVOICES", bg="#e0e0e0")
                return

            self.log(f"Found {len(pdf_files)} PDF file(s)")
            successful = failed = 0

            for i, file in enumerate(pdf_files, 1):
                self.log(f"\n[{i}/{len(pdf_files)}] Processing: {file}")
                try:
                    src  = os.path.join(source_path, file)
                    dest = os.path.join(target_path, file)
                    shutil.copy2(src, dest)

                    # Extract fields from original before overlay is applied
                    doc  = fitz.open(src)
                    text = doc[0].get_text("text")
                    doc.close()

                    # Detect format from file content
                    fmt = detect_format(text)
                    if not fmt:
                        self.log(f"  ✗ Could not detect format (no SALES PERSON or ITIN NO found)")
                        failed += 1
                        continue

                    self.log(f"  Format: {fmt.upper()}")
                    agent, invoice_no, last_name = extract_fields(text, fmt)

                    overlay_desc = "full overlay" if fmt == FOLDER_ITIN else "full p1 / footer-only p2+"
                    self.log(f"  Applying overlay ({overlay_desc}) and back page...")
                    if self.apply_overlay_and_backside(dest, fmt):
                        self.log("  ✓ Overlay & back page applied")
                    else:
                        self.log("  ✗ Overlay failed, continuing with rename...")

                    if invoice_no and last_name:
                        new_name = build_filename(agent, invoice_no, last_name)
                        new_path = os.path.join(target_path, new_name)
                        os.rename(dest, new_path)
                        self.log(f"  ✓ Renamed to: {new_name}")
                        successful += 1
                    else:
                        missing = []
                        if not invoice_no: missing.append("invoice number")
                        if not last_name:  missing.append("last name")
                        self.log(f"  ✗ Could not extract: {', '.join(missing)}")
                        failed += 1

                except Exception as e:
                    self.log(f"  ✗ Error: {e}")
                    failed += 1

            self.log(f"\n{'='*50}")
            self.log("SUMMARY:")
            self.log(f"  Successfully processed : {successful}")
            self.log(f"  Failed                 : {failed}")
            self.log(f"  Output folder          : {target_path}")

            if successful > 0:
                messagebox.showinfo(
                    "Complete",
                    f"Processing complete!\n"
                    f"Successfully processed: {successful} file(s)\n"
                    f"Failed: {failed} file(s)\n\n"
                    f"Check the 'processed_invoices' folder.",
                )

        except Exception as e:
            self.log(f"Fatal error: {e}")
            messagebox.showerror("Error", f"An error occurred: {e}")
        finally:
            self.process_btn.config(state="normal", text="▶  PROCESS INVOICES", bg="#e0e0e0")

    def run(self):
        self.root.mainloop()


if __name__ == "__main__":
    app = PDFRenamerGUI()
    app.run()