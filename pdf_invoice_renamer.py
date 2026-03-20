#!/usr/bin/env python3
"""
PDF Invoice Renamer - Travel Wizards
Supports both invoice formats (ITIN/INVOICE and ITIN NO).
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
    """Return the absolute path to a bundled asset."""
    if getattr(sys, "frozen", False):
        # PyInstaller sets sys._MEIPASS to the temp extraction directory
        base = sys._MEIPASS
    else:
        base = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base, filename)


OVERLAY_PATH = _asset("overlay.pdf")
BACKSIDE_PATH = _asset("backside.pdf")


# ---------------------------------------------------------------------------
# Regex helpers – handle both invoice layouts
# ---------------------------------------------------------------------------
def extract_fields(text):
    """
    Extract (agent_initials, invoice_no, last_name) from invoice text.

    Layout A  – original invoice
        SALES PERSON: 89SSV   ITIN/INVOICE NO.   328016
        FOR: POOLE/EDWARD …

    Layout B  – airline itinerary
        ITIN NO:   360085
        FAHRNEY/DONALD WILLIAM   (first non-trivial NAME/ line at top)
        No SALES PERSON line.
    """
    # --- Invoice / ITIN number ---
    # Layout A:  ITIN/INVOICE NO.  328016  (may have trailing "01")
    inv = re.search(r'ITIN/INVOICE NO\.\s+(\d+)', text)
    if not inv:
        # Layout B:  ITIN NO:  360085
        inv = re.search(r'ITIN NO:\s*(\d+)', text)
    invoice_no = inv.group(1) if inv else None

    # --- Sales person / agent (optional) ---
    agent_match = re.search(r'SALES PERSON:\s*(\S+)', text)
    if agent_match:
        raw = agent_match.group(1)   # e.g. "89SSV"
        agent_initials = raw[-2:]    # last 2 chars → "SV"
    else:
        agent_initials = None        # not present in Layout B

    # --- Last name ---
    # Layout A:  FOR: POOLE/EDWARD GRAY
    last = re.search(r'FOR:\s+([A-Z]+)/', text)
    if not last:
        # Layout B: first line that is  LASTNAME/FIRSTNAME (len ≥ 3)
        last = re.search(r'^\s{0,10}([A-Z]{3,})/', text, re.MULTILINE)
    last_name = last.group(1) if last else None

    return agent_initials, invoice_no, last_name


def build_filename(agent_initials, invoice_no, last_name):
    """
    {agent}{invoice}_{last_name}.pdf
    Agent is omitted (with its underscore) when not present.
    """
    if agent_initials:
        return f"{invoice_no}_{last_name}_{agent_initials}.pdf"
    else:
        return f"{invoice_no}_{last_name}.pdf"


# ---------------------------------------------------------------------------
# GUI
# ---------------------------------------------------------------------------
class PDFRenamerGUI:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("PDF Invoice Renamer")
        self.root.geometry("800x300")

        self.source_folder = tk.StringVar()
        self.setup_ui()

    def setup_ui(self):
        title = tk.Label(self.root, text="Modify Invoices", font=("Arial", 16, "bold"))
        title.pack(pady=10)

        instructions = tk.Label(
            self.root,
            text="Select the folder containing your PDF invoices.\n"
                 "Processed files will be saved in a 'processed_invoices' subfolder.",
            wraplength=500,
        )
        instructions.pack(pady=5)

        # Folder selection
        folder_frame = tk.Frame(self.root)
        folder_frame.pack(pady=10, padx=20, fill="x")
        tk.Label(folder_frame, text="Folder:").pack(side="left")
        tk.Entry(folder_frame, textvariable=self.source_folder, width=40,
                 state="readonly").pack(side="left", padx=5, fill="x", expand=True)
        tk.Button(folder_frame, text="Browse",
                  command=self.browse_folder).pack(side="right")

        # Process button
        self.process_btn = tk.Button(
            self.root, text="Process PDFs",
            command=self.start_processing,
            bg="navy", fg="black",
            font=("Arial", 12, "bold"),
            height=2,
        )
        self.process_btn.pack(pady=20)

        # Progress / log
        tk.Label(self.root, text="Progress Log:").pack(anchor="w", padx=20)
        self.log_text = scrolledtext.ScrolledText(self.root, height=15, width=70)
        self.log_text.pack(pady=5, padx=20, fill="both", expand=True)

    def browse_folder(self):
        folder = filedialog.askdirectory(title="Select folder containing PDF invoices")
        if folder:
            self.source_folder.set(folder)
            self.log("Selected folder: " + folder)

    def log(self, message):
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.root.update()

    def start_processing(self):
        if not self.source_folder.get():
            messagebox.showerror("Error", "Please select a folder first!")
            return

        # Verify bundled assets exist
        for path, label in [(OVERLAY_PATH, "overlay.pdf"), (BACKSIDE_PATH, "backside.pdf")]:
            if not os.path.exists(path):
                messagebox.showerror(
                    "Missing Asset",
                    f"Bundled asset not found: {label}\n"
                    f"Expected at: {path}\n\n"
                    "Please ensure the file is in the same directory as this script.",
                )
                return

        self.process_btn.config(state="disabled", text="Processing...")
        self.log_text.delete(1.0, tk.END)

        thread = threading.Thread(target=self.process_pdfs)
        thread.daemon = True
        thread.start()

    # ------------------------------------------------------------------
    @staticmethod
    def _make_bottom_overlay(overlay_path: str) -> fitz.Document:
        """
        Return an in-memory fitz.Document containing ONLY the footer strip
        of the overlay, placed at its original position on a blank page.
        Nothing is drawn in the top portion — existing page text is untouched.
        """
        src = fitz.open(overlay_path)
        page_rect = src[0].rect  # 612 x 792

        # Clip to just the footer region (y=730 to bottom)
        footer_rect = fitz.Rect(0, 730, page_rect.width, page_rect.height)

        out = fitz.open()
        new_page = out.new_page(width=page_rect.width, height=page_rect.height)
        # target_rect == footer_rect so content lands at the exact same position
        new_page.show_pdf_page(footer_rect, src, 0, clip=footer_rect)
        src.close()
        return out  # caller must close

    def apply_overlay_and_backside(self, pdf_path, has_sales_person: bool):
        """
        Stamp overlay on every page, then append the back page.
        - Page 0  : full overlay (logo + footer)
        - Pages 1+: bottom-only overlay (footer only) — only for Layout B
                    (documents without a sales person line)
        """
        temp_path = pdf_path.replace(".pdf", "_temp.pdf")
        try:
            background = fitz.open(pdf_path)
            overlay_full = fitz.open(OVERLAY_PATH)
            backside = fitz.open(BACKSIDE_PATH)

            if len(background) < 1 or len(overlay_full) < 1:
                self.log("    Warning: background or overlay has no pages")
                background.close(); overlay_full.close(); backside.close()
                return False

            # Build bottom-only overlay once if needed
            overlay_bottom = None
            if not has_sales_person and len(background) > 1:
                overlay_bottom = self._make_bottom_overlay(OVERLAY_PATH)

            for i, page in enumerate(background):
                if i == 0 or has_sales_person:
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

            if not os.path.exists(target_path):
                os.makedirs(target_path)
                self.log(f"Created directory: {target_path}")

            pdf_files = [
                f for f in os.listdir(source_path)
                if f.lower().endswith(".pdf") and os.path.isfile(os.path.join(source_path, f))
            ]

            if not pdf_files:
                self.log("No PDF files found in the selected folder!")
                self.process_btn.config(state="normal", text="Process PDFs")
                return

            self.log(f"Found {len(pdf_files)} PDF file(s)")
            successful = failed = 0

            for i, file in enumerate(pdf_files, 1):
                self.log(f"\n[{i}/{len(pdf_files)}] Processing: {file}")
                try:
                    src = os.path.join(source_path, file)
                    dest = os.path.join(target_path, file)
                    shutil.copy2(src, dest)

                    # Extract text from original (no overlay interference)
                    doc = fitz.open(src)
                    text = doc[0].get_text("text")
                    doc.close()

                    agent, invoice_no, last_name = extract_fields(text)

                    # Overlay + back page (pass format flag so page 2+ get footer-only overlay)
                    has_sales_person = agent is not None
                    overlay_type = "full" if has_sales_person else "footer-only on pages 2+"
                    self.log(f"  Applying overlay ({overlay_type}) and back page...")
                    if self.apply_overlay_and_backside(dest, has_sales_person):
                        self.log("  ✓ Overlay & back page applied")
                    else:
                        self.log("  ✗ Overlay failed, continuing with rename...")

                    if invoice_no and last_name:
                        new_name = build_filename(agent, invoice_no, last_name)
                        new_path = os.path.join(target_path, new_name)
                        os.rename(dest, new_path)
                        self.log(f"  ✓ Renamed to: {new_name}")
                        if not agent:
                            self.log("    (no SALES PERSON found – agent omitted from name)")
                        successful += 1
                    else:
                        missing = []
                        if not invoice_no:
                            missing.append("invoice number")
                        if not last_name:
                            missing.append("last name")
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
            self.process_btn.config(state="normal", text="Process PDFs")

    def run(self):
        self.root.mainloop()


if __name__ == "__main__":
    app = PDFRenamerGUI()
    app.run()