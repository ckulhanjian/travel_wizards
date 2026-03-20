#!/usr/bin/env python3
"""
PDF Invoice Renamer - Travel Wizards
"""

import os
import sys
import shutil
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import threading
import re

try:
    import fitz
except ImportError as e:
    root = tk.Tk(); root.withdraw()
    messagebox.showerror("Missing Dependencies",
                         f"Missing required library: {e}\n\npip install PyMuPDF")
    sys.exit(1)


def _asset(filename):
    if getattr(sys, "frozen", False):
        base = sys._MEIPASS
    else:
        base = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base, filename)


def _center_window(win, w, h):
    win.update_idletasks()
    sw = win.winfo_screenwidth()
    sh = win.winfo_screenheight()
    win.geometry(f"{w}x{h}+{(sw-w)//2}+{(sh-h)//2}")


OVERLAY_PATH  = _asset("overlay.pdf")
BACKSIDE_PATH = _asset("backside.pdf")
FOLDER_ITIN    = "itin"
FOLDER_TIPITIN = "tipitin"


def detect_format(text: str):
    if re.search(r'SALES PERSON:', text):
        return FOLDER_ITIN
    if re.search(r'ITIN NO:', text):
        return FOLDER_TIPITIN
    return None


def extract_fields(text, fmt: str):
    if fmt == FOLDER_ITIN:
        inv   = re.search(r'ITIN/INVOICE NO\.\s+(\d+)', text)
        agent = re.search(r'SALES PERSON:\s*(\S+)', text)
        last  = re.search(r'FOR:\s+([A-Z]+)/', text)
        invoice_no     = inv.group(1)        if inv   else None
        agent_initials = agent.group(1)[-2:] if agent else None
        last_name      = last.group(1)       if last  else None
    else:
        inv   = re.search(r'ITIN NO:\s*(\d+)', text)
        last  = re.search(r'^\s{0,10}([A-Z]{3,})/', text, re.MULTILINE)
        invoice_no     = inv.group(1)  if inv  else None
        agent_initials = None
        last_name      = last.group(1) if last else None
    return agent_initials, invoice_no, last_name


def build_filename(agent_initials, invoice_no, last_name):
    if agent_initials:
        return f"{invoice_no} {last_name} {agent_initials}.pdf"
    return f"{invoice_no} {last_name}.pdf"


class PDFRenamerGUI:
    CLR_BG        = "#ffffff"
    CLR_PANEL     = "#f5f5f5"
    CLR_ACCENT    = "#000000"
    CLR_TEXT      = "#000000"
    CLR_MUTED     = "#555555"
    CLR_BORDER    = "#cccccc"
    CLR_BTN_BG    = "#e0e0e0"
    CLR_BTN_ACT   = "#cccccc"
    CLR_LOG_BG    = "#f9f9f9"
    CLR_LOG_FG    = "#111111"

    def __init__(self, parent=None):
        if parent is None:
            self.root = tk.Tk()
        else:
            self.root = tk.Toplevel(parent)
        self.root.title("Travel Wizards — Invoice Processor")
        self.root.resizable(True, True)
        self.root.configure(bg=self.CLR_BG)
        self.root.option_add("*Button.Background", "#e0e0e0")
        self.root.option_add("*Button.Foreground", "#000000")
        self.root.option_add("*Button.activeBackground", "#cccccc")
        self.root.option_add("*Button.activeForeground", "#000000")
        self.root.option_add("*Button.relief", "flat")
        _center_window(self.root, 860, 520)

        self.source_folder = tk.StringVar()
        self.detected_fmt  = tk.StringVar(value="—")
        self._draw_logo()
        self.setup_ui()

    def _draw_logo(self):
        c = tk.Canvas(self.root, height=74, bg=self.CLR_BG,
                      highlightthickness=0, bd=0)
        c.pack(fill="x", padx=0, pady=0)
        c.create_line(0, 73, 2000, 73, fill=self.CLR_BORDER, width=1)
        c.create_text(430, 26, text="TRAVEL  WIZARDS",
                      font=("Georgia", 26, "bold"),
                      fill=self.CLR_ACCENT, anchor="center")
        c.create_text(430, 52, text="M A G I C A L   J O U R N E Y S",
                      font=("Arial", 10), fill="#cc0000", anchor="center")
        if isinstance(self.root, tk.Toplevel):
            tk.Button(self.root, text="⌂  Home", command=self.root.destroy,
                      relief="flat", cursor="hand2",
                      bg=self.CLR_BG, fg="#000000",
                      activebackground=self.CLR_BG,
                      font=("Arial", 9, "bold"),
                      padx=10, pady=4, bd=0).place(x=10, y=10)

    def setup_ui(self):
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
                 relief="flat", bg=self.CLR_PANEL, fg=self.CLR_TEXT,
                 readonlybackground=self.CLR_PANEL,
                 insertbackground=self.CLR_TEXT,
                 font=("Consolas", 11), bd=0).pack(side="left", padx=10, pady=8,
                                                    fill="x", expand=True)
        tk.Button(folder_row, text="Browse", command=self.browse_folder,
                  relief="flat", cursor="hand2",
                  bg="#e0e0e0", fg="#000000",
                  activebackground="#cccccc",
                  font=("Arial", 10, "bold"),
                  padx=14, pady=6, bd=0).pack(side="right", padx=6, pady=4)

        fmt_row = tk.Frame(self.root, bg=self.CLR_BG)
        fmt_row.pack(fill="x", padx=24, pady=(2, 4))
        tk.Label(fmt_row, text="Format:", font=("Arial", 8),
                 bg=self.CLR_BG, fg=self.CLR_MUTED).pack(side="left")
        tk.Label(fmt_row, textvariable=self.detected_fmt,
                 font=("Arial", 10, "bold"),
                 bg=self.CLR_BG, fg=self.CLR_ACCENT).pack(side="left", padx=6)

        btn_frame = tk.Frame(self.root, bg=self.CLR_BG)
        btn_frame.pack(fill="x", padx=24, pady=(8, 8))
        self.process_btn = tk.Button(
            btn_frame, text="▶  PROCESS INVOICES",
            command=self.start_processing,
            relief="flat", cursor="hand2",
            bg="#e0e0e0", fg="#000000",
            activebackground="#cccccc",
            font=("Arial", 13, "bold"),
            pady=10, bd=0)
        self.process_btn.pack(fill="x")

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
            log_outer, relief="flat", bd=0,
            bg=self.CLR_LOG_BG, fg=self.CLR_LOG_FG,
            insertbackground=self.CLR_LOG_FG,
            font=("Consolas", 11), wrap="word")
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
                messagebox.showerror("Missing Asset",
                                     f"Bundled asset not found: {label}\n"
                                     f"Expected at: {path}")
                return
        self.process_btn.config(state="disabled", text="⏳  Processing...", bg="#aaaaaa")
        self.log_text.delete(1.0, tk.END)
        thread = threading.Thread(target=self.process_pdfs)
        thread.daemon = True
        thread.start()

    @staticmethod
    def _make_bottom_overlay(overlay_path: str) -> fitz.Document:
        src = fitz.open(overlay_path)
        page_rect  = src[0].rect
        footer_rect = fitz.Rect(0, 730, page_rect.width, page_rect.height)
        out = fitz.open()
        new_page = out.new_page(width=page_rect.width, height=page_rect.height)
        new_page.show_pdf_page(footer_rect, src, 0, clip=footer_rect)
        src.close()
        return out

    def apply_overlay_and_backside(self, pdf_path: str, fmt: str) -> bool:
        temp_path = pdf_path.replace(".pdf", "_temp.pdf")
        try:
            background   = fitz.open(pdf_path)
            overlay_full = fitz.open(OVERLAY_PATH)
            backside     = fitz.open(BACKSIDE_PATH)
            if len(background) < 1 or len(overlay_full) < 1:
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
            background.close(); overlay_full.close(); backside.close()
            if overlay_bottom:
                overlay_bottom.close()
            os.remove(pdf_path)
            os.rename(temp_path, pdf_path)
            return True
        except Exception as e:
            self.log(f"    ✗ Overlay error: {e}")
            if os.path.exists(temp_path):
                try: os.remove(temp_path)
                except Exception: pass
            return False

    def process_pdfs(self):
        try:
            source_path = self.source_folder.get()
            target_path = os.path.join(source_path, "processed_invoices")
            if not os.path.exists(target_path):
                os.makedirs(target_path)
                self.log(f"Created directory: {target_path}")
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
                    doc  = fitz.open(src)
                    text = doc[0].get_text("text")
                    doc.close()
                    fmt = detect_format(text)
                    if not fmt:
                        self.log(f"  ✗ Could not detect format")
                        failed += 1
                        continue
                    self.log(f"  Format: {fmt.upper()}")
                    agent, invoice_no, last_name = extract_fields(text, fmt)
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
                        self.log(f"  ✗ Could not extract fields")
                        failed += 1
                except Exception as e:
                    self.log(f"  ✗ Error: {e}")
                    failed += 1
            self.log(f"\n{'='*50}\nSUMMARY:\n  Successfully processed : {successful}\n  Failed : {failed}")
            if successful > 0:
                messagebox.showinfo("Complete",
                                    f"Processing complete!\n✓ {successful} file(s)\n✗ {failed} failed")
        except Exception as e:
            self.log(f"Fatal error: {e}")
            messagebox.showerror("Error", f"An error occurred: {e}")
        finally:
            self.process_btn.config(state="normal", text="▶  PROCESS INVOICES", bg="#e0e0e0")

    def run(self):
        if isinstance(self.root, tk.Tk):
            self.root.mainloop()


if __name__ == "__main__":
    app = PDFRenamerGUI()
    app.run()