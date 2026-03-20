#!/usr/bin/env python3
"""
Travel Wizards — Invoice Portal

Icons by Flaticon (https://www.flaticon.com):
  Customer Review — flaticon.com/free-icon/customer-review_8743903
  Hotel           — flaticon.com/free-icon/hotel_3086454
"""

import os
import sys
import math
import tkinter as tk
from tkinter import messagebox

try:
    from PIL import Image, ImageTk, ImageDraw
except ImportError:
    root = tk.Tk(); root.withdraw()
    messagebox.showerror("Missing Dependency", "pip install Pillow")
    sys.exit(1)


# ---------------------------------------------------------------------------
# Asset helper
# ---------------------------------------------------------------------------
def _asset(filename):
    base = sys._MEIPASS if getattr(sys, "frozen", False) \
           else os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base, filename)


# ---------------------------------------------------------------------------
# Icon drawing
# ---------------------------------------------------------------------------
def _make_customer_icon(size=96) -> Image.Image:
    img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    s, c = size, size // 2
    r = s * 0.14
    d.ellipse([c - r, s*0.08, c + r, s*0.08 + r*2], fill="#2c3e50")
    bw = s * 0.32
    d.ellipse([c - bw, s*0.42, c + bw, s*0.77], fill="#2c3e50")
    cx, cy, ro, ri = s*0.74, s*0.74, s*0.18, s*0.08
    pts = []
    for i in range(10):
        angle = math.radians(i * 36 - 90)
        rv = ro if i % 2 == 0 else ri
        pts.append((cx + rv * math.cos(angle), cy + rv * math.sin(angle)))
    d.polygon(pts, fill="#e67e22")
    return img


def _make_hotel_icon(size=96) -> Image.Image:
    img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    s = size
    p = s * 0.1
    d.rectangle([p, s*0.28, s - p, s*0.92], fill="#2c3e50")
    d.polygon([(s//2, s*0.06), (p - s*0.04, s*0.30),
               (s - p + s*0.04, s*0.30)], fill="#1a252f")
    dw = s * 0.14
    d.rectangle([s//2 - dw//2, s*0.68, s//2 + dw//2, s*0.92], fill="#ecf0f1")
    ww, wh = s*0.10, s*0.09
    for ry in [s*0.38, s*0.56]:
        for rx in [s*0.22, s*0.50, s*0.78]:
            d.rectangle([rx - ww//2, ry, rx + ww//2, ry + wh], fill="#f39c12")
    return img


def _load_icon(filename, size, fallback_fn):
    path = _asset(filename)
    try:
        img = Image.open(path).convert("RGBA")
        import numpy as np
        if np.array(img).max() == 0:
            raise ValueError("blank image")
        img = img.resize((size, size), Image.LANCZOS)
    except Exception:
        img = fallback_fn(size)
    return ImageTk.PhotoImage(img)


# ---------------------------------------------------------------------------
# Portal card
# ---------------------------------------------------------------------------
class PortalCard(tk.Frame):
    BG       = "#ffffff"
    BG_HOVER = "#f0f0f0"
    BORDER   = "#cccccc"
    TEXT     = "#1a1a1a"
    TEXT_DIM = "#888888"
    SUB      = "#666666"

    def __init__(self, parent, title, subtitle, icon_img, command):
        super().__init__(parent, bg=self.BG,
                         highlightbackground=self.BORDER,
                         highlightthickness=1,
                         cursor="hand2",
                         width=210, height=230)
        self.pack_propagate(False)
        self._cmd = command
        self._all = [self]

        t = tk.Label(self, text=title, font=("Georgia", 14, "bold"),
                     bg=self.BG, fg=self.TEXT, justify="center")
        t.pack(pady=(26, 2))
        self._all.append(t)

        s = tk.Label(self, text=subtitle, font=("Arial", 8),
                     bg=self.BG, fg=self.SUB)
        s.pack()
        self._all.append(s)

        ico = tk.Label(self, image=icon_img, bg=self.BG)
        ico.image = icon_img
        ico.pack(pady=(14, 0))
        self._all.append(ico)

        for w in self._all:
            w.bind("<Enter>",    self._enter)
            w.bind("<Leave>",    self._leave)
            w.bind("<Button-1>", lambda e: command())

    def _enter(self, _=None):
        for w in self._all:
            w.config(bg=self.BG_HOVER)
        self.config(highlightbackground="#999999")
        # Dim the title on hover
        self._all[1].config(fg=self.TEXT_DIM)

    def _leave(self, _=None):
        for w in self._all:
            w.config(bg=self.BG)
        self.config(highlightbackground=self.BORDER)
        self._all[1].config(fg=self.TEXT)


# ---------------------------------------------------------------------------
# Portal
# ---------------------------------------------------------------------------
class InvoicePortal:
    BG = "#fafafa"

    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Travel Wizards — Invoice Portal")
        self.root.configure(bg=self.BG)
        self.root.resizable(False, False)
        _center(self.root, 600, 460)
        self._build()

    def _build(self):
        # Header
        tk.Label(self.root, text="TRAVEL  WIZARDS",
                 font=("Georgia", 28, "bold"),
                 bg=self.BG, fg="#000000").pack(pady=(38, 2))
        tk.Label(self.root, text="I N V O I C E   P O R T A L",
                 font=("Arial", 11),
                 bg=self.BG, fg="#cc0000").pack()

        tk.Frame(self.root, bg="#cccccc", height=1).pack(
            fill="x", padx=44, pady=(20, 0))

        # Cards
        row = tk.Frame(self.root, bg=self.BG)
        row.pack(expand=True, pady=28)

        c_img = _load_icon("customer-review.png", 80, _make_customer_icon)
        h_img = _load_icon("hotel.png",           80, _make_hotel_icon)

        PortalCard(row, "Customer\nInvoices", "PDF invoice processor",
                   c_img, self._open_pdf).pack(side="left", padx=20)
        PortalCard(row, "Hotel\nInvoices",    "Hotel invoice processor",
                   h_img, self._open_hotel).pack(side="left", padx=20)

        # Credit
        cr = tk.Label(self.root,
                      text="Icons by Flaticon  ·  flaticon.com",
                      font=("Arial", 7), bg=self.BG, fg="#bbbbbb",
                      cursor="hand2")
        cr.pack(side="bottom", pady=8)
        cr.bind("<Button-1>", lambda e: _open_url("https://www.flaticon.com"))

    def _open_pdf(self):
        try:
            from v5.invoice_processor import PDFRenamerGUI
            PDFRenamerGUI(parent=self.root)
        except ImportError:
            messagebox.showerror("Not Found",
                                 "invoice_processor.py not found in the same folder.")

    def _open_hotel(self):
        try:
            from v5.hotel_invoice_processor import HotelInvoiceGUI
            HotelInvoiceGUI(parent=self.root)
        except ImportError:
            messagebox.showerror("Not Found",
                                 "hotel_invoice_processor.py not found in the same folder.")

    def run(self):
        self.root.mainloop()


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
def _center(win, w, h):
    win.update_idletasks()
    sw = win.winfo_screenwidth()
    sh = win.winfo_screenheight()
    win.geometry(f"{w}x{h}+{(sw-w)//2}+{(sh-h)//2}")


def _open_url(url):
    import subprocess
    try:
        if sys.platform == "win32":    os.startfile(url)
        elif sys.platform == "darwin": subprocess.run(["open", url])
        else:                          subprocess.run(["xdg-open", url])
    except Exception:
        pass


# ---------------------------------------------------------------------------
if __name__ == "__main__":
    InvoicePortal().run()
