"""
compare_view.py — Side-by-side PDF comparison viewer.

Opens both the original and processed invoice so the user
can visually verify no information is missing.
"""

import os
import sys
import tkinter as tk
from tkinter import ttk

try:
    import fitz
    from PIL import Image, ImageTk
except ImportError as e:
    print(f"Missing dependency: {e}")
    print("pip install PyMuPDF Pillow")
    sys.exit(1)


class CompareViewer:
    """Show two PDFs side by side with page navigation."""

    BG = "#f0f0f0"

    def __init__(self, original_path, processed_path, parent=None):
        if parent:
            self.root = tk.Toplevel(parent)
        else:
            self.root = tk.Tk()

        self.root.title("Compare — Original vs Processed")
        self.root.configure(bg=self.BG)

        # Center and size window
        sw = self.root.winfo_screenwidth()
        sh = self.root.winfo_screenheight()
        w = min(sw - 100, 1400)
        h = min(sh - 100, 900)
        self.root.geometry(f"{w}x{h}+{(sw-w)//2}+{(sh-h)//2}")

        self.orig_path = original_path
        self.proc_path = processed_path

        # Load PDFs
        self.orig_doc = fitz.open(original_path)
        self.proc_doc = fitz.open(processed_path)
        self.max_pages = max(len(self.orig_doc), len(self.proc_doc))
        self.current_page = 0
        self.zoom = 1.3

        # Cache rendered images
        self._cache = {}

        self._build_ui()
        self._render_page()

    def _build_ui(self):
        # Header
        header = tk.Frame(self.root, bg="#ffffff", pady=8)
        header.pack(fill="x")

        tk.Label(header, text="ORIGINAL",
                 font=("Arial", 11, "bold"), bg="#ffffff", fg="#cc0000"
                 ).pack(side="left", padx=(20, 0))
        tk.Label(header, text=os.path.basename(self.orig_path),
                 font=("Arial", 9), bg="#ffffff", fg="#888888"
                 ).pack(side="left", padx=(8, 0))

        # Nav in center
        nav = tk.Frame(header, bg="#ffffff")
        nav.pack(side="left", expand=True)

        self.prev_btn = tk.Button(nav, text="◀ Prev", command=self._prev,
                                   font=("Arial", 10), relief="flat",
                                   bg="#e0e0e0", padx=10, pady=4)
        self.prev_btn.pack(side="left", padx=4)

        self.page_label = tk.Label(nav, text="1 / 1",
                                    font=("Arial", 10, "bold"), bg="#ffffff")
        self.page_label.pack(side="left", padx=8)

        self.next_btn = tk.Button(nav, text="Next ▶", command=self._next,
                                   font=("Arial", 10), relief="flat",
                                   bg="#e0e0e0", padx=10, pady=4)
        self.next_btn.pack(side="left", padx=4)

        tk.Label(header, text="PROCESSED",
                 font=("Arial", 11, "bold"), bg="#ffffff", fg="#005e8d"
                 ).pack(side="right", padx=(0, 20))

        # Separator
        tk.Frame(self.root, bg="#cccccc", height=1).pack(fill="x")

        # Scrollable canvas area
        container = tk.Frame(self.root, bg=self.BG)
        container.pack(fill="both", expand=True)

        self.canvas = tk.Canvas(container, bg=self.BG, highlightthickness=0)
        self.scrollbar = ttk.Scrollbar(container, orient="vertical",
                                        command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        self.scrollbar.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both", expand=True)

        self.inner = tk.Frame(self.canvas, bg=self.BG)
        self.canvas_window = self.canvas.create_window((0, 0), window=self.inner,
                                                         anchor="nw")

        self.inner.bind("<Configure>",
                        lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.bind("<Configure>", self._on_canvas_resize)

        # Mouse wheel scrolling
        self.canvas.bind_all("<MouseWheel>",
                             lambda e: self.canvas.yview_scroll(int(-e.delta/120), "units"))
        # macOS trackpad
        self.canvas.bind_all("<Button-4>",
                             lambda e: self.canvas.yview_scroll(-3, "units"))
        self.canvas.bind_all("<Button-5>",
                             lambda e: self.canvas.yview_scroll(3, "units"))

        # Left and right image labels
        self.left_label = tk.Label(self.inner, bg=self.BG)
        self.left_label.pack(side="left", padx=(10, 5), pady=10, anchor="n")

        self.right_label = tk.Label(self.inner, bg=self.BG)
        self.right_label.pack(side="left", padx=(5, 10), pady=10, anchor="n")

    def _on_canvas_resize(self, event):
        self.canvas.itemconfig(self.canvas_window, width=event.width)

    def _render_pdf_page(self, doc, page_num):
        """Render a PDF page to a PhotoImage."""
        cache_key = (id(doc), page_num, self.zoom)
        if cache_key in self._cache:
            return self._cache[cache_key]

        if page_num >= len(doc):
            # Return a blank placeholder
            img = Image.new("RGB", (int(612 * self.zoom), int(792 * self.zoom)), "#f8f8f8")
            photo = ImageTk.PhotoImage(img)
            self._cache[cache_key] = photo
            return photo

        page = doc[page_num]
        mat = fitz.Matrix(self.zoom, self.zoom)
        pix = page.get_pixmap(matrix=mat)
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        photo = ImageTk.PhotoImage(img)
        self._cache[cache_key] = photo
        return photo

    def _render_page(self):
        left_img = self._render_pdf_page(self.orig_doc, self.current_page)
        right_img = self._render_pdf_page(self.proc_doc, self.current_page)

        self.left_label.configure(image=left_img)
        self.left_label.image = left_img

        self.right_label.configure(image=right_img)
        self.right_label.image = right_img

        self.page_label.configure(text=f"{self.current_page + 1} / {self.max_pages}")
        self.prev_btn.configure(state="normal" if self.current_page > 0 else "disabled")
        self.next_btn.configure(state="normal" if self.current_page < self.max_pages - 1 else "disabled")

        self.canvas.yview_moveto(0)

    def _prev(self):
        if self.current_page > 0:
            self.current_page -= 1
            self._render_page()

    def _next(self):
        if self.current_page < self.max_pages - 1:
            self.current_page += 1
            self._render_page()

    def close(self):
        self.orig_doc.close()
        self.proc_doc.close()
        self.root.destroy()

    def run(self):
        self.root.protocol("WM_DELETE_WINDOW", self.close)
        if isinstance(self.root, tk.Tk):
            self.root.mainloop()


def open_file_in_viewer(filepath):
    """Open a file in the system's default viewer."""
    import subprocess
    try:
        if sys.platform == "win32":
            os.startfile(filepath)
        elif sys.platform == "darwin":
            subprocess.Popen(["open", filepath])
        else:
            subprocess.Popen(["xdg-open", filepath])
    except Exception:
        pass


if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Usage: python compare_view.py <original.pdf> <processed.pdf>")
        sys.exit(1)
    viewer = CompareViewer(sys.argv[1], sys.argv[2])
    viewer.run()