#!/usr/bin/env python3
"""
Airport Database Manager - Travel Wizards

Lets staff browse, search, and UPDATE existing airport entries in
airport_lookup.py's IATA dict — the same database airport_resolver.py's
"unknown airport" prompt writes to during invoice processing, and the same
one invoice_generator.py reads from. This screen only edits values on
airports that already exist; it doesn't add or delete entries (new ones
come in through the "unknown airport" prompt while processing invoices).

All writes go through airport_resolver.update_airport_entry(), which is
the same validated, safe-write, verified-live-update path used everywhere
else this file gets edited — so a save made here is subject to the exact
same corruption/consistency protections as an add made during processing,
and is immediately visible to the next invoice processed, with no restart.
"""

import tkinter as tk
from tkinter import messagebox

import airport_lookup
import airport_resolver

FIELD_LABELS = {
    "name": "Airport Name",
    "city": "City",
}
FIELD_ORDER = ["name", "city"]


def _center_window(win, w, h):
    win.update_idletasks()
    sw = win.winfo_screenwidth()
    sh = win.winfo_screenheight()
    win.geometry(f"{w}x{h}+{(sw-w)//2}+{(sh-h)//2}")


class AirportManagerGUI:
    CLR_BG         = "#ffffff"
    CLR_PANEL      = "#f5f5f5"
    CLR_ACCENT     = "#000000"
    CLR_TEXT       = "#000000"
    CLR_MUTED      = "#555555"
    CLR_BORDER     = "#cccccc"
    CLR_ROW_SEL    = "#dceefb"
    CLR_SAVE_IDLE  = "#e8e8e8"
    CLR_SAVE_DIRTY = "#2e8b46"   # green — shown only when something changed

    def __init__(self, parent=None):
        if parent is None:
            self.root = tk.Tk()
        else:
            self.root = tk.Toplevel(parent)
        self.root.title("Travel Wizards — Airport Database")
        self.root.configure(bg=self.CLR_BG)
        self.root.resizable(True, True)
        self.root.option_add("*Button.relief", "flat")
        _center_window(self.root, 1000, 640)
        self.root.minsize(820, 480)

        self.selected_code = None
        self.original_values = {}
        self.field_vars = {}
        self.field_entries = {}
        self._suspend_dirty_check = False

        self._draw_logo()
        self._build_ui()
        self._refresh_list()
        # Pick up entries added elsewhere (e.g. via the "unknown airport"
        # prompt during processing) as soon as this window regains focus.
        self.root.bind("<FocusIn>", lambda e: self._refresh_list())

    # ------------------------------------------------------------------
    def _draw_logo(self):
        c = tk.Canvas(self.root, height=64, bg=self.CLR_BG,
                      highlightthickness=0, bd=0)
        c.pack(fill="x", padx=0, pady=0)
        c.create_line(0, 63, 2000, 63, fill=self.CLR_BORDER, width=1)
        c.create_text(500, 22, text="TRAVEL  WIZARDS",
                      font=("Georgia", 20, "bold"),
                      fill=self.CLR_ACCENT, anchor="center")
        c.create_text(500, 44, text="A I R P O R T   D A T A B A S E",
                      font=("Arial", 9), fill="#cc0000", anchor="center")
        if isinstance(self.root, tk.Toplevel):
            tk.Button(self.root, text="⌂  Home", command=self.root.destroy,
                      relief="flat", cursor="hand2",
                      bg=self.CLR_BG, fg="#000000",
                      activebackground=self.CLR_BG,
                      font=("Arial", 9, "bold"),
                      padx=10, pady=4, bd=0).place(x=10, y=8)

    def _build_ui(self):
        search_row = tk.Frame(self.root, bg=self.CLR_BG)
        search_row.pack(fill="x", padx=20, pady=(14, 8))
        tk.Label(search_row, text="SEARCH", font=("Arial", 9, "bold"),
                 bg=self.CLR_BG, fg=self.CLR_MUTED).pack(side="left")
        tk.Label(search_row, text="(IATA code, airport name, or city)",
                 font=("Arial", 8), bg=self.CLR_BG,
                 fg="#999999").pack(side="left", padx=(6, 0))

        search_box = tk.Frame(self.root, bg=self.CLR_PANEL,
                              highlightbackground=self.CLR_BORDER,
                              highlightthickness=1)
        search_box.pack(fill="x", padx=20, pady=(0, 12))
        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", lambda *a: self._refresh_list())
        entry = tk.Entry(search_box, textvariable=self.search_var,
                         relief="flat", bg=self.CLR_PANEL, fg=self.CLR_TEXT,
                         insertbackground=self.CLR_TEXT,
                         font=("Consolas", 12), bd=0)
        entry.pack(side="left", fill="x", expand=True, padx=10, pady=8)
        entry.focus_set()

        body = tk.Frame(self.root, bg=self.CLR_BG)
        body.pack(fill="both", expand=True, padx=20, pady=(0, 16))
        body.columnconfigure(0, weight=2)
        body.columnconfigure(1, weight=3)
        body.rowconfigure(0, weight=1)

        self._build_left_panel(body)
        self._build_right_panel(body)

    # ------------------------------------------------------------------
    # LEFT PANEL
    # ------------------------------------------------------------------
    def _build_left_panel(self, parent):
        outer = tk.Frame(parent, bg=self.CLR_BG,
                         highlightbackground=self.CLR_BORDER,
                         highlightthickness=1)
        outer.grid(row=0, column=0, sticky="nsew", padx=(0, 10))

        header = tk.Frame(outer, bg=self.CLR_PANEL)
        header.pack(fill="x")
        tk.Label(header, text="AIRPORTS", font=("Arial", 10, "bold"),
                 bg=self.CLR_PANEL, fg=self.CLR_MUTED).pack(
                     side="left", padx=10, pady=6)
        self.count_label = tk.Label(header, text="", font=("Arial", 9),
                                    bg=self.CLR_PANEL, fg="#999999")
        self.count_label.pack(side="right", padx=10)

        list_frame = tk.Frame(outer, bg=self.CLR_BG)
        list_frame.pack(fill="both", expand=True)
        scrollbar = tk.Scrollbar(list_frame)
        scrollbar.pack(side="right", fill="y")

        self.listbox = tk.Listbox(
            list_frame, relief="flat", bd=0,
            bg=self.CLR_BG, fg=self.CLR_TEXT,
            selectbackground=self.CLR_ROW_SEL, selectforeground=self.CLR_TEXT,
            activestyle="none", font=("Consolas", 11),
            yscrollcommand=scrollbar.set, highlightthickness=0)
        self.listbox.pack(side="left", fill="both", expand=True, padx=(2, 0))
        scrollbar.config(command=self.listbox.yview)
        self.listbox.bind("<<ListboxSelect>>", self._on_select_airport)

    # ------------------------------------------------------------------
    # RIGHT PANEL
    # ------------------------------------------------------------------
    def _build_right_panel(self, parent):
        outer = tk.Frame(parent, bg=self.CLR_BG,
                         highlightbackground=self.CLR_BORDER,
                         highlightthickness=1)
        outer.grid(row=0, column=1, sticky="nsew")
        outer.rowconfigure(1, weight=1)
        outer.columnconfigure(0, weight=1)

        header = tk.Frame(outer, bg=self.CLR_PANEL)
        header.grid(row=0, column=0, sticky="ew")
        tk.Label(header, text="AIRPORT DETAILS", font=("Arial", 10, "bold"),
                 bg=self.CLR_PANEL, fg=self.CLR_MUTED).pack(
                     side="left", padx=10, pady=6)

        self.detail_frame = tk.Frame(outer, bg=self.CLR_BG)
        self.detail_frame.grid(row=1, column=0, sticky="nsew", padx=20, pady=16)

        self.placeholder = tk.Label(
            self.detail_frame,
            text="Select an airport from the list on the left to view\nand edit its details.",
            font=("Arial", 11), fg="#999999", bg=self.CLR_BG, justify="left")
        self.placeholder.pack(anchor="w", pady=20)

        self.save_bar = tk.Frame(outer, bg=self.CLR_SAVE_IDLE, height=44,
                                 cursor="arrow")
        self.save_bar.grid(row=2, column=0, sticky="ew")
        self.save_bar.grid_propagate(False)
        self.save_label = tk.Label(self.save_bar, text="",
                                   font=("Arial", 11, "bold"),
                                   bg=self.CLR_SAVE_IDLE, fg="#ffffff")
        self.save_label.pack(expand=True)
        self._save_bar_active = False

    # ------------------------------------------------------------------
    # List population / filtering — reads airport_lookup.IATA live
    # ------------------------------------------------------------------
    def _current_records(self):
        """{code: {'name':..., 'city':...}} straight from the live module —
        always current, including anything added/edited elsewhere."""
        return {code: {"name": name, "city": city}
                for code, (name, city) in airport_lookup.IATA.items()}

    def _matching_codes(self):
        self.airports = self._current_records()
        query = self.search_var.get().strip().lower()
        codes = list(self.airports.keys())
        codes.sort(key=lambda c: (self.airports[c].get("name") or "").lower())

        if not query:
            return codes

        def matches(code):
            rec = self.airports[code]
            haystacks = [code, rec.get("name", ""), rec.get("city", "")]
            return any(query in (h or "").lower() for h in haystacks)

        return [c for c in codes if matches(c)]

    def _refresh_list(self):
        codes = self._matching_codes()
        self.listbox.delete(0, "end")
        self._visible_codes = codes
        for code in codes:
            rec = self.airports[code]
            self.listbox.insert(
                "end", f"{rec.get('name','')}  ({code}) — {rec.get('city','')}")
        self.count_label.config(text=f"{len(codes)} of {len(self.airports)}")

        if self.selected_code and self.selected_code in codes:
            idx = codes.index(self.selected_code)
            self.listbox.selection_set(idx)
            self.listbox.see(idx)

    # ------------------------------------------------------------------
    # Selecting an airport
    # ------------------------------------------------------------------
    def _on_select_airport(self, event=None):
        sel = self.listbox.curselection()
        if not sel:
            return
        code = self._visible_codes[sel[0]]
        if code == self.selected_code:
            return

        if self._is_dirty():
            if not messagebox.askyesno(
                    "Unsaved changes",
                    "You have unsaved changes to this airport.\n"
                    "Discard them and switch airports?", parent=self.root):
                if self.selected_code in self._visible_codes:
                    idx = self._visible_codes.index(self.selected_code)
                    self.listbox.selection_clear(0, "end")
                    self.listbox.selection_set(idx)
                return

        self._load_airport(code)

    def _load_airport(self, code):
        self.selected_code = code
        rec = self.airports[code]
        self.original_values = {f: str(rec.get(f, "")) for f in FIELD_ORDER}
        self._render_detail_form(code, rec)
        self._set_save_bar(dirty=False)

    def _render_detail_form(self, code, rec):
        for w in self.detail_frame.winfo_children():
            w.destroy()
        self.field_vars = {}
        self.field_entries = {}

        code_row = tk.Frame(self.detail_frame, bg=self.CLR_BG)
        code_row.pack(fill="x", pady=(0, 16))
        tk.Label(code_row, text="IATA CODE", font=("Arial", 8, "bold"),
                 bg=self.CLR_BG, fg=self.CLR_MUTED).pack(anchor="w")
        tk.Label(code_row, text=code, font=("Georgia", 22, "bold"),
                 bg=self.CLR_BG, fg=self.CLR_ACCENT).pack(anchor="w")
        tk.Label(code_row, text="(not editable here — used as the lookup key)",
                 font=("Arial", 8), bg=self.CLR_BG, fg="#999999").pack(anchor="w")

        self._suspend_dirty_check = True
        for field in FIELD_ORDER:
            row = tk.Frame(self.detail_frame, bg=self.CLR_BG)
            row.pack(fill="x", pady=6)
            tk.Label(row, text=FIELD_LABELS.get(field, field.title()),
                     font=("Arial", 9, "bold"), bg=self.CLR_BG,
                     fg=self.CLR_MUTED, width=14, anchor="w").pack(side="left")

            var = tk.StringVar(value=str(rec.get(field, "")))
            var.trace_add("write", lambda *a: self._check_dirty())
            self.field_vars[field] = var

            box = tk.Frame(row, bg=self.CLR_PANEL,
                           highlightbackground=self.CLR_BORDER,
                           highlightthickness=1)
            box.pack(side="left", fill="x", expand=True)
            entry = tk.Entry(box, textvariable=var, relief="flat",
                             bg=self.CLR_PANEL, fg=self.CLR_TEXT,
                             insertbackground=self.CLR_TEXT,
                             font=("Consolas", 12), bd=0)
            entry.pack(fill="x", padx=8, pady=6)
            self.field_entries[field] = entry
        self._suspend_dirty_check = False

    # ------------------------------------------------------------------
    # Dirty-state tracking + save bar
    # ------------------------------------------------------------------
    def _is_dirty(self):
        if not self.selected_code:
            return False
        return any(
            self.field_vars[f].get() != self.original_values.get(f, "")
            for f in FIELD_ORDER)

    def _check_dirty(self):
        if self._suspend_dirty_check:
            return
        self._set_save_bar(dirty=self._is_dirty())

    def _set_save_bar(self, dirty: bool):
        self._save_bar_active = dirty
        if dirty:
            self.save_bar.config(bg=self.CLR_SAVE_DIRTY, cursor="hand2")
            self.save_label.config(bg=self.CLR_SAVE_DIRTY,
                                   text="●  SAVE CHANGES", fg="#ffffff")
            self.save_bar.bind("<Button-1>", lambda e: self._save())
            self.save_label.bind("<Button-1>", lambda e: self._save())
        else:
            self.save_bar.config(bg=self.CLR_SAVE_IDLE, cursor="arrow")
            self.save_label.config(bg=self.CLR_SAVE_IDLE, text="", fg="#ffffff")
            self.save_bar.unbind("<Button-1>")
            self.save_label.unbind("<Button-1>")

    # ------------------------------------------------------------------
    # Save — routed through airport_resolver's validated write path
    # ------------------------------------------------------------------
    def _save(self):
        if not self.selected_code or not self._save_bar_active:
            return
        code = self.selected_code
        new_name = self.field_vars["name"].get().strip()
        new_city = self.field_vars["city"].get().strip()

        try:
            airport_resolver.update_airport_entry(code, new_name, new_city)
        except airport_resolver.LookupUpdateError as e:
            messagebox.showerror("Save failed", str(e), parent=self.root)
            return

        self.original_values = {"name": new_name, "city": new_city}
        self._refresh_list()   # name/city may have changed → re-sort/re-render
        self._flash_saved()

    def _flash_saved(self):
        self.save_bar.config(bg=self.CLR_SAVE_DIRTY, cursor="arrow")
        self.save_label.config(bg=self.CLR_SAVE_DIRTY, text="✓  SAVED", fg="#ffffff")
        self.save_bar.unbind("<Button-1>")
        self.save_label.unbind("<Button-1>")
        self.root.after(1100, lambda: self._set_save_bar(dirty=False))

    # ------------------------------------------------------------------
    def run(self):
        if isinstance(self.root, tk.Tk):
            self.root.mainloop()


if __name__ == "__main__":
    AirportManagerGUI().run()