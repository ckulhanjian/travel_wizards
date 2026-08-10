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

    def __init__(self, parent=None, container=None, on_back=None):
        """
        Two modes:
          - container given: EMBEDDED. Builds into that existing Frame —
            no new window, no Toplevel. on_back(), if given, is called
            when the user clicks "← Back" (the caller is responsible for
            hiding this screen and restoring whatever was there before).
          - container not given: STANDALONE (original behavior) — opens
            its own Tk() root or Toplevel(parent), with a "⌂ Home" button
            that just destroys the window. Kept for `python3
            airport_manager.py` direct use.
        """
        self.embedded = container is not None
        self.on_back = on_back

        if self.embedded:
            self.container = container
            self.root = container.winfo_toplevel()
        elif parent is None:
            self.root = tk.Tk()
            self.container = self.root
        else:
            self.root = tk.Toplevel(parent)
            self.container = self.root

        if not self.embedded:
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
        self._search_after_id = None
        self.airports = {}

        self._draw_logo()
        self._build_ui()
        # Build a local cache of airport records once on open to avoid
        # re-creating large dicts on every keystroke or refresh.
        self._refresh_cache()
        self._refresh_list()

        if not self.embedded:
            # Pick up entries added elsewhere (e.g. via the "unknown
            # airport" prompt during processing) as soon as this window
            # regains focus. Only bound in standalone mode — in embedded
            # mode this shares a root with other screens, and binding
            # FocusIn there would silently clobber their own bindings.
            self.root.bind("<FocusIn>", lambda e: self._refresh_list())

    # ------------------------------------------------------------------
    def _draw_logo(self):
        c = tk.Canvas(self.container, height=64, bg=self.CLR_BG,
                      highlightthickness=0, bd=0)
        c.pack(fill="x", padx=0, pady=0)
        c.create_line(0, 63, 2000, 63, fill=self.CLR_BORDER, width=1)
        c.create_text(500, 22, text="TRAVEL  WIZARDS",
                      font=("Georgia", 20, "bold"),
                      fill=self.CLR_ACCENT, anchor="center")
        c.create_text(500, 44, text="A I R P O R T   D A T A B A S E",
                      font=("Arial", 9), fill="#cc0000", anchor="center")
        if self.embedded:
            tk.Button(self.container, text="←  Back", command=self._handle_back,
                      relief="flat", cursor="hand2",
                      bg=self.CLR_BG, fg="#000000",
                      activebackground=self.CLR_BG,
                      font=("Arial", 9, "bold"),
                      padx=10, pady=4, bd=0).place(x=10, y=8)
        elif isinstance(self.root, tk.Toplevel):
            tk.Button(self.container, text="⌂  Home", command=self.root.destroy,
                      relief="flat", cursor="hand2",
                      bg=self.CLR_BG, fg="#000000",
                      activebackground=self.CLR_BG,
                      font=("Arial", 9, "bold"),
                      padx=10, pady=4, bd=0).place(x=10, y=8)

    def _handle_back(self):
        if self._is_dirty():
            if not messagebox.askyesno(
                    "Unsaved changes",
                    "You have unsaved changes to this airport.\n"
                    "Discard them and go back?", parent=self.root):
                return
        # Prefer caller-provided handler, but be tolerant if it doesn't
        # remove/hide the embedded container (some callers might not).
        if self.on_back:
            try:
                self.on_back()
            except Exception:
                # If the caller's handler failed, try to at least clean up
                # the embedded frame so the UI doesn't get stuck.
                try:
                    if hasattr(self, "container") and self.container.winfo_exists():
                        self.container.destroy()
                except Exception:
                    pass
            else:
                # If handler returned successfully but the container still
                # exists (caller didn't remove it), attempt to destroy it.
                try:
                    if hasattr(self, "container") and self.container.winfo_exists():
                        self.container.destroy()
                except Exception:
                    pass
            return

    def _build_ui(self):
        search_row = tk.Frame(self.container, bg=self.CLR_BG)
        search_row.pack(fill="x", padx=20, pady=(14, 8))
        tk.Label(search_row, text="SEARCH", font=("Arial", 9, "bold"),
                 bg=self.CLR_BG, fg=self.CLR_MUTED).pack(side="left")
        tk.Label(search_row, text="(IATA code, airport name, or city)",
                 font=("Arial", 8), bg=self.CLR_BG,
                 fg="#999999").pack(side="left", padx=(6, 0))

        search_box = tk.Frame(self.container, bg=self.CLR_PANEL,
                              highlightbackground=self.CLR_BORDER,
                              highlightthickness=1)
        search_box.pack(fill="x", padx=20, pady=(0, 12))
        self.search_var = tk.StringVar()
        # Debounce search updates so typing doesn't force immediate full
        # list rebuild on every keystroke.
        self.search_var.trace_add("write", lambda *a: self._on_search_change())
        entry = tk.Entry(search_box, textvariable=self.search_var,
                         relief="flat", bg=self.CLR_PANEL, fg=self.CLR_TEXT,
                         insertbackground=self.CLR_TEXT,
                         font=("Consolas", 12), bd=0)
        entry.pack(side="left", fill="x", expand=True, padx=10, pady=8)
        entry.focus_set()

        body = tk.Frame(self.container, bg=self.CLR_BG)
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
        # Loading indicator shown while the list/cache is being refreshed
        self.loading_label = tk.Label(header, text="", font=("Arial", 9, "italic"),
                                      bg=self.CLR_PANEL, fg="#999999")
        self.loading_label.pack(side="right", padx=(10, 0))
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
        """Return a fresh mapping used to build the cache.

        This is kept separate so callers can refresh the cache on-demand.
        """
        return {code: {"name": name, "city": city}
                for code, (name, city) in airport_lookup.IATA.items()}

    def _refresh_cache(self):
        """Populate `self.airports` from the live lookup module once."""
        # Show loading indicator while we read the live module
        try:
            if hasattr(self, "loading_label"):
                self.loading_label.config(text="Loading…")
                self.root.update_idletasks()
            self.airports = self._current_records()
        except Exception:
            # Defensive: fall back to empty mapping on error
            self.airports = {}
        finally:
            try:
                if hasattr(self, "loading_label"):
                    self.loading_label.config(text="")
            except Exception:
                pass

    def _matching_codes(self):
        # Use cached `self.airports` (refresh via _refresh_cache when needed)
        query = self.search_var.get().strip().lower()
        codes = list(self.airports.keys())
        # Sort once by name (case-insensitive)
        codes.sort(key=lambda c: (self.airports[c].get("name") or "").lower())
        if not query:
            return codes

        def matches(code):
            rec = self.airports[code]
            haystacks = [code, rec.get("name", ""), rec.get("city", "")]
            return any(query in (h or "").lower() for h in haystacks)

        return [c for c in codes if matches(c)]

    def _refresh_list(self):
        if hasattr(self, "loading_label"):
            try:
                self.loading_label.config(text="Loading…")
                self.root.update_idletasks()
            except Exception:
                pass

        codes = self._matching_codes()
        # Build display strings in a list and insert in bulk to avoid
        # the per-item cost of many Tk calls.
        display_items = []
        for code in codes:
            rec = self.airports[code]
            display_items.append(f"{rec.get('name','')}  ({code}) — {rec.get('city','')}")

        self.listbox.delete(0, "end")
        if display_items:
            self.listbox.insert("end", *display_items)
        self._visible_codes = codes
        self.count_label.config(text=f"{len(codes)} of {len(self.airports)}")

        if self.selected_code and self.selected_code in codes:
            idx = codes.index(self.selected_code)
            self.listbox.selection_set(idx)
            self.listbox.see(idx)

        # If the search is empty, refresh the cache so future opens/refreshes
        # see any edits added elsewhere.
        if not self.search_var.get().strip():
            try:
                self._refresh_cache()
            except Exception:
                pass

        try:
            if hasattr(self, "loading_label"):
                self.loading_label.config(text="")
        except Exception:
            pass

    def _on_search_change(self):
        """Debounced handler for search text changes."""
        if self._search_after_id:
            try:
                self.root.after_cancel(self._search_after_id)
            except Exception:
                pass
        # Schedule refresh after short idle so rapid typing won't thrash UI
        self._search_after_id = self.root.after(150, lambda: self._refresh_list())

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