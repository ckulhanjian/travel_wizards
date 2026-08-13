"""
airport_resolver.py - Resolve unknown airports by prompting the user
and saving new entries to a persistent overrides file.

IMPORTANT: this no longer edits airport_lookup.py's source code. In a
packaged/frozen build, the .py file that's actually imported typically
lives in a temp extraction folder (e.g. Windows'
...\\Temp\\_MEIxxxxxx\\airport_lookup.py) that's recreated fresh on every
launch and deleted afterward — writes there can appear to succeed and
even take effect for the rest of that one run, but vanish the moment the
app is closed and reopened. Since "the user will only have access to the
executable" (no access to a real, persistent copy of the source), that
approach can never produce a durable database no matter how carefully the
write itself is done.

Instead, every add/update/rename here goes through
airport_lookup.load_overrides() / save_overrides() / reload_overrides(),
which read and write a small JSON file in a proper per-user, persistent,
always-writable data directory (see airport_lookup.py's "Persistent,
writable overrides layer" section) — completely separate from wherever
the app's own bundled code happens to be extracted to.
"""

import os
import sys
import re

import airport_lookup


class LookupUpdateError(Exception):
    """Raised when a new/updated airport entry can't be saved."""
    pass


def _extract_code(raw: str):
    if not raw:
        return None
    tail = raw.strip().split("/")[-1].strip().upper()
    return tail if re.fullmatch(r"[A-Z]{3}", tail) else None


def _fallback_city(raw: str) -> str:
    if not raw:
        return ""
    city_part = raw.strip().split("/")[0]
    return " ".join(w.capitalize() for w in city_part.lower().split())


def check_unknown_airports(data: dict) -> list:
    """
    Scan parsed invoice data for airports not in the lookup.
    Returns list of unknown truncated city names (deduplicated).
    """
    unknown = []
    seen = set()

    for fl in data.get("flights", []):
        for city in [fl.get("departure_city", ""), fl.get("arrival_city", "")]:
            key = city.strip().upper()
            if not key or key in seen:
                continue
            seen.add(key)
            if airport_lookup.lookup_airport(city) is None:
                unknown.append(city)

    return unknown


def add_airport(iata_code: str, airport_name: str, city: str, truncated_name: str = None) -> bool:
    """
    Add (or update) an airport in the persistent overrides file, and make
    it live in this process immediately. If truncated_name is given, also
    saves it as an alias pointing at this code.
    """
    code = iata_code.strip().upper()
    name = airport_name.strip()
    city = city.strip()

    if not re.fullmatch(r"[A-Z]{3}", code):
        raise LookupUpdateError(f"'{iata_code}' isn't a valid 3-letter IATA code.")
    if not name or not city:
        raise LookupUpdateError("Airport Name and City can't be empty.")

    ov = airport_lookup.load_overrides()
    ov["iata_updates"][code] = {"name": name, "city": city}
    if code in ov["iata_removed"]:
        ov["iata_removed"].remove(code)

    if truncated_name:
        trunc_upper = truncated_name.strip().upper()
        ov["truncated_updates"][trunc_upper] = code
        if trunc_upper in ov["truncated_removed"]:
            ov["truncated_removed"].remove(trunc_upper)

    if not airport_lookup.save_overrides(ov):
        raise LookupUpdateError(
            f"Could not write to {airport_lookup.overrides_path()}. "
            "Check that the folder is writable.")

    airport_lookup.reload_overrides()

    if airport_lookup.IATA.get(code) != (name, city):
        raise LookupUpdateError(
            "Saved, but the running app still doesn't show the change. "
            f"Check {airport_lookup.overrides_path()} directly.")

    return True


def link_alias(iata_code: str, truncated_name: str) -> bool:
    """
    Point an additional raw-text string at an EXISTING airport, without
    touching that airport's own name/city. Use this — instead of
    add_airport — when the same real-world airport just shows up under a
    different raw string in another invoice (e.g. an invoice's "NYC/
    KENNEDY" and another's "NEW YORK/JOHN F KENNEDY" both meaning JFK):
    every airport can have any number of these strings pointing at it.
    """
    code = iata_code.strip().upper()
    if code not in airport_lookup.IATA:
        raise LookupUpdateError(
            f"'{code}' isn't a known airport in the database.")

    trunc_upper = truncated_name.strip().upper()
    if not trunc_upper:
        raise LookupUpdateError("Nothing to link — the original text was empty.")

    ov = airport_lookup.load_overrides()
    ov["truncated_updates"][trunc_upper] = code
    if trunc_upper in ov["truncated_removed"]:
        ov["truncated_removed"].remove(trunc_upper)

    if not airport_lookup.save_overrides(ov):
        raise LookupUpdateError(
            f"Could not write to {airport_lookup.overrides_path()}. "
            "Check that the folder is writable.")

    airport_lookup.reload_overrides()

    if airport_lookup.lookup_airport(truncated_name) is None:
        raise LookupUpdateError(
            "Saved, but the running app still doesn't show the change. "
            f"Check {airport_lookup.overrides_path()} directly.")

    return True


def update_airport_entry(iata_code: str, airport_name: str, city: str, new_code: str = None) -> bool:
    """
    Update the Airport Name / City for an EXISTING IATA entry, and
    optionally rename its code (used by the airport database editor
    screen). A rename also repoints any TRUNCATED aliases that pointed at
    the old code, and — if the old code was one of the bundled defaults —
    records it as removed so it stops showing up under the old code.

    Raises LookupUpdateError if the old code doesn't currently resolve,
    the new code is already used by a different airport, or the save
    can't be verified to have taken effect.
    """
    old_code = iata_code.strip().upper()
    new_code_upper = (new_code or iata_code).strip().upper()
    name = airport_name.strip()
    city = city.strip()

    if not re.fullmatch(r"[A-Z]{3}", old_code):
        raise LookupUpdateError(f"'{iata_code}' isn't a valid 3-letter IATA code.")
    if not re.fullmatch(r"[A-Z]{3}", new_code_upper):
        raise LookupUpdateError(f"'{new_code}' isn't a valid 3-letter IATA code.")
    if not name or not city:
        raise LookupUpdateError("Airport Name and City can't be empty.")
    if old_code not in airport_lookup.IATA:
        raise LookupUpdateError(
            f"'{old_code}' isn't an existing entry — this screen only "
            "edits airports that are already there.")

    is_rename = new_code_upper != old_code
    if is_rename and new_code_upper in airport_lookup.IATA:
        existing_name, existing_city = airport_lookup.IATA[new_code_upper]
        raise LookupUpdateError(
            f"'{new_code_upper}' is already used by {existing_name}, "
            f"{existing_city}. Choose a different code, or edit that "
            "entry instead.")

    ov = airport_lookup.load_overrides()

    ov["iata_updates"][new_code_upper] = {"name": name, "city": city}
    if new_code_upper in ov["iata_removed"]:
        ov["iata_removed"].remove(new_code_upper)

    if is_rename:
        # Retire the old code: drop any override for it, and — if it was
        # one of the bundled defaults — mark it removed so the merged
        # view stops showing it.
        ov["iata_updates"].pop(old_code, None)
        if old_code in airport_lookup._BUILTIN_IATA and old_code not in ov["iata_removed"]:
            ov["iata_removed"].append(old_code)

        # Repoint any alias (built-in or override) that pointed at the
        # old code, so old lookups keep working instead of going stale.
        for key, val in list(airport_lookup._BUILTIN_TRUNCATED.items()):
            if val == old_code and key not in ov["truncated_updates"]:
                ov["truncated_updates"][key] = new_code_upper
        for key, val in list(ov["truncated_updates"].items()):
            if val == old_code:
                ov["truncated_updates"][key] = new_code_upper

    if not airport_lookup.save_overrides(ov):
        raise LookupUpdateError(
            f"Could not write to {airport_lookup.overrides_path()}. "
            "Check that the folder is writable.")

    airport_lookup.reload_overrides()

    if airport_lookup.IATA.get(new_code_upper) != (name, city):
        raise LookupUpdateError(
            "Saved, but the running app still doesn't show the change. "
            f"Check {airport_lookup.overrides_path()} directly.")
    if is_rename and old_code in airport_lookup.IATA:
        raise LookupUpdateError(
            f"Saved '{new_code_upper}', but '{old_code}' is still showing "
            "up too — please report this.")

    return True


def prompt_and_save(truncated_name: str, parent=None, source_pdf=None) -> tuple:
    """
    Show a tkinter dialog for one unknown airport with two ways to
    resolve it:
      - Search existing airports and Link this text to one of them (no
        new airport created — just another string pointing at an
        airport that's already in the database).
      - Or, if it's genuinely not there yet, fill in IATA Code / Airport
        Name / City and Add to Lookup to create a new entry.
      - Skip: saves nothing; returns airport_lookup.resolve_city(truncated_name)
        — exactly the fallback the invoice would already show — "the
        existing name listed" — never None.

    Returns (display_string, added_new_airport). added_new_airport is True
    only for "Add to Lookup" creating a genuinely new entry — False for
    Skip and for "Link to Selected Airport" (which points at an airport
    that already existed). Callers that want to report "N airports added"
    for a batch should sum this flag, not just count non-skip results.
    """
    from tkinter import (Toplevel, Label, Entry, Button, StringVar, Frame,
                          Listbox, messagebox)

    result = {"display": None, "added": False}

    def _do_link():
        sel = search_list.curselection()
        if not sel:
            return
        code = search_results[sel[0]][0]
        try:
            link_alias(code, truncated_name)
        except LookupUpdateError as e:
            messagebox.showerror("Couldn't save", str(e), parent=dialog)
            return
        info = airport_lookup.lookup_airport(truncated_name)
        result["display"] = info["display"] if info else code
        result["added"] = False  # pointed at an EXISTING airport, not a new one
        dialog.destroy()

    def _on_search_write(*_a):
        query = search_var.get()
        search_results.clear()
        search_list.delete(0, "end")
        if query.strip():
            for code, name, city in airport_lookup.search_airports(query):
                search_results.append((code, name, city))
                search_list.insert("end", f"{name}  ({code}) — {city}")
        link_btn.config(state="disabled")

    def _on_search_select(_event=None):
        link_btn.config(state="normal" if search_list.curselection() else "disabled")

    def _submit():
        code = iata_var.get().strip().upper()
        name = name_var.get().strip()
        city = city_var.get().strip()

        if not (code and name and city):
            messagebox.showerror(
                "Missing info",
                "IATA Code, Airport Name, and City are all required to add "
                "a new entry. Leave them blank and click Skip instead if "
                "you don't want to add this one.",
                parent=dialog)
            return

        try:
            add_airport(code, name, city, truncated_name=truncated_name)
        except LookupUpdateError as e:
            messagebox.showerror("Couldn't save", str(e), parent=dialog)
            return

        info = airport_lookup.lookup_airport(truncated_name) or airport_lookup.lookup_airport(code)
        result["display"] = info["display"] if info else f"{name}, {city} ({code})"
        result["added"] = True
        dialog.destroy()

    def _skip():
        result["display"] = airport_lookup.resolve_city(truncated_name)
        result["added"] = False
        dialog.destroy()

    if source_pdf:
        try:
            import subprocess
            if sys.platform == "win32":
                os.startfile(source_pdf)
            elif sys.platform == "darwin":
                subprocess.Popen(["open", source_pdf])
            else:
                subprocess.Popen(["xdg-open", source_pdf])
        except Exception:
            pass

    dialog = Toplevel(parent)
    dialog.title("New Airport")
    dialog.resizable(False, False)
    dialog.configure(bg="#ffffff")
    dialog.grab_set()

    if parent:
        dialog.transient(parent)

    pad = {"padx": 16}

    Label(dialog, text="Unknown airport found in invoice (original PDF opened for reference):",
          font=("Arial", 9), bg="#ffffff", fg="#555555").pack(**pad, anchor="w", pady=(14, 0))

    Label(dialog, text=truncated_name,
          font=("Consolas", 12, "bold"), bg="#ffffff", fg="#000000").pack(**pad, anchor="w", pady=2)

    # ── Search existing airports ────────────────────────────────
    Label(dialog, text="Search existing airports (e.g. JFK or LGA) — an "
                        "airport can have several names pointing at it:",
          font=("Arial", 9), bg="#ffffff", fg="#555555",
          wraplength=420, justify="left").pack(**pad, anchor="w", pady=(14, 4))

    search_var = StringVar()
    search_var.trace_add("write", _on_search_write)
    search_entry = Entry(dialog, textvariable=search_var, font=("Arial", 10),
                         relief="solid", bd=1)
    search_entry.pack(fill="x", padx=16)
    search_entry.focus_set()

    search_results = []
    search_list = Listbox(dialog, height=4, font=("Consolas", 10),
                          relief="solid", bd=1, activestyle="none",
                          exportselection=False)
    search_list.pack(fill="x", padx=16, pady=(4, 4))
    search_list.bind("<<ListboxSelect>>", _on_search_select)
    search_list.bind("<Double-Button-1>", lambda e: _do_link())

    link_btn = Button(dialog, text="Link to Selected Airport", command=_do_link,
                      font=("Arial", 10, "bold"), bg="#2e8b46", fg="#ffffff",
                      activebackground="#256e38", activeforeground="#ffffff",
                      relief="flat", padx=16, pady=6, cursor="hand2",
                      state="disabled")
    link_btn.pack(padx=16, pady=(0, 6), anchor="w")

    sep = Frame(dialog, bg="#dddddd", height=1)
    sep.pack(fill="x", padx=16, pady=(8, 10))

    # ── Or add a brand-new airport ───────────────────────────────
    Label(dialog, text="Not there? Add it as a new airport:",
          font=("Arial", 9), bg="#ffffff", fg="#555555").pack(**pad, anchor="w", pady=(0, 4))

    fields = Frame(dialog, bg="#ffffff")
    fields.pack(fill="x", padx=16, pady=4)

    iata_var = StringVar()
    name_var = StringVar()
    city_var = StringVar()

    for row, (label, var, hint) in enumerate([
        ("IATA Code:", iata_var, "e.g. SFO"),
        ("Airport Name:", name_var, "e.g. San Francisco Intl"),
        ("City:", city_var, "e.g. San Francisco"),
    ]):
        Label(fields, text=label, font=("Arial", 9, "bold"),
              bg="#ffffff", fg="#333333").grid(row=row, column=0, sticky="e", pady=3, padx=(0, 8))
        e = Entry(fields, textvariable=var, font=("Arial", 10), width=30,
                  relief="solid", bd=1)
        e.grid(row=row, column=1, sticky="w", pady=3)
        Label(fields, text=hint, font=("Arial", 8),
              bg="#ffffff", fg="#aaaaaa").grid(row=row, column=2, sticky="w", padx=(6, 0))

    btn_row = Frame(dialog, bg="#ffffff")
    btn_row.pack(pady=(10, 14))

    Button(btn_row, text="Add to Lookup", command=_submit,
           font=("Arial", 10, "bold"), bg="#e0e0e0", fg="#000000",
           relief="flat", padx=16, pady=6, cursor="hand2").pack(side="left", padx=6)

    Button(btn_row, text="Skip", command=_skip,
           font=("Arial", 10), bg="#ffffff", fg="#888888",
           relief="flat", padx=16, pady=6, cursor="hand2").pack(side="left", padx=6)

    dialog.protocol("WM_DELETE_WINDOW", _skip)

    dialog.wait_window()
    return result["display"], result["added"]