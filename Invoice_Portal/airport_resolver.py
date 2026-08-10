"""
airport_resolver.py - Resolve unknown airports by prompting the user
and writing new entries directly into airport_lookup.py.

The lookup file grows over time as new airports are encountered.

Fixes vs. the previous version:
  - Skip no longer re-derives its own fallback text. It calls
    airport_lookup.resolve_city() directly, so "the existing name listed"
    is guaranteed to be exactly what the rest of the app (invoice
    generation, later prompts, etc.) would already show for this airport —
    one formula, one place, always in agreement.
  - New entries are generated with repr()/safe escaping instead of raw
    f-string interpolation, so a stray quote or special character typed
    into Airport Name / City / Country can never produce invalid Python.
  - The new file content is validated with ast.parse() BEFORE it's written
    to disk. If it wouldn't be valid Python, nothing is written at all —
    the existing, working lookup file is never put at risk.
  - A reload failure is no longer swallowed silently. If it happens, the
    dialog reports it clearly instead of quietly leaving stale data in
    memory while claiming success.
  - "Add to Lookup" validates its inputs before closing the dialog. If
    something's missing/invalid it shows an error and stays open, instead
    of silently degrading into skip-like behavior with no indication
    anything went wrong.

Because the whole point of this module is that airport_lookup.py's module
globals (IATA / TRUNCATED) get reloaded in place, any code that already
did `from airport_lookup import lookup_airport` earlier in the same run
(e.g. invoice_generator.py, imported once at startup) sees the update
immediately too — Python functions look up module globals by name at call
time, not at definition time, so this doesn't require restarting anything
or re-importing anywhere else. That part already worked correctly; this
fix is about making sure the write it depends on can never silently fail
or corrupt the file it's editing.
"""

import os
import sys
import re
import ast
import json
import importlib


def _lookup_path():
    """
    Path to the airport_lookup.py that's actually imported and live in this
    process — NOT independently guessed from sys.executable/__file__.

    The previous version computed this from os.path.dirname(sys.executable)
    when frozen. That's wrong for a packaged build: PyInstaller extracts
    bundled modules to a temp folder (sys._MEIPASS), not to the folder next
    to the .exe, so that guess pointed at a file that was never the one
    actually imported. Writes silently landed somewhere the running app
    never reads from — reload() would "succeed" on an unrelated (or
    nonexistent) file, and a later invoice in the same batch would look up
    the new airport against the *original*, untouched in-memory data and
    still call it unknown, even though the dialog reported success.

    If airport_lookup is already imported (it will be, since this module
    imports it too), sys.modules gives us its real, guaranteed-correct file
    path directly — no guessing involved.
    """
    mod = sys.modules.get("airport_lookup")
    if mod is not None and getattr(mod, "__file__", None):
        return mod.__file__

    # Not imported yet — fall back to a best-effort guess (dev/first run).
    if getattr(sys, "frozen", False):
        base = os.path.dirname(sys.executable)
    else:
        base = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base, "airport_lookup.py")


class LookupUpdateError(Exception):
    """Raised when a new entry can't be safely added to airport_lookup.py."""
    pass


def _write_and_verify(path: str, new_content: str, verify_fn):
    """
    Shared safety sequence for any write to airport_lookup.py:
      1. Refuse to write anything that isn't valid Python.
      2. Write atomically (temp file + os.replace).
      3. Reload the live module.
      4. Call verify_fn(airport_lookup) and require it to return True —
         confirms the change is actually visible through the real,
         actually-imported module, not just present in some file on disk.
    Raises LookupUpdateError at the first problem; never leaves the
    lookup file in a broken or silently-stale state.
    """
    try:
        ast.parse(new_content)
    except SyntaxError as e:
        raise LookupUpdateError(
            f"This change would break airport_lookup.py ({e}). "
            "Nothing was written — the file is unchanged.")

    tmp_path = path + ".tmp"
    with open(tmp_path, "w", encoding="utf-8") as f:
        f.write(new_content)
    os.replace(tmp_path, path)

    try:
        import airport_lookup
        importlib.reload(airport_lookup)
    except Exception as e:
        raise LookupUpdateError(
            f"Saved to airport_lookup.py, but reloading it failed: {e}. "
            "A restart may be needed for the change to take effect.")

    if not verify_fn(airport_lookup):
        raise LookupUpdateError(
            f"Wrote to {path} and reloaded without error, but the running "
            "app doesn't reflect the change. This usually means that file "
            "isn't actually the one this app imports airport_lookup from — "
            "check for a second copy of airport_lookup.py, or a packaged "
            "build extracting it somewhere unexpected."
        )


def _add_to_lookup_file(iata_code: str, airport_name: str, city: str, truncated_name: str):
    """
    Write a new entry into airport_lookup.py by inserting into the IATA
    and TRUNCATED dicts, then reload the module so it's live immediately.

    Raises LookupUpdateError (instead of silently failing or silently
    "succeeding") if:
      - the inputs are invalid
      - the resulting file wouldn't be valid Python (nothing is written)
      - the reload throws
      - OR — the critical case this function used to get wrong — the
        write+reload both appear to succeed, but the live, actually-
        imported module STILL doesn't resolve the new airport afterward.
        That combination can only mean the write and the import didn't
        actually target the same file, so it's treated as a hard failure
        rather than a false "added" confirmation.
    """
    path = _lookup_path()
    if not os.path.exists(path):
        raise LookupUpdateError(f"airport_lookup.py not found at {path}")

    iata_upper = iata_code.strip().upper()
    trunc_upper = truncated_name.strip().upper()

    if not re.fullmatch(r"[A-Z]{3}", iata_upper):
        raise LookupUpdateError(f"'{iata_code}' isn't a valid 3-letter IATA code.")
    if not airport_name.strip() or not city.strip():
        raise LookupUpdateError("Airport Name and City can't be empty.")

    with open(path, "r", encoding="utf-8") as f:
        content = f.read()

    # json.dumps always produces a double-quoted string literal — matching
    # the rest of the file's style — with any special characters (quotes,
    # backslashes, unicode) safely escaped. Unlike manually wrapping user
    # input in f'"{...}"', this can't produce broken syntax no matter what
    # the user types, and unlike repr() it won't switch to single quotes
    # (which would look inconsistent and break this module's own "does
    # this code already exist" double-quote-based text search next time).
    iata_entry = f"    {json.dumps(iata_upper)}: ({json.dumps(airport_name.strip())}, {json.dumps(city.strip())}),\n"
    truncated_entry = f"    {json.dumps(trunc_upper)}: {json.dumps(iata_upper)},\n"

    # Already present? Nothing to do.
    if f'"{iata_upper}":' in content and f'"{trunc_upper}":' in content:
        return True

    if f'"{iata_upper}":' not in content:
        lines = content.split("\n")
        new_lines = []
        in_iata = False
        inserted_iata = False

        for line in lines:
            if "IATA = {" in line or (not in_iata and re.match(r'^IATA\s*=\s*\{', line)):
                in_iata = True
                new_lines.append(line)
                continue

            if in_iata:
                if line.strip() == "}":
                    if not inserted_iata:
                        new_lines.append(iata_entry.rstrip())
                        inserted_iata = True
                    in_iata = False
                    new_lines.append(line)
                    continue

                code_match = re.match(r'\s*"([A-Z]{3})":', line)
                if code_match and not inserted_iata:
                    existing_code = code_match.group(1)
                    if iata_upper < existing_code:
                        new_lines.append(iata_entry.rstrip())
                        inserted_iata = True

            new_lines.append(line)

        content = "\n".join(new_lines)

    if f'"{trunc_upper}":' not in content:
        lines = content.split("\n")
        new_lines = []
        in_trunc = False
        inserted_trunc = False

        for line in lines:
            if "TRUNCATED = {" in line or re.match(r'^TRUNCATED\s*=\s*\{', line):
                in_trunc = True
                new_lines.append(line)
                continue

            if in_trunc:
                if line.strip() == "}":
                    if not inserted_trunc:
                        new_lines.append(truncated_entry.rstrip())
                        inserted_trunc = True
                    in_trunc = False
                    new_lines.append(line)
                    continue

                name_match = re.match(r'\s*"([^"]+)":', line)
                if name_match and not inserted_trunc:
                    existing_name = name_match.group(1)
                    if trunc_upper < existing_name:
                        new_lines.append(truncated_entry.rstrip())
                        inserted_trunc = True

            new_lines.append(line)

        content = "\n".join(new_lines)

    _write_and_verify(
        path, content,
        verify_fn=lambda mod: mod.lookup_airport(trunc_upper) is not None)
    return True


def update_airport_entry(iata_code: str, airport_name: str, city: str):
    """
    Update the Airport Name / City for an EXISTING IATA entry (used by the
    airport database editor screen — this never adds a new code, only
    edits values already present, same as the manager UI promises).

    Raises LookupUpdateError if the code doesn't already exist, the inputs
    are invalid, or (as in _add_to_lookup_file) the write can't be safely
    verified to have taken effect in the live, actually-imported module.
    """
    path = _lookup_path()
    if not os.path.exists(path):
        raise LookupUpdateError(f"airport_lookup.py not found at {path}")

    iata_upper = iata_code.strip().upper()
    name = airport_name.strip()
    city = city.strip()

    if not re.fullmatch(r"[A-Z]{3}", iata_upper):
        raise LookupUpdateError(f"'{iata_code}' isn't a valid 3-letter IATA code.")
    if not name or not city:
        raise LookupUpdateError("Airport Name and City can't be empty.")

    with open(path, "r", encoding="utf-8") as f:
        content = f.read()

    pattern = re.compile(
        r'^(\s*"' + re.escape(iata_upper) + r'"\s*:\s*)\("[^"]*",\s*"[^"]*"\)(,?\s*)$',
        re.MULTILINE)
    if not pattern.search(content):
        raise LookupUpdateError(
            f"'{iata_upper}' isn't an existing entry in airport_lookup.py "
            "— this screen only edits airports that are already there.")

    replacement = rf'\g<1>({json.dumps(name)}, {json.dumps(city)})\g<2>'
    new_content = pattern.sub(replacement, content, count=1)

    _write_and_verify(
        path, new_content,
        verify_fn=lambda mod: mod.IATA.get(iata_upper) == (name, city))
    return True




def check_unknown_airports(data: dict) -> list:
    """
    Scan parsed invoice data for airports not in the lookup.
    Returns list of unknown truncated city names (deduplicated).
    """
    try:
        from airport_lookup import lookup_airport
    except ImportError:
        return []

    unknown = []
    seen = set()

    for fl in data.get("flights", []):
        for city in [fl.get("departure_city", ""), fl.get("arrival_city", "")]:
            key = city.strip().upper()
            if not key or key in seen:
                continue
            seen.add(key)
            if lookup_airport(city) is None:
                unknown.append(city)

    return unknown


def prompt_and_save(truncated_name: str, parent=None, source_pdf=None) -> str:
    """
    Show a tkinter dialog asking for IATA code, airport name, and city.
    Writes the new entry into airport_lookup.py.

    - Skip: saves nothing; returns airport_lookup.resolve_city(truncated_name),
      i.e. exactly the fallback name the invoice would already show — "the
      existing name listed" — never None, never a diverging one-off format.
    - Add to Lookup: validates first; on success returns the freshly
      resolved display string (re-fetched from the reloaded module, so it's
      guaranteed to match what lookup_airport() itself would now return).
      On failure, shows the error and keeps the dialog open so the user can
      fix the input or fall back to Skip themselves.
    """
    from tkinter import Toplevel, Label, Entry, Button, StringVar, Frame, messagebox
    from airport_lookup import resolve_city, lookup_airport

    result = {"display": None}

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
            _add_to_lookup_file(code, name, city, truncated_name)
        except LookupUpdateError as e:
            messagebox.showerror("Couldn't save", str(e), parent=dialog)
            return

        # Re-fetch from the just-reloaded module rather than reconstructing
        # the string by hand, so it's always in lockstep with lookup_airport().
        info = lookup_airport(truncated_name) or lookup_airport(code)
        result["display"] = info["display"] if info else f"{name}, {city} ({code})"
        dialog.destroy()

    def _skip():
        # Nothing is written. Whatever the rest of the app would already
        # display for this unresolved airport is exactly what we return.
        result["display"] = resolve_city(truncated_name)
        dialog.destroy()

    # Open the source PDF so the user can look up the IATA code
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

    Label(dialog, text="Enter the airport details to add to the lookup:",
          font=("Arial", 9), bg="#ffffff", fg="#555555").pack(**pad, anchor="w", pady=(10, 4))

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

    # Closing the dialog via the window's own close button is treated the
    # same as Skip, so a caller blocked on the result never hangs forever.
    dialog.protocol("WM_DELETE_WINDOW", _skip)

    dialog.wait_window()
    return result["display"]