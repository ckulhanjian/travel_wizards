"""
airport_resolver.py - Resolve unknown airports by prompting the user
and writing new entries directly into airport_lookup.py.

The lookup file grows over time as new airports are encountered.
"""

import os
import sys
import re


def _lookup_path():
    """Path to airport_lookup.py — same folder as this script/exe."""
    if getattr(sys, "frozen", False):
        base = os.path.dirname(sys.executable)
    else:
        base = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base, "airport_lookup.py")


def _add_to_lookup_file(iata_code: str, airport_name: str, city: str, truncated_name: str):
    """
    Write a new entry into airport_lookup.py by inserting into
    the IATA and TRUNCATED dicts.
    """
    path = _lookup_path()
    if not os.path.exists(path):
        return False

    with open(path, "r", encoding="utf-8") as f:
        content = f.read()

    iata_upper = iata_code.strip().upper()
    trunc_upper = truncated_name.strip().upper()

    # Insert into IATA dict — find the closing brace of IATA = {
    # Add before the last entry's line that precedes the closing }
    iata_entry = f'    "{iata_upper}": ("{airport_name}", "{city}"),\n'
    truncated_entry = f'    "{trunc_upper}": "{iata_upper}",\n'

    # Check if already exists
    if f'"{iata_upper}":' in content and f'"{trunc_upper}":' in content:
        return True  # already there

    # Find IATA dict closing — look for the pattern: lone "}" after IATA entries
    # Strategy: find "IATA = {" then find its closing "}"
    # We insert the new entry right before the closing }

    if f'"{iata_upper}":' not in content:
        # Find a good insertion point in IATA dict — insert alphabetically
        # Find the line with the next alphabetical IATA code after ours
        iata_lines = []
        in_iata = False
        iata_end_idx = None

        lines = content.split("\n")
        new_lines = []
        inserted_iata = False

        for i, line in enumerate(lines):
            if "IATA = {" in line or (not in_iata and re.match(r'^IATA\s*=\s*\{', line)):
                in_iata = True
                new_lines.append(line)
                continue

            if in_iata:
                # Check if this is the closing brace
                if line.strip() == "}":
                    if not inserted_iata:
                        new_lines.append(iata_entry.rstrip())
                        inserted_iata = True
                    in_iata = False
                    new_lines.append(line)
                    continue

                # Check if we should insert before this line (alphabetical)
                code_match = re.match(r'\s*"([A-Z]{3})":', line)
                if code_match and not inserted_iata:
                    existing_code = code_match.group(1)
                    if iata_upper < existing_code:
                        new_lines.append(iata_entry.rstrip())
                        inserted_iata = True

            new_lines.append(line)

        content = "\n".join(new_lines)

    if f'"{trunc_upper}":' not in content:
        # Insert into TRUNCATED dict
        lines = content.split("\n")
        new_lines = []
        in_trunc = False
        inserted_trunc = False

        for i, line in enumerate(lines):
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

                # Insert alphabetically
                name_match = re.match(r'\s*"([^"]+)":', line)
                if name_match and not inserted_trunc:
                    existing_name = name_match.group(1)
                    if trunc_upper < existing_name:
                        new_lines.append(truncated_entry.rstrip())
                        inserted_trunc = True

            new_lines.append(line)

        content = "\n".join(new_lines)

    with open(path, "w", encoding="utf-8") as f:
        f.write(content)

    # Reload the module so the new entry is available immediately
    try:
        import airport_lookup
        import importlib
        importlib.reload(airport_lookup)
    except Exception:
        pass

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
    Returns the display string, or a title-cased fallback.
    """
    from tkinter import Toplevel, Label, Entry, Button, StringVar, Frame

    result = {"display": None}

    def _submit():
        code = iata_var.get().strip().upper()
        name = name_var.get().strip()
        city = city_var.get().strip()
        if code and name and city:
            _add_to_lookup_file(code, name, city, truncated_name)
            result["display"] = f"{name}, {city} ({code})"
        else:
            # Use whatever they typed, or fallback
            result["display"] = name or city or " ".join(
                w.capitalize() for w in truncated_name.lower().split("/")[0].split()
            )
        dialog.destroy()

    def _skip():
        result["display"] = " ".join(
            w.capitalize() for w in truncated_name.lower().split("/")[0].split()
        )
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

    # Center on parent
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

    dialog.wait_window()
    return result["display"]