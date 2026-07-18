"""
state_parser.py — State machine parser for Travel Wizards invoices.

Walks every line in order, tracks which section it's in,
extracts data based on context, and validates must-have fields.

Handles both TIPITIN and ITIN header formats, plus all content types:
  FLIGHTS, HOTELS, CRUISES, TOURS, PACKAGES, INSURANCE,
  TICKETS, FINANCIAL, BAGGAGE, NOTICES
"""

import re
import fitz

# ── States ─────────────────────────────────────────────────────
HEADER       = "HEADER"
PASSENGERS   = "PASSENGERS"
FLIGHT       = "FLIGHT"
HOTEL        = "HOTEL"
CRUISE       = "CRUISE"
TOUR         = "TOUR"
PACKAGE      = "PACKAGE"
INSURANCE    = "INSURANCE"
TICKETS      = "TICKETS"
FINANCIAL    = "FINANCIAL"
BAGGAGE      = "BAGGAGE"
CARRY_ON     = "CARRY_ON"
NOTICES      = "NOTICES"
SKIP         = "SKIP"

# Lines we intentionally skip (no data loss)
SKIP_PATTERNS = [
    r'^\s*$',                           # blank
    r'VIEWTRIP\.TRAVELPORT',            # travelport links
    r'^\s*ADT\s*$',                     # ADT marker
    r'^\s*ELECTRONIC\s*$',              # electronic marker
    r'^\s*-{4,}\s*$',                   # dashed lines
    r'SURFACE TRANSPORTATION',          # surface transport marker
    r'^\s*BFFF\s*$',                    # booking reference
    r'AIR MILEAGE:',                    # mileage (user said skip)
    r'MILEAGE MEMBERSHIP:',            # mileage membership
    r'^\s*PAGE:\s*\d+',                 # page numbers
    r'SF-\d+\s+USD',                    # service fee payment ref line
    r'PLEASE NOTE',                     # generic notice prefix
    r'CANCELLATION FEES APPLY',         # boilerplate cancellation-fee notice — not wanted in output
]

# ── Detection helpers ──────────────────────────────────────────
def _is_date_line(line):
    """Match: '03 DEC 26 - THURSDAY' or '03 DEC 26    - THURSDAY'"""
    m = re.match(r'\s*(\d{2} [A-Z]{3} \d{2})\s+- ([A-Z]+)', line)
    return m.groups() if m else None

def _is_tipitin_airline(line):
    """Match: '  UNITED AIRLINES INC.      362   BUS/ FIRST XCPTNS'"""
    m = re.match(r'\s+(\S.+?(?:INC\.|AIRLINES|AIR LINES).*?)\s+(\d+)\s+(.+?)$', line)
    if m and not any(kw in line for kw in ['HOTEL', 'RESORT', 'GUARD', 'INSURANCE']):
        return m.groups()
    return None

def _is_itin_airline(line):
    """Match: '  AIR   KLM            FLT: 605   COACH CLASS  MEAL'"""
    m = re.match(r'\s+AIR\s+(\S+)\s+FLT:\s*(\d+)\s+(.+?)$', line)
    return m.groups() if m else None

def _is_hotel_line(line):
    """Match: '  WESTIN HOTELS AND RESORTS 02 NT/S - OUT 24JUL CONFIRMED'
       or:    '  HOTEL AVENIDA PALACE 01 NT/S - OUT 02JUN CONFIRMED'
       or:    '  HOTEL INDIGO DENVER DOWN 03 NT/S - OUT 22JUL CONFIRMED'"""
    m = re.match(r'\s*(.+?)\s+(\d+)\s+NT/S\s*-\s*OUT\s+(\S+)\s+(CONFIRMED|WAITLIST)', line)
    return m.groups() if m else None

def _should_skip(line):
    return any(re.search(p, line) for p in SKIP_PATTERNS)

def _detect_format(lines):
    """Detect ITIN vs TIPITIN from first few lines."""
    for line in lines[:5]:
        if re.search(r'SALES PERSON:', line):
            return "ITIN"
        if re.search(r'ITIN NO:', line):
            return "TIPITIN"
    return "UNKNOWN"


# ── Main parser ────────────────────────────────────────────────
def parse(pdf_path: str) -> dict:
    """Parse any Travel Wizards invoice. Returns structured data + validation warnings."""
    doc = fitz.open(pdf_path)
    pages_text = [page.get_text("text") for page in doc]
    doc.close()

    all_lines = []
    for page_text in pages_text:
        all_lines.extend(page_text.split("\n"))

    fmt = _detect_format(all_lines)

    data = {
        "format": fmt,
        "passengers": [],
        "mailing_address": [],
        "booking": {},
        "flights": [],
        "hotels": [],
        "cruises": [],
        "tours": [],
        "packages": [],
        "insurance": [],
        "tickets": [],
        "financial": {},
        "baggage": [],
        "carry_on": [],
        "notices": [],
        "service_fee": None,
        "freq_flyers": [],
        "unrecognized": [],  # lines the parser didn't match
        "warnings": [],      # validation warnings
    }

    state = HEADER
    current_date = None
    current_day = None
    current_flight = None
    current_hotel = None
    current_cruise = None
    current_tour = None
    current_package = None
    header_done = False

    for line_num, raw_line in enumerate(all_lines):
        line = raw_line.rstrip()

        # Skip intentionally ignored lines
        if _should_skip(line):
            continue

        # ── Check for state transitions ───────────────────────

        # Date line — sets current date, doesn't change state yet
        date_match = _is_date_line(line)
        if date_match:
            current_date, current_day = date_match
            header_done = True
            # Check what follows on same line (e.g. "21 JAN 27 - THURSDAY TOUR")
            rest = line[line.find(current_day) + len(current_day):].strip()
            if rest == "TOUR":
                state = TOUR
                current_tour = {"date_raw": current_date, "day_name": current_day,
                                "vendor": None, "amount": None, "confirmation": None,
                                "details": []}
            elif "OTHER ARRANGEMENTS" in rest:
                state = TOUR
                current_tour = {"date_raw": current_date, "day_name": current_day,
                                "vendor": None, "amount": None, "confirmation": None,
                                "details": []}
            continue

        # Cruise arrangements
        if re.match(r'\s*CRUISE ARRANGEMENTS', line):
            state = CRUISE
            current_cruise = {"date_raw": current_date, "day_name": current_day,
                              "details": {}}
            continue

        # Package arrangements
        if re.match(r'\s*PACKAGE ARRANGEMENTS', line):
            state = PACKAGE
            current_package = {"date_raw": current_date, "day_name": current_day,
                               "details": {}}
            continue

        # Standalone TOUR or OTHER ARRANGEMENTS line (after a date line)
        if re.match(r'^\s*TOUR\s*$', line) or re.match(r'^\s*OTHER ARRANGEMENTS\s*$', line):
            # Save previous tour if any
            if current_tour and current_tour.get("vendor"):
                data["tours"].append(current_tour)
            current_tour = {"date_raw": current_date, "day_name": current_day,
                            "vendor": None, "amount": None, "confirmation": None,
                            "details": []}
            state = TOUR
            continue

        # Airline line (TIPITIN format)
        tipitin_air = _is_tipitin_airline(line)
        if tipitin_air:
            if current_flight:
                data["flights"].append(current_flight)
            airline, fnum, cabin = tipitin_air
            current_flight = {
                "date_raw": current_date, "day_name": current_day,
                "airline": airline.strip(), "flight_number": fnum,
                "cabin_class": cabin.strip(), "operated_by": None,
                "departure_city": None, "departure_time": None,
                "arrival_city": None, "arrival_time": None,
                "arrives_next_day": None, "duration": None,
                "nonstop": False, "confirmed": False,
                "baggage_allowance": None, "seats": [],
                "meals": None, "dep_terminal": None, "arr_terminal": None,
                "airline_locator_carrier": None, "airline_locator_code": None,
                "wheelchair": False,
            }
            state = FLIGHT
            header_done = True
            continue

        # Airline line (ITIN format)
        itin_air = _is_itin_airline(line)
        if itin_air:
            if current_flight:
                data["flights"].append(current_flight)
            airline, fnum, cabin_meal = itin_air
            # Split cabin and meal
            parts = re.match(r'(.+?)\s{2,}(\S+)$', cabin_meal)
            cabin = parts.group(1).strip() if parts else cabin_meal
            meals = parts.group(2).strip() if parts else None
            current_flight = {
                "date_raw": current_date, "day_name": current_day,
                "airline": airline.strip(), "flight_number": fnum,
                "cabin_class": cabin, "operated_by": None,
                "departure_city": None, "departure_time": None,
                "arrival_city": None, "arrival_time": None,
                "arrives_next_day": None, "duration": None,
                "nonstop": False, "confirmed": True,
                "baggage_allowance": None, "seats": [],
                "meals": meals, "dep_terminal": None, "arr_terminal": None,
                "airline_locator_carrier": None, "airline_locator_code": None,
                "wheelchair": False,
            }
            state = FLIGHT
            header_done = True
            continue

        # Hotel line
        hotel_match = _is_hotel_line(line)
        if hotel_match:
            if current_hotel:
                data["hotels"].append(current_hotel)
            chain, nights, checkout, status = hotel_match
            current_hotel = {
                "chain": chain.strip(), "nights": int(nights),
                "checkout_date": checkout, "status": status,
                "name": None, "address": None, "city": None,
                "phone": None, "fax": None, "rate_currency": None,
                "rate_amount": None, "confirmation": None,
                "guarantee": None, "approx_total": None,
                "cancel_policy": None, "notes": [], "room_info": None,
            }
            state = HOTEL
            header_done = True
            continue

        # Ticket section
        if re.match(r'\s*TICKET NUMBER/S:', line):
            if current_flight:
                data["flights"].append(current_flight)
                current_flight = None
            state = TICKETS
            continue

        # Baggage
        if re.match(r'BAGGAGE ALLOWANCE', line):
            state = BAGGAGE
            continue

        # Carry on
        if re.match(r'CARRY ON ALLOWANCE', line):
            state = CARRY_ON
            continue

        # Financial markers
        if re.match(r'AIR FARE USD|^\s*SUB TOTAL', line) and state != FINANCIAL:
            state = FINANCIAL
            # Don't continue — process this line below

        # ── Extract data based on current state ───────────────

        if state == HEADER and not header_done:
            # M/M prefix lines are the mailing-address salutation (e.g. "M/M
            # MIGUEL MORENO"), not a passenger. This MUST be checked before the
            # generic passenger regex below — "M/M NAME" also matches the
            # LASTNAME/FIRSTNAME shape ("M" as last name, "/M NAME" as first),
            # which is what previously caused it to be captured as a third
            # passenger.
            if re.match(r'^\s*M/M\s+', line):
                data["mailing_address"].append(line.strip())
                continue

            # TIPITIN header: passengers at top as LASTNAME/FIRSTNAME
            pax = re.match(r'^\s*([A-Z]+)/([A-Z][A-Z ]+)$', line)
            if pax:
                last, first_mid = pax.group(1), pax.group(2).strip()
                parts = first_mid.split()
                data["passengers"].append({
                    "last_name": last,
                    "first_name": parts[0] if parts else "",
                    "middle_name": " ".join(parts[1:]) if len(parts) > 1 else "",
                    "full_slash": f"{last}/{first_mid}",
                })
                continue

            # ITIN header: SALES PERSON line
            sp = re.match(r'SALES PERSON:\s*(\S+)\s+ITIN/INVOICE NO\.\s+(\d+).*?DATE:(.+?)$', line)
            if sp:
                data["booking"]["sales_person"] = sp.group(1)
                data["booking"]["itin_no"] = sp.group(2)
                data["booking"]["date"] = sp.group(3).strip()
                continue

            # CUSTOMER NBR
            cn = re.match(r'CUSTOMER NBR:\s*(\S+)', line)
            if cn:
                data["booking"]["customer_nbr"] = cn.group(1)
                continue

            # TIPITIN booking line
            bk = re.search(r'ITIN NO:\s*(\d+)', line)
            if bk:
                data["booking"]["itin_no"] = bk.group(1)
                rec = re.search(r'RECORD LOCATOR:\s*(\S+)', line)
                if rec:
                    data["booking"]["record_locator"] = rec.group(1)
                dt = re.search(r'DATE:\s*(.+?)$', line)
                if dt:
                    data["booking"]["date"] = dt.group(1).strip()
                header_done = True
                continue

            # TO: address block (ITIN format)
            if re.match(r'\s*TO:\s', line):
                addr = line.replace("TO:", "").strip()
                if addr:
                    data["mailing_address"].append(addr)
                state = HEADER  # stay in header, collecting address
                continue

            # FOR: passenger block (ITIN format)
            pax_for = re.match(r'\s*(?:FOR:\s*)?([A-Z]+)/([A-Z][A-Z ]+)$', line)
            if pax_for and "FOR:" in line or (data["passengers"] and re.match(r'\s+[A-Z]+/[A-Z]', line)):
                last, first_mid = pax_for.group(1), pax_for.group(2).strip()
                parts = first_mid.split()
                data["passengers"].append({
                    "last_name": last,
                    "first_name": parts[0] if parts else "",
                    "middle_name": " ".join(parts[1:]) if len(parts) > 1 else "",
                    "full_slash": f"{last}/{first_mid}",
                })
                continue

            # Mailing address lines (plain text between passengers and booking)
            stripped = line.strip()
            if stripped and not re.match(r'^\s*$', line):
                data["mailing_address"].append(stripped)
                continue

        elif state == FLIGHT and current_flight:
            # Departure
            dep = re.search(r'LV:\s+(.+?)\s{2,}(\d+[AP])', line)
            if dep:
                current_flight["departure_city"] = dep.group(1).strip()
                current_flight["departure_time"] = dep.group(2)
                if "CONFIRMED" in line:
                    current_flight["confirmed"] = True
                if "NON-STOP" in line or "NONSTOP" in line:
                    current_flight["nonstop"] = True
                continue

            # Arrival (TIPITIN: ARR:, ITIN: AR:)
            arr = re.search(r'AR[R]?:\s+(.+?)\s{2,}(\d+[AP])', line)
            if arr:
                current_flight["arrival_city"] = arr.group(1).strip()
                current_flight["arrival_time"] = arr.group(2)
                if "CONFIRMED" in line:
                    current_flight["confirmed"] = True
                if "NON-STOP" in line or "NONSTOP" in line:
                    current_flight["nonstop"] = True
                nxt = re.search(r'ARRIVES-\s*(\d+ [A-Z]+)', line)
                if nxt:
                    current_flight["arrives_next_day"] = nxt.group(1)
                continue

            # Duration (both formats)
            dur = re.search(r'(?:FLIGHT TIME -|ELAPSED TIME-)\s*(.+?)(?:\s{2,}|$)', line)
            if dur:
                current_flight["duration"] = dur.group(1).strip()
                bag = re.search(r'BAGGAGE ALLOWANCE - (\S+)', line)
                if bag:
                    current_flight["baggage_allowance"] = bag.group(1)
                continue

            # Equipment (ITIN)
            eq = re.match(r'\s+EQUIP-(\S+)', line)
            if eq:
                current_flight["equipment"] = eq.group(1)
                dur2 = re.search(r'ELAPSED TIME-\s*(.+?)$', line)
                if dur2:
                    current_flight["duration"] = dur2.group(1).strip()
                continue

            # Operated by
            op = re.search(r'OPERATED BY-(.+)', line)
            if op:
                current_flight["operated_by"] = op.group(1).strip()
                continue

            # Seats (TIPITIN: SEAT ASSIGNED xxx NON SMOKING)
            seat = re.search(r'SEAT ASSIGNED\s+(\S+)', line)
            if seat:
                current_flight["seats"].append(seat.group(1))
                if "CONFIRMED" in line:
                    current_flight["confirmed"] = True
                continue

            # Continuation seat line
            cont_seat = re.match(r'^\s{10,}(\S+)\s+NON SMOKING', line)
            if cont_seat:
                current_flight["seats"].append(cont_seat.group(1))
                continue

            # Continuation seat without NON SMOKING (Forbes: "04F" alone)
            cont_seat2 = re.match(r'^\s{10,}(\d+[A-Z])\s*$', line)
            if cont_seat2:
                current_flight["seats"].append(cont_seat2.group(1))
                continue

            # Seats (ITIN: RESERVED SEATS SEAT- 3E 3F)
            rseat = re.search(r'RESERVED SEATS\s+SEAT-\s*(.+?)$', line)
            if rseat:
                seats_str = rseat.group(1).strip()
                if seats_str:
                    current_flight["seats"].extend(seats_str.split())
                continue
            # RESERVED SEATS alone (no seat numbers)
            if re.match(r'\s+RESERVED SEATS\s*$', line):
                continue

            # Meals
            meal = re.search(r'MEALS SERVED\s+(.+?)$', line)
            if meal:
                current_flight["meals"] = meal.group(1).strip()
                continue

            # Terminals
            dep_t = re.search(r'DEPART - (TERMINAL\s+\S+)', line)
            arr_t = re.search(r'ARRIVE - (TERMINAL\s+\S+)', line)
            if dep_t:
                current_flight["dep_terminal"] = dep_t.group(1).strip()
            if arr_t:
                current_flight["arr_terminal"] = arr_t.group(1).strip()
            if dep_t or arr_t:
                continue

            # Airline locator (TIPITIN)
            loc = re.search(r'AIRLINE LOCATOR:\s*(\S+)\s*-(\S+)', line)
            if loc:
                current_flight["airline_locator_carrier"] = loc.group(1)
                current_flight["airline_locator_code"] = loc.group(2)
                continue

            # Airline confirmation (ITIN)
            conf = re.search(r'AIRLINE CONFIRMATION:\s*(\S+)\s*-(\S+)', line)
            if conf:
                current_flight["airline_locator_carrier"] = conf.group(1)
                current_flight["airline_locator_code"] = conf.group(2)
                continue

            # Frequent flyer
            ff = re.search(r'FREQ FLYER:\s*(.+?)\s+([A-Z]{2})\s{2,}(\S+)', line)
            if ff:
                entry = {"passenger": ff.group(1).strip(),
                         "airline": ff.group(2), "number": ff.group(3)}
                if entry not in data["freq_flyers"]:
                    data["freq_flyers"].append(entry)
                continue

            # Wheelchair
            if re.search(r'WHEELCHAIR', line):
                current_flight["wheelchair"] = True
                continue

            # Service fees
            sf = re.search(r'SERVICE FEES?\s+USD\s+([\d.]+)', line)
            if sf:
                data["service_fee"] = sf.group(1)
                state = NOTICES  # usually followed by notices
                continue

        elif state == HOTEL and current_hotel:
            # Hotel name + guarantee
            nm = re.search(r'^\s*(.+?)\s{2,}GUARANTEE-(.+?)$', line)
            if nm:
                current_hotel["name"] = nm.group(1).strip()
                current_hotel["guarantee"] = nm.group(2).strip()
                continue

            # Address + rate
            addr_rate = re.search(r'^\s+(.+?)\s{2,}RATE-\s*(\S+)\s+([\d.]+)', line)
            if addr_rate:
                current_hotel["address"] = addr_rate.group(1).strip()
                current_hotel["rate_currency"] = addr_rate.group(2)
                current_hotel["rate_amount"] = addr_rate.group(3)
                continue

            # Simple address line (hotels without chain prefix)
            if re.match(r'^\s+\d+\s+', line) and not current_hotel.get("address"):
                current_hotel["address"] = line.strip()
                continue

            # City line
            city = re.match(r'^\s+([A-Z].*?(?:CA|FL|AZ|US|UK|AU))\s', line)
            if city and not current_hotel.get("city"):
                current_hotel["city"] = city.group(1).strip()
                continue

            # Phone/fax
            phone = re.search(r'PHONE NO-(.+?)$', line)
            fax = re.search(r'FAX NO-(.+?)\s{2,}', line)
            if phone:
                current_hotel["phone"] = phone.group(1).strip()
            if fax:
                current_hotel["fax"] = fax.group(1).strip()
            if phone or fax:
                continue

            # Phone number without PHONE NO- prefix (e.g. Kurzrock "351 213 218 100")
            simple_phone = re.match(r'^\s+(\d[\d\s-]+\d)\s{2,}RATE', line)
            if simple_phone:
                current_hotel["phone"] = simple_phone.group(1).strip()
                rate = re.search(r'RATE-\s*(\S+)\s+([\d.]+)', line)
                if rate:
                    current_hotel["rate_currency"] = rate.group(1)
                    current_hotel["rate_amount"] = rate.group(2)
                continue

            # Confirmation
            cf = re.search(r'CONFIRMATION-(\S+)', line)
            if cf:
                current_hotel["confirmation"] = cf.group(1)
                continue

            # Approx total
            ttl = re.search(r'APPROX TTL.*?([\d.]+)(\w+)', line)
            if ttl:
                current_hotel["approx_total"] = f"{ttl.group(1)}{ttl.group(2)}"
                continue

            # Cancel policy / notes
            if re.search(r'CXL:|CANCEL|AAA', line):
                current_hotel["notes"].append(line.strip())
                continue

            # Room info
            if re.match(r'^\s+\d+\s+BC-', line):
                current_hotel["room_info"] = line.strip()
                continue

            # Rate status, guarantee on separate line
            if re.search(r'RATESTATUS', line):
                continue

        elif state == CRUISE and current_cruise:
            # Cruise detail lines
            d = current_cruise["details"]

            # Cruise vendor line: **PRINCESS CRUISES**/CF-MJ8X7W
            vendor = re.search(r'\*\*(.+?)\*\*', line)
            if vendor:
                d["vendor"] = vendor.group(1).strip()
                cf = re.search(r'CF-(\S+)', line)
                if cf:
                    d["confirmation"] = cf.group(1)
                continue

            # Total cost, e.g. "TOTAL COST   USD  3673.90"
            tc = re.search(r'TOTAL COST\s+USD\s+([\d.]+)', line)
            if tc:
                d["total_cost"] = tc.group(1)
                continue

            # Payment already applied, e.g. "30JUN2026 PAYMENT BY VISA  USD  3673.90-"
            pay = re.search(r'(\d{2}[A-Z]{3}(?:\d{2,4})?)\s+PAYMENT BY\s*(.*?)\s+USD\s+([\d.]+)-', line)
            if pay:
                d.setdefault("payments", []).append({
                    "date": pay.group(1), "method": pay.group(2).strip(),
                    "amount": pay.group(3),
                })
                continue

            # Remaining balance, e.g. "BALANCE OF 2899.00 DUE 08JUL2026"
            bal = re.search(r'BALANCE OF\s+([\d.]+)\s+DUE\s+(\S+)', line)
            if bal:
                d["balance_due"] = bal.group(1)
                d["balance_due_date"] = bal.group(2)
                continue

            for pattern, key in [
                (r'SHIP NAME:\s*(.+)', "ship"),
                (r'CABIN NUMBER:\s*(\S+)', "cabin"),
                (r'DECK:\s*(\S*)', "deck"),
                (r'DEPARTURE PORT:\s*(.+)', "port"),
                (r'ITINERARY NAME:\s*(.+)', "itinerary"),
                (r'DINING REQUEST:\s*(.+)', "dining"),
                (r'DEPART DATE\s+(\S+)', "depart_date"),
                (r'RETURN DATE\s+(\S+)', "return_date"),
                (r'ADULT:\s*([\d.]+)\s*X\s*(\d+)', "per_person"),
            ]:
                m = re.search(pattern, line)
                if m:
                    d[key] = m.group(1).strip() if m.lastindex == 1 else m.groups()
                    break
            continue

        elif state == TOUR and current_tour:
            # Tour vendor line: **POSITANO CAR SERVICE**/AMT-550.00/CF-GAETA
            vendor = re.search(r'\*\*(.+?)\*\*', line)
            if vendor:
                current_tour["vendor"] = vendor.group(1).strip()
                amt = re.search(r'AMT-([\d.]+)', line)
                if amt:
                    current_tour["amount"] = amt.group(1)
                cf = re.search(r'CF-(\S+)', line)
                if cf:
                    current_tour["confirmation"] = cf.group(1)
                continue

            # Total cost of the tour/package, e.g. "TOTAL COST   USD  3649.00"
            tc = re.search(r'TOTAL COST\s+USD\s+([\d.]+)', line)
            if tc:
                current_tour["total_cost"] = tc.group(1)
                continue

            # Payment already applied, e.g. "12JUN PAYMENT BY VISA  USD  750.00-"
            # (may or may not name a method between "PAYMENT BY" and the amount)
            pay = re.search(r'(\d{2}[A-Z]{3})\s+PAYMENT BY\s*(.*?)\s+USD\s+([\d.]+)-', line)
            if pay:
                current_tour.setdefault("payments", []).append({
                    "date": pay.group(1), "method": pay.group(2).strip(),
                    "amount": pay.group(3),
                })
                continue

            # Remaining balance, e.g. "BALANCE OF 2899.00 DUE 08JUL2026" — this is
            # the authoritative "what's actually still owed" figure.
            bal = re.search(r'BALANCE OF\s+([\d.]+)\s+DUE\s+(\S+)', line)
            if bal:
                current_tour["balance_due"] = bal.group(1)
                current_tour["balance_due_date"] = bal.group(2)
                continue

            # "TOTAL DUE: 3649.00 BY 08JUL2026" repeats TOTAL COST under a due date
            # rather than reflecting payments already made — it's redundant/misleading
            # next to BALANCE OF, so we only keep the date (as a fallback if there's
            # no BALANCE OF line) and drop the confusing repeated amount.
            td = re.search(r'TOTAL DUE:\s*[\d.]+\s+BY\s+(\S+)', line)
            if td:
                current_tour.setdefault("balance_due_date", td.group(1))
                continue

            # "TYPE OF PKG: BACKROADS BELIZE & GUATEMALA MULTI ADVENTURE" — this
            # describes what the booking actually is (not always "transportation"),
            # so it's captured as its own field rather than dumped into details.
            ty = re.search(r'TYPE OF PKG:\s*(.+)', line)
            if ty:
                current_tour["type"] = ty.group(1).strip()
                continue

            # Tour detail lines
            if re.search(r'PICK UP|DROP OFF', line):
                current_tour["details"].append(line.strip())
                continue

            # Generic detail continuation
            stripped = line.strip()
            if stripped and current_tour.get("vendor"):
                current_tour["details"].append(stripped)
                continue

        elif state == PACKAGE and current_package:
            d = current_package["details"]

            vendor = re.search(r'\*\*(.+?)\*\*', line)
            if vendor:
                d["vendor"] = vendor.group(1).strip()
                amt = re.search(r'AMT-([\d.]+)', line)
                if amt:
                    d["amount"] = amt.group(1)
                cf = re.search(r'CF-(\S+)', line)
                if cf:
                    d["confirmation"] = cf.group(1)
                continue

            # Total cost of the package, e.g. "TOTAL COST   USD  5621.38"
            tc = re.search(r'TOTAL COST\s+USD\s+([\d.]+)', line)
            if tc:
                d["total_cost"] = tc.group(1)
                continue

            # Payment already applied — collected as a list so multiple payments
            # (e.g. a deposit plus a final payment) don't overwrite each other.
            pay = re.search(r'(\d{2}[A-Z]{3})\s+PAYMENT BY\s*(.*?)\s+USD\s+([\d.]+)-', line)
            if pay:
                d.setdefault("payments", []).append({
                    "date": pay.group(1), "method": pay.group(2).strip(),
                    "amount": pay.group(3),
                })
                continue

            # Remaining balance, e.g. "BALANCE OF 2899.00 DUE 08JUL2026"
            bal = re.search(r'BALANCE OF\s+([\d.]+)\s+DUE\s+(\S+)', line)
            if bal:
                d["balance_due"] = bal.group(1)
                d["balance_due_date"] = bal.group(2)
                continue

            # "TOTAL DUE: X BY DATE" repeats TOTAL COST under a due date rather than
            # reflecting payments made — redundant next to BALANCE OF, so only the
            # date is kept (as a fallback), and the confusing repeated amount is dropped.
            td = re.search(r'TOTAL DUE:\s*[\d.]+\s+BY\s+(\S+)', line)
            if td:
                d.setdefault("balance_due_date", td.group(1))
                continue

            ty = re.search(r'TYPE OF PKG:\s*(.+)', line)
            if ty:
                d["type"] = ty.group(1).strip()
                continue

            # Anything else is descriptive text about the package
            stripped = line.strip()
            if stripped:
                d.setdefault("description", []).append(stripped)
            continue

        elif state == TICKETS:
            # TIPITIN: WHEELER/JOHN DANIEL  0167484690269  VIC CARD  USD  612.03
            tkt = re.search(r'([A-Z]+/[A-Z ]+?)\s{2,}(\d{10,}(?:-\d+)?)\s*(\S+\s*\S*)\s+USD\s+([\d.]+)', line)
            if tkt:
                data["tickets"].append({
                    "passenger": tkt.group(1).strip(),
                    "ticket_number": tkt.group(2),
                    "payment_method": tkt.group(3).strip(),
                    "amount_usd": tkt.group(4),
                })
                continue

            # ITIN: AIR TICKET/S  7401640949  AX CARD  3898.44
            itkt = re.search(r'AIR TICKET/S\s+(\d+)\s+(\S+\s*\S*)\s+([\d.]+)', line)
            if itkt:
                data["tickets"].append({
                    "passenger": "",
                    "ticket_number": itkt.group(1),
                    "payment_method": itkt.group(2).strip(),
                    "amount_usd": itkt.group(3),
                })
                continue

            # Exchanged ticket
            if re.search(r'EXCHANGED FOR TICKET', line):
                continue  # next line has the number

            # Exchange ticket number on continuation line
            exch_num = re.match(r'\s+(\d{10,})', line)
            if exch_num:
                data["exchanged_ticket"] = exch_num.group(1)
                continue

            # Financial line encountered while in tickets
            if re.match(r'AIR FARE USD|^\s*SUB TOTAL', line):
                state = FINANCIAL
                # fall through to FINANCIAL processing below

        if state == FINANCIAL:
            for pattern, key in [
                (r'AIR FARE USD\s+([\d.]+)', "air_fare"),
                (r'TAX AND CARRIER FEES USD\s+([\d.]+)', "tax_and_fees"),
                (r'TTL USD\s+([\d.]+)', "total"),
                (r'SUB TOTAL\s+(?:USD\s+)?([\d.]+)', "sub_total"),
                (r'CREDIT CARD PAYMENT\s+USD\s+([\d.]+)', "credit_card_payment"),
                (r'AMOUNT DUE\s+(?:USD\s+)?([\d.]+)', "amount_due"),
                (r'TOTAL AMOUNT\s+([\d.]+)', "amount_due"),
            ]:
                m = re.search(pattern, line)
                if m:
                    data["financial"][key] = m.group(1)
                    break
            # Fare per person (ITIN)
            fare = re.search(r'FARE\.+\s*([\d.]+)', line)
            if fare:
                data["financial"]["fare_per_person"] = fare.group(1)
            continue

        elif state == BAGGAGE:
            route = re.match(r'\s*([A-Z]{2} [A-Z]{3,6})\s+(\d+PC)', line)
            if route:
                data["baggage"].append({
                    "route": route.group(1), "count": route.group(2), "bags": []
                })
                continue
            bag = re.search(r'BAG (\d+) - (.+)', line)
            if bag and data["baggage"]:
                data["baggage"][-1]["bags"].append({
                    "bag_num": bag.group(1), "info": bag.group(2).strip()
                })
                continue

        elif state == CARRY_ON:
            route = re.match(r'\s*([A-Z]{2} [A-Z]{3,6})\s+(\d+PC)', line)
            if route:
                data["carry_on"].append({
                    "route": route.group(1), "count": route.group(2), "bags": []
                })
                continue
            bag = re.search(r'BAG (\d+) - (.+)', line)
            if bag and data["carry_on"]:
                data["carry_on"][-1]["bags"].append({
                    "bag_num": bag.group(1), "info": bag.group(2).strip()
                })
                continue

        # ── Notices (can appear in any state) ─────────────────
        notice = re.match(r'\s*\*\*\s*(.+?)\s*\*\*', line)
        if notice:
            text = notice.group(1).strip()
            if text and text not in data["notices"] and len(text) > 5:
                data["notices"].append(text)
            continue

        # Service fee (can appear in various states)
        sf = re.search(r'SERVICE FEES?\s+USD\s+([\d.]+)', line)
        if sf:
            data["service_fee"] = sf.group(1)
            continue

        # Passport / restriction / baggage discount notices
        if re.search(r'PASSPORT|NON.?REFUNDABLE|PENALTY.*CHANGE|BAGGAGE DISCOUNTS MAY|ONLINE CHECKIN/FORM OF PAYMENT', line):
            stripped = line.strip()
            if stripped and stripped not in data["notices"]:
                data["notices"].append(stripped)
            continue

        # FARE lines (ITIN format)
        fare_line = re.search(r'FARE\.+\s*([\d.]+)', line)
        if fare_line:
            data["financial"]["fare_per_person"] = fare_line.group(1)
            # Capture fare note if present
            note = re.search(r'PER PERSON\s*-?\s*(.+?)$', line)
            if note:
                data["financial"]["fare_note"] = note.group(1).strip()
            continue

        # ROUNDTRIP FARE note
        rt = re.search(r'ROUNDTRIP FARE:\s*(.+?)$', line)
        if rt:
            data["financial"]["fare_note"] = rt.group(1).strip()
            continue

        # Standalone total amount (just a number on its own line)
        if re.match(r'^\s+[\d.]+\s*$', line) and state in (TICKETS, FINANCIAL):
            continue  # total line, already captured elsewhere

        # Insurance lines
        if re.search(r'ALLIANZ|TRAVEL GUARD|POLICY TYPE|INSURED TRIP|PREMIUM BASED|PAYMENT BY CREDIT CARD|INSURANCE COVERAGE|SICKNESS.BAGGAGE|TRIP CANCELLATION|PAYMENT BY CHECK', line):
            data["insurance"].append(line.strip())
            continue

        # After hours / emergency
        if re.search(r'AFTER HOURS|EMERGENCY', line):
            continue

        # Lines we don't recognize
        stripped = line.strip()
        if stripped and header_done:
            # Skip repeat header lines on page 2+ (ITIN format repeats header)
            if re.search(r'SALES PERSON:|CUSTOMER NBR:', line):
                continue
            if re.match(r'\s*TO:\s', line):
                continue
            if re.match(r'\s*FOR:\s', line):
                continue
            # Skip passenger lines that appear in repeated headers
            if re.match(r'\s+[A-Z]+/[A-Z]', line) and any(
                p["full_slash"] in line for p in data["passengers"]
            ):
                continue
            # Skip repeated address lines
            if stripped in data["mailing_address"]:
                continue
            data["unrecognized"].append(f"L{line_num}: {stripped}")

    # ── Flush pending items ───────────────────────────────────
    if current_flight:
        data["flights"].append(current_flight)
    if current_hotel:
        data["hotels"].append(current_hotel)
    if current_cruise and current_cruise["details"]:
        data["cruises"].append(current_cruise)
    if current_tour and (current_tour.get("vendor") or current_tour.get("details")):
        data["tours"].append(current_tour)
    if current_package and current_package.get("details"):
        data["packages"].append(current_package)

    # Assign passengers to ITIN-format tickets
    if fmt == "ITIN":
        for i, ticket in enumerate(data["tickets"]):
            if not ticket["passenger"] and i < len(data["passengers"]):
                ticket["passenger"] = data["passengers"][i]["full_slash"]

    # Set mileage from first FF entry
    if data["freq_flyers"]:
        ff = data["freq_flyers"][0]
        data["mileage"] = {"airline": ff["airline"], "number": ff["number"]}

    # ── Validate must-haves ───────────────────────────────────
    data["warnings"] = _validate(data)

    return data


def _validate(data: dict) -> list:
    """Check must-have fields. Returns list of warning strings."""
    w = []

    if not data["passengers"]:
        w.append("MISSING: No passengers found")
    if not data["booking"].get("itin_no"):
        w.append("MISSING: No ITIN/invoice number")
    if not data["booking"].get("date"):
        w.append("MISSING: No date")

    for i, fl in enumerate(data["flights"]):
        prefix = f"Flight {i+1}"
        if not fl.get("departure_city"):
            w.append(f"MISSING: {prefix} departure city")
        if not fl.get("departure_time"):
            w.append(f"MISSING: {prefix} departure time")
        if not fl.get("arrival_city"):
            w.append(f"MISSING: {prefix} arrival city")
        if not fl.get("arrival_time"):
            w.append(f"MISSING: {prefix} arrival time")
        if not fl.get("airline"):
            w.append(f"MISSING: {prefix} airline")
        if not fl.get("flight_number"):
            w.append(f"MISSING: {prefix} flight number")
        if not fl.get("cabin_class"):
            w.append(f"MISSING: {prefix} cabin class")
        if not fl.get("date_raw"):
            w.append(f"MISSING: {prefix} date")

    for i, h in enumerate(data["hotels"]):
        prefix = f"Hotel {i+1}"
        name = h.get("name") or h.get("chain")
        if not name:
            w.append(f"MISSING: {prefix} name")
        if not h.get("confirmation"):
            w.append(f"MISSING: {prefix} confirmation number")
        if not h.get("rate_amount"):
            w.append(f"MISSING: {prefix} rate")

    if data["flights"] and not data["tickets"]:
        w.append("MISSING: No tickets found (flights exist)")

    return w


# ── CLI ────────────────────────────────────────────────────────
if __name__ == "__main__":
    import json
    import sys

    path = sys.argv[1] if len(sys.argv) > 1 else None
    if not path:
        print("Usage: python state_parser.py <invoice.pdf>")
        sys.exit(1)

    result = parse(path)

    # Print summary
    print(f"Format: {result['format']}")
    print(f"Passengers: {[p['full_slash'] for p in result['passengers']]}")
    print(f"Booking: {result['booking']}")
    print(f"Flights: {len(result['flights'])}")
    print(f"Hotels: {len(result['hotels'])}")
    print(f"Cruises: {len(result['cruises'])}")
    print(f"Tours: {len(result['tours'])}")
    print(f"Packages: {len(result['packages'])}")
    print(f"Insurance: {len(result['insurance'])}")
    print(f"Tickets: {len(result['tickets'])}")
    print(f"Financial: {result['financial']}")
    print(f"Service fee: {result['service_fee']}")

    if result["warnings"]:
        print(f"\n⚠ WARNINGS ({len(result['warnings'])}):")
        for w in result["warnings"]:
            print(f"  {w}")

    if result["unrecognized"]:
        print(f"\n? UNRECOGNIZED LINES ({len(result['unrecognized'])}):")
        for u in result["unrecognized"][:20]:
            print(f"  {u}")