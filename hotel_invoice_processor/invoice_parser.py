"""
invoice_parser.py - Extract structured data from old-format (tipitin) invoices.
Only uses information present in the PDF text. No enrichment or API calls.
"""

import re
import fitz


def _find_flight_segments(text):
    """Find start positions of each flight segment (airline + flight number line)."""
    pattern = re.compile(
        r'^\s+(\S.+?(?:INC\.|AIRLINES).*?)\s+(\d+)\s+(.+?)$',
        re.MULTILINE
    )
    segments = []
    for m in pattern.finditer(text):
        line = m.group(0).strip()
        if any(kw in line for kw in ['HOTEL', 'RESORT', 'GUARD', 'INSURANCE', 'TRAVEL WIZARDS']):
            continue
        segments.append(m.start())
    return segments


def _parse_flight_block(block, passengers):
    """Parse a single flight segment block."""
    airline_match = re.search(
        r'^\s*(.+?(?:INC\.|AIRLINES).*?)\s+(\d+)\s+(.+?)$',
        block, re.MULTILINE
    )
    if not airline_match:
        return None

    airline = airline_match.group(1).strip()
    flight_num = airline_match.group(2)
    cabin = airline_match.group(3).strip()

    if any(kw in airline for kw in ['HOTEL', 'RESORT', 'GUARD', 'INSURANCE']):
        return None

    depart = re.search(r'LV:\s+(.+?)\s{2,}(\d+[AP])\s', block)
    arrive = re.search(r'ARR:\s+(.+?)\s{2,}(\d+[AP])', block)
    if not depart or not arrive:
        return None

    operated_by = None
    op = re.search(r'OPERATED BY-(.+)', block)
    if op:
        operated_by = op.group(1).strip()

    duration = re.search(r'FLIGHT TIME - (.+?)(?:\s{2,}|$)', block)
    baggage = re.search(r'BAGGAGE ALLOWANCE - (\S+)', block)

    seat_lines = re.findall(
        r'(?:SEAT ASSIGNED\s+|^\s{10,})([A-Z0-9]+)\s+NON SMOKING',
        block, re.MULTILINE
    )
    seat_list = [s for s in seat_lines if s != 'NON']

    if not seat_list:
        unassigned = re.findall(r'NON SMOKING/(?:AISLE|WINDOW)', block)
        n_pax = len(passengers) or 2
        seat_list = ["Unassigned"] * min(len(unassigned), n_pax)

    meals = re.search(r'MEALS SERVED\s+(.+?)$', block, re.MULTILINE)
    dep_terminal = re.search(r'DEPART - (TERMINAL\s+\S+)', block)
    arr_terminal = re.search(r'ARRIVE - (TERMINAL\s+\S+)', block)
    locator = re.search(r'AIRLINE LOCATOR:\s*(\S+)\s*-(\S+)', block)
    nonstop = 'NON-STOP' in block
    confirmed = 'CONFIRMED' in block
    arrives_next = re.search(r'ARRIVES-\s*(\d+ [A-Z]+)', block)
    date_match = re.search(r'(\d{2} [A-Z]{3} \d{2}) - ([A-Z]+)', block)

    return {
        "date_raw": date_match.group(1) if date_match else None,
        "day_name": date_match.group(2) if date_match else None,
        "airline": airline,
        "flight_number": flight_num,
        "cabin_class": cabin,
        "operated_by": operated_by,
        "departure_city": depart.group(1).strip(),
        "departure_time": depart.group(2),
        "arrival_city": arrive.group(1).strip(),
        "arrival_time": arrive.group(2),
        "arrives_next_day": arrives_next.group(1) if arrives_next else None,
        "duration": duration.group(1).strip() if duration else None,
        "nonstop": nonstop,
        "confirmed": confirmed,
        "baggage_allowance": baggage.group(1) if baggage else None,
        "seats": seat_list,
        "meals": meals.group(1).strip() if meals else None,
        "dep_terminal": dep_terminal.group(1).strip() if dep_terminal else None,
        "arr_terminal": arr_terminal.group(1).strip() if arr_terminal else None,
        "airline_locator_carrier": locator.group(1) if locator else None,
        "airline_locator_code": locator.group(2) if locator else None,
    }


def parse_invoice(pdf_path: str) -> dict:
    """Parse an old-format invoice PDF and return structured data."""
    doc = fitz.open(pdf_path)
    pages_text = [page.get_text("text") for page in doc]
    doc.close()

    full_text = "\n".join(pages_text)
    page1 = pages_text[0]

    data = {
        "passengers": [],
        "mailing_address": [],
        "booking": {},
        "flights": [],
        "hotels": [],
        "tickets": [],
        "financial": {},
        "baggage": [],
        "carry_on": [],
        "notices": [],
        "mileage": None,
        "service_fee": None,
    }

    # ── Passengers ─────────────────────────────────────────────
    pax_matches = re.findall(
        r'^\s{2,}([A-Z]+)/([A-Z][A-Z ]+)$', page1, re.MULTILINE
    )
    for last, first_mid in pax_matches:
        parts = first_mid.strip().split()
        data["passengers"].append({
            "last_name": last,
            "first_name": parts[0] if parts else "",
            "middle_name": " ".join(parts[1:]) if len(parts) > 1 else "",
            "full_slash": f"{last}/{first_mid.strip()}",
        })

    # ── Mailing address ───────────────────────────────────────
    addr_match = re.search(
        r'(?:^[A-Z]+/[A-Z ]+\n)+(.+?)(?=\s*ITIN NO:)',
        page1, re.DOTALL | re.MULTILINE
    )
    if addr_match:
        data["mailing_address"] = [
            l.strip() for l in addr_match.group(1).strip().split("\n") if l.strip()
        ]

    # ── Booking info ──────────────────────────────────────────
    itin = re.search(r'ITIN NO:\s*(\d+)', page1)
    rec = re.search(r'RECORD LOCATOR:\s*(\S+)', page1)
    date = re.search(r'DATE:\s*(.+?)$', page1, re.MULTILINE)
    data["booking"] = {
        "itin_no": itin.group(1) if itin else None,
        "record_locator": rec.group(1) if rec else None,
        "date": date.group(1).strip() if date else None,
    }

    # ── Mileage membership ────────────────────────────────────
    mm = re.search(r'MILEAGE MEMBERSHIP:\s*(\S+)\s+(\S+)', full_text)
    if mm:
        data["mileage"] = {"airline": mm.group(1), "number": mm.group(2)}

    # ── Flights ───────────────────────────────────────────────
    seg_starts = _find_flight_segments(full_text)

    for idx, start in enumerate(seg_starts):
        end = seg_starts[idx + 1] if idx + 1 < len(seg_starts) else len(full_text)
        block = full_text[start:end]
        flight = _parse_flight_block(block, data["passengers"])
        if flight:
            # Get date from text PRECEDING this segment (within 300 chars)
            pre = full_text[max(0, start - 300):start]
            date_match = re.search(r'(\d{2} [A-Z]{3} \d{2}) - ([A-Z]+)', pre)
            if date_match:
                flight["date_raw"] = date_match.group(1)
                flight["day_name"] = date_match.group(2)
            elif data["flights"]:
                # Continuation flight, inherit from previous
                flight["date_raw"] = data["flights"][-1]["date_raw"]
                flight["day_name"] = data["flights"][-1]["day_name"]
            data["flights"].append(flight)

    # ── Hotels ────────────────────────────────────────────────
    for hm in re.finditer(
        r'^\s*(.+?(?:HOTELS?|RESORTS?).*?)\s+(\d+)\s+NT/S\s*-\s*OUT\s+(\S+)\s+(CONFIRMED|WAITLIST)',
        full_text, re.MULTILINE
    ):
        remaining = full_text[hm.start():]
        block_end = re.search(
            r'SURFACE TRANSPORTATION|\d{2} [A-Z]{3} \d{2} - [A-Z]+',
            remaining[10:]
        )
        hotel_block = remaining[:block_end.start() + 10] if block_end else remaining[:2000]

        name_m = re.search(r'^\s*(.+?)\s{2,}GUARANTEE', hotel_block, re.MULTILINE)
        addr_m = re.search(r'^\s{2}(\d+.+?)\s{2,}RATE', hotel_block, re.MULTILINE)
        city_m = re.search(r'^\s{2}([A-Z][A-Z\s\d]+(?:CA|US|UK))\s', hotel_block, re.MULTILINE)
        rate_m = re.search(r'RATE-\s*(\S+)\s+([\d.]+)', hotel_block)
        conf_m = re.search(r'CONFIRMATION-(\S+)', hotel_block)
        phone_m = re.search(r'PHONE NO-(.+?)$', hotel_block, re.MULTILINE)
        fax_m = re.search(r'FAX NO-(.+?)\s{2,}', hotel_block)
        total_m = re.search(r'APPROX TTL.*?([\d.]+)(\w+)', hotel_block)
        guar_m = re.search(r'GUARANTEE-(.+?)$', hotel_block, re.MULTILINE)

        notes = []
        for line in hotel_block.split("\n"):
            line = line.strip()
            if line and (line.startswith("WI") or "CXL:" in line or
                         "AAA" in line or "CANCEL" in line):
                notes.append(line)

        data["hotels"].append({
            "chain": hm.group(1).strip(),
            "nights": int(hm.group(2)),
            "checkout_date": hm.group(3),
            "status": hm.group(4),
            "name": name_m.group(1).strip() if name_m else None,
            "address": addr_m.group(1).strip() if addr_m else None,
            "city": city_m.group(1).strip() if city_m else None,
            "rate_currency": rate_m.group(1) if rate_m else None,
            "rate_amount": rate_m.group(2) if rate_m else None,
            "confirmation": conf_m.group(1) if conf_m else None,
            "phone": phone_m.group(1).strip() if phone_m else None,
            "fax": fax_m.group(1).strip() if fax_m else None,
            "approx_total": f"{total_m.group(1)}{total_m.group(2)}" if total_m else None,
            "guarantee": guar_m.group(1).strip() if guar_m else None,
            "notes": notes,
        })

    # ── Tickets ───────────────────────────────────────────────
    for tm in re.finditer(
        r'^\s*([A-Z]+/[A-Z ]+?)\s{2,}(\d{13})\s+(\S+\s*\S*)\s+USD\s+([\d.]+)',
        full_text, re.MULTILINE
    ):
        data["tickets"].append({
            "passenger": tm.group(1).strip(),
            "ticket_number": tm.group(2),
            "payment_method": tm.group(3).strip(),
            "amount_usd": tm.group(4),
        })

    # ── Financial summary ─────────────────────────────────────
    air_fare = re.search(r'AIR FARE USD\s+([\d.]+)', full_text)
    tax = re.search(r'TAX AND CARRIER FEES USD\s+([\d.]+)', full_text)
    ttl = re.search(r'TTL USD\s+([\d.]+)', full_text)
    sub = re.search(r'SUB TOTAL\s+USD\s+([\d.]+)', full_text)
    cc = re.search(r'CREDIT CARD PAYMENT\s+USD\s+([\d.]+)', full_text)
    due = re.search(r'AMOUNT DUE\s+USD\s+([\d.]+)', full_text)
    data["financial"] = {
        "air_fare": air_fare.group(1) if air_fare else None,
        "tax_and_fees": tax.group(1) if tax else None,
        "total": ttl.group(1) if ttl else None,
        "sub_total": sub.group(1) if sub else None,
        "credit_card_payment": cc.group(1) if cc else None,
        "amount_due": due.group(1) if due else None,
    }

    # ── Service fees ──────────────────────────────────────────
    sf = re.search(r'SERVICE FEES?\s+USD\s+([\d.]+)', full_text)
    if sf:
        data["service_fee"] = sf.group(1)

    # ── Baggage allowance ─────────────────────────────────────
    bag_start = full_text.find("BAGGAGE ALLOWANCE")
    carry_start = full_text.find("CARRY ON ALLOWANCE")

    if bag_start >= 0:
        bag_end = carry_start if carry_start > bag_start else len(full_text)
        bag_text = full_text[bag_start:bag_end]
        for route, count, details in re.findall(
            r'([A-Z]{2} [A-Z]{3,6})\s+(\d+PC)\n((?:\s+BAG \d+ - .+\n)*)',
            bag_text
        ):
            bags = re.findall(r'BAG (\d+) - (.+)', details)
            data["baggage"].append({
                "route": route, "count": count,
                "bags": [{"bag_num": b[0], "info": b[1].strip()} for b in bags],
            })

    if carry_start >= 0:
        carry_text = full_text[carry_start:]
        for route, count, details in re.findall(
            r'([A-Z]{2} [A-Z]{3,6})\s+(\d+PC)\n((?:\s+BAG \d+ - .+\n)*)',
            carry_text
        ):
            bags = re.findall(r'BAG (\d+) - (.+)', details)
            data["carry_on"].append({
                "route": route, "count": count,
                "bags": [{"bag_num": b[0], "info": b[1].strip()} for b in bags],
            })

    # ── Notices ───────────────────────────────────────────────
    if "TRAVEL WIZARDS CANCELLATION FEES APPLY" in full_text:
        data["notices"].append("TRAVEL WIZARDS CANCELLATION FEES APPLY")
    if "BAGGAGE DISCOUNTS MAY APPLY" in full_text:
        data["notices"].append(
            "BAGGAGE DISCOUNTS MAY APPLY BASED ON FREQUENT FLYER STATUS/"
            "ONLINE CHECKIN/FORM OF PAYMENT/MILITARY/ETC."
        )

    return data


if __name__ == "__main__":
    import json
    import sys

    path = sys.argv[1] if len(sys.argv) > 1 else "/mnt/user-data/uploads/OTHER_639137689954244936.pdf"
    result = parse_invoice(path)
    print(json.dumps(result, indent=2))
