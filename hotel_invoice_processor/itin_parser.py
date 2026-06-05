"""
itin_parser.py - Extract structured data from ITIN-format invoices.
(The format with SALES PERSON, AIR UNITED FLT:, AR:, etc.)

Returns the same dict structure as invoice_parser.py so the
generator can use either parser interchangeably.
"""

import re
import fitz


def parse_itin_invoice(pdf_path: str) -> dict:
    """Parse an ITIN-format invoice PDF and return structured data."""
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
        "insurance": None,
        "freq_flyers": [],
    }

    # ── Booking info ──────────────────────────────────────────
    inv = re.search(r'ITIN/INVOICE NO\.\s+(\d+)', page1)
    sp = re.search(r'SALES PERSON:\s*(\S+)', page1)
    date = re.search(r'DATE:(.+?)$', page1, re.MULTILINE)
    cust = re.search(r'CUSTOMER NBR:\s*(\S+)', page1)

    data["booking"] = {
        "itin_no": inv.group(1) if inv else None,
        "record_locator": cust.group(1) if cust else None,
        "date": date.group(1).strip() if date else None,
        "sales_person": sp.group(1) if sp else None,
    }

    # ── Mailing address (TO: block) ───────────────────────────
    addr_match = re.search(
        r'TO:\s*(.+?)(?=\nFOR:)',
        page1, re.DOTALL
    )
    if addr_match:
        data["mailing_address"] = [
            l.strip() for l in addr_match.group(1).strip().split("\n") if l.strip()
        ]

    # ── Passengers (FOR: block) ───────────────────────────────
    pax_block = re.search(
        r'FOR:\s*(.+?)(?=\n\d{2} [A-Z]{3} \d{2}|\n\s*\-{4,})',
        page1, re.DOTALL
    )
    if pax_block:
        pax_lines = [l.strip() for l in pax_block.group(1).strip().split("\n") if l.strip()]
        for line in pax_lines:
            m = re.match(r'([A-Z]+)/(.+)', line)
            if m:
                last = m.group(1)
                first_mid = m.group(2).strip()
                parts = first_mid.split()
                data["passengers"].append({
                    "last_name": last,
                    "first_name": parts[0] if parts else "",
                    "middle_name": " ".join(parts[1:]) if len(parts) > 1 else "",
                    "full_slash": f"{last}/{first_mid}",
                })

    # ── Frequent flyers ───────────────────────────────────────
    for fm in re.finditer(
        r'FREQ FLYER:\s*(.+?)\s+([A-Z]{2})\s{2,}(\S+)',
        full_text
    ):
        entry = {
            "passenger": fm.group(1).strip(),
            "airline": fm.group(2).strip(),
            "number": fm.group(3).strip(),
        }
        # Avoid duplicates
        if entry not in data["freq_flyers"]:
            data["freq_flyers"].append(entry)

    # Set mileage from first FF entry (for compatibility with generator)
    if data["freq_flyers"]:
        ff = data["freq_flyers"][0]
        data["mileage"] = {"airline": ff["airline"], "number": ff["number"]}

    # ── Flights ───────────────────────────────────────────────
    # Split on date lines: "26 JUN 26    - FRIDAY"
    flight_splits = re.split(r'(?=\d{2} [A-Z]{3} \d{2}\s+- [A-Z]+)', full_text)

    for block in flight_splits:
        date_match = re.search(r'(\d{2} [A-Z]{3} \d{2})\s+- ([A-Z]+)', block)
        if not date_match:
            continue

        # Find AIR lines in this block
        air_matches = list(re.finditer(
            r'AIR\s+(\S+)\s+FLT:(\d+)\s+(.+?)$',
            block, re.MULTILINE
        ))

        for am in air_matches:
            airline_name = am.group(1).strip()
            flight_num = am.group(2)
            cabin_meal = am.group(3).strip()

            # Get the sub-block for this flight
            am_start = am.start()
            # Find next AIR line or end of block
            next_air = re.search(r'\nAIR\s+\S+\s+FLT:', block[am.end():])
            if next_air:
                sub_block = block[am_start:am.end() + next_air.start()]
            else:
                sub_block = block[am_start:]

            # Parse cabin and meal from "BUS/ FIRST   MEAL"
            cabin_parts = re.match(r'(.+?)\s{2,}(\S+)$', cabin_meal)
            if cabin_parts:
                cabin = cabin_parts.group(1).strip()
                meals = cabin_parts.group(2).strip()
            else:
                cabin = cabin_meal
                meals = None

            # Departure
            dep = re.search(r'LV:\s+(.+?)\s{2,}(\d+[AP])', sub_block)
            # Arrival
            arr = re.search(r'AR:\s+(.+?)\s{2,}(\d+[AP])', sub_block)

            if not dep or not arr:
                continue

            nonstop = 'NONSTOP' in sub_block

            # Duration
            duration = re.search(r'ELAPSED TIME-\s*(.+?)$', sub_block, re.MULTILINE)

            # Equipment
            equip = re.search(r'EQUIP-(\S+)', sub_block)

            # Seats
            seats_match = re.search(r'RESERVED SEATS\s+SEAT-\s*(.+?)$', sub_block, re.MULTILINE)
            seat_list = []
            if seats_match:
                seat_list = seats_match.group(1).strip().split()

            # Airline confirmation
            conf = re.search(r'AIRLINE CONFIRMATION:\s*(\S+)\s*-(\S+)', sub_block)

            # Air mileage
            mileage = re.search(r'AIR MILEAGE:\s*(\d+)', sub_block)

            flight = {
                "date_raw": date_match.group(1),
                "day_name": date_match.group(2),
                "airline": airline_name,
                "flight_number": flight_num,
                "cabin_class": cabin,
                "operated_by": None,
                "departure_city": dep.group(1).strip(),
                "departure_time": dep.group(2),
                "arrival_city": arr.group(1).strip(),
                "arrival_time": arr.group(2),
                "arrives_next_day": None,
                "duration": duration.group(1).strip() if duration else None,
                "nonstop": nonstop,
                "confirmed": True,
                "baggage_allowance": None,
                "seats": seat_list,
                "meals": meals,
                "dep_terminal": None,
                "arr_terminal": None,
                "airline_locator_carrier": conf.group(1) if conf else None,
                "airline_locator_code": conf.group(2) if conf else None,
                "equipment": equip.group(1) if equip else None,
                "air_mileage": mileage.group(1) if mileage else None,
            }
            data["flights"].append(flight)

    # ── Insurance ─────────────────────────────────────────────
    ins_match = re.search(
        r'TRAVEL GUARD\s*\n(.+?)(?=\n\s*\*\*|\n\s*-{4,}|\n\s*FOR EMERGENCY)',
        full_text, re.DOTALL
    )
    if ins_match:
        ins_lines = [l.strip() for l in ins_match.group(1).strip().split("\n") if l.strip()]
        data["insurance"] = ins_lines

        # Extract premium amount
        prem = re.search(r'COVERA\s+([\d.]+)', full_text)
        if prem:
            data["insurance_premium"] = prem.group(1)

    # ── Tickets ───────────────────────────────────────────────
    for tm in re.finditer(
        r'AIR TICKET/S\s+(\d+)\s+(\S+\s*\S*)\s+([\d.]+)',
        full_text
    ):
        data["tickets"].append({
            "passenger": "",  # ITIN format doesn't list passenger per ticket
            "ticket_number": tm.group(1),
            "payment_method": tm.group(2).strip(),
            "amount_usd": tm.group(3),
        })

    # Assign passengers to tickets in order
    for i, ticket in enumerate(data["tickets"]):
        if i < len(data["passengers"]):
            ticket["passenger"] = data["passengers"][i]["full_slash"]

    # ── Fare info ─────────────────────────────────────────────
    fare_match = re.search(r'FARE\.+\s*([\d.]+)\s+PER PERSON\s*-?\s*(.*?)$', full_text, re.MULTILINE)
    if fare_match:
        data["fare_per_person"] = fare_match.group(1)
        data["fare_note"] = fare_match.group(2).strip()

    # ── Financial summary ─────────────────────────────────────
    sub = re.search(r'SUB TOTAL\s+([\d.]+)', full_text)
    total_amt = re.search(r'TOTAL AMOUNT\s+([\d.]+)', full_text)

    data["financial"] = {
        "air_fare": None,
        "tax_and_fees": None,
        "total": None,
        "sub_total": sub.group(1) if sub else None,
        "credit_card_payment": None,
        "amount_due": total_amt.group(1) if total_amt else None,
    }

    # If we have fare_per_person and ticket count, calculate total fare
    if data.get("fare_per_person") and data["tickets"]:
        n_pax = len(data["passengers"]) or 1
        per_person = float(data["fare_per_person"])
        data["financial"]["air_fare"] = f"{per_person:.2f} x{n_pax}"
        data["financial"]["total"] = sub.group(1) if sub else f"{per_person * n_pax:.2f}"

    # ── Service fees ──────────────────────────────────────────
    sf = re.search(r'SERVICE FEES?\s+(?:USD\s+)?([\d.]+)', full_text)
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
    for pattern in [
        r'\*\*\s*(.+?)\s*\*\*',
    ]:
        for m in re.finditer(pattern, full_text):
            notice = m.group(1).strip()
            if notice and notice not in data["notices"] and len(notice) > 10:
                data["notices"].append(notice)

    if "BAGGAGE DISCOUNTS MAY APPLY" in full_text:
        notice = ("BAGGAGE DISCOUNTS MAY APPLY BASED ON FREQUENT FLYER STATUS/"
                  "ONLINE CHECKIN/FORM OF PAYMENT/MILITARY/ETC.")
        if notice not in data["notices"]:
            data["notices"].append(notice)

    return data


if __name__ == "__main__":
    import json
    import sys

    path = sys.argv[1] if len(sys.argv) > 1 else "/mnt/user-data/uploads/OTHER_639155617112913314.pdf"
    result = parse_itin_invoice(path)
    print(json.dumps(result, indent=2))