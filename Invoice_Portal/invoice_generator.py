"""
invoice_generator.py — Generate styled PDF from state_parser output.

Handles: flights, hotels, cruises, tours, packages, insurance,
tickets, financial, baggage, notices.

No header/footer drawn — overlay.pdf provides branding.
"""

import os
import sys

try:
    from airport_lookup import resolve_city, lookup_airport
except ImportError:
    def resolve_city(name):
        return " ".join(w.capitalize() for w in name.lower().split("/")[0].split())
    def lookup_airport(name):
        return None

from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch
from reportlab.lib.colors import HexColor, white
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    KeepTogether, PageBreak
)
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_RIGHT, TA_CENTER

CLR_BLUE    = HexColor("#005e8d")
CLR_BLUE_LT = HexColor("#1a61a9")
CLR_DARK    = HexColor("#222222")
CLR_GREY    = HexColor("#555555")
CLR_WHITE   = white
CLR_BLACK   = HexColor("#000000")
CLR_LINE    = HexColor("#cccccc")

W = 540  # usable width


def _styles():
    s = {}
    s["pax"]        = ParagraphStyle("pax", fontName="Helvetica", fontSize=9.2, textColor=CLR_BLACK, leading=12)
    s["section"]    = ParagraphStyle("section", fontName="Helvetica", fontSize=9.5, textColor=CLR_WHITE, leading=13)
    s["route"]      = ParagraphStyle("route", fontName="Helvetica", fontSize=13, textColor=CLR_BLUE, leading=16, spaceBefore=4, spaceAfter=4)
    s["normal"]     = ParagraphStyle("normal", fontName="Helvetica", fontSize=9.5, textColor=CLR_DARK, leading=12)
    s["bold"]       = ParagraphStyle("bold", fontName="Helvetica-Bold", fontSize=9.5, textColor=CLR_DARK, leading=12)
    s["detail"]     = ParagraphStyle("detail", fontName="Helvetica", fontSize=9.5, textColor=CLR_GREY, leading=12)
    s["detail_b"]   = ParagraphStyle("detail_b", fontName="Helvetica-Bold", fontSize=9.5, textColor=CLR_GREY, leading=12)
    s["sub_label"]  = ParagraphStyle("sub_label", fontName="Helvetica-Bold", fontSize=8.5, textColor=CLR_BLUE_LT, leading=11, spaceBefore=6)
    s["sub_text"]   = ParagraphStyle("sub_text", fontName="Helvetica", fontSize=8.5, textColor=CLR_GREY, leading=11)
    s["notice"]     = ParagraphStyle("notice", fontName="Helvetica-Bold", fontSize=8.5, textColor=CLR_DARK, leading=11, spaceBefore=6)
    s["right"]      = ParagraphStyle("right", fontName="Helvetica", fontSize=9.2, textColor=CLR_BLACK, leading=12, alignment=TA_RIGHT)
    s["company"]    = ParagraphStyle("company", fontName="Helvetica", fontSize=7.5, textColor=CLR_GREY, leading=10, alignment=TA_CENTER)
    return s


def _bar(text, styles):
    t = Table([[Paragraph(text, styles["section"])]], colWidths=[W], rowHeights=[20])
    t.setStyle(TableStyle([("BACKGROUND", (0,0), (-1,-1), CLR_BLUE),
                           ("LEFTPADDING", (0,0), (-1,-1), 8),
                           ("TOPPADDING", (0,0), (-1,-1), 3),
                           ("BOTTOMPADDING", (0,0), (-1,-1), 3)]))
    return t


def _tc(s):
    return " ".join(w.capitalize() for w in s.lower().split())


def _fmt_time(raw):
    if not raw: return ""
    suffix = "AM" if raw.endswith("A") else "PM"
    digits = raw[:-1]
    if len(digits) <= 2: return f"{digits}:00 {suffix}"
    return f"{digits[:-2]}:{digits[-2:]} {suffix}"


def _fmt_date(date_raw, day_name):
    if not date_raw: return ""
    months = {"JAN":"January","FEB":"February","MAR":"March","APR":"April",
              "MAY":"May","JUN":"June","JUL":"July","AUG":"August",
              "SEP":"September","OCT":"October","NOV":"November","DEC":"December"}
    parts = date_raw.split()
    month = months.get(parts[1], parts[1])
    year = f"20{parts[2]}" if len(parts[2]) == 2 else parts[2]
    day_str = _tc(day_name) if day_name else ""
    return f"{day_str} {parts[0]} {month} {year}".strip()


def generate_invoice_pdf(data: dict, output_path: str):
    styles = _styles()
    story = []
    story.append(Spacer(1, 4))

    # ── Passengers + booking ──────────────────────────────────
    pax_lines = []
    for i, p in enumerate(data["passengers"], 1):
        name = f"{_tc(p['last_name'])}, {_tc(p['first_name'])}"
        if p.get("middle_name"):
            name += f" {_tc(p['middle_name'])}"
        ff = ""
        if data.get("mileage") and i == 1:
            ff = f' (FF: {data["mileage"]["airline"]}{data["mileage"]["number"]})'
        pax_lines.append(f"{i}. {name}{ff}")

    # Add frequent flyers for other passengers
    for ff in data.get("freq_flyers", [])[1:]:
        for j, p in enumerate(data["passengers"]):
            if p["full_slash"].startswith(ff["passenger"].split("/")[0]):
                idx = j + 1
                if idx <= len(pax_lines):
                    pax_lines[j] += f' (FF: {ff["airline"]}{ff["number"]})'

    booking = data.get("booking", {})
    right_parts = []
    if booking.get("record_locator"):
        right_parts.append(f'Record Locator: <b>{booking["record_locator"]}</b>')
    if booking.get("itin_no"):
        right_parts.append(f'ITIN: <b>{booking["itin_no"]}</b>')
    if booking.get("date"):
        right_parts.append(f'Date: {booking["date"]}')

    info = Table(
        [[Paragraph("<br/>".join(pax_lines), styles["pax"]),
          Paragraph("<br/>".join(right_parts), styles["right"])]],
        colWidths=[W * 0.6, W * 0.4])
    info.setStyle(TableStyle([("VALIGN", (0,0), (-1,-1), "TOP")]))
    story.append(info)

    # Mailing address
    if data.get("mailing_address"):
        story.append(Spacer(1, 4))
        addr = " | ".join(data["mailing_address"])
        story.append(Paragraph(f'<font color="#555555">{addr}</font>', styles["sub_text"]))

    story.append(Spacer(1, 10))

    # ── Flights ───────────────────────────────────────────────
    if data.get("flights"):
        story.append(_bar("Flight", styles))
        story.append(Spacer(1, 6))

        for fl in data["flights"]:
            elems = []
            dep = fl.get("departure_city", "")
            arr = fl.get("arrival_city", "")
            dep_info = lookup_airport(dep)
            arr_info = lookup_airport(arr)
            dep_display = resolve_city(dep)
            arr_display = resolve_city(arr)
            dep_city = dep_info["city"] if dep_info else _tc(dep.split("/")[0])
            arr_city = arr_info["city"] if arr_info else _tc(arr.split("/")[0])
            elems.append(Paragraph(f"{dep_city} - {arr_city}", styles["route"]))

            carrier = fl.get("airline_locator_carrier", "")
            fnum = fl.get("flight_number", "")
            conf = fl.get("airline_locator_code", "")
            status = "Confirmed" if fl.get("confirmed") else ""
            dep_time = _fmt_time(fl.get("departure_time"))
            arr_time = _fmt_time(fl.get("arrival_time"))
            dep_date = _fmt_date(fl.get("date_raw"), fl.get("day_name"))
            arr_date = dep_date
            if fl.get("arrives_next_day"):
                arr_date += f" (Arrives {fl['arrives_next_day']})"

            left = f'{carrier} {fnum}<br/>CF# {conf}<br/>{status}'
            if fl.get("operated_by"):
                left += f'<br/>Operated by: {_tc(fl["operated_by"])}'

            mid = f'<b>DEPART:</b><br/>{dep_display}<br/>{dep_date}, {dep_time}'
            if fl.get("dep_terminal"):
                mid += f'<br/>{_tc(fl["dep_terminal"])}'

            right = f'<b>ARRIVE:</b><br/>{arr_display}<br/>{arr_date}, {arr_time}'
            if fl.get("arr_terminal"):
                right += f'<br/>{_tc(fl["arr_terminal"])}'

            grid = Table(
                [[Paragraph(left, styles["normal"]),
                  Paragraph(mid, styles["normal"]),
                  Paragraph(right, styles["normal"])]],
                colWidths=[W*0.22, W*0.39, W*0.39])
            grid.setStyle(TableStyle([("VALIGN",(0,0),(-1,-1),"TOP"),
                                      ("TOPPADDING",(0,0),(-1,-1),2),
                                      ("BOTTOMPADDING",(0,0),(-1,-1),4)]))
            elems.append(grid)

            details = []
            if fl.get("duration"): details.append(f'<b>Duration:</b> {fl["duration"]}')
            if fl.get("cabin_class"): details.append(f'<b>Cabin:</b> {fl["cabin_class"]}')
            if fl.get("nonstop"): details.append("<b>Non-stop</b>")
            if fl.get("meals"): details.append(f'<b>Meals:</b> {_tc(fl["meals"])}')
            if fl.get("wheelchair"): details.append('<b>Wheelchair assist</b>')
            if details:
                elems.append(Paragraph("  |  ".join(details), styles["detail"]))

            pax_names = [p["full_slash"] for p in data["passengers"]]
            seats = fl.get("seats", [])
            if pax_names:
                rows = [[Paragraph("<b>Passengers</b>", styles["detail_b"]),
                         Paragraph("<b>Seat</b>", styles["detail_b"])]]
                for j, name in enumerate(pax_names):
                    seat = seats[j] if j < len(seats) else "-"
                    rows.append([Paragraph(name, styles["detail"]),
                                 Paragraph(seat, styles["detail"])])
                pt = Table(rows, colWidths=[W*0.5, W*0.2])
                pt.setStyle(TableStyle([("TOPPADDING",(0,0),(-1,-1),1),
                                        ("BOTTOMPADDING",(0,0),(-1,-1),1),
                                        ("LEFTPADDING",(0,0),(-1,-1),0)]))
                elems.append(Spacer(1, 4))
                elems.append(pt)

            elems.append(Spacer(1, 10))
            story.append(KeepTogether(elems))

    # ── Hotels ────────────────────────────────────────────────
    if data.get("hotels"):
        story.append(_bar("Hotel", styles))
        story.append(Spacer(1, 6))
        for h in data["hotels"]:
            elems = []
            name = h.get("name") or h.get("chain", "")
            elems.append(Paragraph(_tc(name), styles["route"]))

            left = f'<b>{_tc(name)}</b><br/>'
            if h.get("address"): left += f'{_tc(h["address"])}<br/>'
            if h.get("city"): left += f'{_tc(h["city"])}<br/>'
            contacts = []
            if h.get("phone"): contacts.append(f'Ph: {h["phone"]}')
            if h.get("fax"): contacts.append(f'Fax: {h["fax"]}')
            if contacts: left += " | ".join(contacts)

            right = ""
            if h.get("nights"): right += f'{h["nights"]} Night(s) | {h.get("status","")}<br/>'
            if h.get("confirmation"): right += f'<b>Confirmation:</b> {h["confirmation"]}<br/>'
            if h.get("guarantee"): right += f'<b>Guarantee:</b> {_tc(h["guarantee"])}<br/>'
            if h.get("rate_currency") and h.get("rate_amount"):
                right += f'<b>Rate:</b> {h["rate_currency"]} {h["rate_amount"]} /night'

            grid = Table([[Paragraph(left, styles["normal"]),
                           Paragraph(right, styles["normal"])]],
                         colWidths=[W*0.55, W*0.45])
            grid.setStyle(TableStyle([("VALIGN",(0,0),(-1,-1),"TOP")]))
            elems.append(grid)

            if h.get("approx_total"):
                elems.append(Paragraph(f'<b>Approximate Total (incl. taxes):</b> {h["approx_total"]}', styles["detail"]))
            if h.get("cancel_policy"):
                elems.append(Paragraph(h["cancel_policy"], styles["sub_text"]))
            for note in h.get("notes", []):
                elems.append(Paragraph(note, styles["sub_text"]))

            elems.append(Spacer(1, 10))
            story.append(KeepTogether(elems))

    # ── Cruises ───────────────────────────────────────────────
    if data.get("cruises"):
        story.append(_bar("Cruise", styles))
        story.append(Spacer(1, 6))
        for cr in data["cruises"]:
            elems = []
            d = cr.get("details", {})
            ship = d.get("ship", "Cruise")
            elems.append(Paragraph(_tc(ship), styles["route"]))

            left_parts = []
            if d.get("vendor"): left_parts.append(f'<b>{d["vendor"]}</b>')
            if d.get("ship"): left_parts.append(f'Ship: {d["ship"]}')
            if d.get("cabin"): left_parts.append(f'Cabin: {d["cabin"]}')
            if d.get("deck"): left_parts.append(f'Deck: {d["deck"]}')
            if d.get("port"): left_parts.append(f'Port: {_tc(d["port"])}')
            if d.get("dining"): left_parts.append(f'Dining: {_tc(d["dining"])}')

            right_parts = []
            if d.get("itinerary"): right_parts.append(f'<b>Itinerary:</b> {_tc(d["itinerary"])}')
            if d.get("depart_date"): right_parts.append(f'Depart: {d["depart_date"]}')
            if d.get("return_date"): right_parts.append(f'Return: {d["return_date"]}')
            if d.get("total_cost"): right_parts.append(f'<b>Total Cost:</b> USD {d["total_cost"]}')
            if d.get("balance"):
                bal = d["balance"]
                right_parts.append(f'Balance: {bal[0]} due {bal[1]}' if isinstance(bal, tuple) else f'Balance: {bal}')
            if d.get("confirmation"): right_parts.append(f'Confirmation: {d["confirmation"]}')

            grid = Table([[Paragraph("<br/>".join(left_parts), styles["normal"]),
                           Paragraph("<br/>".join(right_parts), styles["normal"])]],
                         colWidths=[W*0.5, W*0.5])
            grid.setStyle(TableStyle([("VALIGN",(0,0),(-1,-1),"TOP")]))
            elems.append(grid)
            elems.append(Spacer(1, 10))
            story.append(KeepTogether(elems))

    # ── Tours ─────────────────────────────────────────────────
    if data.get("tours"):
        story.append(_bar("Tour / Transportation", styles))
        story.append(Spacer(1, 6))
        for tr in data["tours"]:
            elems = []
            vendor = tr.get("vendor", "Tour")
            date = _fmt_date(tr.get("date_raw"), tr.get("day_name"))
            elems.append(Paragraph(f'{_tc(vendor)}', styles["route"]))

            parts = [f'<b>Date:</b> {date}']
            if tr.get("amount"): parts.append(f'<b>Amount:</b> USD {tr["amount"]}')
            if tr.get("confirmation"): parts.append(f'<b>Confirmation:</b> {tr["confirmation"]}')
            elems.append(Paragraph("  |  ".join(parts), styles["detail"]))

            for detail in tr.get("details", []):
                elems.append(Paragraph(detail, styles["sub_text"]))

            elems.append(Spacer(1, 8))
            story.append(KeepTogether(elems))

    # ── Packages ──────────────────────────────────────────────
    if data.get("packages"):
        story.append(_bar("Package", styles))
        story.append(Spacer(1, 6))
        for pk in data["packages"]:
            elems = []
            d = pk.get("details", {})
            vendor = d.get("vendor", "Package")
            elems.append(Paragraph(_tc(vendor), styles["route"]))

            parts = []
            if d.get("type"): parts.append(f'<b>Type:</b> {d["type"]}')
            if d.get("total_cost"): parts.append(f'<b>Total Cost:</b> USD {d["total_cost"]}')
            if d.get("amount"): parts.append(f'<b>Amount:</b> USD {d["amount"]}')
            if d.get("confirmation"): parts.append(f'Confirmation: {d["confirmation"]}')
            if parts:
                elems.append(Paragraph("  |  ".join(parts), styles["detail"]))

            for line in d.get("description", []):
                if not any(kw in line for kw in ["TOTAL COST", "PAYMENT BY", "TYPE OF PKG", "TOTAL DUE"]):
                    elems.append(Paragraph(line, styles["sub_text"]))

            elems.append(Spacer(1, 8))
            story.append(KeepTogether(elems))

    # ── Insurance (minimal) ───────────────────────────────────
    if data.get("insurance"):
        story.append(Spacer(1, 4))
        story.append(Paragraph("<b>Insurance</b>", styles["detail_b"]))
        for line in data["insurance"]:
            story.append(Paragraph(line, styles["sub_text"]))
        story.append(Spacer(1, 8))

    # ── Tickets / Air Fare ────────────────────────────────────
    if data.get("tickets"):
        story.append(_bar("Air Fare", styles))
        story.append(Spacer(1, 6))
        rows = [[Paragraph("<b>Passenger</b>", styles["detail_b"]),
                 Paragraph("<b>Ticket Number</b>", styles["detail_b"]),
                 Paragraph("<b>Payment</b>", styles["detail_b"]),
                 Paragraph("<b>Amount (USD)</b>", styles["detail_b"])]]
        for t in data["tickets"]:
            rows.append([Paragraph(t.get("passenger",""), styles["detail"]),
                         Paragraph(t.get("ticket_number",""), styles["detail"]),
                         Paragraph(t.get("payment_method",""), styles["detail"]),
                         Paragraph(t.get("amount_usd",""), styles["detail"])])
        tbl = Table(rows, colWidths=[W*0.30, W*0.28, W*0.20, W*0.22])
        tbl.setStyle(TableStyle([("TOPPADDING",(0,0),(-1,-1),2),
                                 ("BOTTOMPADDING",(0,0),(-1,-1),2),
                                 ("LINEBELOW",(0,0),(-1,0),0.5,CLR_LINE),
                                 ("LEFTPADDING",(0,0),(-1,-1),0)]))
        story.append(tbl)
        story.append(Spacer(1, 6))

    # ── Financial ─────────────────────────────────────────────
    fin = data.get("financial", {})
    fin_rows = []
    if fin.get("fare_per_person"): fin_rows.append(["Fare:", f'USD {fin["fare_per_person"]}'])
    if fin.get("air_fare"): fin_rows.append(["Air Fare:", f'USD {fin["air_fare"]}'])
    if fin.get("tax_and_fees"): fin_rows.append(["Tax and Carrier Fees:", f'USD {fin["tax_and_fees"]}'])
    if fin.get("total"): fin_rows.append(["Total:", f'USD {fin["total"]}'])
    if data.get("service_fee"): fin_rows.append(["Service Fee:", f'USD {data["service_fee"]}'])
    if fin.get("sub_total"): fin_rows.append(["Sub Total:", f'USD {fin["sub_total"]}'])
    if fin.get("credit_card_payment"): fin_rows.append(["Credit Card Payment:", f'USD {fin["credit_card_payment"]}'])
    if fin.get("amount_due") is not None: fin_rows.append(["Amount Due:", f'USD {fin["amount_due"]}'])

    if fin_rows:
        tbl_rows = []
        for label, val in fin_rows:
            is_due = "Amount Due" in label
            sl = styles["bold"] if is_due else styles["detail"]
            tbl_rows.append([Paragraph(label, sl), Paragraph(val, sl)])
        ft = Table(tbl_rows, colWidths=[W*0.55, W*0.25])
        ft.setStyle(TableStyle([("ALIGN",(1,0),(1,-1),"RIGHT"),
                                ("TOPPADDING",(0,0),(-1,-1),1),
                                ("BOTTOMPADDING",(0,0),(-1,-1),1),
                                ("LEFTPADDING",(0,0),(-1,-1),0),
                                ("LINEABOVE",(0,-1),(-1,-1),0.5,CLR_LINE)]))
        wrapper = Table([[Spacer(1,1), ft]], colWidths=[W*0.2, W*0.8])
        wrapper.setStyle(TableStyle([("VALIGN",(0,0),(-1,-1),"TOP")]))
        story.append(wrapper)

        if fin.get("fare_note"):
            story.append(Paragraph(fin["fare_note"], styles["detail"]))
        story.append(Spacer(1, 10))

    # ── Baggage ───────────────────────────────────────────────
    if data.get("baggage") or data.get("carry_on"):
        story.append(_bar("Baggage Allowance", styles))
        story.append(Spacer(1, 6))
        for b in data.get("baggage", []):
            story.append(Paragraph(f'<b>{b["route"]}</b>  {b["count"]}', styles["detail"]))
            for bag in b.get("bags", []):
                story.append(Paragraph(f'&nbsp;&nbsp;Bag {bag["bag_num"]}: {bag["info"]}', styles["sub_text"]))
        if data.get("carry_on"):
            story.append(Spacer(1, 4))
            story.append(Paragraph("Carry On Allowance:", styles["sub_label"]))
            for b in data["carry_on"]:
                story.append(Paragraph(f'<b>{b["route"]}</b>  {b["count"]}', styles["detail"]))
                for bag in b.get("bags", []):
                    story.append(Paragraph(f'&nbsp;&nbsp;Bag {bag["bag_num"]}: {bag["info"]}', styles["sub_text"]))
        story.append(Spacer(1, 10))

    # ── Notices ───────────────────────────────────────────────
    for notice in data.get("notices", []):
        story.append(Paragraph(f"** {notice} **", styles["notice"]))

    # ── Build ─────────────────────────────────────────────────
    doc = SimpleDocTemplate(output_path, pagesize=letter,
                            leftMargin=36, rightMargin=36,
                            topMargin=65, bottomMargin=75)
    doc.build(story)
    return output_path


if __name__ == "__main__":
    from state_parser import parse
    import json
    path = sys.argv[1] if len(sys.argv) > 1 else None
    if not path:
        print("Usage: python invoice_generator.py <invoice.pdf> [output.pdf]")
        sys.exit(1)
    out = sys.argv[2] if len(sys.argv) > 2 else "output.pdf"
    data = parse(path)
    generate_invoice_pdf(data, out)
    print(f"Generated: {out}")