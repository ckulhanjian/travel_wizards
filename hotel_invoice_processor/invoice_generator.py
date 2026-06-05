"""
invoice_generator.py - Generate a clean, professional PDF invoice
from parsed old-format data. Only uses data present in the source.

Style reference: Travelport+ layout (simplified).
- Company header with logo placeholder
- Blue (#005e8d) section headers with white text on blue bar
- Dark text (#222222) for main content
- Grey (#555555) for secondary/detail content
- Blue (#1a61a9) for sub-section labels
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

# ── Colours ────────────────────────────────────────────────────
CLR_BLUE     = HexColor("#005e8d")
CLR_BLUE_LT  = HexColor("#1a61a9")
CLR_DARK     = HexColor("#222222")
CLR_GREY     = HexColor("#555555")
CLR_LIGHT_BG = HexColor("#f0f5f8")
CLR_WHITE    = white
CLR_BLACK    = HexColor("#000000")
CLR_LINE     = HexColor("#cccccc")


# ── Styles ─────────────────────────────────────────────────────
def _make_styles():
    s = {}
    s["company"] = ParagraphStyle(
        "company", fontName="Helvetica", fontSize=7.5,
        textColor=CLR_GREY, leading=10, alignment=TA_CENTER,
    )
    s["pax"] = ParagraphStyle(
        "pax", fontName="Helvetica", fontSize=9.2,
        textColor=CLR_BLACK, leading=12,
    )
    s["pax_bold"] = ParagraphStyle(
        "pax_bold", fontName="Helvetica-Bold", fontSize=9.2,
        textColor=CLR_BLACK, leading=12,
    )
    s["section_bar"] = ParagraphStyle(
        "section_bar", fontName="Helvetica", fontSize=9.5,
        textColor=CLR_WHITE, leading=13,
    )
    s["route_title"] = ParagraphStyle(
        "route_title", fontName="Helvetica", fontSize=13,
        textColor=CLR_BLUE, leading=16, spaceBefore=4, spaceAfter=4,
    )
    s["normal"] = ParagraphStyle(
        "normal", fontName="Helvetica", fontSize=9.5,
        textColor=CLR_DARK, leading=12,
    )
    s["normal_bold"] = ParagraphStyle(
        "normal_bold", fontName="Helvetica-Bold", fontSize=9.5,
        textColor=CLR_DARK, leading=12,
    )
    s["detail"] = ParagraphStyle(
        "detail", fontName="Helvetica", fontSize=9.5,
        textColor=CLR_GREY, leading=12,
    )
    s["detail_bold"] = ParagraphStyle(
        "detail_bold", fontName="Helvetica-Bold", fontSize=9.5,
        textColor=CLR_GREY, leading=12,
    )
    s["sub_label"] = ParagraphStyle(
        "sub_label", fontName="Helvetica-Bold", fontSize=8.5,
        textColor=CLR_BLUE_LT, leading=11, spaceBefore=6,
    )
    s["sub_text"] = ParagraphStyle(
        "sub_text", fontName="Helvetica", fontSize=8.5,
        textColor=CLR_GREY, leading=11,
    )
    s["notice"] = ParagraphStyle(
        "notice", fontName="Helvetica-Bold", fontSize=8.5,
        textColor=CLR_DARK, leading=11, spaceBefore=6,
    )
    s["footer"] = ParagraphStyle(
        "footer", fontName="Helvetica", fontSize=8,
        textColor=CLR_GREY, leading=10, alignment=TA_CENTER,
    )
    s["right"] = ParagraphStyle(
        "right", fontName="Helvetica", fontSize=9.2,
        textColor=CLR_BLACK, leading=12, alignment=TA_RIGHT,
    )
    s["right_bold"] = ParagraphStyle(
        "right_bold", fontName="Helvetica-Bold", fontSize=9.2,
        textColor=CLR_BLACK, leading=12, alignment=TA_RIGHT,
    )
    return s


def _section_bar(text, styles):
    """Blue bar with white text — section header."""
    t = Table(
        [[Paragraph(text, styles["section_bar"])]],
        colWidths=[540],
        rowHeights=[20],
    )
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), CLR_BLUE),
        ("LEFTPADDING", (0, 0), (-1, -1), 8),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
    ]))
    return t


def _title_case(s):
    """Convert 'JOHN DANIEL' to 'John Daniel'."""
    return " ".join(w.capitalize() for w in s.lower().split())


def _format_time(raw):
    """Convert '751A' to '7:51 AM', '343P' to '3:43 PM'."""
    if not raw:
        return ""
    raw = raw.strip()
    suffix = "AM" if raw.endswith("A") else "PM"
    digits = raw[:-1]
    if len(digits) <= 2:
        return f"{digits}:00 {suffix}"
    minutes = digits[-2:]
    hours = digits[:-2]
    return f"{hours}:{minutes} {suffix}"


def _format_date(date_raw, day_name):
    """Convert '22 JUL 26' + 'WEDNESDAY' to 'Wednesday 22 July 2026'."""
    if not date_raw:
        return ""
    months = {
        "JAN": "January", "FEB": "February", "MAR": "March",
        "APR": "April", "MAY": "May", "JUN": "June",
        "JUL": "July", "AUG": "August", "SEP": "September",
        "OCT": "October", "NOV": "November", "DEC": "December",
    }
    parts = date_raw.split()
    day_num = parts[0]
    month = months.get(parts[1], parts[1])
    year = f"20{parts[2]}" if len(parts[2]) == 2 else parts[2]
    day_str = _title_case(day_name) if day_name else ""
    return f"{day_str} {day_num} {month} {year}".strip()


def generate_invoice_pdf(data: dict, output_path: str):
    """Generate a professional PDF invoice from parsed data."""
    styles = _make_styles()
    story = []

    W = 540  # usable width

    # Space reserved for overlay logo (top) and footer (bottom)
    # No header drawn here — the overlay.pdf provides the branding
    story.append(Spacer(1, 4))

    # ── Passenger list + booking info ─────────────────────────
    pax_lines = []
    for i, p in enumerate(data["passengers"], 1):
        name = f"{_title_case(p['last_name'])}, {_title_case(p['first_name'])}"
        if p.get("middle_name"):
            name += f" {_title_case(p['middle_name'])}"
        # Add FF number if mileage matches this passenger
        ff = ""
        if data.get("mileage") and i == 1:
            ff = f' (FF: {data["mileage"]["airline"]}{data["mileage"]["number"]})'
        pax_lines.append(f"{i}. {name}{ff}")

    booking = data.get("booking", {})
    rec_loc = booking.get("record_locator", "")
    itin_no = booking.get("itin_no", "")

    left_col = "<br/>".join(pax_lines)
    right_col = (
        f'Record Locator: <b>{rec_loc}</b><br/>'
        f'ITIN: <b>{itin_no}</b>'
    )

    info_table = Table(
        [[Paragraph(left_col, styles["pax"]),
          Paragraph(right_col, styles["right"])]],
        colWidths=[W * 0.6, W * 0.4],
    )
    info_table.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
    ]))
    story.append(info_table)
    story.append(Spacer(1, 10))

    # ── Flights ───────────────────────────────────────────────
    if data.get("flights"):
        story.append(_section_bar("Flight", styles))
        story.append(Spacer(1, 6))

        for fl in data["flights"]:
            flight_elems = []

            dep = fl.get("departure_city", "")
            arr = fl.get("arrival_city", "")
            dep_info = lookup_airport(dep)
            arr_info = lookup_airport(arr)
            dep_display = resolve_city(dep)
            arr_display = resolve_city(arr)
            dep_city = dep_info["city"] if dep_info else _title_case(dep.split("/")[0])
            arr_city = arr_info["city"] if arr_info else _title_case(arr.split("/")[0])
            route = f"{dep_city} - {arr_city}"
            flight_elems.append(Paragraph(route, styles["route_title"]))

            # Flight info grid
            carrier = fl.get("airline_locator_carrier", "")
            fnum = fl.get("flight_number", "")
            conf_code = fl.get("airline_locator_code", "")
            status = "Confirmed" if fl.get("confirmed") else ""
            op_by = fl.get("operated_by")

            dep_time = _format_time(fl.get("departure_time"))
            arr_time = _format_time(fl.get("arrival_time"))
            dep_date = _format_date(fl.get("date_raw"), fl.get("day_name"))
            arr_date = dep_date
            if fl.get("arrives_next_day"):
                arr_date += f" (Arrives {fl['arrives_next_day']})"

            left_info = (
                f'{carrier} {fnum}<br/>'
                f'CF# {conf_code}<br/>'
                f'{status}'
            )
            if op_by:
                left_info += f'<br/>Operated by: {_title_case(op_by)}'

            mid_info = (
                f'<b>DEPART:</b><br/>'
                f'{dep_display}<br/>'
                f'{dep_date}, {dep_time}'
            )
            if fl.get("dep_terminal"):
                mid_info += f'<br/>{_title_case(fl["dep_terminal"])}'

            right_info = (
                f'<b>ARRIVE:</b><br/>'
                f'{arr_display}<br/>'
                f'{arr_date}, {arr_time}'
            )
            if fl.get("arr_terminal"):
                right_info += f'<br/>{_title_case(fl["arr_terminal"])}'

            grid = Table(
                [[Paragraph(left_info, styles["normal"]),
                  Paragraph(mid_info, styles["normal"]),
                  Paragraph(right_info, styles["normal"])]],
                colWidths=[W * 0.22, W * 0.39, W * 0.39],
            )
            grid.setStyle(TableStyle([
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("TOPPADDING", (0, 0), (-1, -1), 2),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
            ]))
            flight_elems.append(grid)

            # Duration / cabin / nonstop
            details = []
            if fl.get("duration"):
                details.append(f'<b>Duration:</b> {fl["duration"]}')
            if fl.get("cabin_class"):
                details.append(f'<b>Cabin:</b> {fl["cabin_class"]}')
            if fl.get("nonstop"):
                details.append("<b>Non-stop</b>")
            if fl.get("meals"):
                details.append(f'<b>Meals:</b> {_title_case(fl["meals"])}')
            if details:
                flight_elems.append(Paragraph(
                    "  |  ".join(details), styles["detail"]
                ))

            # Passengers & seats
            pax_names = [p["full_slash"] for p in data["passengers"]]
            seats = fl.get("seats", [])
            if pax_names:
                rows = [
                    [Paragraph("<b>Passengers</b>", styles["detail_bold"]),
                     Paragraph("<b>Seat</b>", styles["detail_bold"])]
                ]
                for j, name in enumerate(pax_names):
                    seat = seats[j] if j < len(seats) else "-"
                    rows.append([
                        Paragraph(name, styles["detail"]),
                        Paragraph(seat, styles["detail"]),
                    ])
                pax_tbl = Table(rows, colWidths=[W * 0.5, W * 0.2])
                pax_tbl.setStyle(TableStyle([
                    ("TOPPADDING", (0, 0), (-1, -1), 1),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 1),
                    ("LEFTPADDING", (0, 0), (-1, -1), 0),
                ]))
                flight_elems.append(Spacer(1, 4))
                flight_elems.append(pax_tbl)

            flight_elems.append(Spacer(1, 10))
            story.append(KeepTogether(flight_elems))

    # ── Hotels ────────────────────────────────────────────────
    if data.get("hotels"):
        story.append(_section_bar("Hotel", styles))
        story.append(Spacer(1, 6))

        for h in data["hotels"]:
            hotel_elems = []
            name = h.get("name") or h.get("chain", "")
            hotel_elems.append(Paragraph(
                _title_case(name), styles["route_title"]
            ))

            # Hotel info grid
            left = f'<b>{_title_case(name)}</b><br/>'
            if h.get("address"):
                left += f'{_title_case(h["address"])}<br/>'
            if h.get("city"):
                left += f'{_title_case(h["city"])}<br/>'
            contacts = []
            if h.get("phone"):
                contacts.append(f'Ph: {h["phone"]}')
            if h.get("fax"):
                contacts.append(f'Fax: {h["fax"]}')
            if contacts:
                left += " | ".join(contacts)

            right = ""
            if h.get("nights"):
                right += f'{h["nights"]} Night(s) | {h["status"]}<br/>'
            if h.get("confirmation"):
                right += f'<b>Confirmation:</b> {h["confirmation"]}<br/>'
            if h.get("guarantee"):
                right += f'<b>Guarantee:</b> {_title_case(h["guarantee"])}<br/>'
            if h.get("rate_currency") and h.get("rate_amount"):
                right += f'<b>Rate:</b> {h["rate_currency"]} {h["rate_amount"]} /night'

            grid = Table(
                [[Paragraph(left, styles["normal"]),
                  Paragraph(right, styles["normal"])]],
                colWidths=[W * 0.55, W * 0.45],
            )
            grid.setStyle(TableStyle([
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ]))
            hotel_elems.append(grid)

            if h.get("approx_total"):
                hotel_elems.append(Spacer(1, 2))
                hotel_elems.append(Paragraph(
                    f'<b>Approximate Total (incl. taxes):</b> {h["approx_total"]}',
                    styles["detail"]
                ))

            if h.get("notes"):
                hotel_elems.append(Paragraph("Room Info:", styles["sub_label"]))
                for note in h["notes"]:
                    hotel_elems.append(Paragraph(note, styles["sub_text"]))

            hotel_elems.append(Spacer(1, 10))
            story.append(KeepTogether(hotel_elems))

    # ── Tickets / Air Fare ────────────────────────────────────
    if data.get("tickets"):
        story.append(_section_bar("Air Fare", styles))
        story.append(Spacer(1, 6))

        rows = [[
            Paragraph("<b>Passenger</b>", styles["detail_bold"]),
            Paragraph("<b>Ticket Number</b>", styles["detail_bold"]),
            Paragraph("<b>Payment</b>", styles["detail_bold"]),
            Paragraph("<b>Amount (USD)</b>", styles["detail_bold"]),
        ]]
        for t in data["tickets"]:
            rows.append([
                Paragraph(t["passenger"], styles["detail"]),
                Paragraph(t["ticket_number"], styles["detail"]),
                Paragraph(t["payment_method"], styles["detail"]),
                Paragraph(t["amount_usd"], styles["detail"]),
            ])

        tkt_tbl = Table(rows, colWidths=[W * 0.30, W * 0.28, W * 0.20, W * 0.22])
        tkt_tbl.setStyle(TableStyle([
            ("TOPPADDING", (0, 0), (-1, -1), 2),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
            ("LINEBELOW", (0, 0), (-1, 0), 0.5, CLR_LINE),
            ("LEFTPADDING", (0, 0), (-1, -1), 0),
        ]))
        story.append(tkt_tbl)
        story.append(Spacer(1, 6))

    # ── Financial summary ─────────────────────────────────────
    fin = data.get("financial", {})
    if fin.get("air_fare"):
        fin_rows = []
        if fin.get("air_fare"):
            fin_rows.append(["Air Fare:", f'USD {fin["air_fare"]}'])
        if fin.get("tax_and_fees"):
            fin_rows.append(["Tax and Carrier Fees:", f'USD {fin["tax_and_fees"]}'])
        if fin.get("total"):
            fin_rows.append(["Total:", f'USD {fin["total"]}'])
        if data.get("service_fee"):
            fin_rows.append(["Service Fee:", f'USD {data["service_fee"]}'])
        if fin.get("sub_total"):
            fin_rows.append(["Sub Total:", f'USD {fin["sub_total"]}'])
        if fin.get("credit_card_payment"):
            fin_rows.append(["Credit Card Payment:", f'USD {fin["credit_card_payment"]}'])
        if fin.get("amount_due"):
            fin_rows.append(["Amount Due:", f'USD {fin["amount_due"]}'])

        fin_table_rows = []
        for label, val in fin_rows:
            is_due = "Amount Due" in label
            s_label = styles["normal_bold"] if is_due else styles["detail"]
            s_val = styles["normal_bold"] if is_due else styles["detail"]
            fin_table_rows.append([
                Paragraph(label, s_label),
                Paragraph(val, s_val),
            ])

        fin_tbl = Table(fin_table_rows, colWidths=[W * 0.55, W * 0.25])
        fin_tbl.setStyle(TableStyle([
            ("ALIGN", (1, 0), (1, -1), "RIGHT"),
            ("TOPPADDING", (0, 0), (-1, -1), 1),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 1),
            ("LEFTPADDING", (0, 0), (-1, -1), 0),
            ("LINEABOVE", (0, -1), (-1, -1), 0.5, CLR_LINE),
        ]))

        # Right-align the whole financial block
        wrapper = Table([[Spacer(1, 1), fin_tbl]], colWidths=[W * 0.2, W * 0.8])
        wrapper.setStyle(TableStyle([
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ]))
        story.append(wrapper)
        story.append(Spacer(1, 10))

    # ── Baggage ───────────────────────────────────────────────
    if data.get("baggage") or data.get("carry_on"):
        story.append(_section_bar("Baggage Allowance", styles))
        story.append(Spacer(1, 6))

        for b in data.get("baggage", []):
            story.append(Paragraph(
                f'<b>{b["route"]}</b>  {b["count"]}', styles["detail"]
            ))
            for bag in b.get("bags", []):
                story.append(Paragraph(
                    f'&nbsp;&nbsp;Bag {bag["bag_num"]}: {bag["info"]}',
                    styles["sub_text"]
                ))

        if data.get("carry_on"):
            story.append(Spacer(1, 4))
            story.append(Paragraph("Carry On Allowance:", styles["sub_label"]))
            for b in data["carry_on"]:
                story.append(Paragraph(
                    f'<b>{b["route"]}</b>  {b["count"]}', styles["detail"]
                ))
                for bag in b.get("bags", []):
                    story.append(Paragraph(
                        f'&nbsp;&nbsp;Bag {bag["bag_num"]}: {bag["info"]}',
                        styles["sub_text"]
                    ))

        story.append(Spacer(1, 10))

    # ── Notices ───────────────────────────────────────────────
    for notice in data.get("notices", []):
        story.append(Paragraph(f"** {notice} **", styles["notice"]))

    # ── Build ─────────────────────────────────────────────────
    doc = SimpleDocTemplate(
        output_path,
        pagesize=letter,
        leftMargin=36, rightMargin=36,
        topMargin=65, bottomMargin=75,
        # Top margin clears overlay logo (ends at y~43pt)
        # Bottom margin clears overlay footer (starts at y~738pt)
    )

    doc.build(story)
    return output_path


if __name__ == "__main__":
    from invoice_parser import parse_invoice
    import json

    path = sys.argv[1] if len(sys.argv) > 1 else "/mnt/user-data/uploads/OTHER_639137689954244936.pdf"
    out = sys.argv[2] if len(sys.argv) > 2 else "/home/claude/output_invoice.pdf"

    data = parse_invoice(path)
    print(json.dumps(data, indent=2)[:200] + "...")
    print(f"\nGenerating PDF: {out}")
    generate_invoice_pdf(data, out)
    print(f"Done: {out}")