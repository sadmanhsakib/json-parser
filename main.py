import json
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

DARK_BLUE = "1F3864"
MID_BLUE = "2E75B6"
LIGHT_BLUE = "B4C7DC"
WHITE = "FFFFFF"
LIGHT_GREY = "F5F5F5"
LIGHT_GREEN = "AFD095"
LIGHT_RED = "FFA6A6"
LIGHT_YELLOW = "FFFFA6"


def main():
    # loading the json file
    with open("hotel_bookings.json", "r") as file:
        data = json.load(file)

    rows = flatten_reservations(data) 

    if not rows:
        raise ValueError("No data to write")

    wb = write_to_excel(rows)
    wb = style_sheet(wb, list(rows[0].keys()))
    wb = write_summary(wb, rows)
    wb.save("hotel_report.xlsx")


def flatten_reservations(data: dict) -> list[dict]:
    try:
        rows = []
        for reservation in data["reservations"]:
            row = {
                "Reservation ID": reservation["reservation_id"],
                "Status": reservation["status"].replace("_", " ").title(),
                "Guest Name": reservation["guest"]["full_name"],
                "Email": reservation["guest"]["email"],
                "Nationality": reservation["guest"]["nationality"],
                "Loyalty Tier": reservation["guest"]["loyalty_tier"],
                "Room Number": reservation["room"]["room_number"],
                "Room Type": reservation["room"]["type"].replace("_", " ").title(),
                "Floor": reservation["room"]["floor"],
                "Beds": reservation["room"]["beds"],
                "Check-in": reservation["dates"]["check_in"],
                "Check-out": reservation["dates"]["check_out"],
                "Nights": reservation["pricing"]["nights"],
                "Rate/Night": reservation["pricing"]["rate_per_night"],
                "Discount %": reservation["pricing"]["discount_pct"],
                "Extras Total": sum(extra["charge"] for extra in reservation["extras"]),
                "Taxes": reservation["pricing"]["taxes"],
                "Total Charged": reservation["pricing"]["total_charged"],
                "Payment Method": reservation["payment_method"].replace("_", "").title(),
                "Notes": reservation["notes"] or "",
            }
            rows.append(row)
        return rows
    except Exception as error:
        print(error)
        return None


def write_to_excel(rows: list[dict]) -> Workbook:
    wb = Workbook()
    ws = wb.active
    ws.title = "Reservations"

    # writing the headers
    headers = list(rows[0].keys())
    ws.append(headers)

    # writing the rows
    for row in rows:
        ws.append([row.get(header, "") for header in headers])

    return wb


def style_sheet(wb: Workbook, headers: list) -> Workbook:
    ws = wb.active

    # adding the header
    for col_num, value in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_num)
        make_header_cell(cell, value)

    ws.row_dimensions[1].height = 30

    # Alternating row colors
    white_fill = PatternFill(fill_type="solid", fgColor=WHITE)
    light_grey_fill = PatternFill(fill_type="solid", fgColor=LIGHT_GREY)

    # Status color coding
    status_col = headers.index("Status") + 1

    status_colors = {
        "Checked Out": LIGHT_GREEN,
        "Checked In": LIGHT_BLUE,
        "Confirmed": LIGHT_YELLOW,
        "Cancelled": LIGHT_RED,
    }

    for row in range(2, ws.max_row + 1):
        # checking for even and odd row number
        fill = white_fill if row % 2 == 0 else light_grey_fill
            
        # adding the row color change
        for col in range(1, ws.max_column + 1):
            ws.cell(row=row, column=col).fill = fill

        status_cell = ws.cell(row=row, column=status_col)
        status_value = status_cell.value

        # changing the color for each cell in status column
        if status_value in status_colors:
            status_cell.fill = PatternFill(fill_type="solid", fgColor=status_colors[status_value])

    # Column widths
    column_widths = {
        "Reservation ID": 15,
        "Status": 12,
        "Guest Name": 20,
        "Email": 25,
        "Nationality": 12,
        "Loyalty Tier": 12,
        "Room Number": 12,
        "Room Type": 15,
        "Floor": 8,
        "Beds": 8,
        "Check-in": 12,
        "Check-out": 12,
        "Nights": 8,
        "Rate/Night": 12,
        "Discount %": 12,
        "Extras Total": 12,
        "Taxes": 10,
        "Total Charged": 14,
        "Payment Method": 18,
        "Notes": 60,
    }

    # freezes the rows before A2 (the header row)
    # when scrolling down the header row always remain visible
    ws.freeze_panes = "A2"
    
    # using the header row as a reference for auto_filler function
    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}1"

    # Number formats
    currency_columns = ["Rate/Night", "Extras Total", "Taxes", "Total Charged"]
    percent_column = "Discount %"

    for col_num, header in enumerate(headers, 1):
        if header in column_widths:
            ws.column_dimensions[get_column_letter(col_num)].width = column_widths[
                header
            ]
        if header in currency_columns:
            for row in range(2, ws.max_row + 1):
                ws.cell(row=row, column=col_num).number_format = "$#,##0.00"
        elif header == percent_column:
            for row in range(2, ws.max_row + 1):
                ws.cell(row=row, column=col_num).number_format = '0"%"'

    return wb


def make_header_cell(cell, text):
    """Style a single header cell: dark blue bg, white bold text, centered."""
    cell.value = text
    cell.font = Font(name="Noto Sans Lisu", bold=True, color=WHITE, size=10)
    cell.fill = PatternFill(fill_type="solid", fgColor=DARK_BLUE)
    cell.alignment = Alignment(horizontal="center", vertical="center")


def write_summary(wb: Workbook, rows: list[dict]) -> Workbook:
    ws = wb.create_sheet("Summary")
    ws.sheet_view.showGridLines = False

    total_reservations = len(rows)
    total_revenue = sum(
        row["Total Charged"] for row in rows if row["Status"] != "Cancelled"
    )
    active_guests = sum(
        1 for row in rows if row["Status"] in ("Checked In", "Confirmed")
    )
    avg_stay = round(sum(row["Nights"] for row in rows) / total_reservations, 1)

    room_stats = {}

    for row in rows:
        room_type = row["Room Type"]

        # First time seeing this room type — initialize its bucket
        if room_type not in room_stats:
            room_stats[room_type] = {
                "Reservations": 0,
                "Total Revenue": 0.0,
                "Rates": [],
            }

        room_stats[room_type]["Reservations"] += 1
        room_stats[room_type]["Total Revenue"] += row["Total Charged"]
        room_stats[room_type]["Rates"].append(row["Rate/Night"])

    guest_stats = {}

    for row in rows:
        name = row["Guest Name"]

        if name not in guest_stats:
            guest_stats[name] = {
                "Stays": 0,
                "Total Spent": 0.0,
                "Loyalty Tier": row["Loyalty Tier"]
            }
        guest_stats[name]["Stays"] += 1
        guest_stats[name]["Total Spent"] += row["Total Charged"]
        # Note: Loyalty Tier is set on first encounter.
        # Safe here because the same guest always has the same tier in this dataset.

    for label, value, label_cell, value_cell in [
        ("Total Reservations", total_reservations, "B1", "B2"),
        ("Total Revenue", f"${total_revenue:,.2f}", "D1", "D2"),
        ("Active Guests", active_guests, "B4", "B5"),
        ("Avg Stay (nights)", round(avg_stay, 1), "D4", "D5"),
    ]:
        lc = ws[label_cell]
        vc = ws[value_cell]

        lc.value = label
        lc.font = Font(name="Noto Sans Lisu", bold=True, size=10, color=WHITE)
        lc.fill = PatternFill("solid", fgColor=DARK_BLUE)
        lc.alignment = Alignment(horizontal="center")

        vc.value = value
        vc.font = Font(name="Bahnschrift", bold=True, size=18, color=DARK_BLUE)
        vc.fill = PatternFill("solid", fgColor=LIGHT_BLUE)
        vc.alignment = Alignment(horizontal="center")

    room_headers = ["Room Type", "Reservations", "Total Revenue", "Avg Rate/Night"]
    # Starts at row 9 (rows 6-8 act as visual breathing room)
    start_row = 9

    # Write header row
    for column_index, header in enumerate(
        room_headers, 2
    ):  # start at column B (index 2)
        make_header_cell(ws.cell(row=start_row, column=column_index), header)

    # Sort room types by Total Revenue descending — most valuable room type first
    sorted_rooms = sorted(
        room_stats.items(), key=lambda x: x[1]["Total Revenue"], reverse=True
    )

    for row_index, (room_type, stats) in enumerate(sorted_rooms, start_row + 1):
        avg_rate = sum(stats["Rates"]) / len(stats["Rates"])  # average from the list we built
        bg = WHITE if row_index % 2 == 0 else LIGHT_GREY  # alternating rows

        data = [
            room_type,
            stats["Reservations"],
            f"${stats['Total Revenue']:,.2f}",
            f"${avg_rate:,.2f}",
        ]
        for column_index, val in enumerate(data, 2):
            cell = ws.cell(row=row_index, column=column_index)
            cell.value = val
            cell.fill = PatternFill("solid", fgColor=bg)
            cell.font = Font(name="Bahnschrift", size=10)
            cell.alignment = Alignment(
                horizontal="right" if column_index > 2 else "left"
            )

    guest_headers = ["Guest Name", "Stays", "Total Spent", "Loyalty Tier"]
    # Starts at row 16 (leaves a gap after the room type table)
    guest_start = 16

    for column_index, header in enumerate(guest_headers, 2):
        make_header_cell(ws.cell(row=guest_start, column=column_index), header
    )

    # Sort by Total Spent descending — biggest spender first
    sorted_guests = sorted(
        guest_stats.items(), key=lambda x: x[1]["Total Spent"], reverse=True
    )

    for row_index, (name, stats) in enumerate(sorted_guests, guest_start + 1):
        bg = WHITE if row_index % 2 == 0 else LIGHT_GREY
        data = [
            name,
            stats["Stays"],
            f"${stats['Total Spent']:,.2f}",
            stats["Loyalty Tier"],
        ]
        for column_index, val in enumerate(data, 2):
            cell = ws.cell(row=row_index, column=column_index)
            cell.value = val
            cell.fill = PatternFill("solid", fgColor=bg)
            cell.font = Font(name="Bahnschrift", size=10)
            cell.alignment = Alignment(
                horizontal="right" if column_index in (3, 4) else "left"
            )

    # ── STEP 5: Column widths ─────────────────────────────────────────────
    # B through E covers all our tables (we started at column 2)
    for col, width in {"B": 22, "C": 15, "D": 16, "E": 14}.items():
        ws.column_dimensions[col].width = width

    return wb

if __name__ == "__main__":
    main()
