import time, json
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

colors = {
    "dark-blue": "1F3864",
    "white": "FFFFFF",
    "light-grey": "F5F5F5",
    "light-green": "C6EFCE",
    "light-blue": "DDEBF7",
    "light-yellow": "FFF2CC",
    "light-red": "FCE4D6",
}

def main():
    # loading the json file
    with open("hotel_bookings.json", "r") as file:
        data = json.load(file)

    rows = flatten_reservations(data)

    if not rows:
        raise ValueError("No data to write")

    write_to_excel(rows, "hotel_report.xlsx")


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
                "Room Type": reservation["room"]["type"],
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
                "Payment Method": reservation["payment_method"],
                "Notes": reservation["notes"] or "",
            }
            rows.append(row)
        return rows
    except Exception as error:
        print(error)
        return None


def write_to_excel(rows: list[dict], filename: str) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Reservations"

    # writing the headers
    headers = list(rows[0].keys())
    ws.append(headers)

    # writing the rows
    for row in rows:
        ws.append([row.get(header, "") for header in headers])

    style_sheet(ws, headers)

    wb.save(filename)


def style_sheet(ws, headers):
    # 3a - Header row styling
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color=colors["dark-blue"], end_color=colors["dark-blue"], fill_type="solid")
    header_alignment = Alignment(horizontal="center", vertical="center")

    for col_num, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_num)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_alignment

    ws.row_dimensions[1].height = 30

    # 3b - Alternating row colors
    white_fill = PatternFill(start_color=colors["white"], end_color=colors["white"], fill_type="solid")
    light_grey_fill = PatternFill(start_color=colors["light-grey"], end_color=colors["light-grey"], fill_type="solid")

    for row in range(2, ws.max_row + 1):
        if row % 2 == 0:  # Even rows
            fill = white_fill
        else:  # Odd rows
            fill = light_grey_fill

        for col in range(1, ws.max_column + 1):
            ws.cell(row=row, column=col).fill = fill

    # 3c - Status color coding
    status_col = headers.index("Status") + 1

    status_colors = {
        "Checked Out": colors["light-green"],
        "Checked In": colors["light-blue"],
        "Confirmed": colors["light-yellow"],
        "Cancelled": colors["light-red"]
    }

    for row in range(2, ws.max_row + 1):
        status_cell = ws.cell(row=row, column=status_col)
        status_value = status_cell.value

        if status_value in status_colors:
            status_fill = PatternFill(start_color=status_colors[status_value], 
                                    end_color=status_colors[status_value], 
                                    fill_type="solid")
            status_cell.fill = status_fill

    # 3d - Column widths
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
        "Notes": 60
    }

    for col_num, header in enumerate(headers, 1):
        if header in column_widths:
            ws.column_dimensions[get_column_letter(col_num)].width = column_widths[header]

    # 3e - Freeze panes and auto-filter
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}1"

    # 3f - Number formats
    currency_columns = ["Rate/Night", "Extras Total", "Taxes", "Total Charged"]
    percent_column = "Discount %"

    for col_num, header in enumerate(headers, 1):
        if header in currency_columns:
            for row in range(2, ws.max_row + 1):
                ws.cell(row=row, column=col_num).number_format = "$#,##0.00"
        elif header == percent_column:
            for row in range(2, ws.max_row + 1):
                ws.cell(row=row, column=col_num).number_format = '0"%"'


if __name__ == "__main__":
    start_time = time.time()
    main()
    end_time = time.time()
    print(f"Execution Time: {end_time-start_time}")
