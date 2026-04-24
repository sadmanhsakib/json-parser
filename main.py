import time, json
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

DARK_BLUE   = "1F3864"
MID_BLUE    = "2E75B6"
LIGHT_BLUE  = "D6E4F0"
WHITE       = "FFFFFF"
LIGHT_GREY  = "F5F5F5"
GREEN       = "E2EFDA"
RED_BG      = "FCE4D6"
YELLOW_BG   = "FFF2CC"


def main():
    # loading the json file
    with open("hotel_bookings.json", "r") as file:
        data = json.load(file)

    rows = flatten_reservations(data) 

    if not rows:
        raise ValueError("No data to write")

    wb = write_to_excel(rows)
    wb = style_sheet(wb, list(rows[0].keys()))
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

    # Header row styling
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill(
        start_color=DARK_BLUE,
        end_color=DARK_BLUE,
        fill_type="solid",
    )
    header_alignment = Alignment(horizontal="center", vertical="center")

    # applying the styling changes for the header row
    for col_num, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_num)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_alignment

    ws.row_dimensions[1].height = 30

    # Alternating row colors
    white_fill = PatternFill(
        start_color=WHITE, end_color=WHITE, fill_type="solid"
    )
    light_grey_fill = PatternFill(
        start_color=LIGHT_GREY,
        end_color=LIGHT_GREY,
        fill_type="solid",
    )

    # Status color coding
    status_col = headers.index("Status") + 1

    status_colors = {
        "Checked Out": GREEN,
        "Checked In": LIGHT_BLUE,
        "Confirmed": YELLOW_BG,
        "Cancelled": RED_BG,
    }

    # applying the row color change and status color coding
    for row in range(2, ws.max_row + 1):
        if row % 2 == 0:  # Even rows
            fill = white_fill
        else:  # Odd rows
            fill = light_grey_fill

        for col in range(1, ws.max_column + 1):
            ws.cell(row=row, column=col).fill = fill

        status_cell = ws.cell(row=row, column=status_col)
        status_value = status_cell.value

        if status_value in status_colors:
            status_fill = PatternFill(
                start_color=status_colors[status_value],
                end_color=status_colors[status_value],
                fill_type="solid",
            )
            status_cell.fill = status_fill

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

    # Freeze panes and auto-filter
    ws.freeze_panes = "A2"
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


if __name__ == "__main__":
    start_time = time.time()
    main()
    end_time = time.time()
    print(f"Execution Time: {end_time-start_time}")
