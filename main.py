from datetime import datetime
import pandas as pd
import json
import parser
from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill


def main():

    with open("dummy-jsons/ecommerce_orders.json", "r") as file:
        data = json.load(file)
        flattened_data = flatten_reservations(data)
        wb = parser.write_to_excel(flattened_data, "Orders")
        wb = parser.generic_style_sheet(wb)
        wb.save("orders.xlsx")

    wb = load_workbook("orders.xlsx")

    wb = custom_style(wb)

    wb.save("orders.xlsx")


def flatten_reservations(data: dict) -> list[dict]:
    try:
        rows = []
        for order in data["payload"]["orders"]:
            row = {
                "Order ID": order["order_id"],
                "Customer Name": order["customer"]["name"],
                "Email": order["customer"]["email"],
                "Country": order["customer"]["country"],
                "Total Order": sum(item["unit_price"] for item in order["items"]),
                "Shipping Method": order["shipping"]["method"].title(),
                "Shipping Cost": order["shipping"]["cost"],
                "ETA(Days)": order["shipping"]["estimated_days"],
                "Payment Method": order["payment"]["method"].replace("_", " ").title(),
                "Payment Status": order["payment"]["status"].title(),
                "Order Status": order["status"].replace("_", " ").title(),
                "Created At": datetime.fromisoformat(order["created_at"]).strftime(
                    "%Y-%m-%d %H:%M:%S"
                ),
            }
            # deduplicationn
            if row not in rows:
                rows.append(row)
        return rows
    except Exception as error:
        print(error)
        return None


def custom_style(wb: Workbook) -> Workbook:
    ws = wb.active
    ws1 = wb.create_sheet("asdasdasd")
    headers = []
    for cell in ws[1]:
        headers.append(cell.value)

    colors = {
        "Paid": parser.LIGHT_GREEN,
        "Pending": parser.LIGHT_YELLOW,
        "Refunded": parser.LIGHT_BLUE,
        "Failed": parser.LIGHT_RED,
    }
    color_column = headers.index("Payment Status") + 1

    for row in range(2, ws.max_row + 1):
        status_cell = ws.cell(row=row, column=color_column)
        status_value = status_cell.value

        # changing the color for each cell in status column
        if status_value in colors:
            status_cell.fill = PatternFill(
                fill_type="solid", fgColor=colors[status_value]
            )
    colors = {
        "Paid": parser.LIGHT_GREEN,
        "Pending": parser.LIGHT_YELLOW,
        "Refunded": parser.LIGHT_BLUE,
        "Failed": parser.LIGHT_RED,
    }
    color_column = headers.index("Payment Status") + 1

    for row in range(2, ws.max_row + 1):
        status_cell = ws.cell(row=row, column=color_column)
        status_value = status_cell.value

        # changing the color for each cell in status column
        if status_value in colors:
            status_cell.fill = PatternFill(
                fill_type="solid", fgColor=colors[status_value]
            )

    # Number formats
    currency_columns = ["Total Order", "Shipping Cost"]
    percent_column = "Discount %"

    # adding the correct format
    for col_num, header in enumerate(headers, 1):
        if header in currency_columns:
            for row in range(2, ws.max_row + 1):
                ws.cell(row=row, column=col_num).number_format = "$#,##0.00"
        elif header == percent_column:
            for row in range(2, ws.max_row + 1):
                ws.cell(row=row, column=col_num).number_format = '0"%"'
    return wb


main()
