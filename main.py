import time, json
from openpyxl import Workbook


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

    wb.save(filename)


if __name__ == "__main__":
    start_time = time.time()
    main()
    end_time = time.time()
    print(f"Execution Time: {end_time-start_time}")
