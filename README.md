# JSON → Excel Converter

> **🛒 Available as a Fiverr service** — Need a custom JSON or XML dataset turned into a clean, presentation-ready Excel report? [Order the gig here](https://www.fiverr.com/s/38XgWrx8)

A professional Python toolkit that converts structured **JSON** (and **XML**) data into polished, fully-styled **Excel workbooks** — complete with dark-blue branded headers, alternating row colours, frozen header rows, auto-filters, auto-sized columns, image embedding, and an optional **Summary dashboard sheet**.

## Table of Contents

- [What It Does](#what-it-does)
- [Sample Schemas Included](#sample-schemas-included)
- [Features](#features)
- [Project Structure](#project-structure)
- [Requirements](#requirements)
- [Quick Start](#quick-start)
- [How It Works](#how-it-works)
  - [parser.py — The Generic Engine](#parserpy--the-generic-engine)
  - [main.py — The Domain Script](#mainpy--the-domain-script)
- [Output](#output)
- [Adapting to a New JSON Schema](#adapting-to-a-new-json-schema)
- [Fiverr Gig](#fiverr-gig)

## What It Does

Raw API responses, database exports, and data dumps are rarely ready to share. This toolkit bridges that gap: you hand it a JSON or XML file, and it hands you back a professional `.xlsx` report that looks like it came out of a BI tool — without involving Excel at all.

The architecture is intentionally split into two layers:

| File | Role |
|------|------|
| `parser.py` | **Generic engine** — format-agnostic utilities for XML conversion, data inspection, Excel writing, and visual styling that work with *any* JSON structure |
| `main.py` | **Domain script** — wires the engine to a specific JSON schema, adds custom column formatting, image embedding, and summary statistics |

You extend the toolkit simply by writing a new domain script for your schema — `parser.py` never needs to change.


## Sample Schemas Included

Three real-world schemas ship with the repository so you can see the converter in action right away:

| File | Description |
|------|-------------|
| `data/ecommerce_orders.json` | Nested e-commerce order payload with items, pricing, and payment status |
| `data/hotel_bookings.json` | Hotel reservation data with guest details, room types, and extras |
| `data/user_data.json` | Large user dataset from [randomuser.me](https://randomuser.me/) — 1000+ records with profile photos |

Pre-generated output files (`orders_report.xlsx`, `hotel_report.xlsx`, `user_data.xlsx`) are included for instant preview.


## Features

- ✅ **Nested JSON flattening** — deeply nested structures mapped to clean tabular rows
- ✅ **XML → JSON conversion** via `xmltodict` before processing
- ✅ **Image embedding** — downloads image URLs and inserts them directly into cells, sizing rows automatically
- ✅ **Duplicate-row deduplication** built into the flattening step
- ✅ **Generic Excel styling** — dark-blue headers, alternating row fills, frozen header row, auto-filter drop-downs, auto-sized column widths
- ✅ **Domain-specific colour-coding** of status columns (Paid / Pending / Refunded / Failed)
- ✅ **Currency and percentage number formats** applied per column
- ✅ **Auto-generated Summary sheet** with KPI cards and ranked breakdown tables
- ✅ **Clean separation of concerns** — swap in a new domain script without touching the engine


## Project Structure

```
convert-json-to-excel/
│
├── parser.py              # Generic parsing & styling engine (reusable)
├── main.py                # Domain script — entry point (e-commerce + user schema)
├── test.py                # Quick workbook inspection helper
│
├── data/
│   ├── ecommerce_orders.json   # Sample e-commerce payload
│   ├── hotel_bookings.json     # Sample hotel reservations payload
│   ├── user_data.json          # 1 000+ user records with photo URLs
│   ├── orders_report.xlsx      # Pre-generated output (ecommerce)
│   ├── hotel_report.xlsx       # Pre-generated output (hotel)
│   └── user_data.xlsx          # Pre-generated output (users with embedded images)
│
├── requirements.txt
└── README.md
```


## Requirements

- Python **3.10+**
- [openpyxl](https://openpyxl.readthedocs.io/) `3.1.5` — Excel workbook creation and styling
- [pandas](https://pandas.pydata.org/) `3.0.2` — Data inspection helpers
- [xmltodict](https://github.com/martinblech/xmltodict) `1.0.4` — XML → dict conversion
- numpy, python-dateutil, tzdata *(installed as transitive dependencies)*


## Quick Start

```bash
# 1. Clone the repository
git clone https://github.com/sadmanhsakib/json-parser.git
cd json-parser

# 2. Create and activate a virtual environment
python -m venv .venv
# Windows
.venv\Scripts\activate
# macOS / Linux
source .venv/bin/activate

# 3. Install dependencies
pip install -r requirements.txt

# 4. Run the converter against the bundled user dataset
python main.py
# → Produces data/user_data.xlsx
```

Open `data/user_data.xlsx` (or any of the pre-generated files) to see the fully styled output.


## How It Works

### `parser.py` — The Generic Engine

`parser.py` provides **format-agnostic** utilities. None of its functions know which JSON schema is being processed — they operate purely on the data structures passed to them.

#### `convert_xml_to_json(file_name)`
Reads an XML file, parses it into a Python dict using `xmltodict`, then writes the result as a sibling `.json` file. Expects the XML to have a `<root>` top-level element.

#### `inspect(data)`
A debugging helper. Prints the type, length, inferred column names, dtypes, and first few rows of any JSON-serialisable object via `pandas.json_normalize`. Useful when exploring an unfamiliar schema.

#### `write_to_excel(rows, ws_title)` → `Workbook`
Accepts the flat `list[dict]` produced by a flattening function and writes it into a new `openpyxl` workbook, deriving headers automatically from dict keys.

#### `img_parser(wb, column_name)` → `Workbook`
Scans the specified column for HTTP image URLs, downloads each image into memory, embeds it into the corresponding cell, auto-sizes the row height, and clears the raw URL text so only the image is visible.

#### `apply_generic_style(wb)` → `Workbook`
Applies the standard visual treatment to the active sheet:

| Feature | Detail |
|---------|--------|
| Header row | Dark-blue background (`#1F3864`), white bold text, centred, 30 pt row height |
| Data rows | Alternating white / light-grey fills |
| Navigation | Header row frozen at `A2` so it stays visible while scrolling |
| Filtering | Auto-filter drop-downs across the entire header row |
| Column widths | Auto-sized to the longest cell value + 4-character padding |

#### `make_header_cell(cell, text)`
Low-level helper that styles a single cell as a header. Reused internally and exposed for domain scripts that build custom sub-tables (e.g. the Summary sheet).

#### Colour constants

Six named hex constants are exported for use in any domain script:

```python
DARK_BLUE    = "1F3864"
WHITE        = "FFFFFF"
LIGHT_BLUE   = "D6E4F0"
LIGHT_GREY   = "F5F5F5"
LIGHT_GREEN  = "E2EFDA"
LIGHT_RED    = "FCE4D6"
LIGHT_YELLOW = "FFF2CC"
```


### `main.py` — The Domain Script

`main.py` is the executable entry point. It calls into `parser.py` for all generic work and adds domain-specific logic on top.

#### `main()`
Orchestrates the full pipeline:

```
Load JSON → flatten rows → write workbook → embed images
         → apply generic style → save .xlsx
```

#### `flatten_reservations(data)` *(user schema)*
Maps the `user_data.json` payload to flat rows. Handles:

- Composing a full name from `title`, `first`, and `last` name fields
- Extracting nested location data (street, city, state, country, postcode, coordinates, timezone)
- Reformatting ISO-8601 timestamps to `YYYY-MM-DD HH:MM:SS`
- Preserving the thumbnail URL in a `Picture` column (later processed by `img_parser`)
- Deduplication of identical rows before writing

#### `apply_custom_style(wb)` → `Workbook`
*(Available for e-commerce schema)*  
Post-processes the workbook to add:

- **Status colour-coding** on the `Payment Status` column:

  | Status | Fill |
  |--------|------|
  | Paid | Light green |
  | Pending | Light yellow |
  | Refunded | Light blue |
  | Failed | Light red |

- **Number formats** — `$#,##0.00` applied to currency columns.

#### `write_summary(wb, raw_json, rows)` → `Workbook`
*(Available for e-commerce schema)*  
Creates a separate **Summary** sheet (grid lines hidden) containing:

- **KPI cards** (rows 1–5): Total Orders, Total Revenue, Items Sold, Average Order Value — dark-blue label cells, large-font value cells
- **Order Status breakdown** (cols B–C, from row 9): ranked by count descending, alternating row colours
- **Items Sold breakdown** (cols E–F, from row 9): ranked by quantity descending, same visual treatment
- Auto-sized columns applied to the Summary sheet


## Output

Running `python main.py` against the bundled user dataset produces **`data/user_data.xlsx`**:

- Styled header row with auto-filter
- Profile photo thumbnails embedded directly in the **Picture** column
- Alternating row fills, frozen header, auto-sized columns
- All contact, location, login, and date-of-birth fields in clearly labelled columns

For the e-commerce schema (`data/ecommerce_orders.json`), the output includes a second **Summary** sheet with KPI cards and ranked breakdown tables.


## Adapting to a New JSON Schema

1. **Create a new domain script** (e.g. `invoices.py`).
2. **Import `parser`** and use its utilities directly.
3. **Write your own flattening function** that maps your JSON fields to the desired column headers.
4. Call the standard pipeline:

```python
import json, parser

def main():
    with open("data/invoices.json") as f:
        data = json.load(f)

    rows = flatten_invoices(data)                    # your custom flattener
    wb = parser.write_to_excel(rows, "Invoices")
    wb = parser.apply_generic_style(wb)              # free generic styling
    # add your own custom_style() here if needed
    wb.save("data/invoices_report.xlsx")

def flatten_invoices(data):
    rows = []
    for invoice in data["invoices"]:
        row = {
            "Invoice ID": invoice["id"],
            "Client":     invoice["client"]["name"],
            # ... map remaining fields
        }
        if row not in rows:
            rows.append(row)
    return rows

if __name__ == "__main__":
    main()
```

`parser.py` never needs to be modified — all schema-specific logic lives in the domain script.


## Fiverr Gig

This toolkit was built as the backbone of a **Fiverr data-conversion service**. If you have a JSON or XML export that you need turned into a clean, branded Excel report — with custom columns, colour-coded status fields, embedded images, or a summary dashboard — you can order the service directly on Fiverr.

**📌 Gig link:** [https://www.fiverr.com/s/38XgWrx8](https://www.fiverr.com/s/38XgWrx8)

What a typical order looks like:

| You provide | You receive |
|-------------|-------------|
| A JSON / XML file (any structure) | A fully styled `.xlsx` workbook |
| Column mapping preferences | Auto-sized columns, frozen header, auto-filters |
| Branding colours (optional) | Header colour-scheme to match your brand |
| Status field labelling (optional) | Colour-coded status cells |
| Dashboard requirements (optional) | A Summary sheet with KPI cards and breakdown tables |

Turnaround is typically **24–48 hours** depending on order complexity.

---

*Built with Python · openpyxl · pandas · xmltodict*
