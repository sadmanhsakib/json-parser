# JSON → Excel Converter

A Python toolkit for converting structured JSON (or XML) data into polished, styled Excel workbooks — complete with alternating row colours, frozen headers, auto-filters, colour-coded status cells, and auto-sized columns.

The project is split into two layers:

| File | Role |
|------|------|
| `parser.py` | **Generic engine** — format-agnostic parsing, flattening helpers, and Excel styling utilities that work for *any* JSON structure |
| `main.py` | **Domain-specific script** — wires the engine to a concrete JSON schema (e-commerce orders), adds custom column formatting, and generates a Summary sheet |

---

## Table of Contents

- [Features](#features)
- [Project Structure](#project-structure)
- [Requirements](#requirements)
- [How It Works](#how-it-works)
  - [parser.py — The Generic Engine](#parserpy--the-generic-engine)
  - [main.py — The Domain Script](#mainpy--the-domain-script)
- [Output](#output)
- [Extending to a New JSON Schema](#extending-to-a-new-json-schema)

---

## Features

- ✅ Parse nested / deeply-nested JSON into flat tabular rows  
- ✅ XML → JSON conversion via `xmltodict` before processing  
- ✅ Duplicate-row deduplication built into the flattening step  
- ✅ Generic Excel styling (dark-blue headers, alternating row fills, frozen header row, auto-filter, auto-column width)  
- ✅ Domain-specific colour-coding of status columns (Paid / Pending / Refunded / Failed)  
- ✅ Currency and percentage number formats applied per column  
- ✅ Auto-generated **Summary sheet** with KPI cards and ranked breakdown tables  
- ✅ Clean separation of concerns — swap `main.py` for a new schema without touching `parser.py`

---

## Project Structure

```
convert-json-to-excel/
│
├── parser.py              # Generic parsing & styling engine (reusable)
├── main.py                # E-commerce orders domain script (entry point)
│
├── dummy-jsons/
│   ├── ecommerce_orders.json   # Sample e-commerce payload used by main.py
│   └── hotel_bookings.json     # Alternative sample (hotel reservations schema)
│
├── requirements.txt       # Python dependencies
└── README.md
```

---

## Requirements

- Python **3.10+**
- [openpyxl](https://openpyxl.readthedocs.io/) — Excel workbook creation and styling  
- [pandas](https://pandas.pydata.org/) — Data inspection helpers  
- [xmltodict](https://github.com/martinblech/xmltodict) — XML → dict conversion  

---

## How It Works

### `parser.py` — The Generic Engine

`parser.py` provides **format-agnostic** utilities. None of its functions know (or care) which JSON schema is being processed; they operate purely on the data structures passed to them.

#### `convert_xml_to_json(file_name)`
Reads an XML file, parses it into a Python dict using `xmltodict`, then writes the result as a sibling `.json` file. Expects the XML to have a `<root>` element at the top level.

#### `inspect(data)`
A quick debugging helper. Prints the type, length, inferred column names, dtypes, and the first few rows of any JSON-serialisable object via `pandas.json_normalize`. Useful when exploring an unfamiliar JSON schema.

#### `flatten_reservations(data)` *(generic reference implementation)*
Demonstrates the recommended flattening pattern for a hotel-bookings schema. It iterates over the records list, maps nested fields to human-readable column names, applies light string cleaning (`.replace("_", " ").title()`), computes derived fields (e.g. `Extras Total`), and deduplicates rows before returning a flat `list[dict]`.

> **Note:** This function in `parser.py` targets the *hotel_bookings* schema. `main.py` re-implements the same pattern for the *ecommerce_orders* schema. When adapting the toolkit to a new schema, you provide your own flattening function in your domain script.

#### `write_to_excel(rows, ws_title)` → `Workbook`
Accepts the flat `list[dict]` produced by a flattening function and writes it into a new `openpyxl` workbook with the headers derived from the dict keys.

#### `generic_style_sheet(wb)` → `Workbook`
Applies the standard visual treatment to the active sheet:

| Feature | Detail |
|---------|--------|
| Header row | Dark-blue background (`#1F3864`), white bold "Noto Sans Lisu" text, centred, 30 pt row height |
| Data rows | Alternating white / light-grey fills |
| Navigation | Header row frozen at `A2` so it stays visible while scrolling |
| Filtering | Auto-filter drop-downs applied across the entire header row |
| Column widths | Auto-sized to the longest cell value + 4 character padding |

#### `make_header_cell(cell, text)`
Low-level helper that styles a single cell as a header. Called internally by `generic_style_sheet` and reused by `main.py` when building the Summary sheet's sub-tables.

#### Colour constants
Six named hex constants are exported so domain scripts can reference them without hard-coding colour codes:

```python
DARK_BLUE   = "1F3864"
WHITE       = "FFFFFF"
LIGHT_BLUE  = "D6E4F0"
LIGHT_GREY  = "F5F5F5"
LIGHT_GREEN = "E2EFDA"
LIGHT_RED   = "FCE4D6"
LIGHT_YELLOW = "FFF2CC"
```

---

### `main.py` — The Domain Script

`main.py` is the executable entry point for the **e-commerce orders** use-case. It calls into `parser.py` for all generic work and adds its own domain-specific logic on top.

#### `main()`
Orchestrates the full pipeline:

```
Load JSON → flatten rows → write workbook → generic style
         → custom style → summary sheet → save file
```

#### `flatten_reservations(data)` *(e-commerce schema)*
Maps the `ecommerce_orders.json` payload to flat rows. Handles:

- Joining multiple item names into a single comma-separated string  
- Computing `Subtotal` and `Order Total` from the nested `items` array  
- Reformatting ISO-8601 timestamps to `YYYY-MM-DD HH:MM:SS`  
- Deduplication of identical rows  

#### `custom_style(wb)` → `Workbook`
Post-processes the workbook produced by `parser.generic_style_sheet`:

- **Status colour-coding** — paints the `Payment Status` column cell-by-cell using the colour constants from `parser.py`:

  | Status | Fill colour |
  |--------|-------------|
  | Paid | Light green |
  | Pending | Light yellow |
  | Refunded | Light blue |
  | Failed | Light red |

- **Number formats** — applies `$#,##0.00` to currency columns.

#### `write_summary(wb, raw_json, rows)` → `Workbook`
Creates a separate **Summary** sheet (grid lines hidden) containing:

- **KPI cards** (rows 1–5): Total Orders, Total Revenue, Items Sold, Average Order Value — each with a dark-blue label cell and a large-font value cell.  
- **Order Status breakdown** (starting row 9, columns B–C): ranked by order count descending, with styled headers and alternating row colours.  
- **Items Sold breakdown** (starting row 9, columns E–F): ranked by quantity sold descending, same visual treatment.  
- Auto-sized columns applied to the Summary sheet as well.

---

## Output

Running `python main.py` produces **`orders_report.xlsx`** with two sheets:

### Sheet 1 — Orders
A fully formatted table of all orders:

- Styled header row with auto-filter  
- Alternating row fills  
- Colour-coded `Payment Status` column  
- Currency formatting on monetary columns  
- Frozen header row for easy scrolling  

### Sheet 2 — Summary
A KPI dashboard with:

- Four headline metrics (Total Orders, Total Revenue, Items Sold, Avg Order Value)  
- Order status breakdown table  
- Top items by quantity sold table  

---

## Extending to a New JSON Schema

1. **Create a new domain script** (e.g. `invoices.py`).  
2. **Import `parser`** and use its utilities directly.  
3. **Write your own flattening function** that maps your JSON fields to the desired column headers.  
4. Call the standard pipeline:

```python
import parser

def main():
    with open("invoices.json") as f:
        data = json.load(f)

    rows = flatten_invoices(data)           # your custom flattener
    wb = parser.write_to_excel(rows, "Invoices")
    wb = parser.generic_style_sheet(wb)     # free generic styling
    # add your own custom_style() here if needed
    wb.save("invoices_report.xlsx")

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
```

`parser.py` never needs to be modified — all schema-specific logic lives in the domain script.
