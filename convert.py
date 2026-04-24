#!/usr/bin/env python3
"""
Audico EC → Peachtree Quote CSV Converter
==========================================

Reads an EC_XX_YYYY.xlsx monthly statement and produces one Peachtree-ready
Sage 50 Sales Journal CSV per yellow-highlighted event, plus summary & audit files.

Usage:
    python convert.py <EC_xlsx> [--config config/hotel_mapping.json] [--out ./out] [--customer-list LISTA.xlsx]

Design:
- Yellow-highlighted CLIENTE rows = items that need a quote.
- Each quote = 1 ITBMS tax line + N product lines + '***' separator + Evento line + Fecha line
- Line amounts are NEGATIVE (Peachtree sales credit convention). AR amount is POSITIVE total.
- Unit price = AUDICO U (column I) which is 50% of customer cost. We invoice the hotel at the 50%.
"""

import argparse
import csv
import json
import re
import sys
from collections import defaultdict
from datetime import date, timedelta
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string

# The 69-column Sage 50 Sales Journal header — this is the exact order Peachtree expects
SALES_HEADER = [
    "Customer ID", "Customer Name", "Invoice/CM #", "Apply to Invoice Number", "Credit Memo",
    "Progress Billing Invoice", "Date", "Ship By", "Quote", "Quote #", "Quote Good Thru Date",
    "Drop Ship", "Ship to Name", "Ship to Address-Line One", "Ship to Address-Line Two",
    "Ship to City", "Ship to State", "Ship to Zipcode", "Ship to Country", "Customer PO",
    "Ship Via", "Ship Date", "Date Due", "Discount Amount", "Discount Date", "Displayed Terms",
    "Sales Representative ID", "Accounts Receivable Account", "Accounts Receivable Amount",
    "Sales Tax ID", "Invoice Note", "Note Prints After Line Items", "Statement Note",
    "Stmt Note Prints Before Ref", "Internal Note", "Beginning Balance Transaction",
    "AR Date Cleared in Bank Rec", "Number of Distributions", "Invoice/CM Distribution",
    "Apply to Invoice Distribution", "Apply To Sales Order", "Apply to Proposal", "Quantity",
    "SO/Proposal Number", "Item ID", "Serial Number", "SO/Proposal Distribution", "Description",
    "G/L Account", "GL Date Cleared in Bank Rec", "Unit Price", "Tax Type", "UPC / SKU", "Weight",
    "Amount", "Inventory Account", "Inv Acnt Date Cleared In Bank Rec", "Cost of Sales Account",
    "COS Acnt Date Cleared In Bank Rec", "Cost of Sales Amount", "Job ID", "Sales Tax Agency ID",
    "Transaction Period", "Transaction Number", "Receipt Number", "Return Authorization",
    "Voided by Transaction", "Recur Number", "Recur Frequency",
]

MONTHS_ES = {
    1: "enero", 2: "febrero", 3: "marzo", 4: "abril", 5: "mayo", 6: "junio",
    7: "julio", 8: "agosto", 9: "septiembre", 10: "octubre", 11: "noviembre", 12: "diciembre",
}
MONTH_NAME_TO_NUM = {
    "enero": 1, "febrero": 2, "marzo": 3, "abril": 4, "mayo": 5, "junio": 6,
    "julio": 7, "agosto": 8, "septiembre": 9, "octubre": 10, "noviembre": 11, "diciembre": 12,
    "jan": 1, "feb": 2, "mar": 3, "apr": 4, "may": 5, "jun": 6, "jul": 7,
    "aug": 8, "sep": 9, "oct": 10, "nov": 11, "dec": 12,
}

YELLOW_RGB = "FFFFFF00"


# ────────────────────────────────────────────────────────────────────────────
# Utilities
# ────────────────────────────────────────────────────────────────────────────

def is_yellow(cell) -> bool:
    """True if the cell has a solid yellow (FFFFFF00) fill."""
    fill = cell.fill
    if fill.patternType != "solid":
        return False
    fg = fill.fgColor
    if fg is None or fg.type != "rgb":
        return False
    return fg.rgb == YELLOW_RGB


def col_letter_to_idx(letter):
    """Accepts 'C' or None and returns column index or None."""
    if letter is None:
        return None
    return column_index_from_string(letter)


def to_float(v, default=0.0):
    if v is None or v == "":
        return default
    if isinstance(v, (int, float)):
        return float(v)
    try:
        s = str(v).replace(",", "").strip()
        return float(s) if s else default
    except (ValueError, TypeError):
        return default


def money(v):
    """Sage 50 money format: always 2 decimals."""
    return f"{v:.2f}"


def extract_header_fecha(sheet):
    """Find the 'Fecha:' in the top-right header of the sheet and return (year, month) if parseable."""
    for row in range(1, 10):
        for col in range(1, sheet.max_column + 1):
            val = sheet.cell(row=row, column=col).value
            if val and isinstance(val, str) and "fecha" in val.lower():
                # e.g., "Fecha: Marzo 15, 2026" or "Fecha: Marzo, 2026" or "Fecha: Marzo 2026"
                m = re.search(
                    r"(enero|febrero|marzo|abril|mayo|junio|julio|agosto|septiembre|octubre|noviembre|diciembre)"
                    r"\D+(\d{4})",
                    val.lower(),
                )
                if m:
                    return int(m.group(2)), MONTH_NAME_TO_NUM[m.group(1)]
    return None


def parse_event_date(raw, fallback_year, fallback_month):
    """
    Parse Audico's freeform date strings into a structured representation.

    Handles:
      '28.2.26'          -> single day Feb 28, 2026
      '2-3.3'            -> range Mar 2–3, fallback year
      '2-3.3.26'         -> range Mar 2–3, 2026
      '02-03.03.26'      -> range Mar 2–3, 2026
      '14.03.26'         -> single day Mar 14, 2026
      '7.03.26'          -> single day Mar 7, 2026
      '2-4.3'            -> range Mar 2–4, fallback year
      date object        -> single day
      None / ''          -> None

    Returns dict with: start_date (date), end_date (date), is_range (bool),
                       original (str), parsed_ok (bool), note_es (str Spanish note)
    """
    result = {
        "start_date": None, "end_date": None, "is_range": False,
        "original": str(raw) if raw is not None else "", "parsed_ok": False, "note_es": "",
    }
    if raw is None or raw == "":
        return result

    # If it's already a datetime/date (Excel might parse some numerics as dates)
    if hasattr(raw, "year") and hasattr(raw, "month") and hasattr(raw, "day"):
        d = date(raw.year, raw.month, raw.day)
        result.update({
            "start_date": d, "end_date": d, "parsed_ok": True,
            "note_es": f"{d.day} de {MONTHS_ES[d.month]} de {d.year}",
        })
        return result

    s = str(raw).strip()
    if not s:
        return result

    # Normalize separators: replace hyphen-minus with - (already), keep dots
    # Pattern 1: "D-D.M.YY" or "DD-DD.MM.YY" — explicit range with year
    m = re.match(r"^\s*(\d{1,2})\s*-\s*(\d{1,2})\s*\.\s*(\d{1,2})\s*\.\s*(\d{2,4})\s*$", s)
    if m:
        d1, d2, mo, yr = int(m.group(1)), int(m.group(2)), int(m.group(3)), int(m.group(4))
        yr = 2000 + yr if yr < 100 else yr
        try:
            sd = date(yr, mo, d1)
            ed = date(yr, mo, d2)
            result.update({
                "start_date": sd, "end_date": ed, "is_range": True, "parsed_ok": True,
                "note_es": f"{d1} al {d2} de {MONTHS_ES[mo]} de {yr}",
            })
            return result
        except ValueError:
            pass

    # Pattern 2: "D-D.M" — range without year (use fallback)
    m = re.match(r"^\s*(\d{1,2})\s*-\s*(\d{1,2})\s*\.\s*(\d{1,2})\s*$", s)
    if m and fallback_year:
        d1, d2, mo = int(m.group(1)), int(m.group(2)), int(m.group(3))
        try:
            sd = date(fallback_year, mo, d1)
            ed = date(fallback_year, mo, d2)
            result.update({
                "start_date": sd, "end_date": ed, "is_range": True, "parsed_ok": True,
                "note_es": f"{d1} al {d2} de {MONTHS_ES[mo]} de {fallback_year}",
            })
            return result
        except ValueError:
            pass

    # Pattern 3: "D.M.YY" — single day with year
    m = re.match(r"^\s*(\d{1,2})\s*\.\s*(\d{1,2})\s*\.\s*(\d{2,4})\s*$", s)
    if m:
        d1, mo, yr = int(m.group(1)), int(m.group(2)), int(m.group(3))
        yr = 2000 + yr if yr < 100 else yr
        try:
            sd = date(yr, mo, d1)
            result.update({
                "start_date": sd, "end_date": sd, "parsed_ok": True,
                "note_es": f"{d1} de {MONTHS_ES[mo]} de {yr}",
            })
            return result
        except ValueError:
            pass

    # Pattern 4: "D.M" — single day without year (use fallback)
    m = re.match(r"^\s*(\d{1,2})\s*\.\s*(\d{1,2})\s*$", s)
    if m and fallback_year:
        d1, mo = int(m.group(1)), int(m.group(2))
        try:
            sd = date(fallback_year, mo, d1)
            result.update({
                "start_date": sd, "end_date": sd, "parsed_ok": True,
                "note_es": f"{d1} de {MONTHS_ES[mo]} de {fallback_year}",
            })
            return result
        except ValueError:
            pass

    # Pattern 5: "D1-D2.M.YY" where D1 > D2 = CROSS-MONTH range, e.g. "31-11.02.26" = Jan 31 → Feb 11, 2026
    # (D1 in prior month, D2 in the specified month)
    m = re.match(r"^\s*(\d{1,2})\s*-\s*(\d{1,2})\s*\.\s*(\d{1,2})\s*\.\s*(\d{2,4})\s*$", s)
    if m:
        d1, d2, mo, yr = int(m.group(1)), int(m.group(2)), int(m.group(3)), int(m.group(4))
        yr = 2000 + yr if yr < 100 else yr
        if d1 > d2:
            prior_mo = mo - 1 if mo > 1 else 12
            prior_yr = yr if mo > 1 else yr - 1
            try:
                sd = date(prior_yr, prior_mo, d1)
                ed = date(yr, mo, d2)
                result.update({
                    "start_date": sd, "end_date": ed, "is_range": True, "parsed_ok": True,
                    "note_es": f"{d1} de {MONTHS_ES[prior_mo]} al {d2} de {MONTHS_ES[mo]} de {yr}",
                })
                return result
            except ValueError:
                pass

    # Pattern 6: "D-M-YY" — dashes instead of dots, e.g. "25-3-26"
    m = re.match(r"^\s*(\d{1,2})\s*-\s*(\d{1,2})\s*-\s*(\d{2,4})\s*$", s)
    if m:
        d1, mo, yr = int(m.group(1)), int(m.group(2)), int(m.group(3))
        yr = 2000 + yr if yr < 100 else yr
        try:
            sd = date(yr, mo, d1)
            result.update({
                "start_date": sd, "end_date": sd, "parsed_ok": True,
                "note_es": f"{d1} de {MONTHS_ES[mo]} de {yr}",
            })
            return result
        except ValueError:
            pass

    # Pattern 7: "D1,D2.M[.YY]" — day list (non-consecutive days), e.g. "10,12.3" = Mar 10 and 12
    m = re.match(r"^\s*(\d{1,2})\s*,\s*(\d{1,2})\s*\.\s*(\d{1,2})(?:\s*\.\s*(\d{2,4}))?\s*$", s)
    if m:
        d1, d2, mo = int(m.group(1)), int(m.group(2)), int(m.group(3))
        yr = int(m.group(4)) if m.group(4) else fallback_year
        if yr and yr < 100:
            yr = 2000 + yr
        if yr:
            try:
                sd = date(yr, mo, d1)
                ed = date(yr, mo, d2)
                result.update({
                    "start_date": sd, "end_date": ed, "is_range": True, "parsed_ok": True,
                    "note_es": f"{d1} y {d2} de {MONTHS_ES[mo]} de {yr}",
                })
                return result
            except ValueError:
                pass

    # Fallback — keep original string as note
    result["note_es"] = s
    return result


# ────────────────────────────────────────────────────────────────────────────
# Core extraction
# ────────────────────────────────────────────────────────────────────────────

def extract_quotes_from_sheet(sheet, sheet_data, tab_config, fallback_year, fallback_month):
    """
    Scan the sheet and group consecutive item rows into quotes, one per yellow CLIENTE.
    `sheet` carries fill/formatting; `sheet_data` carries resolved formula values.
    Returns a list of dicts: each = {
        'client_row': int, 'client_name': str, 'event_date_raw': str,
        'event_date': parse_result, 'items': [{row, description, qty, days, unit_price, total, source_total}],
        'skipped': bool, 'skip_reason': str,
    }
    """
    client_col = col_letter_to_idx(tab_config["client_col"])
    desc_col = col_letter_to_idx(tab_config["desc_col"])
    qty_col = col_letter_to_idx(tab_config["qty_col"])
    days_col = col_letter_to_idx(tab_config["days_col"])
    audico_u_col = col_letter_to_idx(tab_config.get("audico_u_col"))
    total_col = col_letter_to_idx(tab_config["total_col"])
    date_col = 2  # always column B per the spec
    skip_rows = set(tab_config.get("skip_rows", []))

    quotes = []
    current_quote = None

    for row in range(11, sheet.max_row + 1):
        client_cell = sheet.cell(row=row, column=client_col)
        client_val = client_cell.value
        client_val_str = str(client_val).strip() if client_val is not None else ""

        # A new quote starts on a yellow-highlighted CLIENTE row
        if client_val_str and is_yellow(client_cell) and client_val_str.upper() != "CLIENTE":
            # Close any previous quote
            if current_quote is not None:
                quotes.append(current_quote)
            date_raw = sheet_data.cell(row=row, column=date_col).value
            parsed = parse_event_date(date_raw, fallback_year, fallback_month)
            current_quote = {
                "client_row": row,
                "client_name": client_val_str,
                "event_date_raw": date_raw,
                "event_date": parsed,
                "items": [],
                "skipped": row in skip_rows,
                "skip_reason": "Listed in config 'skip_rows'" if row in skip_rows else "",
            }
            # Add this row's item (first item of the quote lives on the same row as CLIENTE)
            _add_item_if_present(sheet_data, row, desc_col, qty_col, days_col, audico_u_col, total_col, current_quote)
            continue

        # A new NON-yellow client stops the current quote (different event, non-billed)
        if client_val_str and client_val_str.upper() != "CLIENTE" and not is_yellow(client_cell):
            if current_quote is not None:
                quotes.append(current_quote)
                current_quote = None
            continue

        # Otherwise, if we're inside a quote and the row has a description, it's a continuation item
        if current_quote is not None:
            desc_val = sheet_data.cell(row=row, column=desc_col).value
            if desc_val is not None and str(desc_val).strip():
                _add_item_if_present(sheet_data, row, desc_col, qty_col, days_col, audico_u_col, total_col, current_quote)

    # Close final quote
    if current_quote is not None:
        quotes.append(current_quote)

    return quotes


def _add_item_if_present(sheet, row, desc_col, qty_col, days_col, audico_u_col, total_col, quote):
    """Read one item row and append to the quote, if it looks like a real item."""
    desc = sheet.cell(row=row, column=desc_col).value
    if desc is None or not str(desc).strip():
        return
    desc_str = str(desc).strip()

    qty = to_float(sheet.cell(row=row, column=qty_col).value, 1.0) or 1.0
    days = to_float(sheet.cell(row=row, column=days_col).value, 1.0) or 1.0
    source_total = to_float(sheet.cell(row=row, column=total_col).value, 0.0)

    if audico_u_col is not None:
        unit_price = to_float(sheet.cell(row=row, column=audico_u_col).value, 0.0)
    else:
        # panamazing case: no AUDICO U column; derive from total / (qty * days)
        denom = qty * days
        unit_price = source_total / denom if denom else 0.0

    quote["items"].append({
        "row": row,
        "description": desc_str,
        "qty": qty,
        "days": days,
        "unit_price": unit_price,
        "computed_total": unit_price * qty * days,
        "source_total": source_total,
    })


# ────────────────────────────────────────────────────────────────────────────
# CSV row builders
# ────────────────────────────────────────────────────────────────────────────

def build_quote_rows(quote, tab_config, defaults, quote_date, good_thru_date):
    """
    Build the list of CSV rows (each a dict of header→value) that represent this quote.
    Follows the exact Peachtree pattern:
      Dist 0: ITBMS tax line (Amount = -tax, G/L 219)
      Dist 1..N: Line items (Amount = -(unit*qty*days), G/L 403, Tax 1)
      Dist N+1: '***' separator
      Dist N+2: 'Evento: <client name>'
      Dist N+3: 'Fecha: <event date range>'
    """
    items = quote["items"]
    n_items = len(items)
    total_dists = n_items + 4  # tax + N items + *** + Evento + Fecha

    subtotal = sum(it["computed_total"] for it in items)
    tax_amount = round(subtotal * defaults["sales_tax_rate"], 2)
    grand_total = round(subtotal + tax_amount, 2)

    def _peachtree_date(d: date) -> str:
        """Peachtree format: M/D/YY with no leading zeros, e.g. 3/12/26"""
        return f"{d.month}/{d.day}/{d.year % 100:02d}"

    date_str = _peachtree_date(quote_date)
    good_thru_str = _peachtree_date(good_thru_date)

    # Common header fields repeated on every row
    common = {
        "Customer ID": tab_config["customer_id"],
        "Customer Name": tab_config["customer_name"],
        "Invoice/CM #": "",  # left blank so Peachtree auto-assigns on conversion
        "Apply to Invoice Number": "",
        "Credit Memo": "FALSE",
        "Progress Billing Invoice": "FALSE",
        "Date": date_str,
        "Ship By": "",
        "Quote": "TRUE",
        "Quote #": "",  # auto-assigned by Peachtree
        "Quote Good Thru Date": good_thru_str,
        "Drop Ship": "FALSE",
        "Ship to Name": tab_config["ship_to_name"],
        "Ship to Address-Line One": tab_config["ship_to_address"],
        "Ship to Address-Line Two": "",
        "Ship to City": tab_config["ship_to_city"],
        "Ship to State": "",
        "Ship to Zipcode": "",
        "Ship to Country": "",
        "Customer PO": "",
        "Ship Via": defaults["ship_via"],
        "Ship Date": "",
        "Date Due": good_thru_str,
        "Discount Amount": "0.00",
        "Discount Date": date_str,
        "Displayed Terms": defaults["displayed_terms"],
        "Sales Representative ID": "",
        "Accounts Receivable Account": defaults["accounts_receivable_account"],
        "Accounts Receivable Amount": money(grand_total),
        "Sales Tax ID": defaults["sales_tax_id"],
        "Invoice Note": "",
        "Note Prints After Line Items": "FALSE",
        "Statement Note": "",
        "Stmt Note Prints Before Ref": "FALSE",
        "Internal Note": "",
        "Beginning Balance Transaction": "FALSE",
        "AR Date Cleared in Bank Rec": "",
        "Number of Distributions": str(total_dists),
        "Apply to Invoice Distribution": "0",
        "Apply To Sales Order": "FALSE",
        "Apply to Proposal": "FALSE",
        "SO/Proposal Number": "",
        "Item ID": "",
        "Serial Number": "",
        "SO/Proposal Distribution": "0",
        "GL Date Cleared in Bank Rec": "",
        "UPC / SKU": "",
        "Weight": "0.00",
        "Inventory Account": "",
        "Inv Acnt Date Cleared In Bank Rec": "",
        "Cost of Sales Account": "",
        "COS Acnt Date Cleared In Bank Rec": "",
        "Cost of Sales Amount": "0.00",
        "Job ID": "",
        "Transaction Period": "",
        "Transaction Number": "",
        "Receipt Number": "",
        "Return Authorization": "",
        "Voided by Transaction": "",
        "Recur Number": "0",
        "Recur Frequency": "0",
    }

    rows_out = []

    # Dist 0 — ITBMS tax line
    r0 = dict(common)
    r0.update({
        "Invoice/CM Distribution": "0",
        "Quantity": "0.00",
        "Description": "ITBMS",
        "G/L Account": defaults["gl_tax_line"],
        "Unit Price": "0.00",
        "Tax Type": "0",
        "Amount": money(-tax_amount),
        "Sales Tax Agency ID": "ITBMS",
    })
    rows_out.append(r0)

    # Dist 1..N — product lines
    for idx, item in enumerate(items, start=1):
        r = dict(common)
        r.update({
            "Invoice/CM Distribution": str(idx),
            "Quantity": money(item["qty"]),
            "Description": item["description"],
            "G/L Account": defaults["gl_product_line"],
            "Unit Price": money(item["unit_price"] * item["days"]),  # Peachtree's "Unit Price" = per-line unit × days basis
            "Tax Type": "1",
            "Amount": money(-item["computed_total"]),
            "Sales Tax Agency ID": "",
        })
        rows_out.append(r)

    # Dist N+1 — '***' separator
    r_sep = dict(common)
    r_sep.update({
        "Invoice/CM Distribution": str(n_items + 1),
        "Quantity": "0.00",
        "Description": "***",
        "G/L Account": defaults["gl_product_line"],
        "Unit Price": "0.00",
        "Tax Type": "1",
        "Amount": "0.00",
        "Sales Tax Agency ID": "",
    })
    rows_out.append(r_sep)

    # Dist N+2 — Evento: <client name>
    r_evento = dict(common)
    r_evento.update({
        "Invoice/CM Distribution": str(n_items + 2),
        "Quantity": "0.00",
        "Description": f"Evento: {quote['client_name']}",
        "G/L Account": defaults["gl_product_line"],
        "Unit Price": "0.00",
        "Tax Type": "1",
        "Amount": "0.00",
        "Sales Tax Agency ID": "",
    })
    rows_out.append(r_evento)

    # Dist N+3 — Fecha: <event dates>
    fecha_note = quote["event_date"]["note_es"] or "(sin fecha)"
    r_fecha = dict(common)
    r_fecha.update({
        "Invoice/CM Distribution": str(n_items + 3),
        "Quantity": "0.00",
        "Description": f"Fecha: {fecha_note}",
        "G/L Account": defaults["gl_product_line"],
        "Unit Price": "0.00",
        "Tax Type": "1",
        "Amount": "0.00",
        "Sales Tax Agency ID": "",
    })
    rows_out.append(r_fecha)

    return rows_out, subtotal, tax_amount, grand_total


def slugify(name, maxlen=40):
    """Filename-safe slug from a client/event name."""
    s = re.sub(r"[^A-Za-z0-9\-_ ]+", "", name).strip()
    s = re.sub(r"\s+", "_", s)
    return s[:maxlen] or "UNKNOWN"


def write_quote_csv(rows, path):
    """Write a quote to a CSV file with the 69-column header."""
    with open(path, "w", encoding="latin-1", newline="") as f:
        w = csv.DictWriter(f, fieldnames=SALES_HEADER, extrasaction="ignore", quoting=csv.QUOTE_MINIMAL)
        w.writeheader()
        for r in rows:
            w.writerow(r)


# ────────────────────────────────────────────────────────────────────────────
# Main
# ────────────────────────────────────────────────────────────────────────────

def run_conversion(xlsx_path, config_path="config/hotel_mapping.json", out_dir="./out",
                    customer_list_path=None, run_date=None):
    """
    Core conversion logic as a library function.
    Returns dict with: quotes_written, not_processed (list), warnings (list),
                       summary_rows (list), out_dir (Path).
    """
    with open(config_path) as f:
        config = json.load(f)

    tabs_cfg = config["tabs"]
    ignore_tabs = set(config.get("ignore_tabs", []))
    defaults = config["defaults"]

    # Optional: load customer list for verification
    known_customer_ids = None
    if customer_list_path:
        wb_cust = load_workbook(customer_list_path, data_only=True)
        sheet_cust = wb_cust.active
        known_customer_ids = set()
        for row in range(2, sheet_cust.max_row + 1):
            cid = sheet_cust.cell(row=row, column=1).value
            if cid:
                known_customer_ids.add(str(cid).strip())

    # Run date
    if run_date is None:
        run_date = date.today()
    elif isinstance(run_date, str):
        run_date = date.fromisoformat(run_date)
    good_thru_date = run_date + timedelta(days=defaults["quote_good_thru_days"])

    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    # Load xlsx TWICE: once with formulas (for yellow fill detection) and once data_only (for computed formula values)
    wb = load_workbook(xlsx_path)
    wb_data = load_workbook(xlsx_path, data_only=True)

    # Collectors
    summary_rows = []
    not_processed_rows = []
    warnings = []

    total_quotes_written = 0

    for sheet_name in wb.sheetnames:
        norm = sheet_name.lower().strip()
        if norm in ignore_tabs:
            continue
        if norm not in tabs_cfg:
            warnings.append(f"Tab '{sheet_name}' is not in hotel_mapping.json — skipped.")
            continue

        tab_cfg = tabs_cfg[norm]

        # Customer verification
        if known_customer_ids is not None and tab_cfg["customer_id"] not in known_customer_ids:
            warnings.append(
                f"Tab '{sheet_name}': Customer ID {tab_cfg['customer_id']} NOT found in current "
                f"customer list. Import may fail until this customer is created in Peachtree."
            )

        sheet = wb[sheet_name]          # has fill info
        sheet_data = wb_data[sheet_name]  # has computed formula values
        header_date = extract_header_fecha(sheet)
        fallback_year = header_date[0] if header_date else run_date.year
        fallback_month = header_date[1] if header_date else run_date.month

        quotes = extract_quotes_from_sheet(sheet, sheet_data, tab_cfg, fallback_year, fallback_month)

        tab_out_dir = out_dir / slugify(sheet_name, 30)
        tab_out_dir.mkdir(exist_ok=True)

        for q in quotes:
            # Skip if configured
            if q["skipped"]:
                not_processed_rows.append({
                    "tab": sheet_name, "row": q["client_row"], "client": q["client_name"],
                    "event_date_raw": str(q["event_date_raw"] or ""),
                    "items": len(q["items"]),
                    "reason": q["skip_reason"],
                })
                continue

            # Skip empty quotes (yellow row but no items found)
            if not q["items"]:
                not_processed_rows.append({
                    "tab": sheet_name, "row": q["client_row"], "client": q["client_name"],
                    "event_date_raw": str(q["event_date_raw"] or ""),
                    "items": 0,
                    "reason": "Yellow CLIENTE row has no line items below it",
                })
                continue

            # Safety check: recomputed vs source totals
            for it in q["items"]:
                if it["source_total"] and abs(it["computed_total"] - it["source_total"]) > 0.01:
                    warnings.append(
                        f"{sheet_name} R{it['row']} '{q['client_name']}' — item '{it['description']}': "
                        f"computed (unit×qty×days = {it['computed_total']:.2f}) "
                        f"disagrees with source J column ({it['source_total']:.2f})."
                    )

            # Warn on unparseable dates
            if not q["event_date"]["parsed_ok"] and q["event_date_raw"]:
                warnings.append(
                    f"{sheet_name} R{q['client_row']} '{q['client_name']}' — could not parse date "
                    f"'{q['event_date_raw']}'; passing through as-is into note."
                )

            rows, subtotal, tax, total = build_quote_rows(q, tab_cfg, defaults, run_date, good_thru_date)

            # Filename
            event_date_part = (
                q["event_date"]["start_date"].strftime("%Y-%m-%d")
                if q["event_date"]["start_date"] else "nodate"
            )
            fname = f"{slugify(sheet_name, 20)}_R{q['client_row']}_{slugify(q['client_name'], 30)}_{event_date_part}.csv"
            fpath = tab_out_dir / fname
            write_quote_csv(rows, fpath)

            summary_rows.append({
                "tab": sheet_name,
                "row": q["client_row"],
                "client": q["client_name"],
                "event_date": q["event_date"]["note_es"],
                "n_items": len(q["items"]),
                "subtotal": money(subtotal),
                "itbms": money(tax),
                "total": money(total),
                "file": str(fpath.relative_to(out_dir)),
            })
            total_quotes_written += 1

    # Write summary
    with open(out_dir / "_summary.csv", "w", encoding="utf-8", newline="") as f:
        if summary_rows:
            w = csv.DictWriter(f, fieldnames=list(summary_rows[0].keys()))
            w.writeheader()
            for r in summary_rows:
                w.writerow(r)

    # Write not-processed
    with open(out_dir / "_not_processed.csv", "w", encoding="utf-8", newline="") as f:
        if not_processed_rows:
            w = csv.DictWriter(f, fieldnames=list(not_processed_rows[0].keys()))
            w.writeheader()
            for r in not_processed_rows:
                w.writerow(r)
        else:
            f.write("tab,row,client,event_date_raw,items,reason\n")

    # Write warnings
    with open(out_dir / "_warnings.txt", "w", encoding="utf-8") as f:
        if warnings:
            for w in warnings:
                f.write(w + "\n")
        else:
            f.write("No warnings.\n")

    # Return structured result instead of printing
    return {
        "quotes_written": total_quotes_written,
        "summary_rows": summary_rows,
        "not_processed": not_processed_rows,
        "warnings": warnings,
        "out_dir": out_dir,
    }


def main():
    """CLI entry point."""
    ap = argparse.ArgumentParser(description="Convert Audico EC xlsx to Peachtree quote CSVs.")
    ap.add_argument("xlsx", help="Path to EC_XX_YYYY.xlsx")
    ap.add_argument("--config", default="config/hotel_mapping.json", help="Hotel mapping JSON")
    ap.add_argument("--out", default="./out", help="Output directory")
    ap.add_argument("--customer-list", default=None,
                    help="Optional LISTA_DE_CLIENTES_POR_ID.xlsx for customer ID verification")
    ap.add_argument("--run-date", default=None, help="Override run date as YYYY-MM-DD (default: today)")
    args = ap.parse_args()

    result = run_conversion(
        xlsx_path=args.xlsx,
        config_path=args.config,
        out_dir=args.out,
        customer_list_path=args.customer_list,
        run_date=args.run_date,
    )

    print(f"Wrote {result['quotes_written']} quote CSVs to {result['out_dir']}")
    print(f"Summary: {result['out_dir'] / '_summary.csv'}")
    print(f"Not processed: {result['out_dir'] / '_not_processed.csv'} ({len(result['not_processed'])} entries)")
    print(f"Warnings: {result['out_dir'] / '_warnings.txt'} ({len(result['warnings'])} messages)")


if __name__ == "__main__":
    main()
