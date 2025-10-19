"""
parser.py
----------
Main logic for extracting tables from PDF invoices.
Uses Camelot + pdfplumber for maximum robustness.
Outputs Excel file with live formulas.
"""

import os
import re
import io
import camelot
import pdfplumber
import openpyxl
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

from extractor.utils import normalize_number, safe_float

# ---------- CONFIGURABLE DEFAULTS ----------
MIN_COLUMNS = ["Art. No", "Qty", "Price"]
EXTRA_COLUMNS = ["Page", "AU / Invoice", "Sum"]
HEADER_KEYS = ["Art. No", "Article", "EAN", "Qty", "Price"]


def extract_pdf_to_excel(pdf_path: str, output_path: str):
    """
    Extracts the main product tables from a text-based invoice PDF.
    - Scans every page
    - Finds main table via header detection
    - Normalizes values
    - Builds one Excel file with 3 tables (main, page summary, AU/Invoice summary)
    """
    all_rows = []
    au_data = []
    page_index = 0

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            page_index += 1
            text = page.extract_text() or ""

            # ---- AU/Invoice detection ----
            au_matches = re.findall(r"\b(AU\d{7,8})\b", text)
            inv_matches = re.findall(r"(?:Invoice\s*No[:\s]*)([A-Z0-9\-]+)", text, re.IGNORECASE)
            au_invoice = ", ".join(au_matches + inv_matches)

            # ---- Header presence check ----
            header_present = any(h.lower() in text.lower() for h in ["art. no", "qty", "price"])
            if not header_present:
                continue  # skip non-table pages

            # ---- Try Camelot (lattice first, then stream) ----
            try:
                tables = camelot.read_pdf(pdf_path, pages=str(page_index), flavor="lattice")
                if not tables:
                    tables = camelot.read_pdf(pdf_path, pages=str(page_index), flavor="stream")
            except Exception:
                tables = []

            if not tables:
                continue

            # Pick the biggest table (most rows)
            table = max(tables, key=lambda t: len(t.df))
            df = table.df.rename(columns=lambda x: str(x).strip())
            headers = list(df.iloc[0])
            df = df[1:]
            df.columns = headers

            # Keep only columns that exist in minimal list
            for _, row in df.iterrows():
                try:
                    art = str(row.get("Art. No", "")).strip()
                    qty_raw = str(row.get("Qty", "")).strip()
                    price_raw = str(row.get("Price", "")).strip()

                    # Interpret unavailable items
                    if re.search(r"not available|sold out", qty_raw, re.IGNORECASE):
                        qty = 0
                    else:
                        qty = safe_float(normalize_number(qty_raw)) or 0

                    price = safe_float(normalize_number(price_raw)) or 0
                    all_rows.append({
                        "Art. No": art,
                        "Qty": qty,
                        "Price": price,
                        "Page": page_index,
                        "AU / Invoice": au_invoice,
                    })
                except Exception:
                    continue

    # ----- Write to Excel -----
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Main"

    headers = MIN_COLUMNS + ["Page", "AU / Invoice", "Sum"]
    ws.append(headers)

    for r in all_rows:
        ws.append([
            r["Art. No"],
            r["Qty"],
            r["Price"],
            r["Page"],
            r["AU / Invoice"],
            f"=IFERROR(C{ws.max_row+1}*B{ws.max_row+1};0)"
        ])

    # Format as table
    tab = Table(displayName="MainTable", ref=f"A1:{get_column_letter(len(headers))}{ws.max_row}")
    style = TableStyleInfo(name="TableStyleMedium9", showRowStripes=True)
    tab.tableStyleInfo = style
    ws.add_table(tab)

    # Add spacing columns
    offset = len(headers) + 2

    # Page summary
    ws.cell(row=1, column=offset, value="Page")
    ws.cell(row=1, column=offset+1, value="Sum")
    page_col = get_column_letter(headers.index("Page")+1)
    sum_col = get_column_letter(headers.index("Sum")+1)
    for p in sorted(set(r["Page"] for r in all_rows)):
        r = ws.max_row + 1
        ws.cell(row=r, column=offset, value=p)
        ws.cell(row=r, column=offset+1, value=f"=SUMIF({page_col}:{page_col};A{r};{sum_col}:{sum_col})")

    # AU summary
    offset2 = offset + 4
    ws.cell(row=1, column=offset2, value="AU / Invoice")
    ws.cell(row=1, column=offset2+1, value="Sum")
    au_col = get_column_letter(headers.index("AU / Invoice")+1)
    for au in sorted(set(r["AU / Invoice"] for r in all_rows if r["AU / Invoice"])):
        r = ws.max_row + 1
        ws.cell(row=r, column=offset2, value=au)
        ws.cell(row=r, column=offset2+1, value=f"=SUMIF({au_col}:{au_col};A{r};{sum_col}:{sum_col})")

    wb.save(output_path)
    return output_path
