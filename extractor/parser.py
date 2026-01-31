#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Robust PDF → Excel extractor for Einhell-style order confirmations/invoices.

Key fixes vs. first cut:
- Header normalization (maps header variants to canonical names)
- AU/Invoice de-duplication
- Safer parsing of Qty/Price (handles blanks, text like "not available")
- Poka-yoke: skip pages without main table, but log reasons
"""

import re
from typing import Dict, List, Tuple

import camelot
import pdfplumber
import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

# ---- Canonical schema (always exported) ----
MIN_COLUMNS = ["Art. No", "Qty", "Price"]
ALWAYS_COLUMNS = ["Page", "AU / Invoice", "Sum"]
CANONICAL = MIN_COLUMNS + ALWAYS_COLUMNS

# Header patterns (lenient, case-insensitive)
HEADER_PATTERNS = {
    "Art. No": re.compile(r"\bart\.?\s*no\.?\b|\bart(?:icle)?\s*no\b|\bartnr\b", re.I),
    "Article": re.compile(r"\barticle\b|\bdescription\b", re.I),
    "EAN":     re.compile(r"\bean\b|\bean[-\s]*code\b", re.I),
    "Qty":     re.compile(r"\bqty\b|\bquantity\b", re.I),
    "Price":   re.compile(r"\bprice\b|\bunit\s*price\b|\bpreis\b", re.I),
    # We don’t export Discount/Amount now, but detector tolerates them.
}

UNAVAILABLE_RX = re.compile(r"not\s+available|sold\s+out", re.I)
NUM_SANITIZE_RX = re.compile(r"[^\d,.\-]")  # keep digits, comma, dot, minus


# ------------- Helpers -------------
def _normalize_number(text: str):
    """Normalize numbers like '1 109,20' → '1109.20' ; returns None if not parseable."""
    if text is None:
        return None
    s = str(text)
    s = s.replace("\u00a0", " ").strip()
    # if the cell contains obvious unavailability text → return None (caller sets 0)
    if UNAVAILABLE_RX.search(s):
        return None
    # strip all but digits, comma, dot, minus
    s = NUM_SANITIZE_RX.sub("", s)
    if not s:
        return None
    # heuristic: if both comma and dot present, assume comma=decimals if comma comes last
    if "," in s and "." in s:
        # remove thousand separators heuristically
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "")
            s = s.replace(",", ".")
        else:
            s = s.replace(",", "")
    else:
        # common EU: spaces already stripped; turn comma into dot
        s = s.replace(",", ".")
    try:
        return float(s)
    except Exception:
        return None


def _dedupe_join(values: List[str]) -> str:
    seen, out = set(), []
    for v in values:
        v = v.strip()
        if not v:
            continue
        if v not in seen:
            seen.add(v)
            out.append(v)
    return ", ".join(out)


def _map_headers(raw_headers: List[str]) -> Dict[int, str]:
    """
    Map raw header cells to our canonical names using regex.
    Returns: {col_index: canonical_name}
    """
    mapped: Dict[int, str] = {}
    for idx, h in enumerate(raw_headers):
        h_str = " ".join(str(h).split())  # collapse whitespace/newlines
        for canon, rx in HEADER_PATTERNS.items():
            if rx.search(h_str):
                # prefer the first match per canonical field
                if canon not in mapped.values():
                    mapped[idx] = canon
                break
    return mapped


def _score_header_map(colmap: Dict[int, str]) -> Tuple[int, int]:
    """Return (min_columns_matched, total_mapped)."""
    mapped = set(colmap.values())
    min_score = sum(1 for need in MIN_COLUMNS if need in mapped)
    return (min_score, len(mapped))


def _find_header_row(df, max_scan: int = 6) -> Tuple[int, Dict[int, str], Tuple[int, int]]:
    """
    Scan the first few rows for a likely header row.
    Returns (row_index, colmap, score_tuple).
    """
    best = (-1, {}, (0, 0))
    scan_rows = min(max_scan, len(df))
    for r in range(scan_rows):
        raw_headers = [(" ".join(str(x).split())).strip() for x in list(df.iloc[r])]
        colmap = _map_headers(raw_headers)
        score = _score_header_map(colmap)
        if score > best[2]:
            best = (r, colmap, score)
    return best


def _formula_sep() -> str:
    """Allow override of Excel formula separator via env, default to ';'."""
    import os
    return os.getenv("EXCEL_FORMULA_SEP", ";")


def _extract_au_invoice(page_text_top: str) -> str:
    au = re.findall(r"\bAU\d{7,8}\b", page_text_top)  # AUxxxxxxx
    inv = re.findall(r"(?:Invoice\s*No[:\s]*)([A-Z0-9\-]+)", page_text_top, flags=re.I)
    return _dedupe_join(au + inv)


# ------------- Main entry -------------
def extract_pdf_to_excel(pdf_path: str, output_path: str) -> str:
    """
    Extracts item rows from all pages and writes a single Excel with:
      Main table + Page summary + AU/Invoice summary
    """
    all_rows: List[Dict[str, object]] = []
    logs: List[str] = []

    with pdfplumber.open(pdf_path) as pdf:
        for i, page in enumerate(pdf.pages, start=1):
            text = page.extract_text() or ""

            # AU/Invoice only from top 25% of page
            h = int(page.height * 0.25)
            crop = page.within_bbox((0, page.height - h, page.width, page.height))
            au_invoice = _extract_au_invoice(crop.extract_text() or "")
            if not au_invoice:
                au_invoice = _extract_au_invoice(text)

            # Camelot robust path: lattice first then stream
            try:
                tables = camelot.read_pdf(pdf_path, pages=str(i), flavor="lattice", strip_text="\n")
                if len(tables) == 0:
                    tables = camelot.read_pdf(pdf_path, pages=str(i), flavor="stream", row_tol=10, edge_tol=200)
            except Exception as e:
                logs.append(f"p{i}: camelot error: {e}")
                tables = []

            if not tables:
                logs.append(f"p{i}: no tables found")
                continue

            # pick the table with the strongest header match
            best = None
            for t in tables:
                df_try = t.df.copy()
                if df_try.empty or len(df_try) < 2:
                    continue
                header_row, colmap, score = _find_header_row(df_try)
                if best is None or score > best["score"] or (score == best["score"] and len(df_try) > best["rows"]):
                    best = {"table": t, "df": df_try, "header_row": header_row, "colmap": colmap, "score": score, "rows": len(df_try)}

            if not best:
                logs.append(f"p{i}: table empty/short")
                continue

            # require at least our minimal set present
            if best["score"][0] < len(MIN_COLUMNS):
                logs.append(f"p{i}: header mapping incomplete → {best['colmap']}")
                continue

            df = best["df"]
            colmap = best["colmap"]
            header_row = best["header_row"]

            # rename columns to canonical names
            df = df.iloc[header_row + 1:].reset_index(drop=True)  # drop header row
            # Build new columns: start with canonical we care about; ignore others
            # build a small accessor map: canon -> series
            acc = {}
            for idx, canon in colmap.items():
                acc[canon] = df.iloc[:, idx].astype(str).apply(lambda s: " ".join(s.split())).reset_index(drop=True)  # collapse spaces

            # iterate rows
            n_rows = len(df)
            added = 0
            for r in range(n_rows):
                art = acc["Art. No"].iloc[r].strip() if "Art. No" in acc else ""
                qty_raw = acc["Qty"].iloc[r].strip() if "Qty" in acc else ""
                price_raw = acc["Price"].iloc[r].strip() if "Price" in acc else ""

                # detect 'not available' anywhere across row text (fallback)
                row_concat = " ".join(col.iloc[r] for col in acc.values()).lower()
                qty_val = 0.0
                price_val = 0.0

                if UNAVAILABLE_RX.search(qty_raw) or UNAVAILABLE_RX.search(row_concat):
                    qty_val = 0.0
                else:
                    nv = _normalize_number(qty_raw)
                    qty_val = nv if nv is not None else 0.0

                pv = _normalize_number(price_raw)
                price_val = pv if pv is not None else 0.0

                # Skip rows that clearly have no Art.No and no numeric signals
                if not art and qty_val == 0 and price_val == 0:
                    continue

                all_rows.append({
                    "Art. No": art,
                    "Qty": int(qty_val) if qty_val.is_integer() else qty_val,
                    "Price": price_val,
                    "Page": i,
                    "AU / Invoice": au_invoice,
                })
                added += 1

            logs.append(f"p{i}: added {added} rows")

    # ---------- Write Excel ----------
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Main"

    # header
    ws.append(CANONICAL)

    # data + live Sum formulas (respecting ; as separator for your Excel)
    formula_sep = _formula_sep()
    for row_idx, r in enumerate(all_rows, start=2):
        ws.cell(row=row_idx, column=1, value=r["Art. No"])
        ws.cell(row=row_idx, column=2, value=r["Qty"])
        ws.cell(row=row_idx, column=3, value=r["Price"])
        ws.cell(row=row_idx, column=4, value=r["Page"])
        ws.cell(row=row_idx, column=5, value=r["AU / Invoice"])
        # Sum formula uses local ; and references current row
        ws.cell(row=row_idx, column=6, value=f"=IFERROR(B{row_idx}*C{row_idx}{formula_sep}0)")

    # style as table
    if len(all_rows) > 0:
        tab = Table(displayName="MainTable", ref=f"A1:{get_column_letter(len(CANONICAL))}{len(all_rows)+1}")
        tab.tableStyleInfo = TableStyleInfo(name="TableStyleMedium9", showRowStripes=True, showColumnStripes=False)
        ws.add_table(tab)

    # two blank cols then Page summary
    offset_col = len(CANONICAL) + 2
    ws.cell(row=1, column=offset_col, value="Page")
    ws.cell(row=1, column=offset_col + 1, value="Sum")

    if len(all_rows) > 0:
        pages_sorted = sorted({r["Page"] for r in all_rows})
        page_col_letter = get_column_letter(CANONICAL.index("Page") + 1)
        sum_col_letter = get_column_letter(CANONICAL.index("Sum") + 1)
        criteria_col_letter = get_column_letter(offset_col)
        for idx, p in enumerate(pages_sorted, start=2):
            ws.cell(row=idx, column=offset_col, value=p)
            ws.cell(row=idx, column=offset_col + 1,
                    value=f"=SUMIF({page_col_letter}:{page_col_letter}{formula_sep}{criteria_col_letter}{idx}{formula_sep}{sum_col_letter}:{sum_col_letter})")

    # two blank cols, then AU/Invoice summary (skip empty AU values)
    offset_col2 = offset_col + 4
    ws.cell(row=1, column=offset_col2, value="AU / Invoice")
    ws.cell(row=1, column=offset_col2 + 1, value="Sum")

    if len(all_rows) > 0:
        au_col_letter = get_column_letter(CANONICAL.index("AU / Invoice") + 1)
        criteria_col_letter = get_column_letter(offset_col2)
        # unique non-empty AU/Invoice values
        au_vals = [r["AU / Invoice"] for r in all_rows if str(r.get("AU / Invoice", "")).strip()]
        uniq_au = sorted(set(au_vals))
        base_row = 2
        for j, au in enumerate(uniq_au, start=base_row):
            ws.cell(row=j, column=offset_col2, value=au)
            ws.cell(row=j, column=offset_col2 + 1,
                    value=f"=SUMIF({au_col_letter}:{au_col_letter}{formula_sep}{criteria_col_letter}{j}{formula_sep}{sum_col_letter}:{sum_col_letter})")

    # save file
    wb.save(output_path)
    return output_path
