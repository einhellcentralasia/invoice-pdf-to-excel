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
from typing import Dict, List, Tuple, Optional

import camelot
import pdfplumber
import openpyxl

# ---- Canonical schema (detected) ----
MIN_COLUMNS = ["Art. No", "Qty", "Price"]
CANONICAL_ORDER = [
    "Pos",
    "Art. No",
    "Article",
    "EAN",
    "Qty",
    "Price",
    "Discount",
    "Cust. ItemNo",
    "Amount",
]
ALWAYS_COLUMNS = ["Page"]

# Header patterns (lenient, case-insensitive)
HEADER_PATTERNS = {
    "Pos": re.compile(r"\bpos\b|\bposition\b", re.I),
    "Art. No": re.compile(r"\bart\.?\s*no\.?\b|\bart(?:icle)?\s*no\b|\bartnr\b", re.I),
    "Article": re.compile(r"\barticle\b|\bdescription\b", re.I),
    "EAN":     re.compile(r"\bean\b|\bean[-\s]*code\b", re.I),
    "Qty":     re.compile(r"\bqty\b|\bquantity\b", re.I),
    "Price":   re.compile(r"\bprice\b|\bunit\s*price\b|\bpreis\b", re.I),
    "Discount": re.compile(r"\bdiscount\b|\brabatt\b", re.I),
    "Amount": re.compile(r"\bamount\b|\btotal\s*amount\b|\btotal\b|\bsum\b|\bgesamt\b", re.I),
    "Cust. ItemNo": re.compile(r"\bcust\.?\s*item\s*no\b|\bcustomer\s*item\s*no\b|\bcust\.?\s*itemno\b", re.I),
    # We don’t export Discount/Amount now, but detector tolerates them.
}

_BLOCKED_HEADER_RX = re.compile(r"\bamount\b|total\s*sum|total\s*amount|\bsum\b", re.I)
_NUM_RX = re.compile(r"[^\d,.\-]")  # keep digits, comma, dot, minus


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


def _format_decimal_comma(text: str) -> str:
    """Normalize numeric text to comma-decimal format, e.g. 1234,56."""
    if text is None:
        return ""
    s = str(text).replace("\u00a0", " ").strip()
    if not s:
        return ""
    s = _NUM_RX.sub("", s)
    if not s:
        return text
    if "," in s and "." in s:
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "")
            s = s.replace(",", ".")
        else:
            s = s.replace(",", "")
    else:
        s = s.replace(",", ".")
    try:
        val = float(s)
    except Exception:
        return text
    return f"{val:.2f}".replace(".", ",")


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


def _find_header_row(df, max_scan: int = 50) -> Tuple[int, Dict[int, str], Tuple[int, int]]:
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


def _extract_au_invoice(page_text_top: str) -> str:
    au = re.findall(r"\bAU\d{7,8}\b", page_text_top)  # AUxxxxxxx
    inv = re.findall(r"(?:Invoice\s*No[:\s]*)([A-Z0-9\-]+)", page_text_top, flags=re.I)
    return _dedupe_join(au + inv)


def is_blocked_header(header: str) -> bool:
    return bool(_BLOCKED_HEADER_RX.search(header or ""))


# ------------- Main entry -------------
def extract_pdf_rows(pdf_path: str) -> Tuple[List[Dict[str, str]], List[str]]:
    """Extract item rows from all pages and return (rows, headers_found)."""
    all_rows: List[Dict[str, object]] = []
    logs: List[str] = []
    headers_found: set = set()

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
            headers_found.update(colmap.values())
            # build a small accessor map: canon -> series
            acc = {}
            for idx, canon in colmap.items():
                acc[canon] = df.iloc[:, idx].astype(str).apply(lambda s: " ".join(s.split())).reset_index(drop=True)  # collapse spaces

            # iterate rows
            n_rows = len(df)
            added = 0
            for r in range(n_rows):
                row_data = {}
                for canon, series in acc.items():
                    value = series.iloc[r].strip()
                    if canon in ("Price", "Discount", "Amount"):
                        value = _format_decimal_comma(value)
                    row_data[canon] = value

                # Skip rows that clearly have no data
                if not any(v for v in row_data.values()):
                    continue

                row_data["Page"] = i
                row_data["AU / Invoice"] = au_invoice
                all_rows.append(row_data)
                added += 1

            logs.append(f"p{i}: added {added} rows")

    headers_ordered = [h for h in CANONICAL_ORDER if h in headers_found]
    if any(r.get("AU / Invoice") for r in all_rows):
        headers_ordered.append("AU / Invoice")
    return all_rows, headers_ordered


def extract_pdf_to_excel(
    pdf_path: str,
    output_path: str,
    selected_columns: Optional[List[str]] = None,
) -> str:
    """Extracts item rows from all pages and writes a single Excel with plain values."""
    all_rows, headers = extract_pdf_rows(pdf_path)

    # Filter selected columns, but always keep Page
    if selected_columns:
        selected_set = [c for c in selected_columns if c in headers]
    else:
        selected_set = headers

    if "Page" in selected_set:
        selected_set = [c for c in selected_set if c != "Page"]

    final_headers = selected_set + ["Page"]

    # ---------- Write Excel ----------
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Main"

    # header
    ws.append(final_headers)

    # data rows (plain text values)
    for row_idx, r in enumerate(all_rows, start=2):
        for col_idx, header in enumerate(final_headers, start=1):
            ws.cell(row=row_idx, column=col_idx, value=r.get(header, ""))

    # save file
    wb.save(output_path)
    return output_path
