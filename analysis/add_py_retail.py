#!/usr/bin/env python3
"""Append a 'Previous Year Health-Retail' GDP column (Rs. Crs) to both EOM
workbooks, populated on each insurer's Health row.

Source: 'Segment wise report March 2026.xlsx' -> 'Health Portfolio' sheet. For
each insurer, the 'Previous Year' row directly beneath its current-year row
holds prior-year figures; column B = Health-Retail (already in Crs). Values are
taken as-is from the authoritative sheet (no NL-4 reconciliation exists for the
previous year in the EOM files).
"""
import re
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

import transform_eom as T  # reuse norm_name / core_tokens / match logic

SEG_REPORT = "/projects/sandbox/NL_20/Segment wise report March 2026.xlsx"

# ---- load previous-year Health-Retail per insurer ----
def load_py_retail():
    wb = openpyxl.load_workbook(SEG_REPORT, data_only=True)
    ws = wb["Health Portfolio"]
    skip = ("sub total", "previous", "% ", "industry", "market",
            "general insurers", "stand-alone", "grand total")
    out = {}
    for r in range(5, ws.max_row + 1):
        a = ws.cell(r, 1).value
        if not a:
            continue
        name = str(a).strip()
        low = name.lower()
        if ws.cell(r, 2).value is None or any(k in low for k in skip):
            continue
        nxt = ws.cell(r + 1, 1).value
        if nxt and str(nxt).strip().lower() == "previous year":
            out[T.norm_name(name)] = {"py_retail": ws.cell(r + 1, 2).value, "src_name": name}
    return out

PY = load_py_retail()

def match_py(eom_name):
    key = T.norm_name(eom_name)
    if key in PY:
        return PY[key]
    kt = T.core_tokens(key)
    best = None; best_score = 0.0
    for hk, hv in PY.items():
        score = len(kt & T.core_tokens(hk)) / max(1, len(kt | T.core_tokens(hk)))
        if score > best_score:
            best_score = score; best = hv
    return best if best_score >= 0.5 else None

thin = Side(style="thin", color="BBBBBB")
border = Border(left=thin, right=thin, top=thin, bottom=thin)
hdr_fill = PatternFill("solid", fgColor="1F4E78")
hdr_font = Font(bold=True, color="FFFFFF")
brk_fill = PatternFill("solid", fgColor="FCE4D6")  # light orange to distinguish PY

def add_col(path, label):
    wb = openpyxl.load_workbook(path)
    ws = wb[wb.sheetnames[0]]
    col = ws.max_column + 1
    # guard against re-running: don't add twice
    for c in range(1, ws.max_column + 1):
        if isinstance(ws.cell(1, c).value, str) and "Previous Year Health-Retail" in ws.cell(1, c).value:
            print(f"[{label}] PY Health-Retail column already present (col {get_column_letter(c)}); skipping")
            return
    h = ws.cell(1, col, "Previous Year Health-Retail\nGDP Upto Mar 2025 (Rs. Crs)")
    h.font = hdr_font; h.fill = hdr_fill
    h.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.column_dimensions[get_column_letter(col)].width = 18

    filled = 0; unmatched = []
    last_ins = None
    for r in range(2, ws.max_row + 1):
        a = ws.cell(r, 1).value
        b = ws.cell(r, 2).value
        if a:
            last_ins = a
        if not last_ins or not b or str(b).strip().lower() != "health":
            continue
        cell = ws.cell(r, col)
        cell.border = border; cell.fill = brk_fill
        cell.alignment = Alignment(horizontal="right")
        m = match_py(last_ins)
        if m is None or not isinstance(m["py_retail"], (int, float)):
            cell.value = ""
            unmatched.append(last_ins)
        else:
            cell.value = round(m["py_retail"], 2)
            cell.number_format = "#,##0.00"
            filled += 1
    wb.save(path)
    print(f"[{label}] added PY Health-Retail (col {get_column_letter(col)}); filled {filled} insurers")
    if unmatched:
        print(f"   UNMATCHED: {unmatched}")

if __name__ == "__main__":
    add_col("/projects/sandbox/NL_20/EOM 25-26 (Premium-Exp-Comm) with NL20 Ratios.xlsx", "GI")
    add_col("/projects/sandbox/NL_20/SAHI/EOM 25-26 SAHI with NL20 Ratios.xlsx", "SAHI")
