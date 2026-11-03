#!/usr/bin/env python3
"""For both EOM workbooks (GI and SAHI):
  1. Convert absolute-value columns from Rs. Lakhs to Rs. Crs (divide by 100)
     and relabel headers Lakhs -> Crs. Ratio/percent & formula columns untouched
     (ratios are scale-invariant; derived formula cols flow through automatically).
  2. Add 3 Health GDP breakdown columns (from 'Segment wise report March 2026.xlsx'
     -> 'Health Portfolio' sheet, values already in Crs), populated only on the
     'Health' LOB row for each insurer:
        Health-Retail, Health-Group, Health-Government schemes
"""
import re
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter, column_index_from_string

SEG_REPORT = "/projects/sandbox/NL_20/Segment wise report March 2026.xlsx"

# Absolute-value columns (Rs. Lakhs) to convert to Crs by /100. Only raw-value
# cells are divided; formula cells are left alone (they reference these and so
# convert automatically). Percentage columns (J,K,L,M,N,S,T,U,V) are untouched.
ABS_COLS = ["C", "D", "E", "F", "G", "H", "I"]

# ---------- load Health Portfolio breakdown (Crs) ----------
NOISE = {"limited", "ltd", "co", "company", "insurance", "general", "the",
         "and", "health", "india", "assurance", "mahindra"}

def norm_name(s):
    s = (s or "").strip().lower()
    s = s.replace("&", " and ")
    s = re.sub(r"[^a-z0-9]+", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s

def core_tokens(norm):
    """Distinctive tokens (drop generic corporate noise words)."""
    toks = [t for t in norm.split() if t not in NOISE]
    return set(toks) if toks else set(norm.split())

def load_health_breakdown():
    wb = openpyxl.load_workbook(SEG_REPORT, data_only=True)
    ws = wb["Health Portfolio"]
    data = {}
    skip = ("sub total", "previous year", "% growth", "industry total",
            "% market share", "general insurers", "stand-alone", "grand total")
    for r in range(5, ws.max_row + 1):
        a = ws.cell(r, 1).value
        if not a:
            continue
        a = str(a).strip()
        low = a.lower()
        if any(k in low for k in skip):
            continue
        retail = ws.cell(r, 2).value
        if retail is None:
            continue  # header/section row
        group = ws.cell(r, 3).value
        govt = ws.cell(r, 4).value
        data[norm_name(a)] = {
            "retail": retail, "group": group, "govt": govt,
            "src_name": a,
        }
    return data

HEALTH = load_health_breakdown()

def match_insurer(eom_name):
    key = norm_name(eom_name)
    if key in HEALTH:
        return HEALTH[key]
    kt = core_tokens(key)
    best = None; best_score = 0.0
    for hk, hv in HEALTH.items():
        ht = core_tokens(hk)
        inter = len(kt & ht)
        # Jaccard on distinctive tokens; require the distinctive words to align
        score = inter / max(1, len(kt | ht))
        if score > best_score:
            best_score = score; best = hv
    return best if best_score >= 0.5 else None

# ---------- transform one workbook ----------
thin = Side(style="thin", color="BBBBBB")
border = Border(left=thin, right=thin, top=thin, bottom=thin)
hdr_fill = PatternFill("solid", fgColor="1F4E78")
hdr_font = Font(bold=True, color="FFFFFF")
brk_fill = PatternFill("solid", fgColor="E2EFDA")

def transform(src, out, label):
    wb = openpyxl.load_workbook(src)  # keep formulas
    ws = wb[wb.sheetnames[0]]
    max_col = ws.max_column

    # 1) relabel headers Lakhs -> Crs
    for c in range(1, max_col + 1):
        hv = ws.cell(1, c).value
        if isinstance(hv, str) and "Lakh" in hv:
            ws.cell(1, c).value = hv.replace("Rs. Lakhs", "Rs. Crs").replace("Rs. Lakh", "Rs. Crs").replace("Lakhs", "Crs")

    # 2) convert absolute columns /100 (raw numeric cells only)
    abs_idx = [column_index_from_string(c) for c in ABS_COLS if column_index_from_string(c) <= max_col]
    converted = 0
    for r in range(2, ws.max_row + 1):
        for ci in abs_idx:
            v = ws.cell(r, ci).value
            if isinstance(v, (int, float)) and not isinstance(v, bool):
                ws.cell(r, ci).value = round(v / 100.0, 4)
                converted += 1

    # 3) add breakdown columns (Crs) on the Health row
    # place them after the current last column
    base = max_col
    cols = {
        "retail": base + 1,
        "group": base + 2,
        "govt": base + 3,
    }
    titles = {
        "retail": "Health-Retail\nGDP Upto Mar 2026 (Rs. Crs)",
        "group": "Health-Group\nGDP Upto Mar 2026 (Rs. Crs)",
        "govt": "Health-Government schemes\nGDP Upto Mar 2026 (Rs. Crs)",
    }
    for key, ci in cols.items():
        c = ws.cell(1, ci, titles[key])
        c.font = hdr_font; c.fill = hdr_fill
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.column_dimensions[get_column_letter(ci)].width = 16

    unmatched = []
    filled = 0
    last_ins = None
    for r in range(2, ws.max_row + 1):
        a = ws.cell(r, 1).value
        b = ws.cell(r, 2).value
        if a:
            last_ins = a
        if not last_ins or not b:
            continue
        if str(b).strip().lower() != "health":
            continue
        hb = match_insurer(last_ins)
        for key, ci in cols.items():
            cell = ws.cell(r, ci)
            cell.border = border; cell.fill = brk_fill
            cell.alignment = Alignment(horizontal="right")
            if hb is None:
                cell.value = ""
            else:
                val = hb[key]
                cell.value = round(val, 2) if isinstance(val, (int, float)) else ""
                cell.number_format = "#,##0.00"
        if hb is None:
            unmatched.append(last_ins)
        else:
            filled += 1

    wb.save(out)
    print(f"[{label}] converted {converted} abs cells to Crs; Health breakdown filled for {filled} insurers")
    if unmatched:
        print(f"   UNMATCHED insurers (no Health Portfolio row): {unmatched}")
    print(f"   wrote {out}")

if __name__ == "__main__":
    transform(
        "/projects/sandbox/NL_20/EOM 25-26 (Premium-Exp-Comm) with NL20 Ratios.xlsx",
        "/projects/sandbox/NL_20/EOM 25-26 (Premium-Exp-Comm) with NL20 Ratios.xlsx",
        "GI",
    )
    transform(
        "/projects/sandbox/NL_20/SAHI/EOM 25-26 SAHI with NL20 Ratios.xlsx",
        "/projects/sandbox/NL_20/SAHI/EOM 25-26 SAHI with NL20 Ratios.xlsx",
        "SAHI",
    )
