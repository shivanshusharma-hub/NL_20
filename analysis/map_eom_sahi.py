#!/usr/bin/env python3
"""Map NL-20 segment-wise ratios into the SAHI EOM workbook as columns U & V.
U = Net Incurred Claims to Net Earned Premium (%)
V = Combined Ratio (%)
Source: NL-20 ANALYTICAL RATIOS (segmental) inside each insurer's NL-7 ALLFORMS PDF.
Matched by Insurer (col A) + LOB (col B). Values stored as fraction + percent format.
"""
import re
import pdfplumber
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

import parse  # reuse hardened helpers: locate_segment_block, find_col, to_float, is_current_row, norm

SAHI = "/projects/sandbox/NL_20/SAHI"
SRC = f"{SAHI}/EOM 25-26 SAHI.xlsx"
OUT = f"{SAHI}/EOM 25-26 SAHI with NL20 Ratios.xlsx"

# insurer (as it appears in EOM col A) -> (pdf filename, decimal_scale)
# decimal_scale = 100 if the NL-20 reports ratios as fractions (Care), else 1.
INSURERS = [
    ("Aditya Birla Health Insurance Co. Limited",
     "Aditya_Birla_Health_Insurance_Co_Ltd_NL-7_ALLFORMS_FY2025-26_Q4.pdf", 1),
    ("ManipalCigna Health Insurance Company Limited",
     "ManipalCigna_Health_Insurance_Co_Ltd_NL-7_ALLFORMS_FY2025-26_Q4.pdf", 1),
    ("Niva Bupa Health Insurance Company Limited",
     "Niva_Bupa_Health_Insurance_Company_Ltd_NL-7_FY2025-26_Q4.pdf", 1),
    ("Care Health Insurance Limited",
     "Care_Health_Insurance_Company_Ltd_NL-7_FY2025-26_Q4.pdf.pdf", 100),
    ("Star Health and Allied Insurance Company Limited",
     "Star_Health_Allied_Insurance_Co_Ltd_NL-7_ALLFORMS_FY2025-26_Q4.pdf", 1),
]

# EOM LOB (lowercased) -> NL-20 segment label (lowercased, as emitted by parse.norm)
EOM_LOB_MAP = {
    "health": "health",
    "personal accident": "personal accident",
    "travel": "travel insurance",
    "total": "total health",   # health insurers' overall total == Total Health
}

def extract(pdf_path, scale):
    """Return {seg_label_lower: {'nic':pp, 'cr':pp}} in percentage points."""
    out = {}
    with pdfplumber.open(pdf_path) as pdf:
        loc = parse.locate_segment_block(pdf)
        if loc is None:
            return out
        _, _, t, hi, lc = loc
        hdr = t[hi]
        ci_nic = parse.find_col(hdr, "nic")
        ci_cr = parse.find_col(hdr, "cr")
        cur = None
        for r in t[hi + 1:]:
            lab = parse.norm(r[lc]) if lc < len(r) else ""
            if not lab:
                continue
            if parse.is_segment_name(lab):
                cur = lab
            elif parse.is_current_row(lab) and cur:
                nic = parse.to_float(r[ci_nic]) if ci_nic is not None and ci_nic < len(r) else None
                crv = parse.to_float(r[ci_cr]) if ci_cr is not None and ci_cr < len(r) else None
                out[cur] = {
                    "nic": nic * scale if nic is not None else None,
                    "cr": crv * scale if crv is not None else None,
                }
                cur = None
    return out

# build LUT
LUT = {}
for name, fn, scale in INSURERS:
    LUT[name] = extract(f"{SAHI}/{fn}", scale)
    print(f"{name[:40]:<40} segments: {sorted(LUT[name].keys())}")

# write into workbook
wb = openpyxl.load_workbook(SRC)
ws = wb[wb.sheetnames[0]]
hdr_font = Font(bold=True, color="FFFFFF")
hdr_fill = PatternFill("solid", fgColor="1F4E78")
thin = Side(style="thin", color="BBBBBB")
border = Border(left=thin, right=thin, top=thin, bottom=thin)
na_fill = PatternFill("solid", fgColor="F2F2F2")

ws.cell(1, 21, "Net Incurred Claims to Net Earned Premium % (NL-20, upto Mar 2026)")
ws.cell(1, 22, "Combined Ratio % (NL-20, upto Mar 2026)")
for col in (21, 22):
    c = ws.cell(1, col)
    c.font = hdr_font; c.fill = hdr_fill
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
ws.column_dimensions["U"].width = 22
ws.column_dimensions["V"].width = 18

matched = 0
unmatched = []
last_ins = None
for r in range(2, ws.max_row + 1):
    a = ws.cell(r, 1).value
    b = ws.cell(r, 2).value
    if a:
        last_ins = a
    ins = last_ins
    uc = ws.cell(r, 21); vc = ws.cell(r, 22)
    uc.border = border; vc.border = border
    uc.alignment = Alignment(horizontal="right"); vc.alignment = Alignment(horizontal="right")
    if not ins or not b:
        continue
    seg_key = EOM_LOB_MAP.get(str(b).strip().lower())
    seg = LUT.get(ins, {}).get(seg_key) if seg_key else None
    if seg is None:
        unmatched.append((r, ins, b))
        uc.value = ""; vc.value = ""
        uc.fill = na_fill; vc.fill = na_fill
        continue
    nic = seg["nic"]; cr = seg["cr"]
    uc.value = round(nic / 100.0, 6) if nic is not None else ""
    vc.value = round(cr / 100.0, 6) if cr is not None else ""
    uc.number_format = "0.0%"; vc.number_format = "0.0%"
    matched += 1

wb.save(OUT)
print("\nWrote", OUT)
print("Matched rows:", matched)
if unmatched:
    print("Unmatched rows:")
    for r, ins, b in unmatched:
        print(f"  row {r}: {ins} | {b}")
