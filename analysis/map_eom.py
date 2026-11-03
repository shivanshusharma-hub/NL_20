#!/usr/bin/env python3
"""Map NL-20 segment-wise ratios into the EOM workbook as new columns U and V.
U = Net Incurred Claims to Net Earned Premium (%)
V = Combined Ratio (%)
Matched by Insurer (col A) + LOB (col B).
Writes a new file; reports any rows left unmatched and insurers with no NL-20.
"""
import json, re
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

SRC = "/projects/sandbox/NL_20/EOM 25-26 (Premium-Exp-Comm).xlsx"
OUT = "/projects/sandbox/NL_20/EOM 25-26 (Premium-Exp-Comm) with NL20 Ratios.xlsx"
PARSED = json.load(open("nl20_parsed.json"))

# ---- NL-20 segment canonicalisation (same as build_excel) ----
def canon_seg_nl20(name):
    n = name.strip().lower().replace("\u2019", "'")
    mp = {
        "fire": "FIRE", "marine cargo": "Marine Cargo", "marine hull": "Marine Hull",
        "total marine": "Total Marine", "motor od": "Motor OD", "motor tp": "Motor TP",
        "total motor": "Total Motor", "health": "Health", "personal accident": "Personal Accident",
        "travel insurance": "Travel Insurance", "total health": "Total Health",
        "public/ product liability": "Public/ Product Liability", "engineering": "Engineering",
        "aviation": "Aviation", "crop insurance": "Crop Insurance",
        "crop insurance (incl. rural)": "Crop Insurance", "credit insurance": "Credit Insurance",
        "total miscellaneous": "Total Miscellaneous", "total": "TOTAL",
    }
    if n in mp: return mp[n]
    if n.startswith("workmen") or n.startswith("wc"): return "Workmen's Compensation"
    if n.startswith("other"): return "Other"
    return name.strip()

# ---- EOM LOB -> canonical NL-20 segment key ----
EOM_LOB_MAP = {
    "total fire": "FIRE",
    "marine cargo": "Marine Cargo",
    "marine hull": "Marine Hull",
    "total marine": "Total Marine",
    "motor od": "Motor OD",
    "motor tp": "Motor TP",
    "motor total": "Total Motor",
    "health": "Health",
    "personal accident": "Personal Accident",
    "travel": "Travel Insurance",
    "total health": "Total Health",
    "workmen's compensation": "Workmen's Compensation",
    "public/product liability": "Public/ Product Liability",
    "engineering": "Engineering",
    "aviation": "Aviation",
    "crop insurance": "Crop Insurance",
    "credit insurance": "Credit Insurance",
    "other misc segments": "Other",
    "total misc": "Total Miscellaneous",
    "grand total": "TOTAL",
}

def norm_ins(s):
    return re.sub(r"\s+", " ", (s or "").strip().lower())

# build LUT: insurer(normalized) -> { seg_key: {nic, combined} }
LUT = {}
NL20_INSURERS = set()
for d in PARSED:
    ins = d["insurer"]
    NL20_INSURERS.add(ins)
    key = norm_ins(ins)
    LUT.setdefault(key, {})
    for s in d["segments"]:
        LUT[key][canon_seg_nl20(s["segment"])] = s

wb = openpyxl.load_workbook(SRC)   # keep formulas/styles
ws = wb[wb.sheetnames[0]]

# headers for U and V
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

unmatched_rows = []
insurers_no_nl20 = set()
matched = 0
lob_unmatched = set()

last_ins = None
for r in range(2, ws.max_row + 1):
    a = ws.cell(r, 1).value
    b = ws.cell(r, 2).value
    if a:
        last_ins = a
    ins = last_ins
    if not ins or not b:
        continue
    ins_key = norm_ins(ins)
    seg_key = EOM_LOB_MAP.get(str(b).strip().lower())
    uc = ws.cell(r, 21); vc = ws.cell(r, 22)
    uc.border = border; vc.border = border
    uc.alignment = Alignment(horizontal="right"); vc.alignment = Alignment(horizontal="right")
    if seg_key is None:
        lob_unmatched.add(str(b).strip())
        uc.value = ""; vc.value = ""
        continue
    if ins_key not in LUT:
        insurers_no_nl20.add(ins)
        uc.value = ""; vc.value = ""
        uc.fill = na_fill; vc.fill = na_fill
        continue
    seg = LUT[ins_key].get(seg_key)
    if seg is None:
        unmatched_rows.append((r, ins, b, seg_key))
        uc.value = ""; vc.value = ""
        uc.fill = na_fill; vc.fill = na_fill
        continue
    nic = seg.get("nic_nep"); cr = seg.get("combined")
    uc.value = round(nic, 2) if nic is not None else ""
    vc.value = round(cr, 2) if cr is not None else ""
    uc.number_format = "0.00"; vc.number_format = "0.00"
    matched += 1

wb.save(OUT)
print("Wrote", OUT)
print("Matched rows (value written):", matched)
print("\nInsurers in EOM with NO NL-20 uploaded (U/V left blank):")
for i in sorted(insurers_no_nl20):
    print("  -", i)
print("\nEOM LOB labels with no mapping (skipped):", sorted(lob_unmatched) or "none")
print("\nRows where insurer HAS NL-20 but that segment not found:")
for r, ins, b, sk in unmatched_rows:
    print(f"  row {r}: {ins} | {b} -> {sk}")
