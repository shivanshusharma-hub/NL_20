#!/usr/bin/env python3
"""Parse NL-20 segment-wise ratios from the uploaded PDFs.
Targets (current period = 'Up to the period ended 31 Mar 2026', segment-wise):
  - Net Incurred Claims to Net Earned Premium
  - Combined Ratio
Uses pdfplumber extract_tables() and matches columns by header text so
column-order / layout differences across insurers are handled automatically.
"""
import os, re, json, glob
import pdfplumber

SRC_DIR = "/projects/sandbox/NL_20"

FILE_MAP = {
    "BAJAJ":    "Bajaj General Insurance Limited",
    "Chola":    "Cholamandalam MS General Insurance Co Ltd",
    "Digit":    "Go Digit General Insurance Ltd",
    "GC":       "Generali Central Insurance Company Limited",
    "HDFC":     "HDFC Ergo General Insurance Co Ltd",
    "ICICI":    "ICICI Lombard General Insurance Co Ltd",
    "Indus":    "IndusInd General Insurance Company Limited",
    "Kotak":    "Zurich Kotak Mahindra General Insurance Co Ltd",
    "Magma":    "Magma General Insurance Limited",
    "National": "National Insurance Co Ltd",
    "Oriental": "The Oriental Insurance Co Ltd",
    "Royal":    "Royal Sundaram General Insurance Co Ltd",
    "SBI":      "SBI General Insurance Co Ltd",
    "TATA":     "Tata AIG General Insurance Co Ltd",
    "United":   "United India Insurance Co Ltd",
    "IFFCO":    "IFFCO-Tokio General Insurance Co Ltd",
    "LIBERTY":  "Liberty General Insurance Co. Ltd",
    "NIA":      "The New India Assurance Co Ltd",
    "SGI":      "Shriram General Insurance Co Ltd",
    "USGI":     "Universal Sompo General Insurance Co Ltd",
    "ZUNO":     "Zuno General Insurance Co Ltd",
}

# labels that denote the CURRENT (FY2025-26 up-to-Mar-2026) data row within a segment
CURRENT_LABELS = ("current period", "current year", "current")
PREVIOUS_LABELS = ("previous period", "previous year", "previous")

def norm(s):
    return re.sub(r"\s+", " ", (s or "").replace("\n", " ")).strip().lower()

def is_current_row(label):
    l = norm(label)
    if l in CURRENT_LABELS:
        return True
    # IndusInd: "Up to Q4 2025-26"; Zuno: "For FY 2025-26"
    if re.match(r"up to q4 2025", l) or re.match(r"for fy 2025", l):
        return True
    return False

SEGMENT_NAME_HINTS = (
    "fire", "marine cargo", "marine hull", "total marine", "marine",
    "motor od", "motor tp", "total motor", "motor",
    "health", "personal accident", "travel", "total health",
    "workmen", "public/ product", "public/product", "engineering",
    "aviation", "crop", "credit", "other segments", "other misc",
    "other liability", "total miscellaneous", "total misc",
)

def is_segment_name(label):
    l = norm(label)
    return any(l.startswith(h) or l == h for h in SEGMENT_NAME_HINTS)

def is_previous_row(label):
    l = norm(label)
    if l in PREVIOUS_LABELS:
        return True
    if re.match(r"up to q4 2024", l) or re.match(r"for fy 2024", l):
        return True
    return False

def find_col(headers, kind):
    """kind: 'nic' or 'cr'. Tolerant to wrapped/split header text."""
    for i, h in enumerate(headers):
        hn = norm(h)
        if kind == "nic":
            if "incurred claims" in hn and ("earned premium" in hn or hn.strip() == "net incurred claims"):
                return i
            # some filings wrap so the cell only says 'net incurred'
            if hn.strip() in ("net incurred claims", "net incurred"):
                return i
        elif kind == "cr":
            if re.search(r"combine ?d rati", hn):
                return i
    return None

def to_float(v):
    if v is None:
        return None
    s = str(v).strip().replace("%", "").replace(",", "")
    s = s.replace("(", "-").replace(")", "")
    if s in ("", "-", "--", "na", "n.a.", "nil", "*"):
        return None
    try:
        return float(s)
    except ValueError:
        return None

def locate_segment_block(pdf):
    """Find (table, header_row_index, label_col) for the segment-wise block.
    Handles: header not necessarily in row 0; a blank spacer col 0 with labels
    in col 1 (IFFCO, NIA)."""
    best = None  # (has_cr, width, table, hdr_idx, label_col)
    for p in pdf.pages:
        for t in p.extract_tables():
            if not t:
                continue
            for hi, hrow in enumerate(t):
                if find_col(hrow, "nic") is None:
                    continue
                # body after header should contain a segment name
                body = " ".join(norm(c) for row in t[hi + 1:hi + 8] for c in row if c)
                if not any(k in body for k in ("fire", "current period", "current year",
                                               "up to q4", "for fy 2025", "motor")):
                    continue
                # determine label column: col with the most text in body rows
                col0_txt = sum(1 for row in t[hi + 1:] if norm(row[0]) and not any(norm(c) for c in row[1:]))
                col1_txt = sum(1 for row in t[hi + 1:]
                               if len(row) > 1 and norm(row[1]) and not any(norm(c) for c in row[2:] if c is not None))
                label_col = 1 if col1_txt > col0_txt else 0
                has_cr = find_col(hrow, "cr") is not None
                cand = (has_cr, len(hrow), t, hi, label_col)
                if best is None or (cand[0], cand[1]) > (best[0], best[1]):
                    best = cand
    return best

def parse_iffco(path, result):
    """IFFCO-Tokio NL-20: segment detail spans page index 2 (Fire..PA) and
    page index 3 (Health, Travel, Total Health, Crop, Other Misc, Total Misc,
    Grand Total). Label column = 1; header on page2 is row index 1.
    Columns: 1=label 2=GDP 3=NetRet 4=NetComm 5=EoM-GDP 6=EoM-GWP
             7=NetIncClaims 8=ClaimsPaid 9=Combined 10=TechRes 11=UW."""
    CI_NIC, CI_CR, LC = 7, 9, 1
    result["col_nic"], result["col_cr"], result["label_col"] = CI_NIC, CI_CR, LC
    # segment labels we keep (skip the Fire/Marine/Miscellaneous/Total *summary* block
    # that appears before the detailed rows on page2)
    def getnum(row, ci):
        return to_float(row[ci]) if ci is not None and ci < len(row) else None
    with pdfplumber.open(path) as pdf:
        rows = []
        for pi in (2, 3):
            if pi < len(pdf.pages):
                for t in pdf.pages[pi].extract_tables():
                    rows.extend(t)
        current_seg = None
        seen_detail = False
        fire_cap = None
        total_cap = None
        for row in rows:
            label = norm(row[LC]) if len(row) > LC else ""
            if not label:
                continue
            # detail block begins at 'Marine Cargo' (after the summary block)
            if label == "marine cargo":
                seen_detail = True
            if is_current_row(label):
                vals = {"nic_nep": getnum(row, CI_NIC), "combined": getnum(row, CI_CR)}
                if current_seg == "Fire" and fire_cap is None:
                    fire_cap = vals
                elif current_seg == "Total" and total_cap is None:
                    total_cap = vals
                elif current_seg and seen_detail:
                    result["segments"].append({"segment": current_seg, **vals})
                continue
            if is_previous_row(label):
                continue
            # segment name row
            current_seg = str(row[LC]).replace("\n", " ").strip()
        # prepend FIRE and append TOTAL from the summary block
        if fire_cap:
            result["segments"].insert(0, {"segment": "FIRE", **fire_cap})
        if total_cap:
            result["segments"].append({"segment": "TOTAL", **total_cap})
    return result

def parse_file(path):
    stem = os.path.basename(path).split(" ")[0]
    insurer = FILE_MAP.get(stem, f"UNKNOWN({stem})")
    result = {"file": os.path.basename(path), "insurer": insurer, "segments": [], "notes": []}
    if stem == "IFFCO":
        return parse_iffco(path, result)
    with pdfplumber.open(path) as pdf:
        loc = locate_segment_block(pdf)
        if loc is None:
            result["notes"].append("segment table not found")
            return result
        _, _, t, hdr_idx, lc = loc
        headers = t[hdr_idx]
        ci_nic = find_col(headers, "nic")
        ci_cr = find_col(headers, "cr")
        result["col_nic"] = ci_nic
        result["col_cr"] = ci_cr
        result["label_col"] = lc
        if ci_cr is None:
            result["notes"].append("Combined Ratio NOT reported segment-wise in this filing")

        def getnum(row, ci):
            return to_float(row[ci]) if ci is not None and ci < len(row) else None

        current_seg = None
        for row in t[hdr_idx + 1:]:
            label = norm(row[lc]) if lc < len(row) else ""
            # values present in columns to the right of label col?
            rest_vals = [c for ci, c in enumerate(row) if ci > lc and norm(c)]
            is_total_prefix = label.startswith("total-") or label.startswith("total -")
            total_is_current = bool(re.search(r"(current|2025|upto q4 2025|for fy 2025)", label))
            total_is_previous = bool(re.search(r"(previous|2024|upto q4 2024|for fy 2024)", label))
            already_have_total = any(s["segment"] == "TOTAL" for s in result["segments"])
            if is_total_prefix and (total_is_current or (not total_is_previous and not already_have_total)):
                result["segments"].append({"segment": "TOTAL", "nic_nep": getnum(row, ci_nic), "combined": getnum(row, ci_cr)})
                current_seg = None
            elif is_total_prefix and total_is_previous:
                continue
            elif is_current_row(label):
                if current_seg is None:
                    continue
                result["segments"].append({"segment": current_seg, "nic_nep": getnum(row, ci_nic), "combined": getnum(row, ci_cr)})
            elif is_previous_row(label):
                continue
            elif label and (not rest_vals or is_segment_name(label)):
                # segment-name row. Also catch names where header-wrap text bled
                # into other cells (e.g. Liberty FIRE row) by matching known names,
                # provided the row has no numeric value in the NIC/CR columns.
                if is_segment_name(label) and (getnum(row, ci_nic) is not None or getnum(row, ci_cr) is not None):
                    # actually a data row w/ inline label -> record it
                    result["segments"].append({"segment": str(row[lc]).replace("\n", " ").strip(),
                                                "nic_nep": getnum(row, ci_nic), "combined": getnum(row, ci_cr)})
                else:
                    current_seg = str(row[lc]).replace("\n", " ").strip()
            elif label in ("total",) and rest_vals:
                result["segments"].append({"segment": "TOTAL", "nic_nep": getnum(row, ci_nic), "combined": getnum(row, ci_cr)})
    return result

def main():
    files = sorted(glob.glob(os.path.join(SRC_DIR, "*.pdf")))
    out = []
    for f in files:
        r = parse_file(f)
        out.append(r)
        nseg = len(r["segments"])
        print(f"\n=== {r['insurer']}  ({r['file']}) cols nic={r.get('col_nic')} cr={r.get('col_cr')} segs={nseg} ===")
        if r["notes"]:
            print("  NOTES:", r["notes"])
        for s in r["segments"]:
            print(f"    {s['segment']:<44} NIC/NEP={s['nic_nep']!s:>8}  Combined={s['combined']!s:>8}")
    json.dump(out, open("nl20_parsed.json", "w"), indent=2)
    print(f"\nParsed {len(out)} files -> nl20_parsed.json")

if __name__ == "__main__":
    main()
