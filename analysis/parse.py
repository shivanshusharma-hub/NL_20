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
    # IndusInd style: "Up to Q4 2025-26"
    if re.match(r"up to q4 2025", l):
        return True
    return False

def is_previous_row(label):
    l = norm(label)
    if l in PREVIOUS_LABELS:
        return True
    if re.match(r"up to q4 2024", l):
        return True
    return False

def find_col(headers, kind):
    """kind: 'nic' or 'cr'. Tolerant to wrapped/split header text."""
    for i, h in enumerate(headers):
        hn = norm(h)
        if kind == "nic":
            if "incurred claims" in hn and "earned premium" in hn:
                return i
        elif kind == "cr":
            # handle 'combine d ratio', 'combined ratio'
            if re.search(r"combine ?d ratio", hn):
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

def pick_segment_table(pdf):
    """Return the segment table: the one whose header has a NIC column AND
    rows contain segment names like FIRE / Motor. Prefer one with a CR column."""
    candidates = []
    for pi, p in enumerate(pdf.pages):
        for t in p.extract_tables():
            if not t or not t[0]:
                continue
            hdr = t[0]
            if find_col(hdr, "nic") is None:
                continue
            bodytext = " ".join(norm(c) for row in t[1:6] for c in row if c)
            if "fire" in bodytext or "current period" in bodytext or "current year" in bodytext or "up to q4" in bodytext:
                has_cr = find_col(hdr, "cr") is not None
                candidates.append((has_cr, len(t[0]), t))
    if not candidates:
        return None
    # prefer table that has a combined-ratio column, then the widest
    candidates.sort(key=lambda x: (x[0], x[1]), reverse=True)
    return candidates[0][2]

def parse_file(path):
    stem = os.path.basename(path).split(" ")[0]
    insurer = FILE_MAP.get(stem, f"UNKNOWN({stem})")
    result = {"file": os.path.basename(path), "insurer": insurer, "segments": [], "notes": []}
    with pdfplumber.open(path) as pdf:
        t = pick_segment_table(pdf)
        if t is None:
            result["notes"].append("segment table not found")
            return result
        headers = t[0]
        ci_nic = find_col(headers, "nic")
        ci_cr = find_col(headers, "cr")
        result["col_nic"] = ci_nic
        result["col_cr"] = ci_cr
        if ci_cr is None:
            result["notes"].append("Combined Ratio NOT reported segment-wise in this filing")
        current_seg = None
        for row in t[1:]:
            c0 = norm(row[0])
            rest_vals = [c for c in row[1:] if norm(c)]
            if c0.startswith("total-current") or c0.startswith("total-upto q4 2025"):
                nic = to_float(row[ci_nic]) if ci_nic is not None and ci_nic < len(row) else None
                cr = to_float(row[ci_cr]) if ci_cr is not None and ci_cr < len(row) else None
                result["segments"].append({"segment": "TOTAL", "nic_nep": nic, "combined": cr})
                current_seg = None
            elif c0.startswith("total-previous") or c0.startswith("total-upto q4 2024"):
                continue
            elif is_current_row(c0):
                if current_seg is None:
                    continue
                nic = to_float(row[ci_nic]) if ci_nic is not None and ci_nic < len(row) else None
                cr = to_float(row[ci_cr]) if ci_cr is not None and ci_cr < len(row) else None
                result["segments"].append({"segment": current_seg, "nic_nep": nic, "combined": cr})
            elif is_previous_row(c0):
                continue
            elif c0 and not rest_vals:
                current_seg = row[0].replace("\n", " ").strip()
            elif c0 in ("total",) and rest_vals:
                # SBI style inline total
                nic = to_float(row[ci_nic]) if ci_nic is not None and ci_nic < len(row) else None
                cr = to_float(row[ci_cr]) if ci_cr is not None and ci_cr < len(row) else None
                result["segments"].append({"segment": "TOTAL", "nic_nep": nic, "combined": cr})
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
