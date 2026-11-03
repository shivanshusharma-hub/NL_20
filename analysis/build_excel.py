#!/usr/bin/env python3
"""Build NL-20 segment-wise ratio workbook with charts.
Two metrics, full-year 'Up to the period ended 31 Mar 2026':
  - Net Incurred Claims to Net Earned Premium (%)
  - Combined Ratio (%)
Sheets:
  - 'Net Inc Claims to NEP' : insurers x segments matrix + bar chart (TOTAL)
  - 'Combined Ratio'        : insurers x segments matrix + bar chart (TOTAL)
  - 'Coverage & Notes'      : which insurers present / missing, caveats
"""
import json
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.chart import BarChart, Reference
from openpyxl.utils import get_column_letter

DATA = json.load(open("nl20_parsed.json"))

# 26 GI insurers (from NL7-GI data.json) to compute what is missing
ALL_GI = [
    "Acko General Insurance Ltd","Bajaj General Insurance Limited",
    "Cholamandalam MS General Insurance Co Ltd","Generali Central Insurance Company Limited",
    "Go Digit General Insurance Ltd","HDFC Ergo General Insurance Co Ltd",
    "ICICI Lombard General Insurance Co Ltd","IFFCO-Tokio General Insurance Co Ltd",
    "IndusInd General Insurance Company Limited","Kiwi General Insurance Ltd",
    "Kshema General Insurance","Liberty General Insurance Co. Ltd",
    "Magma General Insurance Limited","National Insurance Co Ltd",
    "Navi General Insurance Co. Ltd","Raheja QBE General Insurance Co Ltd",
    "Royal Sundaram General Insurance Co Ltd","SBI General Insurance Co Ltd",
    "Shriram General Insurance Co Ltd","Tata AIG General Insurance Co Ltd",
    "The New India Assurance Co Ltd","The Oriental Insurance Co Ltd",
    "United India Insurance Co Ltd","Universal Sompo General Insurance Co Ltd",
    "Zuno General Insurance Co Ltd","Zurich Kotak Mahindra General Insurance Co Ltd",
]

# canonical segment order for the matrix
SEG_ORDER = [
    "FIRE","Marine Cargo","Marine Hull","Total Marine",
    "Motor OD","Motor TP","Total Motor",
    "Health","Personal Accident","Travel Insurance","Total Health",
    "Workmen's Compensation","Public/ Product Liability","Engineering",
    "Aviation","Crop Insurance","Credit Insurance","Other",
    "Total Miscellaneous","TOTAL",
]

def canon_seg(name):
    n = name.strip().lower().replace("’","'")
    mp = {
        "fire":"FIRE","marine cargo":"Marine Cargo","marine hull":"Marine Hull",
        "total marine":"Total Marine","motor od":"Motor OD","motor tp":"Motor TP",
        "total motor":"Total Motor","health":"Health","personal accident":"Personal Accident",
        "travel insurance":"Travel Insurance","total health":"Total Health",
        "public/ product liability":"Public/ Product Liability","engineering":"Engineering",
        "aviation":"Aviation","crop insurance":"Crop Insurance",
        "crop insurance (incl. rural)":"Crop Insurance","credit insurance":"Credit Insurance",
        "total miscellaneous":"Total Miscellaneous","total":"TOTAL",
    }
    if n in mp: return mp[n]
    if n.startswith("workmen") or n.startswith("wc"): return "Workmen's Compensation"
    if n.startswith("other"): return "Other"
    return name.strip()

# insurers present (sorted by display name)
present = sorted({d["insurer"] for d in DATA if d["segments"]})
missing = [x for x in ALL_GI if x not in present]

# build lookup: insurer -> {canon_seg: {nic, combined}}
LUT = {}
NOTE = {}
for d in DATA:
    ins = d["insurer"]
    LUT.setdefault(ins, {})
    for s in d["segments"]:
        cs = canon_seg(s["segment"])
        LUT[ins][cs] = s
    if d.get("notes"):
        NOTE[ins] = "; ".join(d["notes"])

thin = Side(style="thin", color="BBBBBB")
border = Border(left=thin,right=thin,top=thin,bottom=thin)
hdr_fill = PatternFill("solid", fgColor="1F4E78")
tot_fill = PatternFill("solid", fgColor="FFF2CC")
na_fill  = PatternFill("solid", fgColor="F2F2F2")
hdr_font = Font(bold=True, color="FFFFFF", size=10)

def write_matrix(ws, metric_key, title):
    ws.cell(1,1,title).font = Font(bold=True, size=13)
    ws.cell(2,1,"Segment-wise — Up to the period ended 31 March 2026 (%). Parentheses/negatives shown as negative.").font = Font(italic=True, size=9, color="666666")
    r0 = 4
    # header row: Segment | insurer1 | insurer2 ...
    ws.cell(r0,1,"Segment").font = hdr_font
    ws.cell(r0,1).fill = hdr_fill
    ws.cell(r0,1).alignment = Alignment(horizontal="center", vertical="center")
    for j,ins in enumerate(present):
        c = ws.cell(r0, 2+j, ins)
        c.font = hdr_font; c.fill = hdr_fill
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True, text_rotation=90)
    for i,seg in enumerate(SEG_ORDER):
        rr = r0+1+i
        sc = ws.cell(rr,1,seg)
        sc.font = Font(bold=(seg.startswith("Total") or seg=="TOTAL" or seg=="FIRE"))
        sc.border = border
        if seg=="TOTAL" or seg.startswith("Total"):
            sc.fill = tot_fill
        for j,ins in enumerate(present):
            cell = ws.cell(rr, 2+j)
            cell.border = border
            v = LUT.get(ins,{}).get(seg)
            val = v.get(metric_key) if v else None
            if val is None:
                cell.value = "—"
                cell.fill = na_fill
                cell.alignment = Alignment(horizontal="center")
            else:
                cell.value = round(val,2)
                cell.number_format = '0.0'
                cell.alignment = Alignment(horizontal="right")
            if seg=="TOTAL" or seg.startswith("Total"):
                if val is not None:
                    cell.fill = tot_fill
    ws.column_dimensions["A"].width = 30
    for j in range(len(present)):
        ws.column_dimensions[get_column_letter(2+j)].width = 9
    ws.row_dimensions[r0].height = 150
    ws.freeze_panes = ws.cell(r0+1,2)
    return r0

def add_total_chart(ws, matrix_r0, metric_title, anchor):
    """Bar chart of the TOTAL row across insurers."""
    total_row = matrix_r0 + 1 + SEG_ORDER.index("TOTAL")
    chart = BarChart()
    chart.type = "col"
    chart.title = f"{metric_title} — Company TOTAL by Insurer (Up to 31 Mar 2026)"
    chart.y_axis.title = "%"
    chart.x_axis.title = "Insurer"
    chart.height = 9
    chart.width = 30
    data = Reference(ws, min_col=2, max_col=1+len(present), min_row=total_row, max_row=total_row)
    cats = Reference(ws, min_col=2, max_col=1+len(present), min_row=matrix_r0, max_row=matrix_r0)
    chart.add_data(data, from_rows=True, titles_from_data=False)
    chart.set_categories(cats)
    chart.legend = None
    ws.add_chart(chart, anchor)

def main():
    wb = Workbook()
    # Sheet 1: NIC/NEP
    ws1 = wb.active
    ws1.title = "Net Inc Claims to NEP"
    r0 = write_matrix(ws1, "nic_nep", "NL-20  |  Net Incurred Claims to Net Earned Premium — Segment-wise")
    add_total_chart(ws1, r0, "Net Incurred Claims / Net Earned Premium", f"{get_column_letter(3+len(present))}4")

    # Sheet 2: Combined Ratio
    ws2 = wb.create_sheet("Combined Ratio")
    r0b = write_matrix(ws2, "combined", "NL-20  |  Combined Ratio — Segment-wise")
    add_total_chart(ws2, r0b, "Combined Ratio", f"{get_column_letter(3+len(present))}4")

    # Sheet 3: coverage & notes
    ws3 = wb.create_sheet("Coverage & Notes")
    ws3.cell(1,1,"NL-20 Segment-wise Ratios — Coverage").font = Font(bold=True, size=13)
    ws3.cell(3,1,f"Insurers plotted: {len(present)} of 26 GI insurers").font = Font(bold=True)
    ws3.cell(4,1,"Source column used: 'Up to the period / quarter ended 31 March 2026' (Current Period), segment-wise table.")
    ws3.cell(5,1,"Values are percentages as filed. Negatives / parentheses in source are shown as negative.")
    rr = 7
    ws3.cell(rr,1,"Present (plotted)").font = hdr_font; ws3.cell(rr,1).fill = hdr_fill
    ws3.cell(rr,2,"File").font = hdr_font; ws3.cell(rr,2).fill = hdr_fill
    ws3.cell(rr,3,"Caveat").font = hdr_font; ws3.cell(rr,3).fill = hdr_fill
    for d in sorted(DATA, key=lambda x:x["insurer"]):
        if not d["segments"]: continue
        rr += 1
        ws3.cell(rr,1,d["insurer"])
        ws3.cell(rr,2,d["file"])
        ws3.cell(rr,3,NOTE.get(d["insurer"],""))
    rr += 2
    ws3.cell(rr,1,"MISSING (no NL-20 uploaded)").font = Font(bold=True, color="C00000")
    for m in missing:
        rr += 1
        ws3.cell(rr,1,m).font = Font(color="C00000")
    ws3.column_dimensions["A"].width = 42
    ws3.column_dimensions["B"].width = 22
    ws3.column_dimensions["C"].width = 60

    out = "NL20_Segmentwise_Ratios_FY2526_Q4.xlsx"
    wb.save(out)
    print("Wrote", out)
    print("Present:", len(present), "| Missing:", len(missing))
    print("Missing insurers:")
    for m in missing: print("  -", m)

if __name__ == "__main__":
    main()
