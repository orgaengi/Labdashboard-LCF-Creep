"""
generators.py — All generate functions, no tkinter dependency
"""
import os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import lab_core as core


# ============================================================
# generate_lcf_creep
# ============================================================

"""
Tool 1 — LCF & Creep Lab Dashboard
====================================
Tracks LCF (capacity=50/yr) and Creep (capacity=22/yr) labs.
Run: python tool1_lcf_creep.py
"""

import os, sys, threading


import openpyxl

# Add current directory to path so lab_core is found



# ── Tool-specific config ──────────────────────────────────
TOOL_TITLE   = "LCF & Creep Lab Dashboard"
ALLOWED_TYPES = ["LCF", "Creep"]
OUTPUT_FILE   = "LCF_Creep_Dashboard.xlsx"

DEFAULT_CAPS = {
    "LCF":   50,   # 50 samples/year max
    "Creep": 22,   # 22 samples/year max
}

# ═══════════════════════════════════════════════════════════
#  DASHBOARD GENERATOR
# ═══════════════════════════════════════════════════════════

def generate_lcf_creep(input_path, capacities, theme, progress_cb=None):
    def p(msg):
        if progress_cb: progress_cb(msg)

    p("📖 Reading file…")
    df, col_map, errors, warns = core.load_and_filter(input_path, ALLOWED_TYPES)
    if errors:
        raise ValueError("\n".join(errors))

    p("⚙️  Computing weekly demand…")
    weekly_df, util_df, types, years = core.build_weekly(df, col_map, capacities)

    p("📗 Building workbook…")
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    p("📄 Sheet 1 — Weekly Planner…")
    core.write_weekly_planner(wb, weekly_df, types, TOOL_TITLE)

    p("🎨 Sheet 2 — Utilization…")
    core.write_utilization_sheet(wb, util_df, types, theme)

    p("📋 Sheet 3 — Summary…")
    cap_list = " | ".join(f"{k}={v}/yr" for k,v in capacities.items())
    core.write_summary_sheet(wb, weekly_df, util_df, types, capacities, years,
                              TOOL_TITLE, cap_note=cap_list)

    p("📈 Sheet 4 — Utilization Chart…")
    core.write_utilization_chart(wb, util_df, types, years)

    p("📊 Sheet 5 — Capacity vs Demand Chart…")
    core.write_capacity_chart(wb, weekly_df, types, capacities, years)

    p("📊 Sheet 6 — Year-on-Year…")
    core.write_yoy_chart(wb, weekly_df, types, years)

    p("🗓️  Sheet 7 — Gantt All Years…")
    core.write_gantt_all_years(wb, weekly_df, types, capacities, years)

    p("🔥 Sheet 8 — Gantt Heatmap…")
    core.write_gantt_heatmap(wb, weekly_df, types, capacities, years)

    p("💾 Saving…")
    out = core.save_workbook(wb, input_path, OUTPUT_FILE)
    return out, warns


# ═══════════════════════════════════════════════════════════
#  GUI
# ═══════════════════════════════════════════════════════════


# ============================================================
# generate_coating
# ============================================================

"""
Tool 2 — Coating Labs Dashboard (Cold Spray, HVOF, Plasma)
============================================================
Combined capacity = 350 samples/year across all 3 coating labs.
Individual + combined utilization tracking.
Run: python tool2_coating_labs.py
"""

import os, sys, threading, math


import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart.series import SeriesLabel
from openpyxl.utils import get_column_letter




# ── Tool-specific config ──────────────────────────────────
TOOL_TITLE    = "Coating Labs Dashboard (Cold Spray, HVOF, Plasma)"
ALLOWED_TYPES = ["Cold Spray", "HVOF", "Plasma"]
OUTPUT_FILE   = "Coating_Labs_Dashboard.xlsx"
COMBINED_CAP  = 350   # total samples/year across all 3 coating labs

# Individual capacities for single-lab utilization display
# These sum to 350 (proportional split based on typical usage)
DEFAULT_CAPS = {
    "Cold Spray": 140,  # largest portion
    "HVOF":       120,
    "Plasma":      90,
}


# ═══════════════════════════════════════════════════════════
#  COMBINED UTILIZATION SHEET
# ═══════════════════════════════════════════════════════════

def write_combined_sheet(wb, weekly_df, types, years, combined_cap):
    """Shows total combined demand across all coating labs vs 350 cap."""
    ws = wb.create_sheet("Combined_Utilization")
    core.banner(ws, 1,
                f"Combined Coating Labs — Total Demand vs {combined_cap}/yr Capacity",
                cols=20)

    weekly_cap = combined_cap / 52

    # Build combined weekly
    headers = ["Year", "Week", "Total Demand", "Capacity/Wk", "Utilization %",
               "Status"] + types
    for c, h in enumerate(headers, 1):
        core.hdr(ws.cell(3, c, h), bg="1F3864", sz=9)
        core.cw(ws, c, 14)
    core.cw(ws, 1, 8); core.cw(ws, 2, 7); core.cw(ws, 3, 14)
    core.cw(ws, 4, 12); core.cw(ws, 5, 14); core.cw(ws, 6, 12)

    red_f  = PatternFill("solid", start_color="FF4444")
    yel_f  = PatternFill("solid", start_color="FFD700")
    grn_f  = PatternFill("solid", start_color="70AD47")

    for r, row in enumerate(weekly_df.itertuples(index=False), 4):
        vals     = list(row)
        year_v   = vals[0]; week_v = vals[1]
        type_vals = {t: vals[2 + i] for i, t in enumerate(types)}
        total    = sum(type_vals.values())
        util     = total / weekly_cap if weekly_cap > 0 else 0.0

        ws.cell(r, 1, year_v); core.fmt(ws.cell(r, 1))
        ws.cell(r, 2, week_v); core.fmt(ws.cell(r, 2))

        tc = ws.cell(r, 3, round(total, 3)); core.fmt(tc, number_format="0.000")
        if util > 1.0:   tc.fill = red_f;  tc.font = Font(name="Arial", size=9, bold=True, color="FFFFFF")
        elif util >= 0.8: tc.fill = yel_f
        else:             tc.fill = grn_f;  tc.font = Font(name="Arial", size=9, color="FFFFFF")

        ws.cell(r, 4, round(weekly_cap, 2)); core.fmt(ws.cell(r, 4), number_format="0.00")

        uc = ws.cell(r, 5, round(util, 4)); core.fmt(uc, number_format="0.0%")
        if util > 1.0:   uc.fill = red_f;  uc.font = Font(name="Arial", size=9, bold=True, color="FFFFFF")
        elif util >= 0.8: uc.fill = yel_f
        else:             uc.fill = grn_f;  uc.font = Font(name="Arial", size=9, color="FFFFFF")

        status = "🔴 OVERLOAD" if util > 1.0 else "🟡 Near Cap" if util >= 0.8 else "🟢 OK"
        ws.cell(r, 6, status); core.fmt(ws.cell(r, 6))

        for ci, t in enumerate(types, 7):
            ws.cell(r, ci, round(type_vals[t], 3)); core.fmt(ws.cell(r, ci), number_format="0.000")

    ws.freeze_panes = "C4"
    ws.auto_filter.ref = f"A3:{get_column_letter(len(headers))}3"


def write_combined_chart(wb, weekly_df, types, years, combined_cap):
    """Line chart: combined weekly demand vs combined capacity."""
    ws = wb.create_sheet("Chart_Combined")
    core.banner(ws, 1,
                f"Combined Coating Labs — Weekly Demand Trend vs {combined_cap}/yr Cap",
                cols=30)

    # Build monthly aggregated combined demand
    months = ["Jan","Feb","Mar","Apr","May","Jun",
              "Jul","Aug","Sep","Oct","Nov","Dec"]
    weekly_cap = combined_cap / 52

    HR = 3
    ws.cell(HR, 1, "Month").font = Font(bold=True, name="Arial", size=9)
    ws.cell(HR, 1).fill = PatternFill("solid", start_color="D9E1F2")

    yr_cols = {}
    for i, y in enumerate(years):
        col = 2 + i
        h = ws.cell(HR, col, str(y))
        h.font = Font(bold=True, name="Arial", size=9)
        h.fill = PatternFill("solid", start_color="D9E1F2")
        h.alignment = Alignment(horizontal="center")
        core.cw(ws, col, 14)
        yr_cols[y] = col
    core.cw(ws, 1, 10)

    # Capacity line column
    cap_col = 2 + len(years)
    hc = ws.cell(HR, cap_col, f"Cap/Wk ({combined_cap}/yr)")
    hc.font = Font(bold=True, name="Arial", size=9)
    hc.fill = PatternFill("solid", start_color="FFD700")
    hc.alignment = Alignment(horizontal="center")
    core.cw(ws, cap_col, 16)

    for mi, month in enumerate(months):
        row = HR + 1 + mi
        ws.cell(row, 1, month).font = Font(bold=True, name="Arial", size=9)
        ws.cell(row, 1).fill = PatternFill("solid", start_color="EBF3FB")

        for y in years:
            yd = weekly_df[weekly_df["Year"] == y].copy()
            yd["_m"] = yd["Week"].apply(core.week_to_month)
            month_rows = yd[yd["_m"] == mi + 1]
            total_avg  = sum(month_rows[t].mean() for t in types
                             if not month_rows.empty)
            if math.isnan(total_avg): total_avg = 0.0
            cell = ws.cell(row, yr_cols[y], round(total_avg, 3))
            cell.number_format = "0.000"
            cell.font      = Font(name="Arial", size=9)
            cell.alignment = Alignment(horizontal="center")

        # Capacity reference
        cc = ws.cell(row, cap_col, round(weekly_cap, 2))
        cc.number_format = "0.00"
        cc.font      = Font(name="Arial", size=9, bold=True)
        cc.fill      = PatternFill("solid", start_color="FFF2CC")
        cc.alignment = Alignment(horizontal="center")

    # Line chart
    chart = LineChart()
    chart.title        = f"Combined Coating Labs Weekly Demand (Cap={combined_cap}/yr)"
    chart.y_axis.title = "Weekly Demand (samples)"
    chart.x_axis.title = "Month"
    chart.style        = 10
    chart.height       = 16
    chart.width        = 30

    # Add each year series
    first_col = min(yr_cols.values())
    first_ref = Reference(ws, min_col=first_col, max_col=first_col,
                          min_row=HR, max_row=HR + 12)
    chart.add_data(first_ref, titles_from_data=True)
    cats = Reference(ws, min_col=1, min_row=HR+1, max_row=HR+12)
    chart.set_categories(cats)

    for y in list(years)[1:]:
        col = yr_cols[y]
        ref = Reference(ws, min_col=col, max_col=col, min_row=HR, max_row=HR+12)
        chart.add_data(ref, titles_from_data=True)

    # Capacity reference line
    cap_ref = Reference(ws, min_col=cap_col, max_col=cap_col,
                        min_row=HR, max_row=HR+12)
    chart.add_data(cap_ref, titles_from_data=True)

    ws.add_chart(chart, f"A{HR + 15}")

    # Also add stacked bar chart per type
    chart2 = BarChart()
    chart2.type     = "col"
    chart2.grouping = "stacked"
    chart2.title    = "Demand by Type — Stacked (Monthly)"
    chart2.y_axis.title = "Weekly Demand"
    chart2.x_axis.title = "Month"
    chart2.style    = 10
    chart2.height   = 16
    chart2.width    = 28

    # Write per-type monthly data for last year
    last_yr = max(years)
    type_start_col = cap_col + 2
    ws.cell(HR, type_start_col - 1, f"Per-Type {last_yr}").font = \
        Font(bold=True, name="Arial", size=9)
    for ti, t in enumerate(types):
        col = type_start_col + ti
        h2  = ws.cell(HR, col, t)
        h2.font = Font(bold=True, name="Arial", size=9)
        h2.fill = PatternFill("solid", start_color="D9E1F2")
        h2.alignment = Alignment(horizontal="center")
        core.cw(ws, col, 13)
        yd = weekly_df[weekly_df["Year"] == last_yr].copy()
        yd["_m"] = yd["Week"].apply(core.week_to_month)
        for mi in range(12):
            row  = HR + 1 + mi
            grp  = yd[yd["_m"] == mi + 1][t]
            avg  = grp.mean() if not grp.empty else 0.0
            cell = ws.cell(row, col, round(avg if not math.isnan(avg) else 0.0, 3))
            cell.number_format = "0.000"
            cell.alignment     = Alignment(horizontal="center")
            cell.font          = Font(name="Arial", size=9)

    for ti, t in enumerate(types):
        col  = type_start_col + ti
        ref2 = Reference(ws, min_col=col, max_col=col, min_row=HR, max_row=HR+12)
        chart2.add_data(ref2, titles_from_data=True)
    chart2.set_categories(cats)
    chart2_col = get_column_letter(type_start_col)
    ws.add_chart(chart2, f"{chart2_col}{HR + 15}")


# ═══════════════════════════════════════════════════════════
#  MAIN GENERATOR
# ═══════════════════════════════════════════════════════════

def generate_coating(input_path, individual_caps, theme, progress_cb=None):
    def p(msg):
        if progress_cb: progress_cb(msg)

    p("📖 Reading file…")
    df, col_map, errors, warns = core.load_and_filter(input_path, ALLOWED_TYPES)
    if errors:
        raise ValueError("\n".join(errors))

    p("⚙️  Computing weekly demand…")
    weekly_df, util_df, types, years = core.build_weekly(df, col_map, individual_caps)

    p("📗 Building workbook…")
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    p("📄 Sheet 1 — Weekly Planner…")
    core.write_weekly_planner(wb, weekly_df, types, TOOL_TITLE)

    p("🎨 Sheet 2 — Utilization…")
    core.write_utilization_sheet(wb, util_df, types, theme)

    p("📋 Sheet 3 — Summary…")
    cap_note2 = " | ".join(f"{k}={v}/yr" for k,v in individual_caps.items())
    core.write_summary_sheet(wb, weekly_df, util_df, types, individual_caps, years,
                              TOOL_TITLE, cap_note=cap_note2)

    p("📈 Sheet 4 — Utilization Chart…")
    core.write_utilization_chart(wb, util_df, types, years)

    p("📊 Sheet 5 — Capacity vs Demand Chart…")
    core.write_capacity_chart(wb, weekly_df, types, individual_caps, years)

    p("📊 Sheet 6 — Year-on-Year…")
    core.write_yoy_chart(wb, weekly_df, types, years)

    p("🗓️  Sheet 7 — Gantt All Years…")
    core.write_gantt_all_years(wb, weekly_df, types, individual_caps, years)

    p("🔥 Sheet 8 — Gantt Heatmap…")
    core.write_gantt_heatmap(wb, weekly_df, types, individual_caps, years)

    p("💾 Saving…")
    out = core.save_workbook(wb, input_path, OUTPUT_FILE)
    return out, warns


# ═══════════════════════════════════════════════════════════
#  GUI
# ═══════════════════════════════════════════════════════════


# ============================================================
# generate_comparison
# ============================================================

"""
Tool 3 — Lab Comparison Dashboard  (Multi-File Edition)
=========================================================
Supports ONE or MULTIPLE Excel files — data is merged automatically.
  Mode A: Single file with all lab types
  Mode B: Two files (one per group)
  Mode C: Multiple files — all merged by lab type + year

Run: python tool3_comparison.py
"""

import os, sys, threading, math


import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart.series import SeriesLabel
from openpyxl.utils import get_column_letter




TOOL_TITLE  = "Lab Comparison Dashboard — Mechanical vs Coating Labs"
OUTPUT_FILE = "Lab_Comparison_Dashboard.xlsx"

GROUP_A = {"name":"Mechanical Labs","types":["LCF","Creep"],
           "caps":{"LCF":50,"Creep":22},"color":"1A4E8A","cap_total":72}
GROUP_B = {"name":"Coating Labs","types":["Cold Spray","HVOF","Plasma"],
           "caps":{"Cold Spray":140,"HVOF":120,"Plasma":90},"color":"1F5C1A","cap_total":350}
ALL_TYPES = GROUP_A["types"] + GROUP_B["types"]


# ────────────────────────────────────────────────────────────
#  MULTI-FILE MERGE ENGINE
# ────────────────────────────────────────────────────────────

def merge_files(file_paths):
    """
    Load >=1 Excel files, normalise columns, merge into single DataFrame.
    Duplicate Year+Type rows across files are SUMMED.
    Returns: (merged_df, col_map, errors, file_log)
    """
    frames = []
    errors = []
    file_log = []   # one line per file for the Data_Sources sheet

    for path in file_paths:
        fname = os.path.basename(path)
        try:
            raw = pd.read_excel(path)
        except Exception as e:
            errors.append(f"Cannot open '{fname}': {e}")
            continue

        if raw.empty:
            file_log.append(f"SKIPPED (empty): {fname}")
            continue

        col_map = core.detect_columns(raw)
        missing_cols = [k for k in ["year","type","value"] if k not in col_map]
        if missing_cols:
            errors.append(
                f"'{fname}' missing required columns: {missing_cols}. "
                f"Found: {list(raw.columns)}"
            )
            continue

        yc, tc, vc = col_map["year"], col_map["type"], col_map["value"]
        df = raw[[yc, tc, vc]].copy()
        df.columns = ["Year", "Type", "Value"]
        df["Value"] = pd.to_numeric(df["Value"], errors="coerce").fillna(0)
        df["Year"]  = pd.to_numeric(df["Year"],  errors="coerce")
        df = df.dropna(subset=["Year","Type","Value"])
        df["Year"]  = df["Year"].astype(int)
        df["Type"]  = df["Type"].astype(str).str.strip()

        # Normalise type names (case-insensitive)
        known = {t.lower(): t for t in ALL_TYPES}
        df["Type"] = df["Type"].apply(lambda x: known.get(x.lower(), x))

        rows_kept = len(df)
        frames.append(df)
        file_log.append(f"OK ({rows_kept} rows): {fname}")

    if errors:
        return None, {}, errors, file_log

    if not frames:
        errors.append("No valid data found in any file.")
        return None, {}, errors, file_log

    merged = pd.concat(frames, ignore_index=True)
    # Sum duplicates (same Year+Type from multiple files)
    merged = merged.groupby(["Year","Type"], as_index=False)["Value"].sum()

    found   = set(merged["Type"].unique())
    missing = set(ALL_TYPES) - found
    if missing:
        file_log.append(f"WARNING: Lab types not found in any file: {sorted(missing)}")

    col_map_out = {"year":"Year","type":"Type","value":"Value"}
    return merged, col_map_out, [], file_log


# ────────────────────────────────────────────────────────────
#  SHEET WRITERS
# ────────────────────────────────────────────────────────────

def write_data_sources(wb, file_paths, file_log):
    ws = wb.create_sheet("Data_Sources")
    core.banner(ws, 1, "Data Sources Loaded", cols=12)
    for c,h in enumerate(["#","File Name","Status","Full Path"],1):
        core.hdr(ws.cell(3,c,h), bg="1F3864", sz=9)
    core.cw(ws,1,5); core.cw(ws,2,35); core.cw(ws,3,18); core.cw(ws,4,60)
    for i,path in enumerate(file_paths,4):
        log = file_log[i-4] if i-4 < len(file_log) else ""
        status = "✓ OK" if log.startswith("OK") else "⚠ WARNING" if "WARNING" in log else "✗ ERROR"
        ws.cell(i,1,i-3).border=core.bdr()
        ws.cell(i,1).font=Font(name="Arial",size=9)
        c2=ws.cell(i,2,os.path.basename(path))
        c2.font=Font(name="Arial",size=9,bold=True); c2.border=core.bdr()
        c2.alignment=Alignment(horizontal="left")
        c3=ws.cell(i,3,status)
        c3.font=Font(name="Arial",size=9); c3.border=core.bdr()
        c3.fill=PatternFill("solid",start_color="C6EFCE" if "OK" in status else "FFD700")
        c3.alignment=Alignment(horizontal="center")
        c4=ws.cell(i,4,path)
        c4.font=Font(name="Arial",size=8,color="666666"); c4.border=core.bdr()
        c4.alignment=Alignment(horizontal="left")


def write_overview(wb, wdf_a, wdf_b, years):
    ws = wb.create_sheet("Overview")
    core.banner(ws, 1, f"Lab Comparison — Overview", cols=24, sz=13)
    R = 3
    for grp,wdf in [(GROUP_A,wdf_a),(GROUP_B,wdf_b)]:
        ws.merge_cells(start_row=R,start_column=1,end_row=R,end_column=3+len(years))
        c=ws.cell(R,1,f"  {grp['name']}  —  Capacity: {grp['cap_total']} samples/yr")
        c.font=Font(bold=True,size=11,name="Arial",color="FFFFFF")
        c.fill=PatternFill("solid",start_color=grp["color"])
        c.alignment=Alignment(horizontal="left",vertical="center")
        ws.row_dimensions[R].height=20; R+=1
        tgt_bg="2E75B6" if grp==GROUP_A else "375623"
        for ci,h in enumerate(["Lab Type"]+[str(y) for y in years]+["Total"],1):
            core.hdr(ws.cell(R,ci,h),bg=tgt_bg,sz=9); core.cw(ws,ci,14)
        R+=1
        for t in grp["types"]:
            ws.cell(R,1,t).font=Font(name="Arial",size=9,bold=True)
            ws.cell(R,1).border=core.bdr()
            ws.cell(R,1).alignment=Alignment(horizontal="left")
            row_total=0
            for ci,y in enumerate(years,2):
                val=wdf[wdf["Year"]==y][t].sum() if t in wdf.columns else 0
                cell=ws.cell(R,ci,round(val,1)); core.fmt(cell,number_format="#,##0.0")
                row_total+=val
            tc=ws.cell(R,2+len(years),round(row_total,1))
            core.fmt(tc,number_format="#,##0.0",bold=True)
            tc.fill=PatternFill("solid",start_color="BDD7EE" if grp==GROUP_A else "C6EFCE")
            R+=1
        ws.cell(R,1,"GROUP TOTAL")
        ws.cell(R,1).fill=PatternFill("solid",start_color=grp["color"])
        ws.cell(R,1).font=Font(name="Arial",size=9,bold=True,color="FFFFFF")
        ws.cell(R,1).border=core.bdr()
        ws.cell(R,1).alignment=Alignment(horizontal="left")
        grand=0
        for ci,y in enumerate(years,2):
            yr_tot=sum(wdf[wdf["Year"]==y][t].sum() for t in grp["types"] if t in wdf.columns)
            cell=ws.cell(R,ci,round(yr_tot,1)); core.fmt(cell,number_format="#,##0.0",bold=True)
            util=yr_tot/grp["cap_total"] if grp["cap_total"]>0 else 0
            cell.fill=PatternFill("solid",start_color="FF4444" if util>1 else "FFD700" if util>=0.8 else "C6EFCE")
            if util>1: cell.font=Font(name="Arial",size=9,bold=True,color="FFFFFF")
            grand+=yr_tot
        tc=ws.cell(R,2+len(years),round(grand,1)); core.fmt(tc,number_format="#,##0.0",bold=True)
        tc.fill=PatternFill("solid",start_color=grp["color"])
        tc.font=Font(name="Arial",size=9,bold=True,color="FFFFFF"); R+=3
    # Util summary
    core.section_hdr(ws,R,"Annual Utilization vs Capacity",3+len(years)); R+=1
    for ci,h in enumerate(["Lab Group","Cap/yr"]+[str(y) for y in years],1):
        core.hdr(ws.cell(R,ci,h),bg="C00000",sz=9); core.cw(ws,ci,16)
    R+=1
    for grp,wdf in [(GROUP_A,wdf_a),(GROUP_B,wdf_b)]:
        ws.cell(R,1,grp["name"])
        ws.cell(R,1).fill=PatternFill("solid",start_color=grp["color"])
        ws.cell(R,1).font=Font(name="Arial",size=9,bold=True,color="FFFFFF")
        ws.cell(R,1).border=core.bdr(); ws.cell(R,1).alignment=Alignment(horizontal="left")
        ws.cell(R,2,grp["cap_total"]); core.fmt(ws.cell(R,2),number_format="#,##0")
        for ci,y in enumerate(years,3):
            tot=sum(wdf[wdf["Year"]==y][t].sum() for t in grp["types"] if t in wdf.columns)
            util=tot/grp["cap_total"] if grp["cap_total"]>0 else 0
            cell=ws.cell(R,ci,round(util,4)); core.fmt(cell,number_format="0.0%")
            if util>1.0: cell.fill=PatternFill("solid",start_color="FF4444"); cell.font=Font(name="Arial",size=9,bold=True,color="FFFFFF")
            elif util>=0.8: cell.fill=PatternFill("solid",start_color="FFD700"); cell.font=Font(name="Arial",size=9,bold=True)
            else: cell.fill=PatternFill("solid",start_color="70AD47"); cell.font=Font(name="Arial",size=9,color="FFFFFF")
        R+=1


def write_comparison_chart(wb, wdf_a, wdf_b, years):
    ws = wb.create_sheet("Chart_Comparison")
    core.banner(ws,1,"Group Comparison — Annual Demand vs Capacity",cols=24)
    months=["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]
    HR=3; D0=4
    for ci,h in enumerate(["Year",f"{GROUP_A['name']} Demand",f"Cap({GROUP_A['cap_total']})",
                            f"{GROUP_B['name']} Demand",f"Cap({GROUP_B['cap_total']})"],1):
        core.hdr(ws.cell(HR,ci,h),bg="1F3864",sz=9); core.cw(ws,ci,22)
    DE=D0+len(years)-1
    for ri,y in enumerate(years,D0):
        tot_a=sum(wdf_a[wdf_a["Year"]==y][t].sum() for t in GROUP_A["types"] if t in wdf_a.columns)
        tot_b=sum(wdf_b[wdf_b["Year"]==y][t].sum() for t in GROUP_B["types"] if t in wdf_b.columns)
        ws.cell(ri,1,y); core.fmt(ws.cell(ri,1))
        for ci,tot,cap in [(2,tot_a,GROUP_A["cap_total"]),(4,tot_b,GROUP_B["cap_total"])]:
            cell=ws.cell(ri,ci,round(tot,1)); core.fmt(cell,number_format="#,##0.0")
            util=tot/cap if cap>0 else 0
            cell.fill=PatternFill("solid",start_color="FF4444" if util>1 else "FFD700" if util>=0.8 else "C6EFCE")
            if util>1: cell.font=Font(name="Arial",size=9,bold=True,color="FFFFFF")
        ws.cell(ri,3,GROUP_A["cap_total"]); core.fmt(ws.cell(ri,3),number_format="#,##0")
        ws.cell(ri,5,GROUP_B["cap_total"]); core.fmt(ws.cell(ri,5),number_format="#,##0")
    chart=BarChart(); chart.type="col"; chart.grouping="clustered"
    chart.title="Annual Demand — Mechanical vs Coating Labs"
    chart.y_axis.title="Samples/Year"; chart.x_axis.title="Year"
    chart.style=10; chart.height=15; chart.width=28
    cats=Reference(ws,min_col=1,min_row=D0,max_row=DE)
    first=Reference(ws,min_col=2,max_col=2,min_row=HR,max_row=DE)
    chart.add_data(first,titles_from_data=True); chart.set_categories(cats)
    for ci in [4,3,5]:
        ref=Reference(ws,min_col=ci,max_col=ci,min_row=HR,max_row=DE)
        chart.add_data(ref,titles_from_data=True)
    ws.add_chart(chart,f"A{DE+4}")
    # Monthly util line
    UR=DE+22
    core.section_hdr(ws,UR,"Monthly Utilization % (last 2 years)",10); UR+=1
    ws.cell(UR,1,"Month").font=Font(bold=True,name="Arial",size=9)
    ws.cell(UR,1).fill=PatternFill("solid",start_color="D9E1F2")
    plot_years=list(years[-2:]) if len(years)>=2 else list(years)
    col_a={y:2+2*i for i,y in enumerate(plot_years)}
    col_b={y:3+2*i for i,y in enumerate(plot_years)}
    for y in plot_years:
        for col,label in [(col_a[y],f"Mech {y}"),(col_b[y],f"Coating {y}")]:
            h=ws.cell(UR,col,label); h.font=Font(bold=True,name="Arial",size=8)
            h.fill=PatternFill("solid",start_color="D9E1F2")
            h.alignment=Alignment(horizontal="center"); core.cw(ws,col,14)
    core.cw(ws,1,10)
    for mi,month in enumerate(months):
        row=UR+1+mi
        ws.cell(row,1,month).font=Font(bold=True,name="Arial",size=9)
        ws.cell(row,1).fill=PatternFill("solid",start_color="EBF3FB")
        for y in plot_years:
            for wdf,grp,col in [(wdf_a,GROUP_A,col_a[y]),(wdf_b,GROUP_B,col_b[y])]:
                yd=wdf[wdf["Year"]==y].copy(); yd["_m"]=yd["Week"].apply(core.week_to_month)
                tot=sum(yd[yd["_m"]==mi+1][t].mean() for t in grp["types"] if t in yd.columns)
                util=tot/(grp["cap_total"]/52) if grp["cap_total"]>0 else 0
                if math.isnan(util): util=0.0
                cell=ws.cell(row,col,round(util,4)); cell.number_format="0.0%"
                cell.font=Font(name="Arial",size=9); cell.alignment=Alignment(horizontal="center")
    chart2=LineChart()
    chart2.title="Monthly Utilization — Mechanical vs Coating Labs"
    chart2.y_axis.title="Utilization %"; chart2.x_axis.title="Month"
    chart2.y_axis.numFmt="0%"; chart2.y_axis.scaling.min=0
    chart2.style=10; chart2.height=14; chart2.width=28
    all_cols=list(col_a.values())+list(col_b.values())
    first_col=min(all_cols)
    first_ref=Reference(ws,min_col=first_col,max_col=first_col,min_row=UR,max_row=UR+12)
    chart2.add_data(first_ref,titles_from_data=True)
    cats2=Reference(ws,min_col=1,min_row=UR+1,max_row=UR+12); chart2.set_categories(cats2)
    for col in all_cols:
        if col==first_col: continue
        ref=Reference(ws,min_col=col,max_col=col,min_row=UR,max_row=UR+12)
        chart2.add_data(ref,titles_from_data=True)
    ws.add_chart(chart2,f"A{UR+15}")


def write_all_labs(wb, wdf_a, wdf_b, years, types_a, types_b):
    ws=wb.create_sheet("All_Labs_Summary")
    core.banner(ws,1,"All Labs — Combined Annual Demand Summary",cols=20)
    all_types=types_a+types_b
    all_wdfs={t:wdf_a for t in types_a}; all_wdfs.update({t:wdf_b for t in types_b})
    all_caps={**GROUP_A["caps"],**GROUP_B["caps"]}
    HR=3; D0=4
    hdrs=["Year"]+all_types+["Mech Total","Coating Total","Grand Total"]
    for ci,h in enumerate(hdrs,1):
        bg="1F3864" if ci==1 else GROUP_A["color"] if ci<=1+len(types_a) else GROUP_B["color"] if ci<=1+len(all_types) else "C00000"
        core.hdr(ws.cell(HR,ci,h),bg=bg,sz=9); core.cw(ws,ci,13)
    for ri,y in enumerate(years,D0):
        ws.cell(ri,1,y); core.fmt(ws.cell(ri,1))
        mech_tot=coat_tot=0
        for ci,t in enumerate(all_types,2):
            wdf=all_wdfs[t]; val=wdf[wdf["Year"]==y][t].sum() if t in wdf.columns else 0
            cell=ws.cell(ri,ci,round(val,1)); core.fmt(cell,number_format="#,##0.0")
            cap=all_caps.get(t,1); util=val/cap if cap>0 else 0
            cell.fill=PatternFill("solid",start_color="FF4444" if util>1 else "FFD700" if util>=0.8 else "C6EFCE")
            if util>1: cell.font=Font(name="Arial",size=9,bold=True,color="FFFFFF")
            if t in types_a: mech_tot+=val
            else: coat_tot+=val
        lc=1+len(all_types)
        mc=ws.cell(ri,lc+1,round(mech_tot,1)); core.fmt(mc,number_format="#,##0.0",bold=True); mc.fill=PatternFill("solid",start_color="BDD7EE")
        cc=ws.cell(ri,lc+2,round(coat_tot,1)); core.fmt(cc,number_format="#,##0.0",bold=True); cc.fill=PatternFill("solid",start_color="C6EFCE")
        gc=ws.cell(ri,lc+3,round(mech_tot+coat_tot,1)); core.fmt(gc,number_format="#,##0.0",bold=True); gc.fill=PatternFill("solid",start_color="D9D9D9")
    chart=BarChart(); chart.type="col"; chart.grouping="clustered"
    chart.title="All Labs — Annual Demand"; chart.y_axis.title="Samples/Year"
    chart.style=10; chart.height=16; chart.width=34
    cats=Reference(ws,min_col=1,min_row=D0,max_row=D0+len(years)-1)
    first=Reference(ws,min_col=2,max_col=2,min_row=HR,max_row=D0+len(years)-1)
    chart.add_data(first,titles_from_data=True); chart.set_categories(cats)
    for ci in range(3,2+len(all_types)):
        ref=Reference(ws,min_col=ci,max_col=ci,min_row=HR,max_row=D0+len(years)-1)
        chart.add_data(ref,titles_from_data=True)
    ws.add_chart(chart,f"A{D0+len(years)+4}")


# ────────────────────────────────────────────────────────────
#  MAIN GENERATOR
# ────────────────────────────────────────────────────────────

def generate_comparison(file_paths, caps_a, caps_b, theme, progress_cb=None):
    def p(msg):
        if progress_cb: progress_cb(msg)

    p(f"Loading {len(file_paths)} file(s)…")
    merged_df, col_map, errors, file_log = merge_files(file_paths)
    if errors:
        raise ValueError("\n".join(errors))

    tc = col_map["type"]
    df_a = merged_df[merged_df[tc].isin(GROUP_A["types"])].copy()
    df_b = merged_df[merged_df[tc].isin(GROUP_B["types"])].copy()
    if df_a.empty:
        raise ValueError(f"No data for Mechanical labs {GROUP_A['types']}.")
    if df_b.empty:
        raise ValueError(f"No data for Coating labs {GROUP_B['types']}.")

    p("Computing weekly data…")
    wdf_a,udf_a,types_a,years_a = core.build_weekly(df_a, col_map, caps_a)
    wdf_b,udf_b,types_b,years_b = core.build_weekly(df_b, col_map, caps_b)
    years = sorted(set(list(years_a)) | set(list(years_b)))

    p("Building workbook…")
    wb = openpyxl.Workbook(); wb.remove(wb.active)

    p("Sheet 1/8 — Data Sources…")
    write_data_sources(wb, file_paths, file_log)
    p("Sheet 2/8 — Overview…")
    write_overview(wb, wdf_a, wdf_b, years)
    p("Sheet 3/8 — All Labs Summary…")
    write_all_labs(wb, wdf_a, wdf_b, years, types_a, types_b)
    p("Sheet 4/8 — Comparison Charts…")
    write_comparison_chart(wb, wdf_a, wdf_b, years)
    p("Sheet 5/8 — Mech Utilization…")
    core.write_utilization_sheet(wb, udf_a, types_a, theme)
    wb.worksheets[-1].title = "Mech_Utilization"
    p("Sheet 6/8 — Coating Utilization…")
    core.write_utilization_sheet(wb, udf_b, types_b, theme)
    wb.worksheets[-1].title = "Coating_Utilization"
    p("Sheet 7/8 — Mech YoY…")
    core.write_yoy_chart(wb, wdf_a, types_a, years_a)
    wb.worksheets[-1].title = "Mech_YoY"
    p("Sheet 8/8 — Coating YoY…")
    core.write_yoy_chart(wb, wdf_b, types_b, years_b)
    wb.worksheets[-1].title = "Coating_YoY"

    p("Saving…")
    out = core.save_workbook(wb, file_paths[0], OUTPUT_FILE)
    return out, file_log


# ────────────────────────────────────────────────────────────
#  GUI
# ────────────────────────────────────────────────────────────

