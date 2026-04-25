"""
Lab Planning & Occupancy Dashboard — Streamlit Web App
=======================================================
All 3 tools in one web interface.
Run: streamlit run app.py
"""

import os
import sys
import math
import tempfile
import io
import datetime

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import lab_core as core
import generators as gen_module
from generators import merge_files, GROUP_A, GROUP_B

# ──────────────────────────────────────────────────────────
#  PAGE CONFIG
# ──────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Lab Planning Dashboard",
    page_icon="🔬",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ──────────────────────────────────────────────────────────
#  CUSTOM CSS
# ──────────────────────────────────────────────────────────
st.markdown("""
<style>
    /* Main background */
    .stApp { background-color: #F0F4F8; }

    /* Sidebar */
    section[data-testid="stSidebar"] {
        background: linear-gradient(180deg, #1F3864 0%, #2E75B6 100%);
    }
    section[data-testid="stSidebar"] * { color: white !important; }
    section[data-testid="stSidebar"] .stRadio label { color: white !important; }
    section[data-testid="stSidebar"] p { color: #BDD7EE !important; font-size: 13px; }

    /* Tool headers */
    .tool-header {
        background: linear-gradient(90deg, #1F3864 0%, #2E75B6 100%);
        color: white; padding: 18px 24px; border-radius: 10px;
        margin-bottom: 20px;
    }
    .tool-header h2 { color: white !important; margin: 0; font-size: 22px; }
    .tool-header p  { color: #BDD7EE !important; margin: 4px 0 0 0; font-size: 13px; }

    /* Metric cards */
    .metric-card {
        background: white; border-radius: 8px; padding: 16px 20px;
        border-left: 4px solid #2E75B6; margin: 6px 0;
        box-shadow: 0 1px 4px rgba(0,0,0,0.08);
    }
    .metric-card.red   { border-left-color: #FF4444; }
    .metric-card.green { border-left-color: #70AD47; }
    .metric-card.amber { border-left-color: #FFD700; }

    /* Section headers */
    .section-label {
        font-weight: 700; font-size: 14px; color: #1F3864;
        margin: 20px 0 8px 0; padding-bottom: 4px;
        border-bottom: 2px solid #2E75B6;
    }

    /* Capacity input group */
    .cap-group {
        background: white; border-radius: 8px; padding: 14px 16px;
        border: 1px solid #D9E1F2; margin: 6px 0;
    }

    /* Status badges */
    .badge-red    { background:#FF4444; color:white; padding:2px 8px; border-radius:12px; font-size:11px; font-weight:700; }
    .badge-yellow { background:#FFD700; color:#333; padding:2px 8px; border-radius:12px; font-size:11px; font-weight:700; }
    .badge-green  { background:#70AD47; color:white; padding:2px 8px; border-radius:12px; font-size:11px; font-weight:700; }

    /* Download button */
    .stDownloadButton > button {
        background: #1F5C1A !important; color: white !important;
        border-radius: 8px !important; font-weight: 700 !important;
        padding: 10px 24px !important; font-size: 15px !important;
        width: 100% !important;
    }

    /* Generate button */
    .stButton > button {
        background: #1F3864 !important; color: white !important;
        border-radius: 8px !important; font-weight: 700 !important;
        padding: 10px 24px !important; font-size: 15px !important;
        width: 100% !important;
    }

    /* File uploader */
    [data-testid="stFileUploader"] {
        border: 2px dashed #2E75B6 !important;
        border-radius: 10px !important; padding: 10px !important;
    }

    /* Tab styling */
    .stTabs [data-baseweb="tab-list"] { gap: 4px; }
    .stTabs [data-baseweb="tab"] {
        border-radius: 6px 6px 0 0 !important;
        background: #D9E1F2 !important; color: #1F3864 !important;
        font-weight: 600 !important; padding: 8px 20px !important;
    }
    .stTabs [aria-selected="true"] {
        background: #1F3864 !important; color: white !important;
    }

    div[data-testid="stExpander"] {
        border: 1px solid #D9E1F2 !important;
        border-radius: 8px !important;
    }

    /* Chat bubbles */
    .chat-user {
        background: #1F3864; color: white;
        padding: 8px 12px; border-radius: 12px 12px 2px 12px;
        margin: 4px 0; font-size: 12px; max-width: 92%;
        margin-left: auto; text-align: right;
    }
    .chat-bot {
        background: #E8F0FE; color: #1F3864;
        padding: 8px 12px; border-radius: 12px 12px 12px 2px;
        margin: 4px 0; font-size: 12px; max-width: 92%;
        border-left: 3px solid #2E75B6;
    }
    .chat-wrap { max-height: 320px; overflow-y: auto; padding: 4px; }
</style>
""", unsafe_allow_html=True)


# ──────────────────────────────────────────────────────────
#  CONSTANTS
# ──────────────────────────────────────────────────────────
THEME_COLORS = {
    "Default (Red/Yellow/Green)": {"high":"#FF4444","medium":"#FFD700","low":"#70AD47"},
    "Blue Gradient":              {"high":"#1F4E79","medium":"#2E75B6","low":"#BDD7EE"},
    "Pastel":                     {"high":"#FF9999","medium":"#FFEB99","low":"#C6EFCE"},
}

CURRENT_YEAR = datetime.date.today().year
CURRENT_WEEK = datetime.date.today().isocalendar()[1]

GROUP_A = {"name":"Mechanical Labs",  "types":["LCF","Creep"],
           "color":"#1A4E8A","cap_total":72}
GROUP_B = {"name":"Coating Labs",     "types":["Cold Spray","HVOF","Plasma"],
           "color":"#1F5C1A","cap_per_lab":350}  # each lab independently 350/yr


# ──────────────────────────────────────────────────────────
#  HELPERS
# ──────────────────────────────────────────────────────────

def load_uploaded_files(uploaded_files, allowed_types=None):
    """Save uploaded Streamlit files to temp dir, return paths."""
    paths = []
    for f in uploaded_files:
        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
        tmp.write(f.getbuffer())
        tmp.close()
        paths.append(tmp.name)
    return paths

def get_annual_totals(weekly_df, types, years):
    """Return dict {year: {type: annual_total}}."""
    result = {}
    for y in years:
        yd = weekly_df[weekly_df["Year"] == y]
        result[int(y)] = {t: round(yd[t].sum(), 1) for t in types}
    return result

def util_badge(util):
    if util > 1.0:   return f'<span class="badge-red">🔴 Overloaded {util:.0%}</span>'
    if util >= 0.8:  return f'<span class="badge-yellow">🟡 Near Cap {util:.0%}</span>'
    return               f'<span class="badge-green">🟢 OK {util:.0%}</span>'

def generate_excel(tool_fn, *args, **kwargs):
    """Run generator in temp dir, return bytes of Excel file."""
    with tempfile.TemporaryDirectory() as tmpdir:
        dummy_input = os.path.join(tmpdir, "input.xlsx")
        # Create a dummy input path so save_workbook knows where to save
        open(dummy_input, 'wb').close()
        out_path, warns = tool_fn(*args, **kwargs)
        with open(out_path, "rb") as f:
            return f.read(), warns


# ──────────────────────────────────────────────────────────
#  PLOTLY CHART BUILDERS
# ──────────────────────────────────────────────────────────

def chart_yoy_bar_and_pie(annual_totals, types, years, colors):
    """YoY grouped bar chart + pie charts per year."""
    # --- Bar chart ---
    fig_bar = go.Figure()
    bar_colors = ["#1F3864","#2E75B6","#70AD47","#FFD700","#FF4444","#9DC3E6","#C6EFCE"]

    for i, t in enumerate(types):
        y_vals = [annual_totals[int(y)][t] for y in years]
        fig_bar.add_trace(go.Bar(
            name=t,
            x=[str(y) for y in years],
            y=y_vals,
            marker_color=bar_colors[i % len(bar_colors)],
            text=[f"{v:.0f}" for v in y_vals],
            textposition="outside",
        ))

    fig_bar.update_layout(
        barmode="group",
        title="Year-on-Year Demand by Lab Type",
        xaxis_title="Year", yaxis_title="Samples/Year",
        legend=dict(orientation="h", yanchor="top", y=-0.15, x=0),
        height=460, plot_bgcolor="white", paper_bgcolor="white",
        font=dict(family="Arial", size=11),
        margin=dict(t=50, b=110, r=20),
        uniformtext=dict(minsize=8, mode="hide"),
    )
    fig_bar.update_xaxes(showgrid=False)
    fig_bar.update_yaxes(showgrid=True, gridcolor="#F0F0F0")

    # --- Pie charts: one per year ---
    n_years = len(years)
    cols_per_row = min(n_years, 3)
    rows = math.ceil(n_years / cols_per_row)

    fig_pies = make_subplots(
        rows=rows, cols=cols_per_row,
        specs=[[{"type":"pie"}]*cols_per_row]*rows,
        subplot_titles=[str(y) for y in years],
    )
    pie_colors = ["#1F3864","#2E75B6","#70AD47","#FFD700","#FF4444","#9DC3E6","#C6EFCE"]

    for i, year in enumerate(years):
        row = i // cols_per_row + 1
        col = i %  cols_per_row + 1
        vals   = [annual_totals[int(year)][t] for t in types]
        total  = sum(vals)
        labels = [f"{t}<br>{v:.0f} ({v/total:.0%})" if total > 0 else t
                  for t, v in zip(types, vals)]
        fig_pies.add_trace(
            go.Pie(
                labels=types, values=vals,
                name=str(year),
                marker=dict(colors=pie_colors[:len(types)]),
                textinfo="label+percent",
                hovertemplate="<b>%{label}</b><br>%{value:.0f} samples<br>%{percent}<extra></extra>",
            ),
            row=row, col=col,
        )

    fig_pies.update_layout(
        title="Process-wise Demand Share (%) by Year",
        height=320 * rows,
        showlegend=True,
        legend=dict(orientation="h", yanchor="bottom", y=-0.15),
        font=dict(family="Arial", size=12),
        paper_bgcolor="white",
    )

    return fig_bar, fig_pies


def chart_utilization_line(util_df, types, years, colors):
    """Monthly utilization line chart."""
    months = ["Jan","Feb","Mar","Apr","May","Jun",
              "Jul","Aug","Sep","Oct","Nov","Dec"]
    line_colors = ["#1F3864","#2E75B6","#70AD47","#FF4444","#FFD700","#9DC3E6"]

    fig = go.Figure()
    for yi, year in enumerate(years):
        yd = util_df[util_df["Year"] == year].copy()
        yd["_m"] = yd["Week"].apply(core.week_to_month)
        for ti, t in enumerate(types):
            monthly = [yd[yd["_m"] == m+1][t].mean() for m in range(12)]
            monthly = [v if not math.isnan(v) else 0 for v in monthly]
            fig.add_trace(go.Scatter(
                x=months, y=monthly,
                name=f"{t} ({year})",
                line=dict(color=line_colors[ti % len(line_colors)],
                          width=2,
                          dash="solid" if yi == 0 else "dash"),
                mode="lines+markers",
                marker=dict(size=6),
                hovertemplate=f"<b>{t} ({year})</b><br>%{{x}}: %{{y:.1%}}<extra></extra>",
            ))

    # 100% reference line
    fig.add_hline(y=1.0, line_dash="dot", line_color="#FF4444",
                  annotation_text="100% Capacity", annotation_position="right")
    fig.add_hline(y=0.8, line_dash="dot", line_color="#FFD700",
                  annotation_text="80% Threshold", annotation_position="right")

    fig.update_layout(
        title="Monthly Utilization Trend by Lab Type",
        xaxis_title="Month", yaxis_title="Utilization",
        yaxis_tickformat=".0%",
        height=440, plot_bgcolor="white", paper_bgcolor="white",
        font=dict(family="Arial", size=12),
        legend=dict(orientation="h", yanchor="bottom", y=-0.4),
        margin=dict(b=120),
    )
    fig.update_xaxes(showgrid=False)
    fig.update_yaxes(showgrid=True, gridcolor="#F0F0F0")
    return fig


def chart_capacity_bar(weekly_df, types, capacities, years):
    """Capacity vs Avg Weekly Demand bar chart — no text overlap."""
    fig = go.Figure()
    cap_weekly = {t: capacities.get(t, 1) / 52 for t in types}

    # Capacity bars — use "inside" to avoid crowding at top
    fig.add_trace(go.Bar(
        name="Weekly Capacity",
        x=types,
        y=[round(cap_weekly[t], 2) for t in types],
        marker_color="#D9E1F2",
        marker_line_color="#1F3864", marker_line_width=1.5,
        text=[f"Cap: {cap_weekly[t]:.2f}" for t in types],
        textposition="inside",
        insidetextanchor="start",
        textfont=dict(size=9, color="#1F3864"),
    ))

    bar_colors = ["#1F3864","#2E75B6","#70AD47","#FF4444","#FFD700","#9DC3E6","#C6EFCE"]
    for yi, year in enumerate(years):
        yd   = weekly_df[weekly_df["Year"] == year]
        avgs = [round(yd[t].mean(), 3) for t in types]
        fig.add_trace(go.Bar(
            name=f"{year} Avg Demand",
            x=types, y=avgs,
            marker_color=bar_colors[yi % len(bar_colors)],
            text=[f"{v:.2f}" for v in avgs],
            textposition="outside",
            textfont=dict(size=9),
        ))

    # Calculate a clean yaxis max with headroom so outside labels don't clip
    all_vals = [cap_weekly[t] for t in types]
    for y in years:
        yd = weekly_df[weekly_df["Year"] == y]
        all_vals += [yd[t].mean() for t in types]
    y_max = max((v for v in all_vals if not math.isnan(v)), default=1) * 1.25

    fig.update_layout(
        barmode="group",
        title="Average Weekly Demand vs Capacity",
        xaxis_title="Lab Type", yaxis_title="Weekly Units",
        yaxis=dict(range=[0, y_max]),
        height=460, plot_bgcolor="white", paper_bgcolor="white",
        font=dict(family="Arial", size=11),
        legend=dict(orientation="h", yanchor="top", y=-0.15, x=0),
        margin=dict(b=100, t=60, r=20),
        uniformtext=dict(minsize=8, mode="hide"),   # hide labels that cannot fit
    )
    fig.update_xaxes(showgrid=False)
    fig.update_yaxes(showgrid=True, gridcolor="#F0F0F0")
    return fig


def chart_gantt(weekly_df, types, capacities, year, current_week):
    """Gantt heatmap for a given year."""
    weeks = list(range(1, 53))
    yd    = weekly_df[weekly_df["Year"] == year]

    z, labels, text_grid = [], [], []
    for i, t in enumerate(types):
        cap_wk = capacities.get(t, 1) / 52
        row_z, row_txt = [], []
        for w in weeks:
            wrow = yd[yd["Week"] == w]
            dem  = float(wrow[t].values[0]) if not wrow.empty else 0.0
            util = dem / cap_wk if cap_wk > 0 else 0.0
            # clamp to [0, 1.5] for display but store raw for tooltip
            row_z.append(round(min(util, 1.5), 3))
            row_txt.append(f"Week {w} | {t}<br>Demand: {dem:.2f}<br>Util: {util:.1%}")
        z.append(row_z)
        labels.append(t)
        text_grid.append(row_txt)

    # Colorscale values MUST be 0.0–1.0 (plotly requirement)
    # Map: 0→green, 0.53→yellow (=80% of 1.5 scale), 0.67→red (=100%), 1.0→dark red
    colorscale = [
        [0.00, "#C6EFCE"],   # 0% util → light green
        [0.40, "#70AD47"],   # ~60% util → green
        [0.53, "#FFD700"],   # ~80% util → yellow
        [0.67, "#FF4444"],   # ~100% util → red
        [1.00, "#CC0000"],   # ≥150% util → dark red
    ]

    fig = go.Figure(go.Heatmap(
        z=z,
        x=[f"Wk{w}" for w in weeks],
        y=labels,
        colorscale=colorscale,
        zmin=0, zmax=1.5,
        text=text_grid,
        hovertemplate="%{text}<extra></extra>",
        colorbar=dict(
            title=dict(text="Utilization", side="right"),
            tickvals=[0, 0.4, 0.8, 1.0, 1.5],
            ticktext=["0%", "40%", "80%", "100%", "≥150%"],
            len=0.8,
        ),
        showscale=True,
    ))

    # Mark current week
    if year == CURRENT_YEAR and 1 <= current_week <= 52:
        fig.add_vline(
            x=current_week - 1,          # vline on x-axis uses index not label
            line_color="#C00000", line_width=2, line_dash="dash",
            annotation_text=f"Wk {current_week} (Now)",
            annotation_position="top right",
            annotation_font_size=10,
        )

    fig.update_layout(
        title=f"Gantt Heatmap — {year} Weekly Occupancy",
        xaxis_title="Week",
        yaxis_title="Lab Type",
        height=max(280, 90 * len(types) + 140),
        font=dict(family="Arial", size=11),
        paper_bgcolor="white",
        plot_bgcolor="white",
        margin=dict(l=110, r=80, t=60, b=60),
    )
    return fig


def chart_comparison_grouped(annual_a, annual_b, types_a, types_b, years,
                              cap_a_total, cap_b_total):
    """Grouped bar: both groups side by side per year."""
    fig = make_subplots(rows=1, cols=2,
                        subplot_titles=[GROUP_A["name"], GROUP_B["name"]])

    col_map_a = ["#1F3864","#2E75B6","#9DC3E6"]
    col_map_b = ["#1F5C1A","#70AD47","#C6EFCE"]

    for ti, t in enumerate(types_a):
        vals = [annual_a[int(y)].get(t, 0) for y in years]
        fig.add_trace(go.Bar(name=t, x=[str(y) for y in years], y=vals,
                             marker_color=col_map_a[ti % 3],
                             legendgroup="A",
                             text=[f"{v:.0f}" for v in vals],
                             textposition="outside"), row=1, col=1)

    for ti, t in enumerate(types_b):
        vals = [annual_b[int(y)].get(t, 0) for y in years]
        fig.add_trace(go.Bar(name=t, x=[str(y) for y in years], y=vals,
                             marker_color=col_map_b[ti % 3],
                             legendgroup="B",
                             text=[f"{v:.0f}" for v in vals],
                             textposition="outside"), row=1, col=2)

    # Capacity reference lines
    fig.add_hline(y=cap_a_total, line_dash="dot", line_color="#1A4E8A",
                  annotation_text=f"Cap A: {cap_a_total}", row=1, col=1)
    fig.add_hline(y=cap_b_total, line_dash="dot", line_color="#1F5C1A",
                  annotation_text=f"Cap B: {cap_b_total}", row=1, col=2)

    fig.update_layout(barmode="group", height=440,
                      font=dict(family="Arial", size=12),
                      paper_bgcolor="white", plot_bgcolor="white",
                      title="Annual Demand Comparison — Mechanical vs Coating Labs",
                      margin=dict(b=60))
    return fig


def chart_util_comparison_line(wdf_a, wdf_b, types_a, types_b, years,
                                cap_a_total, cap_b_total):
    """Line chart: group-level utilization over time."""
    months = ["Jan","Feb","Mar","Apr","May","Jun",
              "Jul","Aug","Sep","Oct","Nov","Dec"]
    fig = go.Figure()

    for yi, year in enumerate(list(years)[-3:]):  # last 3 years for clarity
        # Group A
        yd_a = wdf_a[wdf_a["Year"] == year].copy()
        yd_a["_m"] = yd_a["Week"].apply(core.week_to_month)
        util_a = []
        for m in range(12):
            tot = sum(yd_a[yd_a["_m"]==m+1][t].mean()
                      for t in types_a if t in yd_a.columns)
            weekly_cap = cap_a_total / 52
            util_a.append(tot / weekly_cap if weekly_cap > 0 and not math.isnan(tot) else 0)

        fig.add_trace(go.Scatter(
            x=months, y=util_a, name=f"Mechanical {year}",
            line=dict(color="#1A4E8A", width=2,
                      dash="solid" if yi == 0 else "dash"),
            mode="lines+markers", marker=dict(size=5),
        ))

        # Group B
        yd_b = wdf_b[wdf_b["Year"] == year].copy()
        yd_b["_m"] = yd_b["Week"].apply(core.week_to_month)
        util_b = []
        for m in range(12):
            tot = sum(yd_b[yd_b["_m"]==m+1][t].mean()
                      for t in types_b if t in yd_b.columns)
            weekly_cap = cap_b_total / 52
            util_b.append(tot / weekly_cap if weekly_cap > 0 and not math.isnan(tot) else 0)

        fig.add_trace(go.Scatter(
            x=months, y=util_b, name=f"Coating {year}",
            line=dict(color="#1F5C1A", width=2,
                      dash="solid" if yi == 0 else "dash"),
            mode="lines+markers", marker=dict(size=5),
        ))

    fig.add_hline(y=1.0, line_dash="dot", line_color="#FF4444",
                  annotation_text="100%")
    fig.add_hline(y=0.8, line_dash="dot", line_color="#FFD700",
                  annotation_text="80%")

    fig.update_layout(
        title="Monthly Utilization % — Mechanical vs Coating Labs (last 3 years)",
        xaxis_title="Month", yaxis_title="Utilization",
        yaxis_tickformat=".0%", height=420,
        font=dict(family="Arial", size=12),
        paper_bgcolor="white", plot_bgcolor="white",
        legend=dict(orientation="h", yanchor="bottom", y=-0.4),
        margin=dict(b=120),
    )
    return fig


# ──────────────────────────────────────────────────────────
#  SHARED UI COMPONENTS
# ──────────────────────────────────────────────────────────

def show_kpi_row(types, annual_totals, capacities, years):
    """Show KPI metric cards for most recent year."""
    last_year = max(years)
    cols = st.columns(len(types))
    for ci, t in enumerate(types):
        cap  = capacities.get(t, 1)
        dem  = annual_totals[int(last_year)].get(t, 0)
        util = dem / cap if cap > 0 else 0
        with cols[ci]:
            cls = "red" if util > 1 else "amber" if util >= 0.8 else "green"
            icon = "🔴" if util > 1 else "🟡" if util >= 0.8 else "🟢"
            st.markdown(f"""
            <div class="metric-card {cls}">
                <div style="font-size:11px;color:#666;font-weight:600;">{t}</div>
                <div style="font-size:22px;font-weight:800;color:#1F3864;">{dem:.0f}</div>
                <div style="font-size:12px;color:#555;">samples in {last_year}</div>
                <div style="font-size:13px;margin-top:4px;">{icon} {util:.0%} of {cap}/yr</div>
            </div>
            """, unsafe_allow_html=True)


def show_util_table(util_df, types, years):
    """Colored utilization summary table."""
    rows = []
    for t in types:
        row = {"Lab Type": t}
        for y in years:
            peak = util_df[util_df["Year"] == y][t].max()
            avg  = util_df[util_df["Year"] == y][t].mean()
            row[f"{y} Peak"] = f"{peak:.1%}"
            row[f"{y} Avg"]  = f"{avg:.1%}"
        rows.append(row)
    df_display = pd.DataFrame(rows)

    def color_util(val):
        if isinstance(val, str) and val.endswith("%"):
            try:
                v = float(val.replace("%","").replace("+","")) / 100
                if v > 1.0: return "background-color:#FFCCCC; color:#CC0000; font-weight:bold"
                if v >= 0.8: return "background-color:#FFF3CC"
                return "background-color:#CCEFCC"
            except: pass
        return ""

    st.dataframe(
        df_display.style.applymap(color_util, subset=[c for c in df_display.columns if c != "Lab Type"]),
        use_container_width=True, hide_index=True,
    )


def make_editable_excel(weekly_df, util_df, annual_totals, types, capacities, years, title="Lab Data"):
    """Build an editable Excel workbook with 3 sheets and return bytes."""
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    wb = openpyxl.Workbook()

    thin = Border(
        left=Side(style="thin", color="CCCCCC"),
        right=Side(style="thin", color="CCCCCC"),
        top=Side(style="thin", color="CCCCCC"),
        bottom=Side(style="thin", color="CCCCCC"),
    )

    def hdr_cell(ws, row, col, val, bg="1F3864", fg="FFFFFF", sz=10, bold=True):
        c = ws.cell(row, col, val)
        c.font = Font(name="Arial", size=sz, bold=bold, color=fg)
        c.fill = PatternFill("solid", start_color=bg)
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = thin
        return c

    def data_cell(ws, row, col, val, fmt=None, bg=None):
        c = ws.cell(row, col, val)
        c.font = Font(name="Arial", size=9)
        c.alignment = Alignment(horizontal="center")
        c.border = thin
        if fmt: c.number_format = fmt
        if bg:  c.fill = PatternFill("solid", start_color=bg)
        return c

    # ── Sheet 1: Weekly Demand (editable) ──────────────────
    ws1 = wb.active
    ws1.title = "Weekly_Demand"
    ws1.cell(1, 1, f"📋 {title} — Editable Weekly Demand").font = Font(
        name="Arial", size=12, bold=True, color="1F3864")
    ws1.cell(2, 1, "✏️  You can edit the values in this sheet. Yellow = editable cells.")
    ws1.cell(2, 1).font = Font(name="Arial", size=9, italic=True, color="7F6000")

    # Header row
    headers = ["Year", "Week"] + list(types)
    for ci, h in enumerate(headers, 1):
        bg = "1F3864" if ci <= 2 else "2E75B6"
        hdr_cell(ws1, 4, ci, h, bg=bg)
        ws1.column_dimensions[get_column_letter(ci)].width = 14 if ci > 2 else 8

    editable_fill = PatternFill("solid", start_color="FFFFE0")  # light yellow
    for ri, row in enumerate(weekly_df.itertuples(index=False), 5):
        vals = list(row)
        data_cell(ws1, ri, 1, vals[0])  # Year
        data_cell(ws1, ri, 2, vals[1])  # Week
        for ci, t in enumerate(types, 3):
            c = data_cell(ws1, ri, ci, round(vals[2 + list(weekly_df.columns[2:]).index(t)], 3),
                          fmt="0.000")
            c.fill = editable_fill  # mark as editable

    ws1.freeze_panes = "A5"
    ws1.auto_filter.ref = f"A4:{get_column_letter(len(headers))}4"

    # ── Sheet 2: Annual Summary ─────────────────────────────
    ws2 = wb.create_sheet("Annual_Summary")
    ws2.cell(1, 1, f"📊 {title} — Annual Summary").font = Font(
        name="Arial", size=12, bold=True, color="1F3864")

    hdrs2 = ["Lab Type", "Capacity/yr", "Weekly Cap"] + [str(y) for y in years] + ["Best Year", "Worst Year"]
    for ci, h in enumerate(hdrs2, 1):
        hdr_cell(ws2, 3, ci, h)
        ws2.column_dimensions[get_column_letter(ci)].width = 14

    for ri, t in enumerate(types, 4):
        cap = capacities.get(t, 1)
        data_cell(ws2, ri, 1, t)
        data_cell(ws2, ri, 2, cap, fmt="#,##0")
        data_cell(ws2, ri, 3, round(cap / 52, 2), fmt="0.00")
        yr_vals = {}
        for ci, y in enumerate(years, 4):
            dem = annual_totals.get(int(y), {}).get(t, 0)
            yr_vals[y] = dem
            util = dem / cap if cap > 0 else 0
            bg = "FF4444" if util > 1 else "FFD700" if util >= 0.8 else "C6EFCE"
            c = data_cell(ws2, ri, ci, round(dem, 1), fmt="#,##0.0", bg=bg)
            if util > 1:
                c.font = Font(name="Arial", size=9, bold=True, color="FFFFFF")
        best_yr = max(yr_vals, key=yr_vals.get, default="")
        worst_yr = min(yr_vals, key=yr_vals.get, default="")
        data_cell(ws2, ri, len(hdrs2), str(worst_yr))
        data_cell(ws2, ri, len(hdrs2) - 1, str(best_yr))

    ws2.freeze_panes = "A4"

    # ── Sheet 3: Utilization Summary ───────────────────────
    ws3 = wb.create_sheet("Utilization_%")
    ws3.cell(1, 1, f"📈 {title} — Utilization % (Green<80% | Yellow 80-100% | Red>100%)").font = Font(
        name="Arial", size=11, bold=True, color="1F3864")

    hdrs3 = ["Lab Type"] + [f"{y} Peak" for y in years] + [f"{y} Avg" for y in years]
    for ci, h in enumerate(hdrs3, 1):
        hdr_cell(ws3, 3, ci, h)
        ws3.column_dimensions[get_column_letter(ci)].width = 13

    for ri, t in enumerate(types, 4):
        data_cell(ws3, ri, 1, t)
        cap = capacities.get(t, 1)
        for ci, y in enumerate(years, 2):
            peak = util_df[util_df["Year"] == y][t].max() if t in util_df.columns else 0
            avg  = util_df[util_df["Year"] == y][t].mean() if t in util_df.columns else 0
            for val, offset in [(peak, 0), (avg, len(years))]:
                col = ci + offset
                bg = "FF4444" if val > 1 else "FFD700" if val >= 0.8 else "C6EFCE"
                c = data_cell(ws3, ri, col, round(val, 4), fmt="0.0%", bg=bg)
                if val > 1:
                    c.font = Font(name="Arial", size=9, bold=True, color="FFFFFF")

    ws3.freeze_panes = "B4"

    # Save to bytes
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()


def capacity_warning(individual_caps, combined_cap):
    """Show warning if individual caps sum > combined cap."""
    total = sum(individual_caps.values())
    if total > combined_cap:
        st.warning(
            f"⚠️ Individual capacities sum to **{total}**, which exceeds "
            f"the combined cap of **{combined_cap}**. "
            "Consider reducing individual values or raising the combined cap."
        )
    elif total < combined_cap:
        st.info(
            f"ℹ️ Individual capacities sum to **{total}** "
            f"(combined cap is **{combined_cap}**). "
            f"**{combined_cap - total}** capacity units are unallocated."
        )
    else:
        st.success(f"✅ Individual capacities sum exactly to combined cap ({combined_cap}).")


# ──────────────────────────────────────────────────────────
#  SESSION STATE — Chat
# ──────────────────────────────────────────────────────────
if "chat_history" not in st.session_state:
    st.session_state.chat_history = []
if "chat_context" not in st.session_state:
    st.session_state.chat_context = {}


# ──────────────────────────────────────────────────────────
#  SIDEBAR
# ──────────────────────────────────────────────────────────

with st.sidebar:
    st.markdown("## 🔬 Lab Dashboard Suite")
    st.markdown("---")
    tool = st.radio(
        "Select Tool",
        ["🔵 Tool 1 — LCF & Creep",
         "🟢 Tool 2 — Coating Labs",
         "🟣 Tool 3 — Comparison"],
        label_visibility="collapsed",
    )
    st.markdown("---")
    st.markdown("**Input Format**")
    st.markdown("Excel file with columns:\n- `Year`\n- `Type` (lab name)\n- `Value` (samples)")
    st.markdown("---")
    st.markdown("**Output**\n\nExcel dashboard with 7–8 sheets including charts, Gantt, and pie charts.")
    st.markdown("---")
    st.caption(f"Current week: **{CURRENT_WEEK}** of {CURRENT_YEAR}")

    # ── AI Chat Assistant in Sidebar ──────────────────────
    st.markdown("---")
    st.markdown("## 🤖 AI Lab Assistant")
    st.caption("Ask questions about your lab data or request changes to capacity settings.")

    # Display chat history with bubble styling
    bubbles_html = '<div class="chat-wrap">'
    for msg in st.session_state.chat_history[-10:]:
        if msg["role"] == "user":
            bubbles_html += f'<div class="chat-user">🧑 {msg["content"]}</div>'
        else:
            # Render line breaks from assistant
            txt = msg["content"].replace("\n", "<br>")
            bubbles_html += f'<div class="chat-bot">🤖 {txt}</div>'
    bubbles_html += "</div>"
    st.markdown(bubbles_html, unsafe_allow_html=True)

    # Chat input
    user_question = st.text_area("Your question:", key="chat_input", height=80,
                                  placeholder="e.g. What is Cold Spray utilization in 2024?\nSuggest capacity for 80% utilization...")
    
    col_send, col_clear = st.columns(2)
    with col_send:
        send_clicked = st.button("💬 Ask", key="chat_send", use_container_width=True)
    with col_clear:
        if st.button("🗑️ Clear", key="chat_clear", use_container_width=True):
            st.session_state.chat_history = []
            st.rerun()

    if send_clicked and user_question.strip():
        # Build context from current session state
        ctx = st.session_state.get("chat_context", {})
        context_str = ""
        if ctx:
            context_str = f"""
Current dashboard context:
- Tool: {ctx.get('tool', 'unknown')}
- Lab types: {ctx.get('types', [])}
- Years in data: {ctx.get('years', [])}
- Capacities (samples/yr): {ctx.get('capacities', {})}
- Latest year demand: {ctx.get('latest_demand', {})}
- Peak utilization per lab: {ctx.get('peak_util', {})}
"""
        system_prompt = """You are an expert lab capacity planning assistant for a coating and mechanical testing lab.
You help users understand utilization charts, interpret lab demand trends, and suggest capacity adjustments.

Labs tracked:
- Cold Spray, HVOF, Plasma (Coating Labs) — each has independent capacity
- LCF, Creep (Mechanical Labs)

Key concepts:
- Utilization = Weekly demand / Weekly capacity. Green <80%, Yellow 80-100%, Red >100%
- Weekly capacity = Annual capacity / 52
- Seasonal demand patterns are applied (summer peaks, winter dips)

When suggesting capacity changes, always give concrete numbers.
Keep answers concise and actionable. Use bullet points for clarity."""

        messages_to_send = []
        for h in st.session_state.chat_history[-8:]:
            messages_to_send.append({"role": h["role"], "content": h["content"]})
        
        user_msg_full = user_question.strip()
        if context_str:
            user_msg_full = f"{context_str}\n\nUser question: {user_question.strip()}"
        
        messages_to_send.append({"role": "user", "content": user_msg_full})
        st.session_state.chat_history.append({"role": "user", "content": user_question.strip()})

        with st.spinner("Thinking…"):
            try:
                import anthropic, os
                api_key = (st.secrets.get("ANTHROPIC_API_KEY", None)
                           if hasattr(st, "secrets") else None) or os.environ.get("ANTHROPIC_API_KEY")
                if not api_key:
                    st.error("❌ ANTHROPIC_API_KEY not set. See HOW_TO_RUN_LOCALLY.md Step 6.")
                else:
                    client = anthropic.Anthropic(api_key=api_key)
                    response = client.messages.create(
                        model="claude-sonnet-4-5-20251001",
                        max_tokens=600,
                        system=system_prompt,
                        messages=messages_to_send,
                    )
                    answer = response.content[0].text
                    st.session_state.chat_history.append({"role": "assistant", "content": answer})
                    st.rerun()
            except ImportError:
                st.error("❌ Install anthropic: `pip install -r requirements.txt`")
            except Exception as e:
                st.error(f"❌ Chat error: {str(e)[:150]}")


# ══════════════════════════════════════════════════════════════════
#  TOOL 1 — LCF & CREEP
# ══════════════════════════════════════════════════════════════════

if "Tool 1" in tool:
    st.markdown("""
    <div class="tool-header">
        <h2>🔵 LCF & Creep Lab Dashboard</h2>
        <p>Mechanical lab occupancy planning — individual capacity tracking</p>
    </div>
    """, unsafe_allow_html=True)

    # ── Config ────────────────────────────────────────────
    col_a, col_b, col_c = st.columns([2, 1, 1])
    with col_a:
        st.markdown('<div class="section-label">📂 Upload Excel File</div>', unsafe_allow_html=True)
        uploaded = st.file_uploader("", type=["xlsx","xls"], key="t1_file",
                                    label_visibility="collapsed")
    with col_b:
        st.markdown('<div class="section-label">LCF Capacity</div>', unsafe_allow_html=True)
        lcf_cap = st.number_input("LCF (samples/yr)", value=50, min_value=1, max_value=9999,
                                   key="t1_lcf", label_visibility="collapsed")
        st.caption("LCF — samples/year")
    with col_c:
        st.markdown('<div class="section-label">Creep Capacity</div>', unsafe_allow_html=True)
        creep_cap = st.number_input("Creep (samples/yr)", value=22, min_value=1, max_value=9999,
                                     key="t1_creep", label_visibility="collapsed")
        st.caption("Creep — samples/year")

    theme_name = st.selectbox("Color Theme", list(core.COLOR_THEMES.keys()), key="t1_theme")
    caps_1 = {"LCF": lcf_cap, "Creep": creep_cap}
    theme_1 = core.COLOR_THEMES[theme_name]

    if uploaded:
        paths_1 = load_uploaded_files([uploaded])
        df_1, col_map_1, err_1, warn_1 = core.load_and_filter(paths_1[0], ["LCF","Creep"])

        if err_1:
            st.error("❌ " + "\n".join(err_1))
        else:
            if warn_1:
                for w in warn_1: st.warning(w)

            wdf_1, udf_1, types_1, years_1 = core.build_weekly(df_1, col_map_1, caps_1)
            annual_1 = get_annual_totals(wdf_1, types_1, years_1)

            # Expose data to AI chat assistant
            last_y1_ctx = int(max(years_1))
            st.session_state.chat_context = {
                "tool": "LCF & Creep (Tool 1)",
                "types": list(types_1),
                "years": [int(y) for y in years_1],
                "capacities": caps_1,
                "latest_demand": annual_1.get(last_y1_ctx, {}),
                "peak_util": {t: f"{udf_1[udf_1['Year']==last_y1_ctx][t].max():.1%}"
                              for t in types_1},
            }

            # KPI cards
            st.markdown('<div class="section-label">📊 Key Metrics — Latest Year</div>', unsafe_allow_html=True)
            show_kpi_row(types_1, annual_1, caps_1, years_1)

            # Charts in tabs
            tabs = st.tabs(["📈 Utilization Trend", "📊 YoY + Pie Charts",
                             "⚡ Capacity vs Demand", "🗓 Gantt", "📋 Data Table"])

            with tabs[0]:
                fig_util = chart_utilization_line(udf_1, types_1, years_1,
                                                   THEME_COLORS[theme_name])
                st.plotly_chart(fig_util, use_container_width=True)

            with tabs[1]:
                fig_bar, fig_pies = chart_yoy_bar_and_pie(annual_1, types_1, years_1,
                                                           THEME_COLORS[theme_name])
                st.plotly_chart(fig_bar, use_container_width=True)
                st.plotly_chart(fig_pies, use_container_width=True)

            with tabs[2]:
                fig_cap = chart_capacity_bar(wdf_1, types_1, caps_1, years_1)
                st.plotly_chart(fig_cap, use_container_width=True)

            with tabs[3]:
                gantt_year = CURRENT_YEAR if CURRENT_YEAR in [int(y) for y in years_1] else int(max(years_1))
                gantt_week = CURRENT_WEEK if gantt_year == CURRENT_YEAR else 52
                fig_gantt = chart_gantt(wdf_1, types_1, caps_1, gantt_year, gantt_week)
                st.plotly_chart(fig_gantt, use_container_width=True)

            with tabs[4]:
                st.markdown("**Utilization Summary (Peak & Avg per year)**")
                show_util_table(udf_1, types_1, years_1)
                st.markdown("---")
                st.markdown("**✏️ Editable Data Export** — download a pre-filled Excel you can edit and re-upload")
                editable_bytes_1 = make_editable_excel(
                    wdf_1, udf_1, annual_1, types_1, caps_1, years_1,
                    title="LCF & Creep Lab"
                )
                st.download_button(
                    "⬇️ Download Editable Excel (Weekly Data + Summary)",
                    data=editable_bytes_1,
                    file_name="LCF_Creep_Editable.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="t1_editable_dl",
                )

            # Generate Excel
            st.markdown("---")
            st.markdown('<div class="section-label">💾 Generate Full Excel Dashboard</div>', unsafe_allow_html=True)
            if st.button("⚡ Generate Full Excel Dashboard", key="t1_gen"):
                with st.spinner("Generating Excel dashboard..."):
                    out1, warns = gen_module.generate_lcf_creep(paths_1[0], caps_1, theme_1)
                    with open(out1, 'rb') as f: excel_bytes = f.read()
                st.download_button(
                    "⬇️ Download LCF_Creep_Dashboard.xlsx",
                    data=excel_bytes, file_name="LCF_Creep_Dashboard.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="t1_dl",
                )
                st.success("✅ Dashboard ready! Click the button above to download.")
    else:
        st.info("👆 Upload an Excel file to begin. The file should contain LCF and/or Creep data.")


# ══════════════════════════════════════════════════════════════════
#  TOOL 2 — COATING LABS
# ══════════════════════════════════════════════════════════════════

elif "Tool 2" in tool:
    st.markdown("""
    <div class="tool-header" style="background: linear-gradient(90deg, #1F5C1A 0%, #70AD47 100%);">
        <h2>🟢 Coating Labs Dashboard</h2>
        <p>Cold Spray · HVOF · Plasma — individual + combined capacity tracking</p>
    </div>
    """, unsafe_allow_html=True)

    # ── File upload ───────────────────────────────────────
    st.markdown('<div class="section-label">📂 Upload Excel File</div>', unsafe_allow_html=True)
    uploaded_2 = st.file_uploader("", type=["xlsx","xls"], key="t2_file",
                                   label_visibility="collapsed")

    # ── Capacity config ───────────────────────────────────
    st.markdown('<div class="section-label">⚙️ Capacity Configuration — Each Lab is Independent</div>', unsafe_allow_html=True)
    st.info("💡 Each lab has its **own independent capacity** of 350 samples/year by default. Edit any lab below.")

    c1, c2, c3 = st.columns(3)
    with c1:
        cs_cap = st.number_input("🔵 Cold Spray (samples/yr)",
                                  value=350, min_value=1, max_value=9999, key="t2_cs")
    with c2:
        hvof_cap = st.number_input("🟠 HVOF (samples/yr)",
                                    value=350, min_value=1, max_value=9999, key="t2_hvof")
    with c3:
        plasma_cap = st.number_input("🟣 Plasma (samples/yr)",
                                      value=350, min_value=1, max_value=9999, key="t2_pl")
    ind_caps_2 = {"Cold Spray": cs_cap, "HVOF": hvof_cap, "Plasma": plasma_cap}
    st.caption(f"Weekly caps → Cold Spray: **{cs_cap/52:.1f}**/wk  |  HVOF: **{hvof_cap/52:.1f}**/wk  |  Plasma: **{plasma_cap/52:.1f}**/wk")

    theme_name_2 = st.selectbox("Color Theme", list(core.COLOR_THEMES.keys()), key="t2_theme")
    theme_2 = core.COLOR_THEMES[theme_name_2]

    if uploaded_2:
        COATING_TYPES = ["Cold Spray", "HVOF", "Plasma"]
        paths_2 = load_uploaded_files([uploaded_2])
        df_2, col_map_2, err_2, warn_2 = core.load_and_filter(paths_2[0], COATING_TYPES)

        if err_2:
            st.error("❌ " + "\n".join(err_2))
        else:
            if warn_2:
                for w in warn_2: st.warning(w)

            wdf_2, udf_2, types_2, years_2 = core.build_weekly(df_2, col_map_2, ind_caps_2)
            annual_2 = get_annual_totals(wdf_2, types_2, years_2)

            # Expose data to AI chat assistant
            last_y2_ctx = int(max(years_2))
            st.session_state.chat_context = {
                "tool": "Coating Labs (Tool 2)",
                "types": list(types_2),
                "years": [int(y) for y in years_2],
                "capacities": ind_caps_2,
                "latest_demand": annual_2.get(last_y2_ctx, {}),
                "peak_util": {t: f"{udf_2[udf_2['Year']==last_y2_ctx][t].max():.1%}"
                              for t in types_2},
            }

            # Combined utilization KPIs
            st.markdown('<div class="section-label">📊 Lab Utilization — Latest Year</div>', unsafe_allow_html=True)
            last_y2 = int(max(years_2))

            col_kpi1, col_kpi2, col_kpi3 = st.columns(3)
            with col_kpi1:
                combined_demand = sum(annual_2[last_y2].get(t, 0) for t in types_2)
                total_cap = sum(ind_caps_2.values())
                comb_util = combined_demand / total_cap if total_cap > 0 else 0
                cls = "red" if comb_util > 1 else "amber" if comb_util >= 0.8 else "green"
                st.markdown(f"""<div class="metric-card {cls}">
                    <div style="font-size:11px;color:#666;font-weight:600;">Total Demand ({last_y2})</div>
                    <div style="font-size:28px;font-weight:800;color:#1F5C1A;">{combined_demand:.0f}</div>
                    <div style="font-size:13px;">across all 3 labs (sum cap: {total_cap})</div>
                </div>""", unsafe_allow_html=True)
            with col_kpi2:
                st.markdown(f"""<div class="metric-card">
                    <div style="font-size:11px;color:#666;font-weight:600;">Weekly Demand Avg</div>
                    <div style="font-size:28px;font-weight:800;color:#1F5C1A;">{combined_demand/52:.1f}</div>
                    <div style="font-size:13px;">samples/week total</div>
                </div>""", unsafe_allow_html=True)
            with col_kpi3:
                peak_wk = sum(wdf_2[wdf_2["Year"]==last_y2][t].max() for t in types_2)
                total_wk_cap = sum(ind_caps_2.values()) / 52
                peak_util = peak_wk / total_wk_cap if total_wk_cap > 0 else 0
                cls3 = "red" if peak_util > 1 else "amber" if peak_util >= 0.8 else "green"
                st.markdown(f"""<div class="metric-card {cls3}">
                    <div style="font-size:11px;color:#666;font-weight:600;">Peak Week Demand ({last_y2})</div>
                    <div style="font-size:28px;font-weight:800;color:#1F5C1A;">{peak_wk:.1f}</div>
                    <div style="font-size:13px;">samples ({peak_util:.0%} of total weekly cap)</div>
                </div>""", unsafe_allow_html=True)

            show_kpi_row(types_2, annual_2, ind_caps_2, years_2)

            tabs2 = st.tabs(["📈 Utilization", "📊 YoY + Pie Charts",
                              "⚡ Capacity vs Demand", "🗓 Gantt", "📋 Data Table"])

            with tabs2[0]:
                fig_u2 = chart_utilization_line(udf_2, types_2, years_2,
                                                 THEME_COLORS[theme_name_2])
                st.plotly_chart(fig_u2, use_container_width=True)

            with tabs2[1]:
                fig_bar2, fig_pies2 = chart_yoy_bar_and_pie(annual_2, types_2, years_2,
                                                              THEME_COLORS[theme_name_2])
                st.plotly_chart(fig_bar2, use_container_width=True)
                st.plotly_chart(fig_pies2, use_container_width=True)

            with tabs2[2]:
                fig_cap2 = chart_capacity_bar(wdf_2, types_2, ind_caps_2, years_2)
                st.plotly_chart(fig_cap2, use_container_width=True)

            with tabs2[3]:
                gantt_year2 = CURRENT_YEAR if CURRENT_YEAR in [int(y) for y in years_2] else int(max(years_2))
                gantt_week2 = CURRENT_WEEK if gantt_year2 == CURRENT_YEAR else 52
                fig_g2 = chart_gantt(wdf_2, types_2, ind_caps_2, gantt_year2, gantt_week2)
                st.plotly_chart(fig_g2, use_container_width=True)

            with tabs2[4]:
                show_util_table(udf_2, types_2, years_2)
                st.markdown("---")
                st.markdown("**✏️ Editable Data Export** — download a pre-filled Excel you can edit and re-upload")
                editable_bytes_2 = make_editable_excel(
                    wdf_2, udf_2, annual_2, types_2, ind_caps_2, years_2,
                    title="Coating Labs"
                )
                st.download_button(
                    "⬇️ Download Editable Excel (Weekly Data + Summary)",
                    data=editable_bytes_2,
                    file_name="Coating_Labs_Editable.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="t2_editable_dl",
                )

            st.markdown("---")
            st.markdown('<div class="section-label">💾 Generate Full Excel Dashboard</div>', unsafe_allow_html=True)
            if st.button("⚡ Generate Full Excel Dashboard", key="t2_gen"):
                with st.spinner("Generating..."):
                    out2, warns2 = gen_module.generate_coating(paths_2[0], ind_caps_2, theme_2)
                    with open(out2, 'rb') as f: excel_bytes2 = f.read()
                st.download_button(
                    "⬇️ Download Coating_Labs_Dashboard.xlsx",
                    data=excel_bytes2, file_name="Coating_Labs_Dashboard.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="t2_dl",
                )
                st.success("✅ Done!")
    else:
        st.info("👆 Upload an Excel file with Cold Spray, HVOF, and/or Plasma data.")


# ══════════════════════════════════════════════════════════════════
#  TOOL 3 — COMPARISON
# ══════════════════════════════════════════════════════════════════

elif "Tool 3" in tool:
    st.markdown("""
    <div class="tool-header" style="background: linear-gradient(90deg, #3D1F6E 0%, #7B5EA7 100%);">
        <h2>🟣 Lab Comparison Dashboard</h2>
        <p>Mechanical Labs (LCF + Creep) vs Coating Labs (Cold Spray + HVOF + Plasma)</p>
    </div>
    """, unsafe_allow_html=True)

    # ── Multi-file upload ─────────────────────────────────
    st.markdown('<div class="section-label">📂 Upload Files (one or more — data merged automatically)</div>',
                unsafe_allow_html=True)
    uploaded_3 = st.file_uploader(
        "",
        type=["xlsx","xls"],
        accept_multiple_files=True,
        key="t3_files",
        label_visibility="collapsed",
        help="Upload one combined file OR separate files per lab group. Duplicates are merged.",
    )
    if uploaded_3:
        st.success(f"✅ {len(uploaded_3)} file(s) loaded: {', '.join(f.name for f in uploaded_3)}")

    # ── Capacities ─────────────────────────────────────────
    cc1, cc2 = st.columns(2)
    with cc1:
        st.markdown('<div class="section-label">🔵 Mechanical Lab Capacities</div>', unsafe_allow_html=True)
        mc1, mc2 = st.columns(2)
        with mc1:
            lcf_cap3 = st.number_input("LCF (samples/yr)", value=50, min_value=1, key="t3_lcf")
        with mc2:
            creep_cap3 = st.number_input("Creep (samples/yr)", value=22, min_value=1, key="t3_creep")
        caps_a3 = {"LCF": lcf_cap3, "Creep": creep_cap3}
        st.caption(f"Mechanical combined: **{sum(caps_a3.values())}** samples/yr")

    with cc2:
        st.markdown('<div class="section-label">🟢 Coating Lab Capacities (each independent)</div>', unsafe_allow_html=True)
        bc1, bc2, bc3 = st.columns(3)
        with bc1: cs3 = st.number_input("Cold Spray", value=350, min_value=1, key="t3_cs")
        with bc2: hv3 = st.number_input("HVOF",       value=350, min_value=1, key="t3_hv")
        with bc3: pl3 = st.number_input("Plasma",     value=350, min_value=1, key="t3_pl")
        caps_b3 = {"Cold Spray": cs3, "HVOF": hv3, "Plasma": pl3}
        comb_cap3 = sum(caps_b3.values())  # for reference only
        st.caption(f"Weekly caps → CS: {cs3/52:.1f}/wk | HVOF: {hv3/52:.1f}/wk | Plasma: {pl3/52:.1f}/wk")

    theme_name_3 = st.selectbox("Color Theme", list(core.COLOR_THEMES.keys()), key="t3_theme")
    theme_3 = core.COLOR_THEMES[theme_name_3]

    if uploaded_3:
        paths_3 = load_uploaded_files(uploaded_3)
        merged_df, col_map_3, errors_3, file_log_3 = merge_files(paths_3)

        for log_line in file_log_3:
            if log_line.startswith("✓") or log_line.startswith("OK"):
                st.success(log_line)
            elif "WARNING" in log_line:
                st.warning(log_line)

        if errors_3:
            st.error("❌ " + "\n".join(errors_3))
        else:
            tc3 = col_map_3["type"]
            df_a3 = merged_df[merged_df[tc3].isin(GROUP_A["types"])].copy()
            df_b3 = merged_df[merged_df[tc3].isin(GROUP_B["types"])].copy()

            if df_a3.empty:
                st.error(f"No Mechanical lab data found. Need: {GROUP_A['types']}")
            elif df_b3.empty:
                st.error(f"No Coating lab data found. Need: {GROUP_B['types']}")
            else:
                wdf_a3, udf_a3, types_a3, years_a3 = core.build_weekly(df_a3, col_map_3, caps_a3)
                wdf_b3, udf_b3, types_b3, years_b3 = core.build_weekly(df_b3, col_map_3, caps_b3)
                years_3 = sorted(set(list(years_a3)) | set(list(years_b3)))
                annual_a3 = get_annual_totals(wdf_a3, types_a3, years_a3)
                annual_b3 = get_annual_totals(wdf_b3, types_b3, years_b3)

                # Summary metrics
                st.markdown('<div class="section-label">📊 Group Utilization Summary</div>',
                            unsafe_allow_html=True)
                last_y3 = int(max(years_3))
                sum_caps_a = sum(caps_a3.values())
                sum_caps_b = comb_cap3

                dem_a3 = sum(annual_a3.get(last_y3, {}).get(t, 0) for t in types_a3)
                dem_b3 = sum(annual_b3.get(last_y3, {}).get(t, 0) for t in types_b3)
                u_a3   = dem_a3 / sum_caps_a if sum_caps_a > 0 else 0
                u_b3   = dem_b3 / sum_caps_b if sum_caps_b > 0 else 0

                col_s1, col_s2 = st.columns(2)
                with col_s1:
                    cls_a = "red" if u_a3 > 1 else "amber" if u_a3 >= 0.8 else "green"
                    st.markdown(f"""<div class="metric-card {cls_a}">
                        <div style="font-size:12px;color:#666;font-weight:700;">🔵 Mechanical Labs ({last_y3})</div>
                        <div style="font-size:26px;font-weight:800;color:#1A4E8A;">{dem_a3:.0f} samples</div>
                        <div style="font-size:13px;">Utilization: {u_a3:.0%} of {sum_caps_a}/yr capacity</div>
                    </div>""", unsafe_allow_html=True)
                with col_s2:
                    cls_b = "red" if u_b3 > 1 else "amber" if u_b3 >= 0.8 else "green"
                    st.markdown(f"""<div class="metric-card {cls_b}">
                        <div style="font-size:12px;color:#666;font-weight:700;">🟢 Coating Labs ({last_y3})</div>
                        <div style="font-size:26px;font-weight:800;color:#1F5C1A;">{dem_b3:.0f} samples</div>
                        <div style="font-size:13px;">Utilization: {u_b3:.0%} of {sum_caps_b}/yr capacity</div>
                    </div>""", unsafe_allow_html=True)

                tabs3 = st.tabs(["📊 Comparison Charts", "🥧 YoY Pie Charts",
                                  "📈 Utilization Trend", "🗓 Gantt", "📋 Data Tables"])

                with tabs3[0]:
                    fig_cmp = chart_comparison_grouped(
                        annual_a3, annual_b3, types_a3, types_b3,
                        years_3, sum_caps_a, sum_caps_b)
                    st.plotly_chart(fig_cmp, use_container_width=True)

                with tabs3[1]:
                    st.subheader("🔵 Mechanical Labs — Process Share")
                    _, fig_pie_a = chart_yoy_bar_and_pie(annual_a3, types_a3, years_a3,
                                                          THEME_COLORS[theme_name_3])
                    st.plotly_chart(fig_pie_a, use_container_width=True)

                    st.subheader("🟢 Coating Labs — Process Share")
                    _, fig_pie_b = chart_yoy_bar_and_pie(annual_b3, types_b3, years_b3,
                                                          THEME_COLORS[theme_name_3])
                    st.plotly_chart(fig_pie_b, use_container_width=True)

                with tabs3[2]:
                    fig_util3 = chart_util_comparison_line(
                        wdf_a3, wdf_b3, types_a3, types_b3,
                        years_3, sum_caps_a, sum_caps_b)
                    st.plotly_chart(fig_util3, use_container_width=True)

                with tabs3[3]:
                    g_yr3 = CURRENT_YEAR if CURRENT_YEAR in [int(y) for y in years_3] else int(max(years_3))
                    g_wk3 = CURRENT_WEEK if g_yr3 == CURRENT_YEAR else 52
                    col_g1, col_g2 = st.columns(2)
                    with col_g1:
                        st.markdown("**🔵 Mechanical**")
                        fig_ga = chart_gantt(wdf_a3, types_a3, caps_a3, g_yr3, g_wk3)
                        st.plotly_chart(fig_ga, use_container_width=True)
                    with col_g2:
                        st.markdown("**🟢 Coating**")
                        fig_gb = chart_gantt(wdf_b3, types_b3, caps_b3, g_yr3, g_wk3)
                        st.plotly_chart(fig_gb, use_container_width=True)

                with tabs3[4]:
                    st.markdown("**🔵 Mechanical Labs**")
                    show_util_table(udf_a3, types_a3, years_a3)
                    st.markdown("**🟢 Coating Labs**")
                    show_util_table(udf_b3, types_b3, years_b3)

                # Excel export
                st.markdown("---")
                st.markdown('<div class="section-label">💾 Generate Excel Comparison Dashboard</div>',
                            unsafe_allow_html=True)
                if st.button("⚡ Generate Full Excel Dashboard", key="t3_gen"):
                    with st.spinner("Generating..."):
                        out3, warns3 = gen_module.generate_comparison(paths_3, caps_a3, caps_b3, theme_3)
                        with open(out3, 'rb') as f: excel_bytes3 = f.read()
                    st.download_button(
                        "⬇️ Download Lab_Comparison_Dashboard.xlsx",
                        data=excel_bytes3,
                        file_name="Lab_Comparison_Dashboard.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="t3_dl",
                    )
                    st.success("✅ Done!")
    else:
        st.info("👆 Upload one or more Excel files. Works with a single combined file or separate files per lab group.")
