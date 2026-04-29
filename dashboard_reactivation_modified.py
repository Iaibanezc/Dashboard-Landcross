import os
import io
import base64
from pathlib import Path

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
import streamlit as st

# ─────────────────────────────────────────────
# PAGE CONFIG
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="Landcros — Fleet Reactivation Dashboard",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─────────────────────────────────────────────
# CONSTANTS
# ─────────────────────────────────────────────
EXCEL_FILE = "Data base Reactivation.xlsx"
TOP_N_TRUCKS = 19

PRIMARY = "#FF6B00"
PRIMARY_LIGHT = "#FF9340"
DARK = "#1A1A1A"
MID = "#888888"
LIGHT = "#F7F7F7"
GRID = "#EFEFEF"
GREEN = "#2E7D32"
YELLOW = "#F9A825"
RED = "#C62828"

KIT_COLS = [
    "Kits 1 A/C System",
    "Kit 2 Lube system",
    "Kit 3 Hydraulic filters & seal kits",
    "Kit 4 Operator Cab Monuting",
    "Kit 5 Drive system ",
    "Kit 6 Drive system elfa fan blower",
    "Kit 7 Drive system braking resistor",
    "Kit 8 Body Mounting & Pads",
    "Kit 9 Accum steer & Brake",
    "Kit 10 Brake Valve & Cooler mounting",
    "Kit 11 Front axle & MTG",
    "Kit 12 Frame Bottom plate optimized",
    "Kit 13 Frame upper plate optimized",
    "Kit 14 Frame left side optimized",
    "Kit 15 Frame right  side optimized ",
    "Kit 16 Hoist plate reinforment optmized",
    "Kit 17 Filtration drive system optmized",
    "Kit 18 Engine & Hardware",
    "Kit 19 Mirror bracket support ",
    "Kit 20 Fuel tank Bracket support",
]

KIT_LABELS = {
    "Kits 1 A/C System": "Kit 1 — A/C System",
    "Kit 2 Lube system": "Kit 2 — Lube System",
    "Kit 3 Hydraulic filters & seal kits": "Kit 3 — Hydraulic Filters",
    "Kit 4 Operator Cab Monuting": "Kit 4 — Operator Cab Mounting",
    "Kit 5 Drive system ": "Kit 5 — Drive System",
    "Kit 6 Drive system elfa fan blower": "Kit 6 — ELFA Fan Blower",
    "Kit 7 Drive system braking resistor": "Kit 7 — Braking Resistor",
    "Kit 8 Body Mounting & Pads": "Kit 8 — Body Mounting & Pads",
    "Kit 9 Accum steer & Brake": "Kit 9 — Accum Steer & Brake",
    "Kit 10 Brake Valve & Cooler mounting": "Kit 10 — Brake Valve & Cooler",
    "Kit 11 Front axle & MTG": "Kit 11 — Front Axle & MTG",
    "Kit 12 Frame Bottom plate optimized": "Kit 12 — Frame Bottom Plate",
    "Kit 13 Frame upper plate optimized": "Kit 13 — Frame Upper Plate",
    "Kit 14 Frame left side optimized": "Kit 14 — Frame Left Side",
    "Kit 15 Frame right  side optimized ": "Kit 15 — Frame Right Side",
    "Kit 16 Hoist plate reinforment optmized": "Kit 16 — Hoist Plate Reinforcement",
    "Kit 17 Filtration drive system optmized": "Kit 17 — Drive Filtration",
    "Kit 18 Engine & Hardware": "Kit 18 — Engine & Hardware",
    "Kit 19 Mirror bracket support ": "Kit 19 — Mirror Bracket Support",
    "Kit 20 Fuel tank Bracket support": "Kit 20 — Fuel Tank Bracket",
}

SEVERITY_COLS = [
    "High Arch Severity",
    "Nose Cone Severity",
    "Inside Web Plates Severity",
    "Hoist Plates Severity",
    "Top & Bottom flange Severity",
]

# ─────────────────────────────────────────────
# CSS
# ─────────────────────────────────────────────
st.markdown(
    f"""
<style>
  @import url('https://fonts.googleapis.com/css2?family=Barlow+Condensed:wght@300;400;600;700;800&family=Barlow:wght@300;400;500;600&display=swap');

  html, body, [class*="css"] {{
    font-family: 'Barlow', sans-serif;
    background-color: #FFFFFF !important;
    color: {DARK};
  }}

  [data-testid="stSidebar"] {{
    background: {DARK} !important;
    border-right: 3px solid {PRIMARY};
  }}
  [data-testid="stSidebar"] * {{ color: #FFFFFF !important; }}
  [data-testid="stSidebar"] label {{
    font-family: 'Barlow Condensed', sans-serif;
    font-weight: 700;
    font-size: 0.82rem;
    letter-spacing: 0.08em;
    text-transform: uppercase;
    color: #CCCCCC !important;
  }}

  .lc-header {{
    display: flex;
    align-items: center;
    justify-content: space-between;
    background: {DARK};
    padding: 18px 32px;
    border-bottom: 4px solid {PRIMARY};
    margin-bottom: 26px;
    border-radius: 0 0 4px 4px;
  }}
  .lc-header-title {{
    font-family: 'Barlow Condensed', sans-serif;
    font-weight: 800;
    font-size: 2.0rem;
    letter-spacing: 0.04em;
    color: #FFFFFF;
    line-height: 1.1;
  }}
  .lc-header-subtitle {{
    font-family: 'Barlow', sans-serif;
    font-weight: 300;
    font-size: 0.85rem;
    color: #AAAAAA;
    margin-top: 2px;
    letter-spacing: 0.06em;
    text-transform: uppercase;
  }}
  .brand-pill {{
    font-family: 'Barlow Condensed', sans-serif;
    color: #FFFFFF;
    background: {PRIMARY};
    padding: 7px 14px;
    border-radius: 99px;
    font-weight: 800;
    letter-spacing: 0.08em;
  }}
  .kpi-grid {{
    display: grid;
    grid-template-columns: repeat(4, 1fr);
    gap: 16px;
    margin-bottom: 24px;
  }}
  .kpi-card {{
    background: {LIGHT};
    border-left: 4px solid {PRIMARY};
    padding: 18px 20px 14px 20px;
    border-radius: 3px;
  }}
  .kpi-label {{
    font-family: 'Barlow Condensed', sans-serif;
    font-weight: 700;
    font-size: 0.72rem;
    letter-spacing: 0.12em;
    text-transform: uppercase;
    color: {MID};
    margin-bottom: 6px;
  }}
  .kpi-value {{
    font-family: 'Barlow Condensed', sans-serif;
    font-weight: 800;
    font-size: 2.05rem;
    color: {DARK};
    line-height: 1;
  }}
  .kpi-sub {{ font-size: 0.75rem; color: {MID}; margin-top: 5px; }}
  .section-title {{
    font-family: 'Barlow Condensed', sans-serif;
    font-weight: 800;
    font-size: 1.18rem;
    color: {DARK};
    text-transform: uppercase;
    letter-spacing: 0.06em;
    border-left: 4px solid {PRIMARY};
    padding-left: 10px;
    margin: 8px 0 12px 0;
  }}
  .small-note {{
    color: {MID};
    font-size: 0.82rem;
    margin-top: -8px;
    margin-bottom: 12px;
  }}
</style>
""",
    unsafe_allow_html=True,
)

# ─────────────────────────────────────────────
# DATA LOADING
# ─────────────────────────────────────────────
@st.cache_data(show_spinner=False)
def load_workbook(uploaded_file=None):
    source = uploaded_file if uploaded_file is not None else Path(EXCEL_FILE)
    if uploaded_file is None and not Path(EXCEL_FILE).exists():
        alt = Path(__file__).with_name(EXCEL_FILE)
        source = alt if alt.exists() else source

    structural = pd.read_excel(source, sheet_name="Structural", engine="openpyxl")
    labour = pd.read_excel(source, sheet_name="Labour", engine="openpyxl")
    comp = pd.read_excel(source, sheet_name="Component $", engine="openpyxl")
    kits = pd.read_excel(source, sheet_name="Kits", engine="openpyxl")
    return structural, labour, comp, kits


def normalize_component_name(name: str) -> str:
    return str(name).strip().replace("  ", " ")


def safe_col(df: pd.DataFrame, preferred: str, fallback=None):
    if preferred in df.columns:
        return preferred
    clean_map = {str(c).strip(): c for c in df.columns}
    return clean_map.get(str(preferred).strip(), fallback)


def safe_numeric_series(s: pd.Series) -> pd.Series:
    """Convert a Series to numeric only when conversion is meaningful.

    Pandas removed support for errors="ignore" in pd.to_numeric.
    This helper preserves categorical/text columns while safely converting
    numeric-looking columns, including percentage strings such as "85%".
    """
    if pd.api.types.is_numeric_dtype(s):
        return s

    text = s.astype(str).str.strip()
    non_empty = text.replace({"": np.nan, "nan": np.nan, "None": np.nan}).notna()

    pct_mask = text.str.endswith("%", na=False)
    cleaned = text.str.replace("%", "", regex=False).str.replace(",", "", regex=False)
    converted = pd.to_numeric(cleaned, errors="coerce")

    denom = max(int(non_empty.sum()), 1)
    numeric_ratio = converted.notna().sum() / denom
    if numeric_ratio >= 0.60 or pct_mask.any():
        return converted
    return s


def build_component_cost_table(comp: pd.DataFrame) -> pd.DataFrame:
    # Component $ sheet is transposed: first column has metric names, other columns are components.
    first_col = comp.columns[0]
    t = comp.set_index(first_col).T.reset_index().rename(columns={"index": "Component"})
    t["Component"] = t["Component"].map(normalize_component_name)
    for c in ["Labour hours", "Labour cost", "Mechanized & Rebuild", "parts", "Chrome tube & rod", "Total"]:
        if c in t.columns:
            t[c] = pd.to_numeric(t[c], errors="coerce").fillna(0)
    t["Labour cost amount"] = t.get("Labour hours", 0) * t.get("Labour cost", 0)
    t["Component total cost"] = (
        t["Labour cost amount"]
        + t.get("Mechanized & Rebuild", 0)
        + t.get("parts", 0)
        + t.get("Chrome tube & rod", 0)
    )
    return t


def build_kit_cost_table(labour: pd.DataFrame) -> pd.DataFrame:
    first_col = labour.columns[0]
    t = labour.set_index(first_col).T.reset_index().rename(columns={"index": "Kit column"})
    t["Kit"] = t["Kit column"].map(lambda x: KIT_LABELS.get(x, str(x).strip()))
    t["Labour hours"] = pd.to_numeric(t.get("Labour", 0), errors="coerce").fillna(0)
    t["Kit cost"] = pd.to_numeric(t.get("Cost", 0), errors="coerce").fillna(0)
    return t[["Kit column", "Kit", "Labour hours", "Kit cost"]]


def prepare_data(structural: pd.DataFrame, comp_costs: pd.DataFrame):
    df = structural.copy()

    # Robust numeric conversion. Do not use errors="ignore" because recent
    # pandas versions raise ValueError: invalid error value specified.
    # Also avoid forcing text/categorical columns to NaN.
    for c in df.columns:
        if c != "DT":
            df[c] = safe_numeric_series(df[c])

    if "Costo Total Camion" not in df.columns:
        if {"Total Labour", "Total cost per truck"}.issubset(df.columns):
            df["Total dashboard cost"] = pd.to_numeric(df["Total cost per truck"], errors="coerce").fillna(0) + pd.to_numeric(df["Total Labour"], errors="coerce").fillna(0)
        else:
            df["Total dashboard cost"] = 0
    else:
        df["Total dashboard cost"] = pd.to_numeric(df["Costo Total Camion"], errors="coerce").fillna(0)

    # Explicit requested definition: total cost variable = existing Total cost per truck + Labour.
    # In this workbook Total Labour is already a monetary backlog/labour value.
    df["Dashboard total cost"] = pd.to_numeric(df.get("Total cost per truck", 0), errors="coerce").fillna(0) + pd.to_numeric(df.get("Total Labour", 0), errors="coerce").fillna(0)

    # Structural has duplicated component names. Pandas mangles duplicates as .1 for flag columns.
    component_names = comp_costs["Component"].tolist()
    component_flag_cols = {}
    component_life_cols = {}
    for comp_name in component_names:
        # flag columns are the duplicate/mangled columns or exact flag-only names after Kit 20
        candidates = [comp_name, f"{comp_name}.1"]
        aliases = {
            "Rear strut left": ["Rear strut left", "Rear strut left.1", "Rear strut right", "Rear strut right.1"],
            "Body reapirs": ["Body reapirs", "Body reapirs.1", "Body", "Body.1"],
        }.get(comp_name, [])
        candidates.extend(aliases)

        flag_col = None
        for cand in candidates:
            if cand in df.columns and (str(cand).endswith(".1") or cand in ["Body reapirs"]):
                flag_col = cand
                break
        if flag_col is None:
            # fall back to the last matching column by stripped name
            matches = [c for c in df.columns if str(c).replace(".1", "").strip() == comp_name]
            flag_col = matches[-1] if matches else None

        life_col = None
        for cand in [comp_name] + aliases:
            if cand in df.columns and not str(cand).endswith(".1"):
                life_col = cand
                break
        if life_col is None:
            matches = [c for c in df.columns if str(c).replace(".1", "").strip() == comp_name]
            life_col = matches[0] if matches else None

        component_flag_cols[comp_name] = flag_col
        component_life_cols[comp_name] = life_col

    return df, component_names, component_flag_cols, component_life_cols


def money(x):
    return f"${x:,.0f}"


def num(x):
    return f"{x:,.0f}"


def pct_life(x):
    if pd.isna(x):
        return "—"
    return f"{x:.1%}" if x <= 1.5 else f"{x:.1f}%"


def life_color(v):
    if pd.isna(v):
        return ""
    # Component life is shown green when high/healthy and red when low/consumed.
    val = v if v <= 1.5 else v / 100
    red = int(198 + (46 - 198) * val)
    green = int(40 + (125 - 40) * val)
    blue = int(40 + (50 - 40) * val)
    return f"background-color: rgba({red},{green},{blue},0.18); color: #1A1A1A; font-weight: 700;"


def plot_layout(fig, height=420, x_title=None, y_title=None):
    fig.update_layout(
        height=height,
        margin=dict(l=10, r=10, t=10, b=10),
        paper_bgcolor="#FFFFFF",
        plot_bgcolor="#FFFFFF",
        font=dict(family="Barlow", color=DARK),
        xaxis=dict(title=x_title, showgrid=False, tickfont=dict(family="Barlow Condensed", size=10)),
        yaxis=dict(title=y_title, showgrid=True, gridcolor=GRID),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
    )
    return fig


def make_component_detail(truck_row, comp_costs, flag_cols, life_cols):
    records = []
    for comp in comp_costs["Component"]:
        flag_col = flag_cols.get(comp)
        life_col = life_cols.get(comp)
        flag = pd.to_numeric(truck_row.get(flag_col, 0), errors="coerce") if flag_col else 0
        if pd.isna(flag):
            flag = 0
        if flag > 0:
            cost_row = comp_costs.loc[comp_costs["Component"] == comp].iloc[0]
            records.append({
                "Component": comp,
                "Required flag": flag,
                "Life percentage": pd.to_numeric(truck_row.get(life_col, np.nan), errors="coerce") if life_col else np.nan,
                "Labour cost": cost_row["Labour cost amount"] * flag,
                "Mechanized & Rebuild": cost_row.get("Mechanized & Rebuild", 0) * flag,
                "Parts": cost_row.get("parts", 0) * flag,
                "Chrome tube & rod": cost_row.get("Chrome tube & rod", 0) * flag,
            })
    out = pd.DataFrame(records)
    if not out.empty:
        out["Total cost"] = out[["Labour cost", "Mechanized & Rebuild", "Parts", "Chrome tube & rod"]].sum(axis=1)
        out = out.sort_values("Total cost", ascending=False)
    return out


def make_kit_detail(truck_row, kit_costs):
    records = []
    for _, k in kit_costs.iterrows():
        col = k["Kit column"]
        qty = pd.to_numeric(truck_row.get(col, 0), errors="coerce") if col in truck_row.index else 0
        if pd.isna(qty):
            qty = 0
        if qty > 0:
            records.append({
                "Kit": k["Kit"],
                "Quantity": qty,
                "Labour hours": k["Labour hours"] * qty,
                "Kit cost": k["Kit cost"] * qty,
                "Total kit value": (k["Kit cost"] + (k["Labour hours"] * 0)) * qty,
            })
    out = pd.DataFrame(records)
    if not out.empty:
        out = out.sort_values("Kit cost", ascending=False)
    return out

# ─────────────────────────────────────────────
# LOAD DATA
# ─────────────────────────────────────────────
with st.sidebar:
    st.markdown("### Data source")
    uploaded = st.file_uploader("Upload updated Excel workbook", type=["xlsx"])

try:
    structural, labour, comp, kits_raw = load_workbook(uploaded)
except Exception as exc:
    st.error(f"Could not load the Excel workbook. Place `{EXCEL_FILE}` next to this Python file or upload it from the sidebar. Error: {exc}")
    st.stop()

component_costs = build_component_cost_table(comp)
kit_costs = build_kit_cost_table(labour)
df_all, component_names, flag_cols, life_cols = prepare_data(structural, component_costs)

# Permanent baseline: first 19 trucks by lowest Weighted criteria; dashboard sorted by total cost.
baseline_ids = (
    df_all.sort_values(["Weighted criteria", "Dashboard total cost"], ascending=[True, True])
    .head(TOP_N_TRUCKS)["DT"]
    .tolist()
)
remaining_ids = [x for x in df_all["DT"].tolist() if x not in baseline_ids]

with st.sidebar:
    st.markdown("### Fleet filter")
    st.caption("Baseline contains the first 19 trucks with the lowest Weighted Criteria. Extra trucks can be added manually.")
    extra_ids = st.multiselect(
        "Add trucks outside baseline",
        options=remaining_ids,
        format_func=lambda x: f"DT {int(x)}" if pd.notna(x) and isinstance(x, (int, float, np.integer, np.floating)) else f"DT {x}",
    )
    selected_ids = baseline_ids + extra_ids

    st.markdown("### Sorting")
    sort_general = st.selectbox(
        "General dashboard sort",
        ["Total cost — ascending", "Weighted criteria — ascending", "Truck number — ascending"],
        index=0,
    )

base_df = df_all[df_all["DT"].isin(selected_ids)].copy()
if sort_general == "Total cost — ascending":
    df = base_df.sort_values("Dashboard total cost", ascending=True)
elif sort_general == "Weighted criteria — ascending":
    df = base_df.sort_values("Weighted criteria", ascending=True)
else:
    df = base_df.sort_values("DT", ascending=True)

# Derived metrics
component_counts = []
for _, row in df.iterrows():
    count = 0
    for comp_name in component_names:
        col = flag_cols.get(comp_name)
        if col and pd.to_numeric(row.get(col, 0), errors="coerce") > 0:
            count += 1
    component_counts.append(count)
df["Required components"] = component_counts
df["Total kit quantity"] = df[[c for c in KIT_COLS if c in df.columns]].apply(pd.to_numeric, errors="coerce").fillna(0).sum(axis=1)

# ─────────────────────────────────────────────
# HEADER
# ─────────────────────────────────────────────
st.markdown(
    """
    <div class="lc-header">
      <div>
        <div class="lc-header-title">FLEET REACTIVATION DASHBOARD</div>
        <div class="lc-header-subtitle">Cost, components, kits and structural severity analysis</div>
      </div>
      <div class="brand-pill">LANDCROS</div>
    </div>
    """,
    unsafe_allow_html=True,
)

tab_summary, tab_truck = st.tabs(["General Fleet Summary", "Cost Analysis per Truck"])

# ─────────────────────────────────────────────
# GENERAL SUMMARY TAB
# ─────────────────────────────────────────────
with tab_summary:
    total_cost = df["Dashboard total cost"].sum()
    avg_cost = df["Dashboard total cost"].mean() if len(df) else 0
    total_components = df["Required components"].sum()
    total_kits = df["Total kit quantity"].sum()

    st.markdown(
        f"""
        <div class="kpi-grid">
          <div class="kpi-card"><div class="kpi-label">Fleet total cost</div><div class="kpi-value">{money(total_cost)}</div><div class="kpi-sub">Selected trucks: {len(df)}</div></div>
          <div class="kpi-card"><div class="kpi-label">Average cost per truck</div><div class="kpi-value">{money(avg_cost)}</div><div class="kpi-sub">Total cost + labour</div></div>
          <div class="kpi-card"><div class="kpi-label">Required components</div><div class="kpi-value">{num(total_components)}</div><div class="kpi-sub">Binary component flags</div></div>
          <div class="kpi-card"><div class="kpi-label">Total kits</div><div class="kpi-value">{num(total_kits)}</div><div class="kpi-sub">Kit quantities across fleet</div></div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    c1, c2 = st.columns([1.25, 1], gap="large")

    with c1:
        st.markdown('<div class="section-title">Progressive Total Cost by Truck</div>', unsafe_allow_html=True)
        st.markdown('<div class="small-note">X-axis: truck. Y-axis: total cost. The chart is sorted to show the progressive cost increase.</div>', unsafe_allow_html=True)
        cost_df = df.sort_values("Dashboard total cost", ascending=True)
        fig = go.Figure()
        fig.add_trace(go.Bar(
            x=cost_df["DT"].astype(str),
            y=cost_df["Dashboard total cost"],
            marker_color=PRIMARY,
            text=[money(v) for v in cost_df["Dashboard total cost"]],
            textposition="outside",
            hovertemplate="Truck DT %{x}<br>Total cost: $%{y:,.0f}<extra></extra>",
        ))
        fig.add_trace(go.Scatter(
            x=cost_df["DT"].astype(str),
            y=cost_df["Dashboard total cost"].cumsum(),
            mode="lines+markers",
            name="Cumulative cost",
            line=dict(color=DARK, width=2),
            marker=dict(size=6),
            yaxis="y2",
            hovertemplate="Truck DT %{x}<br>Cumulative cost: $%{y:,.0f}<extra></extra>",
        ))
        fig.update_layout(
            yaxis=dict(title="Truck total cost (USD)", tickformat="$,.0f", showgrid=True, gridcolor=GRID),
            yaxis2=dict(title="Cumulative cost (USD)", overlaying="y", side="right", tickformat="$,.0f", showgrid=False),
            xaxis=dict(title="Truck DT", showgrid=False),
        )
        plot_layout(fig, height=480)
        st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})

    with c2:
        st.markdown('<div class="section-title">Required Components by Truck</div>', unsafe_allow_html=True)
        fig = go.Figure(go.Bar(
            x=df.sort_values("Required components", ascending=False)["DT"].astype(str),
            y=df.sort_values("Required components", ascending=False)["Required components"],
            marker_color=DARK,
            text=df.sort_values("Required components", ascending=False)["Required components"],
            textposition="outside",
            hovertemplate="Truck DT %{x}<br>Required components: %{y}<extra></extra>",
        ))
        plot_layout(fig, height=480, x_title="Truck DT", y_title="Required components")
        st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})

    c3, c4 = st.columns([1, 1], gap="large")

    with c3:
        st.markdown('<div class="section-title">Total Kits Required by Fleet</div>', unsafe_allow_html=True)
        kit_summary = []
        for col in KIT_COLS:
            if col in df.columns:
                qty = pd.to_numeric(df[col], errors="coerce").fillna(0).sum()
                if qty > 0:
                    unit = kit_costs.loc[kit_costs["Kit column"] == col]
                    cost = unit["Kit cost"].iloc[0] * qty if not unit.empty else 0
                    kit_summary.append({"Kit": KIT_LABELS.get(col, col), "Quantity": qty, "Total cost": cost})
        kit_summary = pd.DataFrame(kit_summary).sort_values("Quantity", ascending=False) if kit_summary else pd.DataFrame(columns=["Kit", "Quantity", "Total cost"])
        fig = go.Figure(go.Bar(
            y=kit_summary["Kit"],
            x=kit_summary["Quantity"],
            orientation="h",
            marker_color=PRIMARY_LIGHT,
            text=kit_summary["Quantity"],
            textposition="outside",
            hovertemplate="%{y}<br>Quantity: %{x:,.0f}<extra></extra>",
        ))
        plot_layout(fig, height=520, x_title="Quantity", y_title=None)
        fig.update_layout(yaxis=dict(autorange="reversed", tickfont=dict(size=9)))
        st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})

    with c4:
        st.markdown('<div class="section-title">Structural Severity by Truck</div>', unsafe_allow_html=True)
        sev_df = df[["DT"] + [c for c in SEVERITY_COLS if c in df.columns]].copy()
        sev_df = sev_df.set_index("DT")
        fig = go.Figure(go.Heatmap(
            z=sev_df.values,
            x=[c.replace(" Severity", "") for c in sev_df.columns],
            y=sev_df.index.astype(str),
            colorscale=[[0, GREEN], [0.5, YELLOW], [1, RED]],
            zmin=0,
            zmax=max(4, np.nanmax(sev_df.values) if sev_df.size else 4),
            text=sev_df.values,
            texttemplate="%{text}",
            hovertemplate="Truck DT %{y}<br>%{x}: %{z}<extra></extra>",
            colorbar=dict(title="Severity"),
        ))
        plot_layout(fig, height=520)
        fig.update_layout(yaxis=dict(autorange="reversed"))
        st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})

    st.markdown('<div class="section-title">Fleet Detail Table</div>', unsafe_allow_html=True)
    detail_cols = [
        "DT", "Hours", "Weighted criteria", "Total Lenghts cracks (mm)",
        "Required components", "Total kit quantity", "Total Labour", "Total cost per truck", "Dashboard total cost",
    ]
    existing_cols = [c for c in detail_cols if c in df.columns]
    out = df[existing_cols].copy()
    rename = {
        "DT": "Truck DT",
        "Hours": "Operating hours",
        "Weighted criteria": "Weighted criteria",
        "Total Lenghts cracks (mm)": "Total crack length (mm)",
        "Required components": "Required components",
        "Total kit quantity": "Total kit quantity",
        "Total Labour": "Labour / backlog cost",
        "Total cost per truck": "Base total cost",
        "Dashboard total cost": "Dashboard total cost",
    }
    out = out.rename(columns=rename)
    st.dataframe(out, use_container_width=True, hide_index=True)

# ─────────────────────────────────────────────
# TRUCK ANALYSIS TAB
# ─────────────────────────────────────────────
with tab_truck:
    truck_options = df_all["DT"].tolist()
    default_truck = df.sort_values("Dashboard total cost", ascending=False)["DT"].iloc[0] if len(df) else truck_options[0]
    selected_truck = st.selectbox(
        "Select truck for individual cost analysis",
        options=truck_options,
        index=truck_options.index(default_truck) if default_truck in truck_options else 0,
        format_func=lambda x: f"DT {int(x)}" if pd.notna(x) and isinstance(x, (int, float, np.integer, np.floating)) else f"DT {x}",
    )

    truck_row = df_all[df_all["DT"] == selected_truck].iloc[0]
    comp_detail = make_component_detail(truck_row, component_costs, flag_cols, life_cols)
    kit_detail = make_kit_detail(truck_row, kit_costs)

    total_truck_cost = pd.to_numeric(truck_row.get("Dashboard total cost", 0), errors="coerce")
    hours = pd.to_numeric(truck_row.get("Hours", 0), errors="coerce")
    weighted = pd.to_numeric(truck_row.get("Weighted criteria", 0), errors="coerce")

    st.markdown(
        f"""
        <div class="kpi-grid">
          <div class="kpi-card"><div class="kpi-label">Selected truck</div><div class="kpi-value">DT {selected_truck}</div><div class="kpi-sub">Individual analysis</div></div>
          <div class="kpi-card"><div class="kpi-label">Truck total cost</div><div class="kpi-value">{money(total_truck_cost)}</div><div class="kpi-sub">Total cost + labour</div></div>
          <div class="kpi-card"><div class="kpi-label">Operating hours</div><div class="kpi-value">{num(hours)}</div><div class="kpi-sub">Backlog operating exposure</div></div>
          <div class="kpi-card"><div class="kpi-label">Weighted criteria</div><div class="kpi-value">{weighted:.3f}</div><div class="kpi-sub">Crack length × HTM severity</div></div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.markdown('<div class="section-title">Applicable Components and Cost Breakdown</div>', unsafe_allow_html=True)
    if comp_detail.empty:
        st.info("No applicable component flags were found for this truck.")
    else:
        comp_table = comp_detail.copy()
        comp_table["Life percentage label"] = comp_table["Life percentage"].map(pct_life)
        show_cols = ["Component", "Life percentage label", "Labour cost", "Mechanized & Rebuild", "Parts", "Chrome tube & rod", "Total cost"]
        styled = comp_table[show_cols].rename(columns={"Life percentage label": "Life percentage"})
        st.dataframe(
            styled.style
                .format({
                    "Labour cost": "${:,.0f}",
                    "Mechanized & Rebuild": "${:,.0f}",
                    "Parts": "${:,.0f}",
                    "Chrome tube & rod": "${:,.0f}",
                    "Total cost": "${:,.0f}",
                })
                .apply(lambda s: [life_color(comp_detail.loc[s.name, "Life percentage"]) if col == "Life percentage" else "" for col in s.index], axis=1),
            use_container_width=True,
            hide_index=True,
        )

        fig = go.Figure()
        cost_parts = ["Labour cost", "Mechanized & Rebuild", "Parts", "Chrome tube & rod"]
        colors = [DARK, PRIMARY, PRIMARY_LIGHT, MID]
        for part, color in zip(cost_parts, colors):
            fig.add_trace(go.Bar(
                x=comp_detail["Component"],
                y=comp_detail[part],
                name=part,
                marker_color=color,
                text=[money(v) if v > 0 else "" for v in comp_detail[part]],
                textposition="inside",
                hovertemplate=f"%{{x}}<br>{part}: $%{{y:,.0f}}<extra></extra>",
            ))
        fig.add_trace(go.Scatter(
            x=comp_detail["Component"],
            y=comp_detail["Total cost"],
            mode="text",
            text=[money(v) for v in comp_detail["Total cost"]],
            textposition="top center",
            showlegend=False,
            textfont=dict(family="Barlow Condensed", size=11, color=DARK),
            hoverinfo="skip",
        ))
        fig.update_layout(barmode="stack", xaxis_tickangle=-35, yaxis=dict(tickformat="$,.0f"))
        plot_layout(fig, height=560, x_title="Component", y_title="Cost (USD)")
        st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})

    st.markdown('<div class="section-title">Kit Analysis</div>', unsafe_allow_html=True)
    if kit_detail.empty:
        st.info("No applicable kits were found for this truck.")
    else:
        kc1, kc2, kc3 = st.columns(3, gap="large")
        with kc1:
            fig = go.Figure(go.Bar(
                y=kit_detail["Kit"],
                x=kit_detail["Quantity"],
                orientation="h",
                marker_color=PRIMARY,
                text=kit_detail["Quantity"],
                textposition="outside",
                hovertemplate="%{y}<br>Quantity: %{x:,.0f}<extra></extra>",
            ))
            plot_layout(fig, height=430, x_title="Quantity", y_title=None)
            fig.update_layout(yaxis=dict(autorange="reversed", tickfont=dict(size=9)))
            st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})
        with kc2:
            fig = go.Figure(go.Bar(
                y=kit_detail["Kit"],
                x=kit_detail["Labour hours"],
                orientation="h",
                marker_color=DARK,
                text=[num(v) for v in kit_detail["Labour hours"]],
                textposition="outside",
                hovertemplate="%{y}<br>Labour hours: %{x:,.0f}<extra></extra>",
            ))
            plot_layout(fig, height=430, x_title="Labour hours", y_title=None)
            fig.update_layout(yaxis=dict(autorange="reversed", tickfont=dict(size=9)))
            st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})
        with kc3:
            fig = go.Figure(go.Bar(
                y=kit_detail["Kit"],
                x=kit_detail["Kit cost"],
                orientation="h",
                marker_color=PRIMARY_LIGHT,
                text=[money(v) for v in kit_detail["Kit cost"]],
                textposition="outside",
                hovertemplate="%{y}<br>Kit cost: $%{x:,.0f}<extra></extra>",
            ))
            plot_layout(fig, height=430, x_title="Kit cost (USD)", y_title=None)
            fig.update_layout(yaxis=dict(autorange="reversed", tickfont=dict(size=9), showticklabels=False), xaxis=dict(tickformat="$,.0f"))
            st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})

        kit_show = kit_detail[["Kit", "Quantity", "Labour hours", "Kit cost"]].copy()
        st.dataframe(
            kit_show.style.format({"Quantity": "{:,.0f}", "Labour hours": "{:,.0f}", "Kit cost": "${:,.0f}"}),
            use_container_width=True,
            hide_index=True,
        )

# ─────────────────────────────────────────────
# FOOTER
# ─────────────────────────────────────────────
st.markdown(
    """
    <div style="margin-top:32px;padding:16px 0;border-top:1px solid #EFEFEF;display:flex;justify-content:space-between;align-items:center;">
      <span style="font-family:'Barlow Condensed';font-size:0.8rem;color:#AAAAAA;letter-spacing:0.08em;text-transform:uppercase;">
        Landcros — Fleet Reactivation Analysis
      </span>
      <span style="font-family:'Barlow';font-size:0.78rem;color:#CCCCCC;">
        Source: Data base Reactivation.xlsx
      </span>
    </div>
    """,
    unsafe_allow_html=True,
)
