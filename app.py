"""
ASHA & Infant Visit Due Monitoring, Healthcare Network & Forecasting Portal
Enterprise Health Analytics Dashboard for Home Based Newborn Care (HBNC)
"""

from datetime import date, datetime, timedelta
from io import BytesIO
import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.platypus import (
    HRFlowable,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)
import streamlit as st

# -----------------------------------------------------------------------------
# 1. Page Configuration & Custom UI Styling
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="ASHA & Baby Visit Monitoring Dashboard",
    page_icon="👶",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
<style>
    .reportview-container .main .block-container {
        padding-top: 1.2rem;
        padding-bottom: 2rem;
    }
    .main-title {
        font-size: 26px;
        font-weight: 800;
        color: #1E3A8A;
        letter-spacing: -0.3px;
        margin-bottom: 2px;
    }
    .sub-title {
        font-size: 13px;
        color: #64748B;
        margin-bottom: 16px;
    }
    .tier-header {
        font-size: 14px;
        font-weight: 700;
        color: #1E293B;
        margin-top: 8px;
        margin-bottom: 8px;
        border-left: 4px solid #2563EB;
        padding-left: 8px;
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }
    .sub-section-header {
        font-size: 15px;
        font-weight: 700;
        color: #0F172A;
        margin-top: 14px;
        margin-bottom: 6px;
    }
    .kpi-card {
        background: #FFFFFF;
        border: 1px solid #E2E8F0;
        border-radius: 10px;
        padding: 12px 14px;
        box-shadow: 0 1px 3px rgba(0, 0, 0, 0.05);
        height: 100%;
        display: flex;
        flex-direction: column;
        justify-content: center;
    }
    .kpi-title {
        font-size: 11px;
        font-weight: 700;
        color: #64748B;
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }
    .kpi-value {
        font-size: 25px;
        font-weight: 800;
        margin-top: 3px;
        line-height: 1.1;
    }
    .kpi-caption {
        font-size: 11px;
        color: #94A3B8;
        margin-top: 3px;
    }
</style>
""",
    unsafe_allow_html=True,
)


# -----------------------------------------------------------------------------
# 2. Export Utility Functions (PDF & Multi-Sheet Excel)
# -----------------------------------------------------------------------------
def generate_pdf_report(
    df_export, columns_to_show, report_title, active_filters_desc
):
    """Generates an executive-formatted landscape A4 PDF report matching active filters."""
    buffer = BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=landscape(A4),
        rightMargin=18,
        leftMargin=18,
        topMargin=18,
        bottomMargin=18,
    )
    styles = getSampleStyleSheet()

    elements = []
    title_style = ParagraphStyle(
        name="DocTitle",
        fontSize=13,
        leading=16,
        textColor=colors.HexColor("#1E3A8A"),
        fontName="Helvetica-Bold",
    )
    meta_style = ParagraphStyle(
        name="DocMeta",
        fontSize=8,
        leading=10,
        textColor=colors.HexColor("#475569"),
    )

    elements.append(Paragraph(report_title, title_style))
    elements.append(Spacer(1, 3))
    elements.append(
        Paragraph(
            f"<b>Filter Scope:</b> {active_filters_desc} &nbsp;|&nbsp; <b>Exported:</b> {datetime.now().strftime('%d-%b-%Y %I:%M %p')} &nbsp;|&nbsp; <b>Total Records:</b> {len(df_export)}",
            meta_style,
        )
    )
    elements.append(Spacer(1, 6))
    elements.append(
        HRFlowable(
            width="100%",
            thickness=1,
            color=colors.HexColor("#CBD5E1"),
            spaceAfter=8,
        )
    )

    valid_cols = [c for c in columns_to_show if c in df_export.columns]
    df_clean = df_export[valid_cols].copy()

    for col in df_clean.columns:
        if pd.api.types.is_datetime64_any_dtype(df_clean[col]):
            df_clean[col] = df_clean[col].dt.strftime("%d-%b-%Y")
        else:
            df_clean[col] = (
                df_clean[col].astype(str).replace({"nan": "-", "NaT": "-"})
            )

    data = [valid_cols] + df_clean.values.tolist()
    cell_style = ParagraphStyle(name="CellTxt", fontSize=7.5, leading=9)
    head_style = ParagraphStyle(
        name="HeadTxt",
        fontSize=8,
        leading=10,
        textColor=colors.whitesmoke,
        fontName="Helvetica-Bold",
    )

    table_data = []
    for r_idx, row in enumerate(data):
        row_cells = []
        for cell in row:
            text = str(cell) if cell is not None else "-"
            if r_idx == 0:
                row_cells.append(Paragraph(f"<b>{text}</b>", head_style))
            else:
                row_cells.append(Paragraph(text, cell_style))
        table_data.append(row_cells)

    pdf_table = Table(table_data, repeatRows=1)
    pdf_table.setStyle(
        TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1E3A8A")),
            ("ALIGN", (0, 0), (-1, -1), "LEFT"),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
            ("TOPPADDING", (0, 0), (-1, -1), 3),
            ("LEFTPADDING", (0, 0), (-1, -1), 4),
            ("RIGHTPADDING", (0, 0), (-1, -1), 4),
            ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#CBD5E1")),
            (
                "ROWBACKGROUNDS",
                (0, 1),
                (-1, -1),
                [colors.white, colors.HexColor("#F8FAFC")],
            ),
        ])
    )

    elements.append(pdf_table)
    doc.build(elements)
    buffer.seek(0)
    return buffer.getvalue()


def generate_single_table_excel(df_table, sheet_name="Data_Export"):
    """Generates an Excel file for a single dataframe."""
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df_table.to_excel(writer, sheet_name=sheet_name[:30], index=False)
    buffer.seek(0)
    return buffer.getvalue()


def generate_excel_workbook(sheets_dict):
    """Generates an Excel workbook containing multiple analytical sheets."""
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        for sheet_name, df_sheet in sheets_dict.items():
            if not df_sheet.empty:
                df_sheet.to_excel(
                    writer, sheet_name=sheet_name[:30], index=False
                )
    buffer.seek(0)
    return buffer.getvalue()


# -----------------------------------------------------------------------------
# 3. Sidebar: File Upload & Ingestion
# -----------------------------------------------------------------------------
st.sidebar.markdown("### 📁 Data Source")
uploaded_file = st.sidebar.file_uploader(
    "Upload HBNC Schedule (.xlsx, .xls, .csv)",
    type=["xlsx", "xls", "csv"],
    help="Upload the HBNC Visit Due Calendar file.",
)

if uploaded_file is None:
    st.info(
        "👋 **Welcome!** Please upload your HBNC Visit Calendar (.xlsx or .csv) from the sidebar to activate the dashboard."
    )
    st.stop()

try:
    if uploaded_file.name.endswith(".csv"):
        df_raw = pd.read_csv(uploaded_file)
    else:
        df_raw = pd.read_excel(uploaded_file)
except Exception as e:
    st.error(f"Error reading file: {e}")
    st.stop()

# -----------------------------------------------------------------------------
# 4. Data Cleaning, Normalization & Standardization
# -----------------------------------------------------------------------------
col_rename = {}
for c in df_raw.columns:
    clean = c.strip().lower()
    if clean in ["chc", "phc", "facility", "block"]:
        col_rename[c] = "PHC"
    elif clean in ["district", "zilla"]:
        col_rename[c] = "District"
    elif clean in ["subcentre", "sub-centre", "sub centre"]:
        col_rename[c] = "Subcentre"
    elif clean in ["asha", "asha name"]:
        col_rename[c] = "ASHA"
    elif clean in ["baby", "baby name", "child name", "child"]:
        col_rename[c] = "Baby"
    elif clean in ["mother", "mother name"]:
        col_rename[c] = "Mother"
    elif clean in ["baby id", "child id", "id"]:
        col_rename[c] = "Baby ID"

df_raw.rename(columns=col_rename, inplace=True)

# Default fallback values for standard fields
if "District" not in df_raw.columns:
    df_raw["District"] = "Fatehpur"
if "PHC" not in df_raw.columns:
    df_raw["PHC"] = "General PHC"
if "Subcentre" not in df_raw.columns:
    df_raw["Subcentre"] = "General Subcentre"
if "ASHA" not in df_raw.columns:
    df_raw["ASHA"] = "Unknown ASHA"
if "Baby ID" not in df_raw.columns:
    df_raw["Baby ID"] = [f"B-{i+1:04d}" for i in range(len(df_raw))]

for c in ["District", "PHC", "Subcentre", "ASHA", "Mother"]:
    df_raw[c] = (
        df_raw[c]
        .astype(str)
        .str.strip()
        .replace({"nan": "Unknown", "": "Unknown"})
    )

df_raw["Baby Display"] = df_raw.apply(
    lambda r: str(r["Baby"]).strip()
    if pd.notna(r.get("Baby")) and str(r.get("Baby")).strip() not in ["", "nan"]
    else f"Baby of {r['Mother']}",
    axis=1,
)

# Standardize date fields
due_cols = sorted([c for c in df_raw.columns if "visit due" in c.lower()])
if not due_cols:
    st.error("Invalid File Format: No 'Visit Due' milestone columns detected.")
    st.stop()

for col in due_cols:
    df_raw[col] = pd.to_datetime(df_raw[col], errors="coerce").dt.date

for dc in ["Registration Date", "DOB", "Last Visit Date"]:
    if dc in df_raw.columns:
        df_raw[dc] = pd.to_datetime(df_raw[dc], errors="coerce").dt.date

# -----------------------------------------------------------------------------
# 5. Sidebar: Evaluation Reference Date & Date Range Configuration
# -----------------------------------------------------------------------------
st.sidebar.markdown("---")
st.sidebar.markdown("### 🗓️ Evaluation Reference Date")
use_custom_date = st.sidebar.checkbox(
    "Set Manual Evaluation Date",
    value=False,
    help="Enable to test historical dates matching the dataset window.",
)

if use_custom_date:
    ref_date = st.sidebar.date_input("Evaluation Date:", value=date(2026, 10, 12))
else:
    ref_date = date.today()

today_date = ref_date
yesterday_date = ref_date - timedelta(days=1)
tomorrow_date = ref_date + timedelta(days=1)

st.sidebar.caption(
    f"Yesterday: **{yesterday_date.strftime('%d-%b-%Y')}**\n\n"
    f"Today: **{today_date.strftime('%d-%b-%Y')}**\n\n"
    f"Tomorrow: **{tomorrow_date.strftime('%d-%b-%Y')}**"
)

# -----------------------------------------------------------------------------
# 6. Accurate Visit-Level Classification Engine
# -----------------------------------------------------------------------------
visit_records = []
for _, row in df_raw.iterrows():
    lvd = (
        row["Last Visit Date"]
        if ("Last Visit Date" in row and pd.notna(row["Last Visit Date"]))
        else None
    )

    for stage in due_cols:
        d_val = row[stage]
        if pd.notna(d_val):
            # Lifecycle classification:
            if d_val < today_date:
                # Past due: Completed if Last Visit Date >= Due Date, otherwise Missed
                v_status = "Completed" if (lvd and lvd >= d_val) else "Missed"
            elif d_val == today_date:
                # Today: Completed if recorded today, otherwise Pending (NOT Missed yet)
                v_status = (
                    "Completed" if (lvd and lvd >= today_date) else "Pending"
                )
            else:
                # Future: Strictly Upcoming (Never classified as Missed)
                v_status = "Upcoming"

            visit_records.append({
                "Baby ID": row["Baby ID"],
                "Baby Display": row["Baby Display"],
                "Mother": row["Mother"],
                "ASHA": row["ASHA"],
                "District": row["District"],
                "PHC": row["PHC"],
                "Subcentre": row["Subcentre"],
                "DOB": row.get("DOB"),
                "Registration Date": row.get("Registration Date"),
                "Last Visit Date": lvd,
                "Visit Stage": stage,
                "Due Date": d_val,
                "Visit Status": v_status,
                "Month Year": d_val.strftime("%B %Y"),
            })

df_all_visits = pd.DataFrame(visit_records)

# Accurate baby-level missed visit calculation (Summed at individual visit milestone level)
baby_missed_series = (
    df_all_visits[df_all_visits["Visit Status"] == "Missed"]
    .groupby("Baby ID")
    .size()
)
df_raw["Missed Visits Count"] = (
    df_raw["Baby ID"].map(baby_missed_series).fillna(0).astype(int)
)

# -----------------------------------------------------------------------------
# 7. Sidebar: Multi-Tier Drill-Down Filters
# -----------------------------------------------------------------------------
st.sidebar.markdown("---")
st.sidebar.markdown("### 🔍 Drill-Down Filters")

# 1. District
dist_opts = ["All"] + sorted(df_raw["District"].unique().tolist())
sel_dist = st.sidebar.selectbox("1. District", dist_opts)
df_scope1 = (
    df_raw if sel_dist == "All" else df_raw[df_raw["District"] == sel_dist]
)

# 2. Block / Facility (PHC/CHC)
phc_opts = ["All"] + sorted(df_scope1["PHC"].unique().tolist())
sel_phc = st.sidebar.selectbox("2. CHC / PHC Facility", phc_opts)
df_scope2 = (
    df_scope1 if sel_phc == "All" else df_scope1[df_scope1["PHC"] == sel_phc]
)

# 3. Subcentre
sub_opts = ["All"] + sorted(df_scope2["Subcentre"].unique().tolist())
sel_sub = st.sidebar.selectbox("3. Subcentre", sub_opts)
df_scope3 = (
    df_scope2 if sel_sub == "All" else df_scope2[df_scope2["Subcentre"] == sel_sub]
)

# 4. ASHA
asha_opts = ["All"] + sorted(df_scope3["ASHA"].unique().tolist())
sel_asha = st.sidebar.selectbox("4. ASHA Worker", asha_opts)

# 5. Scheduled Date Range Filter
min_due_date = df_all_visits["Due Date"].min()
max_due_date = df_all_visits["Due Date"].max()

st.sidebar.markdown("##### 5. Scheduled Date Range")
date_range = st.sidebar.date_input(
    "Filter by Milestone Due Date Range",
    value=(min_due_date, max_due_date),
    min_value=min_due_date,
    max_value=max_due_date,
)

if isinstance(date_range, tuple) and len(date_range) == 2:
    start_date_filter, end_date_filter = date_range
else:
    start_date_filter, end_date_filter = min_due_date, max_due_date

# Filter baby and visit datasets
df_v_filtered = df_all_visits.copy()
df_b_filtered = df_raw.copy()

if sel_dist != "All":
    df_v_filtered = df_v_filtered[df_v_filtered["District"] == sel_dist]
    df_b_filtered = df_b_filtered[df_b_filtered["District"] == sel_dist]
if sel_phc != "All":
    df_v_filtered = df_v_filtered[df_v_filtered["PHC"] == sel_phc]
    df_b_filtered = df_b_filtered[df_b_filtered["PHC"] == sel_phc]
if sel_sub != "All":
    df_v_filtered = df_v_filtered[df_v_filtered["Subcentre"] == sel_sub]
    df_b_filtered = df_b_filtered[df_b_filtered["Subcentre"] == sel_sub]
if sel_asha != "All":
    df_v_filtered = df_v_filtered[df_v_filtered["ASHA"] == sel_asha]
    df_b_filtered = df_b_filtered[df_b_filtered["ASHA"] == sel_asha]

# Filter visits by date range
df_v_filtered = df_v_filtered[
    (df_v_filtered["Due Date"] >= start_date_filter)
    & (df_v_filtered["Due Date"] <= end_date_filter)
]

# Synchronize baby cohort to matching records
cohort_baby_ids = df_v_filtered["Baby ID"].unique()
df_b_filtered = df_b_filtered[df_b_filtered["Baby ID"].isin(cohort_baby_ids)]

active_filters_str = (
    f"District: {sel_dist} | CHC/PHC: {sel_phc} | Subcentre: {sel_sub} | "
    f"ASHA: {sel_asha} | Range: {start_date_filter.strftime('%d-%b-%Y')} to {end_date_filter.strftime('%d-%b-%Y')} | "
    f"Ref Date: {today_date.strftime('%d-%b-%Y')}"
)

# -----------------------------------------------------------------------------
# 8. Compute Required Key Performance Indicators (KPIs)
# -----------------------------------------------------------------------------
df_v_geo = df_all_visits.copy()
if sel_dist != "All":
    df_v_geo = df_v_geo[df_v_geo["District"] == sel_dist]
if sel_phc != "All":
    df_v_geo = df_v_geo[df_v_geo["PHC"] == sel_phc]
if sel_sub != "All":
    df_v_geo = df_v_geo[df_v_geo["Subcentre"] == sel_sub]
if sel_asha != "All":
    df_v_geo = df_v_geo[df_v_geo["ASHA"] == sel_asha]

# Tier 1: Healthcare Network Scale KPIs
kpi_total_ashas = df_b_filtered["ASHA"].nunique()
kpi_total_phc = df_b_filtered["PHC"].nunique()
kpi_total_subcentres = df_b_filtered["Subcentre"].nunique()
kpi_total_babies = len(df_b_filtered)

# Tier 2: Daily Visit Operational Cadence KPIs
v_yesterday = df_v_geo[df_v_geo["Due Date"] == yesterday_date]
v_today = df_v_geo[df_v_geo["Due Date"] == today_date]
v_tomorrow = df_v_geo[df_v_geo["Due Date"] == tomorrow_date]

kpi_yest_scheduled = len(v_yesterday)
kpi_yest_completed = (v_yesterday["Visit Status"] == "Completed").sum()
kpi_yest_missed = (v_yesterday["Visit Status"] == "Missed").sum()
kpi_yest_completion_rate = (
    (kpi_yest_completed / kpi_yest_scheduled * 100.0)
    if kpi_yest_scheduled > 0
    else 0.0
)

kpi_today_scheduled = len(v_today)
kpi_tom_scheduled = len(v_tomorrow)
kpi_tom_babies_count = v_tomorrow["Baby ID"].nunique()

# -----------------------------------------------------------------------------
# 9. Dashboard Header & Prominent KPI Cards
# -----------------------------------------------------------------------------
st.markdown(
    '<div class="main-title">👶 ASHA & Infant Visit Due Monitoring Dashboard</div>',
    unsafe_allow_html=True,
)
st.markdown(
    f'<div class="sub-title"><b>Active Cohort Scope:</b> {active_filters_str}</div>',
    unsafe_allow_html=True,
)

# --- Tier 1: Healthcare Network Scale (4 Cards) ---
st.markdown(
    '<div class="tier-header">🏛️ Healthcare Network Scale & Demographics</div>',
    unsafe_allow_html=True,
)
n1, n2, n3, n4 = st.columns(4)

with n1:
    st.markdown(
        f"""
    <div class="kpi-card" style="border-left: 4px solid #0D9488;">
        <div class="kpi-title">Total ASHAs</div>
        <div class="kpi-value" style="color: #0F766E;">{kpi_total_ashas:,}</div>
        <div class="kpi-caption">Active workers in cohort</div>
    </div>
    """,
        unsafe_allow_html=True,
    )

with n2:
    st.markdown(
        f"""
    <div class="kpi-card" style="border-left: 4px solid #3B82F6;">
        <div class="kpi-title">Total CHC / PHC</div>
        <div class="kpi-value" style="color: #1D4ED8;">{kpi_total_phc:,}</div>
        <div class="kpi-caption">Community / Primary Centres</div>
    </div>
    """,
        unsafe_allow_html=True,
    )

with n3:
    st.markdown(
        f"""
    <div class="kpi-card" style="border-left: 4px solid #8B5CF6;">
        <div class="kpi-title">Total Subcentres</div>
        <div class="kpi-value" style="color: #6D28D9;">{kpi_total_subcentres:,}</div>
        <div class="kpi-caption">Active Subcentres</div>
    </div>
    """,
        unsafe_allow_html=True,
    )

with n4:
    st.markdown(
        f"""
    <div class="kpi-card" style="border-left: 4px solid #2563EB;">
        <div class="kpi-title">Total Registered Babies</div>
        <div class="kpi-value" style="color: #1E40AF;">{kpi_total_babies:,}</div>
        <div class="kpi-caption">Registered infants</div>
    </div>
    """,
        unsafe_allow_html=True,
    )

# --- Tier 2: Daily Visit Operational Cadence (5 Cards) ---
st.markdown(
    '<div class="tier-header">⏱️ Daily Visit Monitoring Cadence (Yesterday, Today & Tomorrow)</div>',
    unsafe_allow_html=True,
)
d1, d2, d3, d4, d5 = st.columns(5)

with d1:
    st.markdown(
        f"""
    <div class="kpi-card" style="border-left: 4px solid #64748B;">
        <div class="kpi-title">Yesterday's Scheduled</div>
        <div class="kpi-value" style="color: #334155;">{kpi_yest_scheduled:,}</div>
        <div class="kpi-caption">{yesterday_date.strftime('%d-%b-%Y')}</div>
    </div>
    """,
        unsafe_allow_html=True,
    )

with d2:
    st.markdown(
        f"""
    <div class="kpi-card" style="border-left: 4px solid #10B981;">
        <div class="kpi-title">Yesterday's Completed</div>
        <div class="kpi-value" style="color: #059669;">{kpi_yest_completed:,}</div>
        <div class="kpi-caption">Rate: <b>{kpi_yest_completion_rate:.1f}%</b></div>
    </div>
    """,
        unsafe_allow_html=True,
    )

with d3:
    st.markdown(
        f"""
    <div class="kpi-card" style="border-left: 4px solid #EF4444;">
        <div class="kpi-title">Yesterday's Missed</div>
        <div class="kpi-value" style="color: #DC2626;">{kpi_yest_missed:,}</div>
        <div class="kpi-caption">Uncompleted milestone visits</div>
    </div>
    """,
        unsafe_allow_html=True,
    )

with d4:
    st.markdown(
        f"""
    <div class="kpi-card" style="border-left: 4px solid #F59E0B;">
        <div class="kpi-title">Today's Scheduled</div>
        <div class="kpi-value" style="color: #D97706;">{kpi_today_scheduled:,}</div>
        <div class="kpi-caption">{today_date.strftime('%d-%b-%Y')} (Visits Due)</div>
    </div>
    """,
        unsafe_allow_html=True,
    )

with d5:
    st.markdown(
        f"""
    <div class="kpi-card" style="border-left: 4px solid #6366F1;">
        <div class="kpi-title">Tomorrow's Scheduled</div>
        <div class="kpi-value" style="color: #4F46E5;">{kpi_tom_scheduled:,}</div>
        <div class="kpi-caption">{tomorrow_date.strftime('%d-%b-%Y')} ({kpi_tom_babies_count} Babies)</div>
    </div>
    """,
        unsafe_allow_html=True,
    )

st.markdown("<br>", unsafe_allow_html=True)

# -----------------------------------------------------------------------------
# 10. Dashboard Functional Tabs
# -----------------------------------------------------------------------------
tab_daily, tab_network, tab_baby_master, tab_trends, tab_export = st.tabs([
    "📊 Daily Visit Overview & Visualizations",
    "👩‍⚕️ ASHA & Healthcare Facility Overview",
    "👶 Baby Master & Missed Visits Ledger",
    "📈 Facility & Operational Trends",
    "📥 Download Filtered Reports",
])

# -----------------------------------------------------------------------------
# TAB 1: Daily Visit Monitoring, Visualizations & Baby-Wise Rosters
# -----------------------------------------------------------------------------
with tab_daily:
    st.subheader(
        "⏱️ Daily Visit Overview: Yesterday, Today & Tomorrow Comparisons"
    )

    row1_c1, row1_c2 = st.columns(2)

    with row1_c1:
        st.markdown(
            f"**1. Yesterday's Performance: Scheduled vs. Completed vs. Missed ({yesterday_date.strftime('%d-%b-%Y')})**"
        )
        df_y_chart = pd.DataFrame({
            "Metric": ["Scheduled", "Completed", "Missed"],
            "Visits": [kpi_yest_scheduled, kpi_yest_completed, kpi_yest_missed],
            "Color": ["#3B82F6", "#10B981", "#EF4444"],
        })
        fig_yest = px.bar(
            df_y_chart,
            x="Metric",
            y="Visits",
            color="Metric",
            color_discrete_map={
                "Scheduled": "#3B82F6",
                "Completed": "#10B981",
                "Missed": "#EF4444",
            },
            text="Visits",
            height=300,
        )
        fig_yest.update_layout(
            margin=dict(l=20, r=20, t=30, b=20),
            showlegend=False,
            yaxis_title="Total Visits",
            plot_bgcolor="rgba(248,250,252,0.5)",
        )
        st.plotly_chart(fig_yest, use_container_width=True)

    with row1_c2:
        st.markdown(
            "**2. Visit Schedule Comparison: Yesterday vs. Today vs. Tomorrow**"
        )
        df_comp = pd.DataFrame({
            "Day": [
                f"Yesterday ({yesterday_date.strftime('%d-%b')})",
                f"Today ({today_date.strftime('%d-%b')})",
                f"Tomorrow ({tomorrow_date.strftime('%d-%b')})",
            ],
            "Scheduled Visits": [
                kpi_yest_scheduled,
                kpi_today_scheduled,
                kpi_tom_scheduled,
            ],
        })
        fig_comp = px.bar(
            df_comp,
            x="Day",
            y="Scheduled Visits",
            text="Scheduled Visits",
            color="Day",
            color_discrete_sequence=["#64748B", "#F59E0B", "#6366F1"],
            height=300,
        )
        fig_comp.update_layout(
            margin=dict(l=20, r=20, t=30, b=20),
            showlegend=False,
            yaxis_title="Scheduled Visits",
            plot_bgcolor="rgba(248,250,252,0.5)",
        )
        st.plotly_chart(fig_comp, use_container_width=True)

    st.markdown("---")

    # -------------------------------------------------------------------------
    # NEW DEDICATED SECTION: Baby-Wise Visit Details & Direct Exports
    # -------------------------------------------------------------------------
    st.markdown(
        '<div class="tier-header">📋 Baby-Wise Detailed Visit Rosters & Direct Exports</div>',
        unsafe_allow_html=True,
    )

    cols_roster_display = [
        "Baby Display",
        "Baby ID",
        "Mother",
        "ASHA",
        "District",
        "PHC",
        "Subcentre",
        "Visit Stage",
        "Due Date",
    ]

    # --- Section A: Yesterday's Missed Visits ---
    st.markdown(
        f'<div class="sub-section-header">A. Yesterday\'s Missed Visits ({yesterday_date.strftime("%d-%b-%Y")})</div>',
        unsafe_allow_html=True,
    )
    df_yest_missed_roster = v_yesterday[
        v_yesterday["Visit Status"] == "Missed"
    ].copy()
    y_missed_total_visits = len(df_yest_missed_roster)
    y_missed_unique_babies = (
        df_yest_missed_roster["Baby ID"].nunique()
        if not df_yest_missed_roster.empty
        else 0
    )

    ym_info_col, ym_btn_col1, ym_btn_col2 = st.columns([2.2, 0.9, 0.9])
    with ym_info_col:
        st.markdown(
            f"**Total Missed Visits:** `{y_missed_total_visits}` &nbsp;|&nbsp; **Affected Unique Babies:** `{y_missed_unique_babies}`"
        )
    with ym_btn_col1:
        if not df_yest_missed_roster.empty:
            ym_excel_bytes = generate_single_table_excel(
                df_yest_missed_roster[cols_roster_display].rename(
                    columns={"Baby Display": "Baby Name", "PHC": "Block/PHC"}
                ),
                sheet_name="Yesterday_Missed",
            )
            st.download_button(
                label="📥 Excel (.xlsx)",
                data=ym_excel_bytes,
                file_name=f"Yesterday_Missed_Visits_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="btn_ym_excel",
                use_container_width=True,
            )
    with ym_btn_col2:
        if not df_yest_missed_roster.empty:
            ym_pdf_bytes = generate_pdf_report(
                df_export=df_yest_missed_roster.rename(
                    columns={"Baby Display": "Baby Name", "PHC": "Block/PHC"}
                ),
                columns_to_show=[
                    "Baby Name",
                    "Baby ID",
                    "Mother",
                    "ASHA",
                    "Block/PHC",
                    "Subcentre",
                    "Visit Stage",
                    "Due Date",
                ],
                report_title=f"HBNC Yesterday's Missed Visits ({yesterday_date.strftime('%d-%b-%Y')})",
                active_filters_desc=active_filters_str,
            )
            st.download_button(
                label="📄 PDF (.pdf)",
                data=ym_pdf_bytes,
                file_name=f"Yesterday_Missed_Visits_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf",
                mime="application/pdf",
                key="btn_ym_pdf",
                use_container_width=True,
            )

    if not df_yest_missed_roster.empty:
        st.dataframe(
            df_yest_missed_roster[cols_roster_display].rename(
                columns={
                    "Baby Display": "Baby Name",
                    "PHC": "Block / PHC Facility",
                    "Due Date": "Scheduled Date",
                }
            ),
            use_container_width=True,
            hide_index=True,
        )
    else:
        st.success(
            f"🎉 No visits were missed yesterday ({yesterday_date.strftime('%d-%b-%Y')}) under the current filter selection."
        )

    st.markdown("---")

    # --- Section B: Today's Scheduled Visits ---
    st.markdown(
        f'<div class="sub-section-header">B. Today\'s Scheduled Visits ({today_date.strftime("%d-%b-%Y")})</div>',
        unsafe_allow_html=True,
    )
    df_today_sched_roster = v_today.copy()
    today_sched_total_visits = len(df_today_sched_roster)
    today_sched_unique_babies = (
        df_today_sched_roster["Baby ID"].nunique()
        if not df_today_sched_roster.empty
        else 0
    )

    t_info_col, t_btn_col1, t_btn_col2 = st.columns([2.2, 0.9, 0.9])
    with t_info_col:
        st.markdown(
            f"**Total Scheduled Visits:** `{today_sched_total_visits}` &nbsp;|&nbsp; **Unique Babies Scheduled:** `{today_sched_unique_babies}`"
        )
    with t_btn_col1:
        if not df_today_sched_roster.empty:
            t_excel_bytes = generate_single_table_excel(
                df_today_sched_roster[cols_roster_display].rename(
                    columns={"Baby Display": "Baby Name", "PHC": "Block/PHC"}
                ),
                sheet_name="Today_Scheduled",
            )
            st.download_button(
                label="📥 Excel (.xlsx)",
                data=t_excel_bytes,
                file_name=f"Today_Scheduled_Visits_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="btn_t_excel",
                use_container_width=True,
            )
    with t_btn_col2:
        if not df_today_sched_roster.empty:
            t_pdf_bytes = generate_pdf_report(
                df_export=df_today_sched_roster.rename(
                    columns={"Baby Display": "Baby Name", "PHC": "Block/PHC"}
                ),
                columns_to_show=[
                    "Baby Name",
                    "Baby ID",
                    "Mother",
                    "ASHA",
                    "Block/PHC",
                    "Subcentre",
                    "Visit Stage",
                    "Due Date",
                ],
                report_title=f"HBNC Today's Scheduled Visits ({today_date.strftime('%d-%b-%Y')})",
                active_filters_desc=active_filters_str,
            )
            st.download_button(
                label="📄 PDF (.pdf)",
                data=t_pdf_bytes,
                file_name=f"Today_Scheduled_Visits_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf",
                mime="application/pdf",
                key="btn_t_pdf",
                use_container_width=True,
            )

    if not df_today_sched_roster.empty:
        st.dataframe(
            df_today_sched_roster[cols_roster_display].rename(
                columns={
                    "Baby Display": "Baby Name",
                    "PHC": "Block / PHC Facility",
                    "Due Date": "Scheduled Date",
                }
            ),
            use_container_width=True,
            hide_index=True,
        )
    else:
        st.info(
            f"ℹ️ No visits are scheduled for today ({today_date.strftime('%d-%b-%Y')}) under the current filter selection."
        )

    st.markdown("---")

    # --- Section C: Tomorrow's Scheduled Visits ---
    st.markdown(
        f'<div class="sub-section-header">C. Tomorrow\'s Scheduled Visits ({tomorrow_date.strftime("%d-%b-%Y")})</div>',
        unsafe_allow_html=True,
    )
    df_tom_sched_roster = v_tomorrow.copy()
    tom_sched_total_visits = len(df_tom_sched_roster)
    tom_sched_unique_babies = (
        df_tom_sched_roster["Baby ID"].nunique()
        if not df_tom_sched_roster.empty
        else 0
    )

    tom_info_col, tom_btn_col1, tom_btn_col2 = st.columns([2.2, 0.9, 0.9])
    with tom_info_col:
        st.markdown(
            f"**Total Scheduled Visits:** `{tom_sched_total_visits}` &nbsp;|&nbsp; **Unique Babies Expected:** `{tom_sched_unique_babies}`"
        )
    with tom_btn_col1:
        if not df_tom_sched_roster.empty:
            tom_excel_bytes = generate_single_table_excel(
                df_tom_sched_roster[cols_roster_display].rename(
                    columns={"Baby Display": "Baby Name", "PHC": "Block/PHC"}
                ),
                sheet_name="Tomorrow_Scheduled",
            )
            st.download_button(
                label="📥 Excel (.xlsx)",
                data=tom_excel_bytes,
                file_name=f"Tomorrow_Scheduled_Visits_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="btn_tom_excel",
                use_container_width=True,
            )
    with tom_btn_col2:
        if not df_tom_sched_roster.empty:
            tom_pdf_bytes = generate_pdf_report(
                df_export=df_tom_sched_roster.rename(
                    columns={"Baby Display": "Baby Name", "PHC": "Block/PHC"}
                ),
                columns_to_show=[
                    "Baby Name",
                    "Baby ID",
                    "Mother",
                    "ASHA",
                    "Block/PHC",
                    "Subcentre",
                    "Visit Stage",
                    "Due Date",
                ],
                report_title=f"HBNC Tomorrow's Scheduled Visits ({tomorrow_date.strftime('%d-%b-%Y')})",
                active_filters_desc=active_filters_str,
            )
            st.download_button(
                label="📄 PDF (.pdf)",
                data=tom_pdf_bytes,
                file_name=f"Tomorrow_Scheduled_Visits_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf",
                mime="application/pdf",
                key="btn_tom_pdf",
                use_container_width=True,
            )

    if not df_tom_sched_roster.empty:
        st.dataframe(
            df_tom_sched_roster[cols_roster_display].rename(
                columns={
                    "Baby Display": "Baby Name",
                    "PHC": "Block / PHC Facility",
                    "Due Date": "Scheduled Date",
                }
            ),
            use_container_width=True,
            hide_index=True,
        )
    else:
        st.info(
            f"ℹ️ No visits are scheduled for tomorrow ({tomorrow_date.strftime('%d-%b-%Y')}) under the current filter selection."
        )

# -----------------------------------------------------------------------------
# TAB 2: ASHA & Healthcare Facility Overview
# -----------------------------------------------------------------------------
with tab_network:
    st.subheader(
        "👩‍⚕️ Healthcare Facility, ASHA Registrations & Operational Overview"
    )

    if not df_b_filtered.empty:
        asha_summary = (
            df_b_filtered.groupby("ASHA")
            .agg(
                District=("District", lambda x: x.iloc[0]),
                Facility_PHC=("PHC", lambda x: ", ".join(sorted(x.unique()))),
                Subcentres=(
                    "Subcentre",
                    lambda x: ", ".join(sorted(x.unique())),
                ),
                Registered_Babies=("Baby ID", "count"),
                Cumulative_Missed_Visits=("Missed Visits Count", "sum"),
            )
            .reset_index()
            .sort_values(
                by=["Registered_Babies", "Cumulative_Missed_Visits"],
                ascending=[False, False],
            )
        )

        st.markdown("**1. Comprehensive ASHA Master Workload & Performance Table**")
        st.dataframe(
            asha_summary.rename(
                columns={
                    "Facility_PHC": "Facility (PHC/CHC)",
                    "Registered_Babies": "Registered Babies",
                    "Cumulative_Missed_Visits": "Cumulative Missed Visits (Exact Sum)",
                }
            ),
            use_container_width=True,
            hide_index=True,
        )

        st.markdown("---")
        st.markdown(
            "**2. Interactive ASHA Drill-Down: Inspect Assigned Infants**"
        )

        selected_drill_asha = st.selectbox(
            "Select an ASHA worker to view her registered babies and visit schedules:",
            options=sorted(df_b_filtered["ASHA"].unique().tolist()),
            key="network_drill_asha",
        )

        df_drill_infants = df_b_filtered[
            df_b_filtered["ASHA"] == selected_drill_asha
        ].copy()
        st.caption(
            f"Displaying **{len(df_drill_infants)}** baby record(s) assigned to **{selected_drill_asha}**."
        )

        st.dataframe(
            df_drill_infants[[
                "Baby Display",
                "Baby ID",
                "Mother",
                "District",
                "PHC",
                "Subcentre",
                "DOB",
                "Registration Date",
                "Last Visit Date",
                "Missed Visits Count",
            ]]
            .rename(
                columns={
                    "Baby Display": "Baby Name",
                    "Missed Visits Count": "Missed Visits (Exact Sum)",
                }
            )
            .sort_values(by="Missed Visits (Exact Sum)", ascending=False),
            use_container_width=True,
            hide_index=True,
        )

        st.markdown("---")
        col_net1, col_net2 = st.columns(2)

        with col_net1:
            st.markdown("**3. Facility-Wise (PHC/CHC) Registration Summary**")
            phc_reg_summary = (
                df_b_filtered.groupby("PHC")
                .agg(
                    Total_ASHAs=("ASHA", "nunique"),
                    Total_Subcentres=("Subcentre", "nunique"),
                    Registered_Babies=("Baby ID", "count"),
                    Total_Missed_Visits=("Missed Visits Count", "sum"),
                )
                .reset_index()
            )
            st.dataframe(
                phc_reg_summary.rename(
                    columns={
                        "Total_ASHAs": "ASHAs",
                        "Total_Subcentres": "Subcentres",
                        "Registered_Babies": "Babies",
                        "Total_Missed_Visits": "Missed Visits (Sum)",
                    }
                ),
                use_container_width=True,
                hide_index=True,
            )

        with col_net2:
            st.markdown("**4. Subcentre-Wise Baby Distribution**")
            sub_counts = df_b_filtered["Subcentre"].value_counts().reset_index()
            sub_counts.columns = ["Subcentre", "Registered Babies"]
            fig_sub = px.pie(
                sub_counts,
                names="Subcentre",
                values="Registered Babies",
                hole=0.45,
                color_discrete_sequence=px.colors.qualitative.Safe,
                height=260,
            )
            fig_sub.update_layout(margin=dict(l=10, r=10, t=10, b=10))
            st.plotly_chart(fig_sub, use_container_width=True)
    else:
        st.warning("No records match the current filter selection.")

# -----------------------------------------------------------------------------
# TAB 3: Baby Master & Accurate Missed Visits Ledger
# -----------------------------------------------------------------------------
with tab_baby_master:
    st.subheader(
        "👶 Individual Baby Master List & Accurate Missed Visit Breakdown"
    )

    st.markdown(
        """
        > **Data Accuracy Rule**:
        > Missed visits are calculated at the **individual visit milestone level**, not by merely counting babies.
        > For example, if Baby A missed 2 visits and Baby B missed 1 visit, the cumulative missed visit total is **$2 + 1 = 3$**, avoiding undercounting and preserving operational fidelity.
        """
    )

    baby_master_cols = [
        "Baby Display",
        "Baby ID",
        "Mother",
        "ASHA",
        "District",
        "PHC",
        "Subcentre",
        "DOB",
        "Registration Date",
        "Last Visit Date",
        "Missed Visits Count",
    ]

    st.dataframe(
        df_b_filtered[baby_master_cols]
        .rename(
            columns={
                "Baby Display": "Baby Name",
                "Missed Visits Count": "Missed Visits (Exact Sum)",
            }
        )
        .sort_values(
            by=["Missed Visits (Exact Sum)", "Baby Name"],
            ascending=[False, True],
        ),
        use_container_width=True,
        hide_index=True,
    )

# -----------------------------------------------------------------------------
# TAB 4: Facility & Operational Trends
# -----------------------------------------------------------------------------
with tab_trends:
    st.subheader("📈 Operational Trends & Multi-Tier Facility Performance")

    tr_c1, tr_c2 = st.columns(2)

    with tr_c1:
        st.markdown("**1. Scheduled Milestone Due Date Timeline Trend**")
        trend_timeline = (
            df_v_filtered.groupby(["Due Date", "Visit Status"])
            .size()
            .unstack(fill_value=0)
            .reset_index()
        )
        trend_timeline["Due Date"] = pd.to_datetime(trend_timeline["Due Date"])
        trend_timeline = trend_timeline.sort_values("Due Date")

        fig_trend = go.Figure()
        for status_label, color in [
            ("Completed", "#10B981"),
            ("Missed", "#EF4444"),
            ("Upcoming", "#6366F1"),
        ]:
            if status_label in trend_timeline.columns:
                fig_trend.add_trace(
                    go.Scatter(
                        x=trend_timeline["Due Date"],
                        y=trend_timeline[status_label],
                        mode="lines+markers",
                        name=status_label,
                        line=dict(color=color, width=2),
                    )
                )

        fig_trend.update_layout(
            height=320,
            margin=dict(l=20, r=20, t=25, b=20),
            xaxis_title="Milestone Due Date",
            yaxis_title="Total Visits",
            legend=dict(
                orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1
            ),
            plot_bgcolor="rgba(248,250,252,0.5)",
        )
        st.plotly_chart(fig_trend, use_container_width=True)

    with tr_c2:
        st.markdown("**2. Facility / Block Workload Distribution**")
        phc_breakdown = (
            df_v_filtered.groupby(["PHC", "Visit Status"])
            .size()
            .unstack(fill_value=0)
            .reset_index()
        )
        fig_phc_chart = px.bar(
            phc_breakdown,
            x="PHC",
            y=[
                c
                for c in ["Completed", "Missed", "Upcoming"]
                if c in phc_breakdown.columns
            ],
            barmode="stack",
            color_discrete_map={
                "Completed": "#10B981",
                "Missed": "#EF4444",
                "Upcoming": "#6366F1",
            },
            height=320,
        )
        fig_phc_chart.update_layout(
            margin=dict(l=20, r=20, t=25, b=20),
            legend=dict(
                orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1
            ),
            xaxis_title="Facility / Block (PHC/CHC)",
            yaxis_title="Total Visits",
            plot_bgcolor="rgba(248,250,252,0.5)",
        )
        st.plotly_chart(fig_phc_chart, use_container_width=True)

# -----------------------------------------------------------------------------
# TAB 5: Filter-Matched Data Export (Comprehensive Excel & PDF)
# -----------------------------------------------------------------------------
with tab_export:
    st.subheader("📥 Export Filter-Matched Analytical Reports")
    st.markdown(
        f"""
        All generated reports will include **strictly only records matching the currently active filters**:
        * **Active Filter Scope:** `{active_filters_str}`
        """
    )

    d_col1, d_col2 = st.columns(2)

    with d_col1:
        st.markdown("#### 📗 Multi-Sheet Excel Workbook (.xlsx)")
        st.caption(
            "Contains Filtered Baby Master, ASHA Workload Summary, and Yesterday/Today/Tomorrow visit rosters."
        )

        export_sheets = {
            "Baby_Master_Filtered": df_b_filtered[baby_master_cols].rename(
                columns={"Baby Display": "Baby Name"}
            ),
            "ASHA_Workload_Summary": asha_summary
            if not df_b_filtered.empty
            else pd.DataFrame(),
            "Yesterday_Missed": df_yest_missed_roster[cols_roster_display].rename(
                columns={"Baby Display": "Baby Name", "PHC": "Block/PHC"}
            ),
            "Today_Scheduled": df_today_sched_roster[cols_roster_display].rename(
                columns={"Baby Display": "Baby Name", "PHC": "Block/PHC"}
            ),
            "Tomorrow_Scheduled": df_tom_sched_roster[cols_roster_display].rename(
                columns={"Baby Display": "Baby Name", "PHC": "Block/PHC"}
            ),
            "All_Filtered_Visits": df_v_filtered[[
                "Due Date",
                "Baby ID",
                "Baby Display",
                "Mother",
                "ASHA",
                "PHC",
                "Subcentre",
                "Visit Stage",
                "Visit Status",
            ]].rename(columns={"Baby Display": "Baby Name"}),
        }

        xlsx_bytes = generate_excel_workbook(export_sheets)
        timestamp_slug = datetime.now().strftime("%Y%m%d_%H%M%S")

        st.download_button(
            label="📥 Download Comprehensive Filtered Excel (.xlsx)",
            data=xlsx_bytes,
            file_name=f"HBNC_Filtered_Report_{timestamp_slug}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    with d_col2:
        st.markdown("#### 📕 Landscape Printable PDF Report (.pdf)")
        st.caption(
            "Clean landscape table formatted for official field circulation and administrative review."
        )

        pdf_view_choice = st.radio(
            "Select Report Content to Export:",
            [
                "Yesterday's Missed Visits",
                "Today's Scheduled Visits",
                "Tomorrow's Scheduled Visits",
                "ASHA Workload & Registration Master Ledger",
                "Baby Master with Individual Missed Visits",
                "Comprehensive Scheduled Visits Roster",
            ],
            index=0,
        )

        if pdf_view_choice == "Yesterday's Missed Visits":
            pdf_source_df = df_yest_missed_roster.rename(
                columns={"Baby Display": "Baby Name", "PHC": "Block/PHC"}
            )
            cols_to_print = [
                "Baby Name",
                "Baby ID",
                "Mother",
                "ASHA",
                "Block/PHC",
                "Subcentre",
                "Visit Stage",
                "Due Date",
            ]
            pdf_title = f"HBNC Yesterday's Missed Visits ({yesterday_date.strftime('%d-%b-%Y')})"
        elif pdf_view_choice == "Today's Scheduled Visits":
            pdf_source_df = df_today_sched_roster.rename(
                columns={"Baby Display": "Baby Name", "PHC": "Block/PHC"}
            )
            cols_to_print = [
                "Baby Name",
                "Baby ID",
                "Mother",
                "ASHA",
                "Block/PHC",
                "Subcentre",
                "Visit Stage",
                "Due Date",
            ]
            pdf_title = f"HBNC Today's Scheduled Visits ({today_date.strftime('%d-%b-%Y')})"
        elif pdf_view_choice == "Tomorrow's Scheduled Visits":
            pdf_source_df = df_tom_sched_roster.rename(
                columns={"Baby Display": "Baby Name", "PHC": "Block/PHC"}
            )
            cols_to_print = [
                "Baby Name",
                "Baby ID",
                "Mother",
                "ASHA",
                "Block/PHC",
                "Subcentre",
                "Visit Stage",
                "Due Date",
            ]
            pdf_title = f"HBNC Tomorrow's Scheduled Visits ({tomorrow_date.strftime('%d-%b-%Y')})"
        elif pdf_view_choice == "ASHA Workload & Registration Master Ledger":
            pdf_source_df = asha_summary
            cols_to_print = [
                "ASHA",
                "Facility_PHC",
                "Subcentres",
                "Registered_Babies",
                "Cumulative_Missed_Visits",
            ]
            pdf_title = "HBNC ASHA Master Registration & Workload Ledger"
        elif pdf_view_choice == "Baby Master with Individual Missed Visits":
            pdf_source_df = df_b_filtered
            cols_to_print = [
                "Baby Display",
                "Mother",
                "ASHA",
                "PHC",
                "Subcentre",
                "DOB",
                "Missed Visits Count",
            ]
            pdf_title = "HBNC Baby Master & Individual Missed Visits Ledger"
        else:
            pdf_source_df = df_v_filtered
            cols_to_print = [
                "Due Date",
                "Baby Display",
                "Mother",
                "ASHA",
                "PHC",
                "Visit Stage",
                "Visit Status",
            ]
            pdf_title = "HBNC Comprehensive Scheduled Visits Master Roster"

        pdf_bytes = generate_pdf_report(
            df_export=pdf_source_df,
            columns_to_show=cols_to_print,
            report_title=pdf_title,
            active_filters_desc=active_filters_str,
        )

        st.download_button(
            label="📄 Download Landscape PDF (.pdf)",
            data=pdf_bytes,
            file_name=f"HBNC_{pdf_view_choice[:14].replace(' ', '_')}_{timestamp_slug}.pdf",
            mime="application/pdf",
            use_container_width=True,
        )