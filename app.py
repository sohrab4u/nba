"""
ASHA & Infant Visit Due Monitoring, Healthcare Network & Video Quality Portal
Enterprise Health Analytics Dashboard for Home Based Newborn Care (HBNC)
Updated: Added Baby DOB alongside Baby Name across all Reports (Preview, Excel & PDF)
"""

import os
from datetime import date, datetime, timedelta
from io import BytesIO
import numpy as np
import pandas as pd

# Safe import for Plotly
try:
    import plotly.express as px
    import plotly.graph_objects as go
    HAS_PLOTLY = True
except ModuleNotFoundError:
    HAS_PLOTLY = False

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
# 1. Page Configuration & Professional Compact CSS
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
    /* Adjust Main Container Padding so Heading is NOT hidden */
    .block-container {
        padding-top: 3.5rem !important;
        padding-bottom: 1.5rem !important;
        padding-left: 1.5rem !important;
        padding-right: 1.5rem !important;
        max-width: 100% !important;
    }
    
    /* Top Header Section */
    .header-box {
        display: flex;
        justify-content: space-between;
        align-items: center;
        background: #FFFFFF;
        padding: 12px 18px;
        border-radius: 8px;
        border: 1px solid #E2E8F0;
        margin-bottom: 12px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.04);
    }
    .main-title {
        font-size: 20px;
        font-weight: 800;
        color: #1E3A8A;
        letter-spacing: -0.3px;
        margin: 0;
    }
    .scope-badge {
        font-size: 11px;
        color: #334155;
        background: #F8FAFC;
        padding: 5px 12px;
        border-radius: 6px;
        border: 1px solid #CBD5E1;
        font-weight: 500;
    }
    
    /* Section Headings */
    .tier-header {
        font-size: 13px;
        font-weight: 700;
        color: #1E293B;
        margin-top: 10px;
        margin-bottom: 6px;
        border-left: 3.5px solid #2563EB;
        padding-left: 8px;
        text-transform: uppercase;
        letter-spacing: 0.4px;
    }
    .sub-section-header {
        font-size: 14px;
        font-weight: 700;
        color: #0F172A;
        margin-top: 10px;
        margin-bottom: 4px;
    }

    /* Optimized KPI Cards */
    .kpi-card {
        background: #FFFFFF;
        border: 1px solid #E2E8F0;
        border-radius: 8px;
        padding: 8px 10px;
        box-shadow: 0 1px 2px rgba(0, 0, 0, 0.03);
        height: 78px;
        display: flex;
        flex-direction: column;
        justify-content: center;
    }
    .kpi-title {
        font-size: 10px;
        font-weight: 700;
        color: #64748B;
        text-transform: uppercase;
        letter-spacing: 0.3px;
        white-space: nowrap;
        overflow: hidden;
        text-overflow: ellipsis;
    }
    .kpi-value {
        font-size: 19px;
        font-weight: 800;
        margin-top: 1px;
        line-height: 1.1;
    }
    .kpi-caption {
        font-size: 9px;
        color: #94A3B8;
        margin-top: 2px;
        white-space: nowrap;
        overflow: hidden;
        text-overflow: ellipsis;
    }
    
    /* Sidebar Spacing Optimization */
    section[data-testid="stSidebar"] {
        width: 310px !important;
    }
    section[data-testid="stSidebar"] .block-container {
        padding-top: 1.5rem !important;
        padding-left: 0.8rem !important;
        padding-right: 0.8rem !important;
    }
    section[data-testid="stSidebar"] div[data-testid="stVerticalBlock"] {
        gap: 0.45rem !important;
    }
    
    /* Compact Tabs */
    .stTabs [data-baseweb="tab-list"] {
        gap: 6px;
    }
    .stTabs [data-baseweb="tab"] {
        padding: 5px 12px;
        font-size: 13px;
        font-weight: 600;
        border-radius: 6px 6px 0 0;
    }
</style>
""",
    unsafe_allow_html=True,
)


# -----------------------------------------------------------------------------
# 2. Export Utility Functions (PDF & Multi-Sheet Excel)
# -----------------------------------------------------------------------------
def generate_pdf_report(df_export, columns_to_show, report_title, active_filters_desc):
    """
    Generates an executive landscape PDF report matching active filters.
    Baby ID is strictly excluded from PDF, while Baby Name and DOB are preserved.
    """
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
        fontSize=13.5,
        leading=16.5,
        textColor=colors.HexColor("#1E3A8A"),
        fontName="Helvetica-Bold",
    )
    meta_style = ParagraphStyle(
        name="DocMeta",
        fontSize=8.5,
        leading=11,
        textColor=colors.HexColor("#475569"),
        fontName="Helvetica",
    )

    elements.append(Paragraph(report_title, title_style))
    elements.append(Spacer(1, 3))
    elements.append(
        Paragraph(
            f"<b>Filter Scope:</b> {active_filters_desc} &nbsp;|&nbsp; <b>Exported:</b> {datetime.now().strftime('%d-%b-%Y %I:%M %p')} &nbsp;|&nbsp; <b>Total Records:</b> {len(df_export)}",
            meta_style,
        )
    )
    elements.append(Spacer(1, 5))
    elements.append(
        HRFlowable(
            width="100%",
            thickness=1.5,
            color=colors.HexColor("#1E3A8A"),
            spaceAfter=8,
        )
    )

    # Strictly filter out 'Baby ID' from PDF download
    valid_cols = [
        c for c in columns_to_show 
        if c in df_export.columns and str(c).strip().lower() not in ["baby id", "baby_id", "babyid", "child id", "id"]
    ]
    df_clean = df_export[valid_cols].copy()

    for col in df_clean.columns:
        if pd.api.types.is_datetime64_any_dtype(df_clean[col]):
            df_clean[col] = df_clean[col].dt.strftime("%d-%b-%Y")
        else:
            df_clean[col] = (
                df_clean[col].astype(str).replace({"nan": "-", "NaT": "-", "None": "-"})
            )

    data = [valid_cols] + df_clean.values.tolist()
    cell_style = ParagraphStyle(
        name="CellTxt",
        fontSize=7.5,
        leading=9.5,
        fontName="Helvetica",
        textColor=colors.HexColor("#1E293B"),
    )
    head_style = ParagraphStyle(
        name="HeadTxt",
        fontSize=8.5,
        leading=10.5,
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
            ("BOTTOMPADDING", (0, 0), (-1, 0), 5),
            ("TOPPADDING", (0, 0), (-1, 0), 5),
            ("BOTTOMPADDING", (0, 1), (-1, -1), 3.5),
            ("TOPPADDING", (0, 1), (-1, -1), 3.5),
            ("LEFTPADDING", (0, 0), (-1, -1), 5),
            ("RIGHTPADDING", (0, 0), (-1, -1), 5),
            ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#CBD5E1")),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F8FAFC")]),
        ])
    )

    elements.append(pdf_table)
    doc.build(elements)
    buffer.seek(0)
    return buffer.getvalue()


def generate_single_table_excel(df_table, sheet_name="Data_Export"):
    """Generates an Excel file for a single dataframe preserving all columns."""
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df_table.to_excel(writer, sheet_name=sheet_name[:30], index=False)
    buffer.seek(0)
    return buffer.getvalue()


def generate_excel_workbook(sheets_dict):
    """Generates an Excel workbook containing multiple analytical sheets preserving all columns."""
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        for sheet_name, df_sheet in sheets_dict.items():
            if not df_sheet.empty:
                df_sheet.to_excel(writer, sheet_name=sheet_name[:30], index=False)
    buffer.seek(0)
    return buffer.getvalue()


# -----------------------------------------------------------------------------
# 3. Sidebar: Compact File Ingestion Panels
# -----------------------------------------------------------------------------
st.sidebar.markdown("### 📁 Data Input Center")

# File 1: Calendar Upload Box
with st.sidebar.expander("📄 1. HBNC Calendar Sheet", expanded=True):
    uploaded_calendar = st.file_uploader(
        "Upload Calendar (.xlsx, .csv)",
        type=["xlsx", "xls", "csv"],
        key="upload_cal",
        label_visibility="collapsed",
    )

# File 2: Feedback Upload Box
with st.sidebar.expander("📹 2. Video Feedback Sheet", expanded=True):
    uploaded_feedback = st.file_uploader(
        "Upload Feedback (.csv, .xlsx)",
        type=["csv", "xlsx", "xls"],
        key="upload_fb",
        label_visibility="collapsed",
    )

# --- Load Calendar File ---
df_raw = None
if uploaded_calendar is not None:
    try:
        if uploaded_calendar.name.endswith(".csv"):
            df_raw = pd.read_csv(uploaded_calendar)
        else:
            df_raw = pd.read_excel(uploaded_calendar)
    except Exception as e:
        st.error(f"Error reading calendar file: {e}")
        st.stop()
else:
    candidate_cals = [
        "Visit_Due_date_Calen_1791782047748.xlsx",
        "Visit_Due_date_Calen.xlsx",
    ]
    for cfile in candidate_cals:
        if os.path.exists(cfile):
            try:
                df_raw = pd.read_excel(cfile)
                break
            except Exception:
                continue

if df_raw is None or not isinstance(df_raw, pd.DataFrame):
    st.info("👋 **Welcome!** Please upload your HBNC Visit Calendar (.xlsx or .csv) from the sidebar to activate the dashboard.")
    st.stop()

# --- Load Feedback Sheet ---
df_fb_raw = None
if uploaded_feedback is not None:
    try:
        if uploaded_feedback.name.endswith(".csv"):
            df_fb_raw = pd.read_csv(uploaded_feedback)
        else:
            df_fb_raw = pd.read_excel(uploaded_feedback)
    except Exception as e:
        st.sidebar.warning(f"Could not read feedback file: {e}")
else:
    candidate_fbs = [
        "Feedback_Sheet_1791967524700_2.csv",
        "Feedback_Sheet_1791967524700.csv",
        "Feedback_Sheet.csv",
    ]
    for fbfile in candidate_fbs:
        if os.path.exists(fbfile):
            try:
                df_fb_raw = pd.read_csv(fbfile)
                break
            except Exception:
                continue

# -----------------------------------------------------------------------------
# 4. Normalization & Calendar Data Preprocessing
# -----------------------------------------------------------------------------
col_rename = {}
for c in df_raw.columns:
    clean = str(c).strip().lower()
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

due_cols = sorted([c for c in df_raw.columns if "visit due" in str(c).lower()])
if not due_cols:
    st.error("Invalid File Format: No 'Visit Due' milestone columns detected.")
    st.stop()

for col in due_cols:
    df_raw[col] = pd.to_datetime(df_raw[col], errors="coerce").dt.date

for dc in ["Registration Date", "DOB", "Last Visit Date"]:
    if dc in df_raw.columns:
        df_raw[dc] = pd.to_datetime(df_raw[dc], errors="coerce").dt.date

# -----------------------------------------------------------------------------
# 5. Feedback Preprocessing: VISIT_ID (COLUMN E) DEDUPLICATION & DOB MAPPING
# -----------------------------------------------------------------------------
df_fb_cleaned = None
has_feedback = False
fb_raw_total_rows = 0
fb_unique_visits_count = 0

# Mapping Mother Name to Baby Name and DOB from Calendar Sheet
mother_to_baby_map = dict(zip(
    df_raw["Mother"].astype(str).str.strip().str.lower(),
    df_raw["Baby Display"]
))

mother_to_dob_map = dict(zip(
    df_raw["Mother"].astype(str).str.strip().str.lower(),
    df_raw["DOB"]
))

if df_fb_raw is not None and isinstance(df_fb_raw, pd.DataFrame) and not df_fb_raw.empty:
    has_feedback = True
    df_fb = df_fb_raw.copy()
    fb_raw_total_rows = len(df_fb)
    
    fb_col_map = {}
    for idx_col, c in enumerate(df_fb.columns):
        c_clean = str(c).strip().lower()
        if "verdict" in c_clean or "annotation" in c_clean:
            fb_col_map[c] = "Annotation Verdict"
        elif "subcenter" in c_clean or "subcentre" in c_clean:
            fb_col_map[c] = "Subcentre"
        elif "visit_id" in c_clean or "visit id" in c_clean or idx_col == 4:
            fb_col_map[c] = "Visit_Id"
        elif "asha" in c_clean:
            fb_col_map[c] = "ASHA"
        elif "mother" in c_clean:
            fb_col_map[c] = "Mother"
        elif "district" in c_clean:
            fb_col_map[c] = "District"
        elif "visit date" in c_clean:
            fb_col_map[c] = "Visit Date"
        elif "sync date" in c_clean:
            fb_col_map[c] = "Sync Date"
    df_fb.rename(columns=fb_col_map, inplace=True)

    if "Visit_Id" in df_fb.columns:
        df_fb["Visit_Id"] = df_fb["Visit_Id"].astype(str).str.strip()
    elif len(df_fb.columns) > 4:
        df_fb["Visit_Id"] = df_fb.iloc[:, 4].astype(str).str.strip()

    if "Mother" in df_fb.columns:
        df_fb["Mother"] = df_fb["Mother"].astype(str).str.strip()
    elif len(df_fb.columns) > 5:
        df_fb["Mother"] = df_fb.iloc[:, 5].astype(str).str.strip()

    if "Visit Date" in df_fb.columns:
        df_fb["Visit Date Clean"] = pd.to_datetime(df_fb["Visit Date"], errors="coerce").dt.date
    elif len(df_fb.columns) > 6:
        df_fb["Visit Date Clean"] = pd.to_datetime(df_fb.iloc[:, 6], errors="coerce").dt.date
    else:
        df_fb["Visit Date Clean"] = None

    # Deduplicate strictly on Visit_Id (Column E)
    if "Visit_Id" in df_fb.columns:
        df_fb_cleaned = df_fb.drop_duplicates(subset=["Visit_Id"], keep="first").copy()
    else:
        df_fb_cleaned = df_fb.drop_duplicates().copy()

    if "Visit Date Clean" in df_fb_cleaned.columns:
        df_fb_cleaned["Visit Date"] = df_fb_cleaned["Visit Date Clean"]
        df_fb_cleaned.drop(columns=["Visit Date Clean"], inplace=True)

    # Attach Baby Name & DOB to Feedback Sheet
    df_fb_cleaned["Baby Name"] = df_fb_cleaned["Mother"].apply(
        lambda m: mother_to_baby_map.get(str(m).strip().lower(), f"Baby of {m}")
    )
    df_fb_cleaned["DOB"] = df_fb_cleaned["Mother"].apply(
        lambda m: mother_to_dob_map.get(str(m).strip().lower(), "-")
    )

    # Compile Comprehensive Rejected Feedback String
    def compile_rejected_reasons(row):
        reasons = []
        wv = str(row.get("Weight Video Feedback", ""))
        if any(w in wv.lower() for w in ["incomplete", "re-read", "not", "issue", "reading"]):
            reasons.append(f"Weight: {wv}")

        arc = str(row.get("Arc Feedback", ""))
        if any(w in arc.lower() for w in ["incomplete", "improvement", "needed", "angle"]):
            reasons.append(f"ARC: {arc}")

        hum = str(row.get("Other Human Feedback", ""))
        if ("another person" in hum.lower() or "present" in hum.lower()) and "no other" not in hum.lower():
            reasons.append(f"Human: {hum}")

        bc = str(row.get("Baby Cropping Feedback", ""))
        if "crop" in bc.lower() and "fully visible" not in bc.lower():
            reasons.append(f"Baby: {bc}")

        rc = str(row.get("Ruler Cropping Feedback", ""))
        if any(w in rc.lower() for w in ["not fully visible", "occlude", "crop"]) and "fully visible" not in rc.lower():
            reasons.append(f"Ruler: {rc}")

        return " | ".join(reasons) if reasons else "Quality criteria not met as per protocol"

    df_fb_cleaned["Rejected Feedback"] = df_fb_cleaned.apply(compile_rejected_reasons, axis=1)
    fb_unique_visits_count = len(df_fb_cleaned)
        
    for c in ["ASHA", "Mother", "District", "Subcentre"]:
        if c in df_fb_cleaned.columns:
            df_fb_cleaned[c] = df_fb_cleaned[c].astype(str).str.strip()

# -----------------------------------------------------------------------------
# 6. Sidebar: Date Range & Evaluation Filters (Two Separate Date Boxes)
# -----------------------------------------------------------------------------
st.sidebar.markdown("---")
st.sidebar.markdown("### 🔍 Filters & Date Range")

# 1. District Filter
dist_opts = ["All"] + sorted(df_raw["District"].unique().tolist())
sel_dist = st.sidebar.selectbox("District", dist_opts)
df_scope1 = df_raw if sel_dist == "All" else df_raw[df_raw["District"] == sel_dist]

# 2. CHC / PHC Facility
phc_opts = ["All"] + sorted(df_scope1["PHC"].unique().tolist())
sel_phc = st.sidebar.selectbox("CHC / PHC Facility", phc_opts)
df_scope2 = df_scope1 if sel_phc == "All" else df_scope1[df_scope1["PHC"] == sel_phc]

# 3. Subcentre
sub_opts = ["All"] + sorted(df_scope2["Subcentre"].unique().tolist())
sel_sub = st.sidebar.selectbox("Subcentre", sub_opts)
df_scope3 = df_scope2 if sel_sub == "All" else df_scope2[df_scope2["Subcentre"] == sel_sub]

# 4. ASHA Worker
asha_opts = ["All"] + sorted(df_scope3["ASHA"].unique().tolist())
sel_asha = st.sidebar.selectbox("ASHA Worker", asha_opts)

# Determine Overall Date Bound
all_calendar_dates = []
for c in due_cols:
    all_calendar_dates.extend(df_raw[c].dropna().tolist())
min_cal_date = min(all_calendar_dates) if all_calendar_dates else date(2026, 9, 1)
max_cal_date = max(all_calendar_dates) if all_calendar_dates else date(2026, 11, 30)

# Two separate boxes for Milestone Date Range
st.sidebar.markdown("**Milestone Due Date Filter**")
col_d1, col_d2 = st.sidebar.columns(2)
with col_d1:
    start_date_filter = st.date_input("Start Date", value=min_cal_date, min_value=min_cal_date, max_value=max_cal_date)
with col_d2:
    end_date_filter = st.date_input("End Date", value=max_cal_date, min_value=min_cal_date, max_value=max_cal_date)

if start_date_filter > end_date_filter:
    st.sidebar.warning("Start Date > End Date. Using full range.")
    start_date_filter, end_date_filter = min_cal_date, max_cal_date

# Manual Evaluation Date Option
use_custom_date = st.sidebar.checkbox(
    "Set Evaluation (Today) Date",
    value=False,
    help="Enable to test historical dates matching dataset timeframe (e.g., 12-Oct-2026).",
)

if use_custom_date:
    ref_date = st.sidebar.date_input("Reference Evaluation Date:", value=date(2026, 10, 12))
else:
    ref_date = date.today()

today_date = ref_date
yesterday_date = ref_date - timedelta(days=1)
day_before_yesterday_date = ref_date - timedelta(days=2)
tomorrow_date = ref_date + timedelta(days=1)

# -----------------------------------------------------------------------------
# 7. Visit Classification Engine
# -----------------------------------------------------------------------------
visit_records = []
for _, row in df_raw.iterrows():
    lvd = row["Last Visit Date"] if ("Last Visit Date" in row and pd.notna(row["Last Visit Date"])) else None
    reg_d = row["Registration Date"] if ("Registration Date" in row and pd.notna(row["Registration Date"])) else None
    dob_val = row["DOB"] if ("DOB" in row and pd.notna(row["DOB"])) else None

    for stage in due_cols:
        d_val = row[stage]
        if pd.notna(d_val):
            is_completed = bool(lvd and lvd >= d_val)
            after_reg = bool(is_completed and reg_d and d_val >= reg_d)

            if d_val < today_date:
                v_status = "Completed" if is_completed else "Missed"
            elif d_val == today_date:
                v_status = "Completed" if is_completed else "Pending"
            else:
                v_status = "Upcoming"

            visit_records.append({
                "Baby ID": row["Baby ID"],
                "Baby Display": row["Baby Display"],
                "Mother": row["Mother"],
                "ASHA": row["ASHA"],
                "District": row["District"],
                "PHC": row["PHC"],
                "Subcentre": row["Subcentre"],
                "DOB": dob_val,
                "Registration Date": reg_d,
                "Last Visit Date": lvd,
                "Visit Stage": stage,
                "Due Date": d_val,
                "Is Completed": is_completed,
                "Is After Registration": after_reg,
                "Visit Status": v_status,
                "Month Year": d_val.strftime("%B %Y"),
            })

df_all_visits = pd.DataFrame(visit_records)

baby_completed_map = df_all_visits[df_all_visits["Is Completed"] == True].groupby("Baby ID").size().to_dict()
baby_after_reg_map = df_all_visits[df_all_visits["Is After Registration"] == True].groupby("Baby ID").size().to_dict()
baby_missed_map = df_all_visits[df_all_visits["Visit Status"] == "Missed"].groupby("Baby ID").size().to_dict()

df_raw["Completed Visits Total"] = df_raw["Baby ID"].map(lambda bid: baby_completed_map.get(bid, 0))
df_raw["Visits After Reg"] = df_raw["Baby ID"].map(lambda bid: baby_after_reg_map.get(bid, 0))
df_raw["Missed Visits Count"] = df_raw["Baby ID"].map(lambda bid: baby_missed_map.get(bid, 0))
df_raw["Pending Visits Total"] = len(due_cols) - df_raw["Completed Visits Total"]

# Apply Active Filters
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

df_v_filtered = df_v_filtered[
    (df_v_filtered["Due Date"] >= start_date_filter)
    & (df_v_filtered["Due Date"] <= end_date_filter)
]

cohort_baby_ids = df_v_filtered["Baby ID"].unique()
df_b_filtered = df_b_filtered[df_b_filtered["Baby ID"].isin(cohort_baby_ids)]

# Apply filters synchronously to Deduplicated Feedback Sheet
df_fb_filtered = pd.DataFrame()
if has_feedback and df_fb_cleaned is not None:
    df_fb_filtered = df_fb_cleaned.copy()
    if sel_dist != "All" and "District" in df_fb_filtered.columns:
        df_fb_filtered = df_fb_filtered[df_fb_filtered["District"] == sel_dist]
    if sel_sub != "All" and "Subcentre" in df_fb_filtered.columns:
        df_fb_filtered = df_fb_filtered[df_fb_filtered["Subcentre"] == sel_sub]
    if sel_asha != "All" and "ASHA" in df_fb_filtered.columns:
        df_fb_filtered = df_fb_filtered[df_fb_filtered["ASHA"] == sel_asha]

active_filters_str = (
    f"District: {sel_dist} | CHC/PHC: {sel_phc} | Subcentre: {sel_sub} | "
    f"ASHA: {sel_asha} | Range: {start_date_filter.strftime('%d-%b')} to {end_date_filter.strftime('%d-%b-%Y')} | "
    f"Ref Date: {today_date.strftime('%d-%b-%Y')}"
)

# -----------------------------------------------------------------------------
# 8. Compute Dashboard Key Performance Indicators (KPIs)
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

kpi_total_babies = len(df_b_filtered)

# Total Visits Completed calculated from Deduplicated Feedback Sheet (Visit_Id based)
if has_feedback and not df_fb_filtered.empty:
    kpi_total_completed_visits = int(len(df_fb_filtered))
    completed_caption_text = f"Feedback Sheet ({fb_unique_visits_count} Single Visits)"
else:
    kpi_total_completed_visits = int(df_b_filtered["Completed Visits Total"].sum())
    completed_caption_text = "Completed up to last visit"

kpi_visits_after_reg = int(df_b_filtered["Visits After Reg"].sum())
kpi_total_pending_visits = int(df_b_filtered["Pending Visits Total"].sum())

# Deduplicated Video Verification KPIs (Visit_Id based)
if has_feedback and not df_fb_filtered.empty and "Annotation Verdict" in df_fb_filtered.columns:
    kpi_videos_accepted = int((df_fb_filtered["Annotation Verdict"].str.strip().str.lower() == "accepted").sum())
    kpi_videos_rejected = int((df_fb_filtered["Annotation Verdict"].str.strip().str.lower() == "rejected").sum())
    total_reviewed_unique = kpi_videos_accepted + kpi_videos_rejected
    video_acc_rate = f"{(kpi_videos_accepted / total_reviewed_unique * 100):.1f}%" if total_reviewed_unique > 0 else "0%"
else:
    kpi_videos_accepted = "N/A"
    kpi_videos_rejected = "N/A"
    video_acc_rate = "Upload Feedback File"

kpi_total_ashas = df_b_filtered["ASHA"].nunique()
kpi_total_phc = df_b_filtered["PHC"].nunique()
kpi_total_subcentres = df_b_filtered["Subcentre"].nunique()

# Daily Monitoring Cadence Subsets:
# 1. Day Before Yesterday (Parso)
v_day_before_yesterday = df_v_geo[df_v_geo["Due Date"] == day_before_yesterday_date]
kpi_dby_scheduled = len(v_day_before_yesterday)
kpi_dby_completed = (v_day_before_yesterday["Visit Status"] == "Completed").sum()
kpi_dby_missed = (v_day_before_yesterday["Visit Status"] == "Missed").sum()
kpi_dby_completion_rate = (
    (kpi_dby_completed / kpi_dby_scheduled * 100.0)
    if kpi_dby_scheduled > 0
    else 0.0
)

# 2. Yesterday
v_yesterday = df_v_geo[df_v_geo["Due Date"] == yesterday_date]
kpi_yest_scheduled = len(v_yesterday)
kpi_yest_completed = (v_yesterday["Visit Status"] == "Completed").sum()
kpi_yest_missed = (v_yesterday["Visit Status"] == "Missed").sum()
kpi_yest_completion_rate = (
    (kpi_yest_completed / kpi_yest_scheduled * 100.0)
    if kpi_yest_scheduled > 0
    else 0.0
)

# 3. Today & Tomorrow
v_today = df_v_geo[df_v_geo["Due Date"] == today_date]
v_tomorrow = df_v_geo[df_v_geo["Due Date"] == tomorrow_date]

kpi_today_scheduled = len(v_today)
kpi_tom_scheduled = len(v_tomorrow)
kpi_tom_babies_count = v_tomorrow["Baby ID"].nunique()

# -----------------------------------------------------------------------------
# 9. Modern Top Header & Two-Tier Compact KPI Cards
# -----------------------------------------------------------------------------
st.markdown(
    f"""
    <div class="header-box">
        <div class="main-title">👶 ASHA & Infant Visit Due Monitoring Dashboard</div>
        <div class="scope-badge">📌 {active_filters_str}</div>
    </div>
    """,
    unsafe_allow_html=True,
)

# --- Tier 1: Registration, Visits & Deduplicated Video Verification ---
st.markdown('<div class="tier-header">📊 Baby Registrations, Completed Visits & Video Verification (Visit_Id Deduplicated)</div>', unsafe_allow_html=True)
r1, r2, r3, r4, r5, r6 = st.columns(6)

with r1:
    st.markdown(f'<div class="kpi-card" style="border-left: 3.5px solid #1D4ED8;"><div class="kpi-title">Registered Babies</div><div class="kpi-value" style="color: #1D4ED8;">{kpi_total_babies:,}</div><div class="kpi-caption">Active cohort</div></div>', unsafe_allow_html=True)
with r2:
    st.markdown(f'<div class="kpi-card" style="border-left: 3.5px solid #059669;"><div class="kpi-title">Visits Completed</div><div class="kpi-value" style="color: #059669;">{kpi_total_completed_visits:,}</div><div class="kpi-caption">{completed_caption_text}</div></div>', unsafe_allow_html=True)
with r3:
    st.markdown(f'<div class="kpi-card" style="border-left: 3.5px solid #0D9488;"><div class="kpi-title">Post-Reg Visits</div><div class="kpi-value" style="color: #0D9488;">{kpi_visits_after_reg:,}</div><div class="kpi-caption">Done ≥ Reg Date</div></div>', unsafe_allow_html=True)
with r4:
    st.markdown(f'<div class="kpi-card" style="border-left: 3.5px solid #D97706;"><div class="kpi-title">Pending Visits</div><div class="kpi-value" style="color: #D97706;">{kpi_total_pending_visits:,}</div><div class="kpi-caption">Remaining cohort</div></div>', unsafe_allow_html=True)
with r5:
    st.markdown(f'<div class="kpi-card" style="border-left: 3.5px solid #10B981;"><div class="kpi-title">Videos Accepted</div><div class="kpi-value" style="color: #10B981;">{kpi_videos_accepted}</div><div class="kpi-caption">Single Visits | Rate: {video_acc_rate}</div></div>', unsafe_allow_html=True)
with r6:
    st.markdown(f'<div class="kpi-card" style="border-left: 3.5px solid #EF4444;"><div class="kpi-title">Videos Rejected</div><div class="kpi-value" style="color: #DC2626;">{kpi_videos_rejected}</div><div class="kpi-caption">Single Visits | Requires re-visit</div></div>', unsafe_allow_html=True)

# --- Tier 2: Healthcare Infrastructure & Daily Cadence ---
st.markdown('<div class="tier-header">⏱️ Operational Daily Schedule & Healthcare Infrastructure</div>', unsafe_allow_html=True)
d1, d2, d3, d4, d5, d6, d7, d8 = st.columns(8)

with d1:
    st.markdown(f'<div class="kpi-card" style="border-left: 3px solid #0D9488;"><div class="kpi-title">Total ASHAs</div><div class="kpi-value" style="font-size: 18px; color: #0F766E;">{kpi_total_ashas:,}</div><div class="kpi-caption">Facilities: {kpi_total_phc}</div></div>', unsafe_allow_html=True)
with d2:
    st.markdown(f'<div class="kpi-card" style="border-left: 3px solid #64748B;"><div class="kpi-title">Parso Sched.</div><div class="kpi-value" style="font-size: 18px; color: #334155;">{kpi_dby_scheduled:,}</div><div class="kpi-caption">{day_before_yesterday_date.strftime("%d-%b")}</div></div>', unsafe_allow_html=True)
with d3:
    st.markdown(f'<div class="kpi-card" style="border-left: 3px solid #EF4444;"><div class="kpi-title">Parso Missed</div><div class="kpi-value" style="font-size: 18px; color: #DC2626;">{kpi_dby_missed:,}</div><div class="kpi-caption">Done: {kpi_dby_completed} ({kpi_dby_completion_rate:.0f}%)</div></div>', unsafe_allow_html=True)
with d4:
    st.markdown(f'<div class="kpi-card" style="border-left: 3px solid #64748B;"><div class="kpi-title">Yest. Sched.</div><div class="kpi-value" style="font-size: 18px; color: #334155;">{kpi_yest_scheduled:,}</div><div class="kpi-caption">{yesterday_date.strftime("%d-%b")}</div></div>', unsafe_allow_html=True)
with d5:
    st.markdown(f'<div class="kpi-card" style="border-left: 3px solid #EF4444;"><div class="kpi-title">Yest. Missed</div><div class="kpi-value" style="font-size: 18px; color: #DC2626;">{kpi_yest_missed:,}</div><div class="kpi-caption">Done: {kpi_yest_completed} ({kpi_yest_completion_rate:.0f}%)</div></div>', unsafe_allow_html=True)
with d6:
    st.markdown(f'<div class="kpi-card" style="border-left: 3px solid #10B981;"><div class="kpi-title">Total Subcentres</div><div class="kpi-value" style="font-size: 18px; color: #059669;">{kpi_total_subcentres:,}</div><div class="kpi-caption">Subcentre coverage</div></div>', unsafe_allow_html=True)
with d7:
    st.markdown(f'<div class="kpi-card" style="border-left: 3px solid #F59E0B;"><div class="kpi-title">Today Sched.</div><div class="kpi-value" style="font-size: 18px; color: #D97706;">{kpi_today_scheduled:,}</div><div class="kpi-caption">{today_date.strftime("%d-%b")} (Due)</div></div>', unsafe_allow_html=True)
with d8:
    st.markdown(f'<div class="kpi-card" style="border-left: 3px solid #6366F1;"><div class="kpi-title">Tom. Sched.</div><div class="kpi-value" style="font-size: 18px; color: #4F46E5;">{kpi_tom_scheduled:,}</div><div class="kpi-caption">{tomorrow_date.strftime("%d-%b")} ({kpi_tom_babies_count})</div></div>', unsafe_allow_html=True)

st.markdown("<div style='height: 8px;'></div>", unsafe_allow_html=True)

# -----------------------------------------------------------------------------
# 10. Functional Tabs
# -----------------------------------------------------------------------------
tab_daily, tab_video, tab_reg_audit, tab_network, tab_baby_master, tab_trends, tab_export = st.tabs([
    "📊 Daily Visit Overview & Visualizations",
    "📹 Video Quality & Rejected Reports",
    "🔍 Registration & Visit Audit (Baby-Wise)",
    "👩‍⚕️ ASHA & Healthcare Facility Overview",
    "👶 Baby Master & Missed Visits Ledger",
    "📈 Facility & Operational Trends",
    "📥 Download Filtered Reports",
])

# Roster columns: Includes Baby Name and DOB
cols_roster_preview = ["Baby Display", "DOB", "Baby ID", "Mother", "ASHA", "District", "PHC", "Subcentre", "Visit Stage", "Due Date"]
cols_roster_pdf_base = ["Baby Name", "DOB", "Mother", "ASHA", "Block/PHC", "Subcentre", "Visit Stage", "Due Date"]

df_dby_all_roster = v_day_before_yesterday.copy()
df_yest_all_roster = v_yesterday.copy()
df_today_sched_roster = v_today.copy()
df_tom_sched_roster = v_tomorrow.copy()

# TAB 1: Daily Visit Monitoring
with tab_daily:
    st.subheader("⏱️ Daily Visit Overview: 4-Day Comparative Performance")
    row1_c1, row1_c2 = st.columns(2)

    with row1_c1:
        st.markdown(f"**1. Historical Daily Performance: Scheduled vs. Completed vs. Missed**")
        df_hist_chart = pd.DataFrame([
            {"Day": f"Parso ({day_before_yesterday_date.strftime('%d-%b')})", "Scheduled": kpi_dby_scheduled, "Completed": kpi_dby_completed, "Missed": kpi_dby_missed},
            {"Day": f"Yesterday ({yesterday_date.strftime('%d-%b')})", "Scheduled": kpi_yest_scheduled, "Completed": kpi_yest_completed, "Missed": kpi_yest_missed},
        ])
        if HAS_PLOTLY:
            fig_hist = go.Figure()
            fig_hist.add_trace(go.Bar(name="Scheduled", x=df_hist_chart["Day"], y=df_hist_chart["Scheduled"], marker_color="#3B82F6", text=df_hist_chart["Scheduled"], textposition="auto"))
            fig_hist.add_trace(go.Bar(name="Completed", x=df_hist_chart["Day"], y=df_hist_chart["Completed"], marker_color="#10B981", text=df_hist_chart["Completed"], textposition="auto"))
            fig_hist.add_trace(go.Bar(name="Missed", x=df_hist_chart["Day"], y=df_hist_chart["Missed"], marker_color="#EF4444", text=df_hist_chart["Missed"], textposition="auto"))
            fig_hist.update_layout(barmode="group", height=260, margin=dict(l=15, r=15, t=25, b=15), legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1), yaxis_title="Visits", plot_bgcolor="rgba(248,250,252,0.5)")
            st.plotly_chart(fig_hist, use_container_width=True)
        else:
            st.bar_chart(df_hist_chart.set_index("Day")[["Scheduled", "Completed", "Missed"]])

    with row1_c2:
        st.markdown("**2. Scheduled Milestone Comparison Across All 4 Days**")
        df_comp = pd.DataFrame({
            "Day": [
                f"Parso ({day_before_yesterday_date.strftime('%d-%b')})",
                f"Yesterday ({yesterday_date.strftime('%d-%b')})",
                f"Today ({today_date.strftime('%d-%b')})",
                f"Tomorrow ({tomorrow_date.strftime('%d-%b')})",
            ],
            "Scheduled Visits": [
                kpi_dby_scheduled,
                kpi_yest_scheduled,
                kpi_today_scheduled,
                kpi_tom_scheduled,
            ],
        })
        if HAS_PLOTLY:
            fig_comp = px.bar(
                df_comp,
                x="Day",
                y="Scheduled Visits",
                text="Scheduled Visits",
                color="Day",
                color_discrete_sequence=["#94A3B8", "#64748B", "#F59E0B", "#6366F1"],
                height=260,
            )
            fig_comp.update_layout(margin=dict(l=15, r=15, t=25, b=15), showlegend=False, yaxis_title="Visits", plot_bgcolor="rgba(248,250,252,0.5)")
            st.plotly_chart(fig_comp, use_container_width=True)
        else:
            st.bar_chart(df_comp.set_index("Day")["Scheduled Visits"])

    st.markdown("---")
    st.markdown('<div class="tier-header">📋 Baby-Wise Detailed Daily Rosters & Direct Exports (With Baby Name & DOB)</div>', unsafe_allow_html=True)

    # Section A: Day Before Yesterday's (Parso) Visits
    st.markdown(f'<div class="sub-section-header">A. Day Before Yesterday\'s (Parso) Visits ({day_before_yesterday_date.strftime("%d-%b-%Y")})</div>', unsafe_allow_html=True)
    dby_info_col, dby_btn_col1, dby_btn_col2 = st.columns([2.2, 0.9, 0.9])
    with dby_info_col:
        st.markdown(
            f"**Scheduled Visits:** `{kpi_dby_scheduled}` &nbsp;|&nbsp; "
            f"**Completed Visits:** `{kpi_dby_completed}` &nbsp;|&nbsp; "
            f"**Missed Visits:** `{kpi_dby_missed}`"
        )
    with dby_btn_col1:
        if not df_dby_all_roster.empty:
            dby_excel_bytes = generate_single_table_excel(
                df_dby_all_roster[cols_roster_preview + ["Visit Status"]].rename(
                    columns={"Baby Display": "Baby Name", "PHC": "Block/PHC", "Visit Status": "Outcome"}
                ),
                sheet_name="Day_Before_Yesterday",
            )
            st.download_button(label="📥 Excel (.xlsx)", data=dby_excel_bytes, file_name=f"Day_Before_Yesterday_Visits_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="btn_dby_excel", use_container_width=True)
    with dby_btn_col2:
        if not df_dby_all_roster.empty:
            dby_pdf_bytes = generate_pdf_report(
                df_export=df_dby_all_roster.rename(columns={"Baby Display": "Baby Name", "PHC": "Block/PHC", "Visit Status": "Outcome"}),
                columns_to_show=cols_roster_pdf_base + ["Outcome"],
                report_title=f"HBNC Day Before Yesterday Visits ({day_before_yesterday_date.strftime('%d-%b-%Y')})",
                active_filters_desc=active_filters_str,
            )
            st.download_button(label="📄 PDF (.pdf)", data=dby_pdf_bytes, file_name=f"Day_Before_Yesterday_Visits_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf", mime="application/pdf", key="btn_dby_pdf", use_container_width=True)

    if not df_dby_all_roster.empty:
        st.dataframe(
            df_dby_all_roster[cols_roster_preview + ["Visit Status"]].rename(
                columns={"Baby Display": "Baby Name", "PHC": "Block / PHC Facility", "Due Date": "Scheduled Date", "Visit Status": "Outcome"}
            ),
            use_container_width=True,
            hide_index=True,
        )
    else:
        st.info(f"ℹ️ No visits were scheduled on the day before yesterday ({day_before_yesterday_date.strftime('%d-%b-%Y')}) under the current filter selection.")

    st.markdown("---")

    # Section B: Yesterday's Visits with Outcome Column (Completed or Missed)
    st.markdown(f'<div class="sub-section-header">B. Yesterday\'s Visits Report ({yesterday_date.strftime("%d-%b-%Y")})</div>', unsafe_allow_html=True)
    yest_info_col, yest_btn_col1, yest_btn_col2 = st.columns([2.2, 0.9, 0.9])
    with yest_info_col:
        st.markdown(
            f"**Scheduled Visits:** `{kpi_yest_scheduled}` &nbsp;|&nbsp; "
            f"**Completed Visits:** `{kpi_yest_completed}` &nbsp;|&nbsp; "
            f"**Missed Visits:** `{kpi_yest_missed}`"
        )
    with yest_btn_col1:
        if not df_yest_all_roster.empty:
            ym_excel_bytes = generate_single_table_excel(
                df_yest_all_roster[cols_roster_preview + ["Visit Status"]].rename(
                    columns={"Baby Display": "Baby Name", "PHC": "Block/PHC", "Visit Status": "Outcome"}
                ),
                sheet_name="Yesterday_Visits",
            )
            st.download_button(label="📥 Excel (.xlsx)", data=ym_excel_bytes, file_name=f"Yesterday_Visits_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="btn_ym_excel", use_container_width=True)
    with yest_btn_col2:
        if not df_yest_all_roster.empty:
            ym_pdf_bytes = generate_pdf_report(
                df_export=df_yest_all_roster.rename(columns={"Baby Display": "Baby Name", "PHC": "Block/PHC", "Visit Status": "Outcome"}),
                columns_to_show=cols_roster_pdf_base + ["Outcome"],
                report_title=f"HBNC Yesterday's Visits & Outcomes ({yesterday_date.strftime('%d-%b-%Y')})",
                active_filters_desc=active_filters_str,
            )
            st.download_button(label="📄 PDF (.pdf)", data=ym_pdf_bytes, file_name=f"Yesterday_Visits_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf", mime="application/pdf", key="btn_ym_pdf", use_container_width=True)

    if not df_yest_all_roster.empty:
        st.dataframe(
            df_yest_all_roster[cols_roster_preview + ["Visit Status"]].rename(
                columns={"Baby Display": "Baby Name", "PHC": "Block / PHC Facility", "Due Date": "Scheduled Date", "Visit Status": "Outcome"}
            ),
            use_container_width=True,
            hide_index=True,
        )
    else:
        st.info(f"ℹ️ No visits were scheduled for yesterday ({yesterday_date.strftime('%d-%b-%Y')}) under the current filter selection.")

    st.markdown("---")

    # Section C: Today's Scheduled Visits
    st.markdown(f'<div class="sub-section-header">C. Today\'s Scheduled Visits ({today_date.strftime("%d-%b-%Y")})</div>', unsafe_allow_html=True)
    t_info_col, t_btn_col1, t_btn_col2 = st.columns([2.2, 0.9, 0.9])
    with t_info_col:
        st.markdown(f"**Total Scheduled Visits:** `{len(df_today_sched_roster)}` &nbsp;|&nbsp; **Unique Babies Scheduled:** `{df_today_sched_roster['Baby ID'].nunique() if not df_today_sched_roster.empty else 0}`")
    with t_btn_col1:
        if not df_today_sched_roster.empty:
            t_excel_bytes = generate_single_table_excel(
                df_today_sched_roster[cols_roster_preview].rename(columns={"Baby Display": "Baby Name", "PHC": "Block/PHC"}),
                sheet_name="Today_Scheduled",
            )
            st.download_button(label="📥 Excel (.xlsx)", data=t_excel_bytes, file_name=f"Today_Scheduled_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="btn_t_excel", use_container_width=True)
    with t_btn_col2:
        if not df_today_sched_roster.empty:
            t_pdf_bytes = generate_pdf_report(
                df_export=df_today_sched_roster.rename(columns={"Baby Display": "Baby Name", "PHC": "Block/PHC"}),
                columns_to_show=cols_roster_pdf_base,
                report_title=f"HBNC Today's Scheduled Visits ({today_date.strftime('%d-%b-%Y')})",
                active_filters_desc=active_filters_str,
            )
            st.download_button(label="📄 PDF (.pdf)", data=t_pdf_bytes, file_name=f"Today_Scheduled_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf", mime="application/pdf", key="btn_t_pdf", use_container_width=True)

    if not df_today_sched_roster.empty:
        st.dataframe(
            df_today_sched_roster[cols_roster_preview].rename(
                columns={"Baby Display": "Baby Name", "PHC": "Block / PHC Facility", "Due Date": "Scheduled Date"}
            ),
            use_container_width=True,
            hide_index=True,
        )
    else:
        st.info(f"ℹ️ No visits are scheduled for today ({today_date.strftime('%d-%b-%Y')}) under the current filter selection.")

    st.markdown("---")

    # Section D: Tomorrow's Scheduled Visits
    st.markdown(f'<div class="sub-section-header">D. Tomorrow\'s Scheduled Visits ({tomorrow_date.strftime("%d-%b-%Y")})</div>', unsafe_allow_html=True)
    tom_info_col, tom_btn_col1, tom_btn_col2 = st.columns([2.2, 0.9, 0.9])
    with tom_info_col:
        st.markdown(f"**Total Scheduled Visits:** `{len(df_tom_sched_roster)}` &nbsp;|&nbsp; **Unique Babies Expected:** `{df_tom_sched_roster['Baby ID'].nunique() if not df_tom_sched_roster.empty else 0}`")
    with tom_btn_col1:
        if not df_tom_sched_roster.empty:
            tom_excel_bytes = generate_single_table_excel(
                df_tom_sched_roster[cols_roster_preview].rename(columns={"Baby Display": "Baby Name", "PHC": "Block/PHC"}),
                sheet_name="Tomorrow_Scheduled",
            )
            st.download_button(label="📥 Excel (.xlsx)", data=tom_excel_bytes, file_name=f"Tomorrow_Scheduled_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="btn_tom_excel", use_container_width=True)
    with tom_btn_col2:
        if not df_tom_sched_roster.empty:
            tom_pdf_bytes = generate_pdf_report(
                df_export=df_tom_sched_roster.rename(columns={"Baby Display": "Baby Name", "PHC": "Block/PHC"}),
                columns_to_show=cols_roster_pdf_base,
                report_title=f"HBNC Tomorrow's Scheduled Visits ({tomorrow_date.strftime('%d-%b-%Y')})",
                active_filters_desc=active_filters_str,
            )
            st.download_button(label="📄 PDF (.pdf)", data=tom_pdf_bytes, file_name=f"Tomorrow_Scheduled_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf", mime="application/pdf", key="btn_tom_pdf", use_container_width=True)

    if not df_tom_sched_roster.empty:
        st.dataframe(
            df_tom_sched_roster[cols_roster_preview].rename(
                columns={"Baby Display": "Baby Name", "PHC": "Block / PHC Facility", "Due Date": "Scheduled Date"}
            ),
            use_container_width=True,
            hide_index=True,
        )
    else:
        st.info(f"ℹ️ No visits are scheduled for tomorrow ({tomorrow_date.strftime('%d-%b-%Y')}) under the current filter selection.")

# TAB 2: Video Quality & Dedicated Rejected Video Report
with tab_video:
    st.subheader("📹 Video Verification & Quality Feedback Center")
    if has_feedback and not df_fb_filtered.empty:
        st.info(
            f"ℹ️ **Column E (Visit_Id) Deduplication Active:** For each visit, double Visit_Ids are counted as a single visit. "
            f"Total Raw Rows: `{fb_raw_total_rows}` | Single Visits: `{fb_unique_visits_count}` | Filtered Single Visits: `{len(df_fb_filtered)}`"
        )

        col_v1, col_v2 = st.columns([1, 1.2])

        with col_v1:
            st.markdown("**1. Annotation Verdict Ratio (Single Unique Visits)**")
            v_counts = df_fb_filtered["Annotation Verdict"].value_counts().reset_index()
            v_counts.columns = ["Verdict", "Unique Visits"]
            if HAS_PLOTLY:
                fig_verdict = px.pie(
                    v_counts,
                    names="Verdict",
                    values="Unique Visits",
                    hole=0.45,
                    color="Verdict",
                    color_discrete_map={"Accepted": "#10B981", "Rejected": "#EF4444"},
                    height=260,
                )
                fig_verdict.update_layout(margin=dict(l=10, r=10, t=10, b=10))
                st.plotly_chart(fig_verdict, use_container_width=True)
            else:
                st.bar_chart(v_counts.set_index("Verdict")["Unique Visits"])

        with col_v2:
            st.markdown("**2. ASHA-Wise Video Quality & Rejection Count**")
            asha_fb = df_fb_filtered.groupby(["ASHA", "Annotation Verdict"]).size().unstack(fill_value=0).reset_index()
            for col_req in ["Accepted", "Rejected"]:
                if col_req not in asha_fb.columns:
                    asha_fb[col_req] = 0
            asha_fb["Total Unique Visits"] = asha_fb["Accepted"] + asha_fb["Rejected"]
            asha_fb["Rejection Rate"] = (asha_fb["Rejected"] / asha_fb["Total Unique Visits"] * 100).round(1).astype(str) + "%"
            st.dataframe(asha_fb.sort_values(by="Rejected", ascending=False), use_container_width=True, hide_index=True)

        st.markdown("---")
        
        # Dedicated Rejected Videos Report
        st.markdown('<div class="tier-header">🚫 Rejected Videos Actionable Report (Column E Visit_Id Deduplicated)</div>', unsafe_allow_html=True)
        
        df_rejected_report = df_fb_filtered[
            df_fb_filtered["Annotation Verdict"].str.strip().str.lower() == "rejected"
        ].copy()

        # Includes ASHA, Baby Name, DOB, Mother, Subcentre, Visit Date, Visit_Id, Rejected Feedback
        rej_cols_display = [
            "ASHA",
            "Baby Name",
            "DOB",
            "Mother",
            "Subcentre",
            "Visit Date",
            "Visit_Id",
            "Rejected Feedback"
        ]

        rej_head_c1, rej_head_c2, rej_head_c3 = st.columns([2.2, 0.9, 0.9])
        with rej_head_c1:
            st.markdown(
                f"**Total Rejected Single Visits:** `{len(df_rejected_report)}` &nbsp;|&nbsp; "
                f"**ASHAs Requiring Re-recording:** `{df_rejected_report['ASHA'].nunique() if not df_rejected_report.empty else 0}`"
            )
        with rej_head_c2:
            if not df_rejected_report.empty:
                rej_excel = generate_single_table_excel(df_rejected_report[rej_cols_display], sheet_name="Rejected_Videos")
                st.download_button(
                    label="📥 Excel (.xlsx)",
                    data=rej_excel,
                    file_name=f"Rejected_Videos_Report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="btn_rej_excel",
                    use_container_width=True,
                )
        with rej_head_c3:
            if not df_rejected_report.empty:
                # PDF report excludes Baby ID, includes Baby Name & DOB
                rej_pdf = generate_pdf_report(
                    df_export=df_rejected_report,
                    columns_to_show=["ASHA", "Baby Name", "DOB", "Mother", "Subcentre", "Visit Date", "Visit_Id", "Rejected Feedback"],
                    report_title="HBNC Actionable Rejected Videos Quality Report",
                    active_filters_desc=active_filters_str,
                )
                st.download_button(
                    label="📄 PDF (.pdf)",
                    data=rej_pdf,
                    file_name=f"Rejected_Videos_Report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf",
                    mime="application/pdf",
                    key="btn_rej_pdf",
                    use_container_width=True,
                )

        if not df_rejected_report.empty:
            st.dataframe(
                df_rejected_report[rej_cols_display].sort_values(by=["ASHA", "Visit Date"], ascending=[True, False]),
                use_container_width=True,
                hide_index=True,
            )
        else:
            st.success("🎉 No rejected videos found for the selected filter combination.")

        st.markdown("---")
        st.markdown("**3. Complete Video Feedback Audit Ledger (Single Visits per Visit_Id)**")
        
        fb_display_cols = [c for c in [
            "Visit_Id", "Annotation Verdict", "ASHA", "Baby Name", "DOB", "Mother", "Subcentre", "Visit Date",
            "Weight Video Feedback", "Arc Feedback", "Other Human Feedback", "Baby Cropping Feedback", "Ruler Cropping Feedback"
        ] if c in df_fb_filtered.columns]

        st.dataframe(df_fb_filtered[fb_display_cols], use_container_width=True, hide_index=True)

        st.markdown("##### 📥 Export Complete Video Audit Ledger")
        fb_col_d1, fb_col_d2 = st.columns(2)
        with fb_col_d1:
            fb_excel = generate_single_table_excel(df_fb_filtered[fb_display_cols], sheet_name="Video_Feedback_Dedup")
            st.download_button(label="📥 Download Complete Audit Excel (.xlsx)", data=fb_excel, file_name=f"Complete_Video_Audit_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
        with fb_col_d2:
            fb_pdf = generate_pdf_report(
                df_export=df_fb_filtered, 
                columns_to_show=["Visit_Id", "Annotation Verdict", "ASHA", "Baby Name", "DOB", "Mother", "Subcentre", "Visit Date"], 
                report_title="HBNC Complete Video Verification & Feedback Audit", 
                active_filters_desc=active_filters_str
            )
            st.download_button(label="📄 Download Complete Audit PDF (.pdf)", data=fb_pdf, file_name=f"Complete_Video_Audit_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf", mime="application/pdf", use_container_width=True)
    else:
        st.info("ℹ️ Video Feedback sheet upload nahi hui hai ya active filters ke matching koi record nahi hai. Kripya sidebar se 'Feedback Sheet' upload karein.")

# TAB 3: Registration & Visit Audit
with tab_reg_audit:
    st.subheader("🔍 Registration Date vs. Visit Date Audit (Baby-Wise Detailed Audit)")
    reg_audit_cols = ["Baby Display", "DOB", "Baby ID", "Mother", "ASHA", "PHC", "Subcentre", "Registration Date", "Last Visit Date", "Completed Visits Total", "Visits After Reg", "Pending Visits Total"]
    st.dataframe(
        df_b_filtered[reg_audit_cols].rename(
            columns={"Baby Display": "Baby Name", "PHC": "Facility", "Completed Visits Total": "Total Visits Done", "Visits After Reg": "Visits After Registration", "Pending Visits Total": "Pending Visits (Remaining)"}
        ).sort_values(by="Visits After Registration", ascending=False),
        use_container_width=True,
        hide_index=True,
    )

# TAB 4: ASHA & Healthcare Facility Overview
with tab_network:
    st.subheader("👩‍⚕️ Healthcare Facility, ASHA Registrations & Operational Overview")
    if not df_b_filtered.empty:
        asha_summary = (
            df_b_filtered.groupby("ASHA")
            .agg(
                District=("District", lambda x: x.iloc[0]),
                Facility_PHC=("PHC", lambda x: ", ".join(sorted(x.unique()))),
                Subcentres=("Subcentre", lambda x: ", ".join(sorted(x.unique()))),
                Registered_Babies=("Baby ID", "count"),
                Total_Completed=("Completed Visits Total", "sum"),
                Visits_After_Reg=("Visits After Reg", "sum"),
                Total_Pending=("Pending Visits Total", "sum"),
                Cumulative_Missed_Visits=("Missed Visits Count", "sum"),
            )
            .reset_index()
            .sort_values(by=["Registered_Babies", "Total_Completed"], ascending=[False, False])
        )

        st.markdown("**1. Comprehensive ASHA Master Workload & Performance Table**")
        st.dataframe(
            asha_summary.rename(columns={"Facility_PHC": "Facility (PHC/CHC)", "Registered_Babies": "Registered Babies", "Total_Completed": "Visits Completed", "Visits_After_Reg": "Visits Post-Reg", "Total_Pending": "Pending Visits", "Cumulative_Missed_Visits": "Overdue Missed Visits"}),
            use_container_width=True,
            hide_index=True,
        )

        st.markdown("---")
        st.markdown("**2. Interactive ASHA Drill-Down: Inspect Assigned Infants**")
        selected_drill_asha = st.selectbox("Select an ASHA worker to view her registered babies and visit schedules:", options=sorted(df_b_filtered["ASHA"].unique().tolist()), key="network_drill_asha")
        df_drill_infants = df_b_filtered[df_b_filtered["ASHA"] == selected_drill_asha].copy()
        st.caption(f"Displaying **{len(df_drill_infants)}** baby record(s) assigned to **{selected_drill_asha}**.")
        st.dataframe(
            df_drill_infants[["Baby Display", "DOB", "Baby ID", "Mother", "District", "PHC", "Subcentre", "Registration Date", "Last Visit Date", "Completed Visits Total", "Visits After Reg", "Pending Visits Total"]].rename(columns={"Baby Display": "Baby Name", "Completed Visits Total": "Completed Visits", "Visits After Reg": "Visits After Reg", "Pending Visits Total": "Pending Visits"}).sort_values(by="Completed Visits", ascending=False),
            use_container_width=True,
            hide_index=True,
        )

        st.markdown("---")
        col_net1, col_net2 = st.columns(2)
        with col_net1:
            st.markdown("**3. Facility-Wise (PHC/CHC) Registration Summary**")
            phc_reg_summary = df_b_filtered.groupby("PHC").agg(Total_ASHAs=("ASHA", "nunique"), Total_Subcentres=("Subcentre", "nunique"), Registered_Babies=("Baby ID", "count"), Total_Completed=("Completed Visits Total", "sum"), Total_Pending=("Pending Visits Total", "sum")).reset_index()
            st.dataframe(phc_reg_summary.rename(columns={"Total_ASHAs": "ASHAs", "Total_Subcentres": "Subcentres", "Registered_Babies": "Babies", "Total_Completed": "Completed Visits", "Total_Pending": "Pending Visits"}), use_container_width=True, hide_index=True)
        with col_net2:
            st.markdown("**4. Subcentre-Wise Baby Distribution**")
            sub_counts = df_b_filtered["Subcentre"].value_counts().reset_index()
            sub_counts.columns = ["Subcentre", "Registered Babies"]
            if HAS_PLOTLY:
                fig_sub = px.pie(sub_counts, names="Subcentre", values="Registered Babies", hole=0.45, color_discrete_sequence=px.colors.qualitative.Safe, height=250)
                fig_sub.update_layout(margin=dict(l=10, r=10, t=10, b=10))
                st.plotly_chart(fig_sub, use_container_width=True)
            else:
                st.bar_chart(sub_counts.set_index("Subcentre")["Registered Babies"])
    else:
        st.warning("No records match the current filter selection.")

# TAB 5: Baby Master Ledger
with tab_baby_master:
    st.subheader("👶 Individual Baby Master List & Accurate Missed Visit Breakdown")
    baby_master_cols = ["Baby Display", "DOB", "Baby ID", "Mother", "ASHA", "District", "PHC", "Subcentre", "Registration Date", "Last Visit Date", "Completed Visits Total", "Visits After Reg", "Pending Visits Total", "Missed Visits Count"]
    st.dataframe(
        df_b_filtered[baby_master_cols].rename(columns={"Baby Display": "Baby Name", "Completed Visits Total": "Completed Visits", "Visits After Reg": "Visits After Reg", "Pending Visits Total": "Pending Visits", "Missed Visits Count": "Missed Visits (Exact Sum)"}).sort_values(by=["Missed Visits (Exact Sum)", "Baby Name"], ascending=[False, True]),
        use_container_width=True,
        hide_index=True,
    )

# TAB 6: Operational Trends
with tab_trends:
    st.subheader("📈 Operational Trends & Multi-Tier Facility Performance")
    tr_c1, tr_c2 = st.columns(2)
    with tr_c1:
        st.markdown("**1. Scheduled Milestone Due Date Timeline Trend**")
        trend_timeline = df_v_filtered.groupby(["Due Date", "Visit Status"]).size().unstack(fill_value=0).reset_index()
        trend_timeline["Due Date"] = pd.to_datetime(trend_timeline["Due Date"])
        trend_timeline = trend_timeline.sort_values("Due Date")
        if HAS_PLOTLY:
            fig_trend = go.Figure()
            for status_label, color in [("Completed", "#10B981"), ("Missed", "#EF4444"), ("Upcoming", "#6366F1")]:
                if status_label in trend_timeline.columns:
                    fig_trend.add_trace(go.Scatter(x=trend_timeline["Due Date"], y=trend_timeline[status_label], mode="lines+markers", name=status_label, line=dict(color=color, width=2)))
            fig_trend.update_layout(height=280, margin=dict(l=15, r=15, t=20, b=15), xaxis_title="Milestone Due Date", yaxis_title="Visits", legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1), plot_bgcolor="rgba(248,250,252,0.5)")
            st.plotly_chart(fig_trend, use_container_width=True)
        else:
            st.line_chart(trend_timeline.set_index("Due Date"))

    with tr_c2:
        st.markdown("**2. Facility / Block Workload Distribution**")
        phc_breakdown = df_v_filtered.groupby(["PHC", "Visit Status"]).size().unstack(fill_value=0).reset_index()
        if HAS_PLOTLY:
            fig_phc_chart = px.bar(phc_breakdown, x="PHC", y=[c for c in ["Completed", "Missed", "Upcoming"] if c in phc_breakdown.columns], barmode="stack", color_discrete_map={"Completed": "#10B981", "Missed": "#EF4444", "Upcoming": "#6366F1"}, height=280)
            fig_phc_chart.update_layout(margin=dict(l=15, r=15, t=20, b=15), legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1), xaxis_title="Facility (PHC/CHC)", yaxis_title="Visits", plot_bgcolor="rgba(248,250,252,0.5)")
            st.plotly_chart(fig_phc_chart, use_container_width=True)
        else:
            st.bar_chart(phc_breakdown.set_index("PHC"))

# TAB 7: Download Reports
with tab_export:
    st.subheader("📥 Export Filter-Matched Analytical Reports")
    d_col1, d_col2 = st.columns(2)
    with d_col1:
        st.markdown("#### 📗 Multi-Sheet Excel Workbook (.xlsx)")
        # Excel workbook preserves all columns (including Baby ID and DOB)
        export_sheets = {
            "Baby_Master_Filtered": df_b_filtered[baby_master_cols].rename(columns={"Baby Display": "Baby Name"}),
            "Day_Before_Yesterday_Visits": df_dby_all_roster[cols_roster_preview + ["Visit Status"]].rename(columns={"Baby Display": "Baby Name", "PHC": "Block/PHC", "Visit Status": "Outcome"}),
            "Yesterday_Visits": df_yest_all_roster[cols_roster_preview + ["Visit Status"]].rename(columns={"Baby Display": "Baby Name", "PHC": "Block/PHC", "Visit Status": "Outcome"}),
            "Today_Scheduled": df_today_sched_roster[cols_roster_preview].rename(columns={"Baby Display": "Baby Name", "PHC": "Block/PHC"}),
            "Tomorrow_Scheduled": df_tom_sched_roster[cols_roster_preview].rename(columns={"Baby Display": "Baby Name", "PHC": "Block/PHC"}),
            "All_Filtered_Visits": df_v_filtered[["Due Date", "Baby ID", "Baby Display", "DOB", "Mother", "ASHA", "PHC", "Subcentre", "Visit Stage", "Visit Status"]].rename(columns={"Baby Display": "Baby Name"}),
        }
        if has_feedback and not df_fb_filtered.empty:
            export_sheets["Video_Feedback_SingleVisits"] = df_fb_filtered
            if not df_rejected_report.empty:
                export_sheets["Rejected_Videos_Report"] = df_rejected_report[rej_cols_display]
            
        xlsx_bytes = generate_excel_workbook(export_sheets)
        timestamp_slug = datetime.now().strftime("%Y%m%d_%H%M%S")
        st.download_button(label="📥 Download Comprehensive Filtered Excel (.xlsx)", data=xlsx_bytes, file_name=f"HBNC_Filtered_Report_{timestamp_slug}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)

    with d_col2:
        st.markdown("#### 📕 Landscape Printable PDF Report (.pdf)")
        export_choices = [
            "Day Before Yesterday's (Parso) Visits",
            "Yesterday's Visits & Outcomes",
            "Today's Scheduled Visits",
            "Tomorrow's Scheduled Visits",
            "Baby Registration vs Visit Audit",
            "Baby Master with Individual Missed Visits",
        ]
        if has_feedback and not df_fb_filtered.empty:
            export_choices.append("Rejected Videos Actionable Report")
            export_choices.append("Complete Video Verification & Quality Feedback")
            
        pdf_view_choice = st.radio("Select Report Content to Export:", export_choices, index=0)

        # PDF reports strictly include Baby Name and DOB, while omitting Baby ID
        if pdf_view_choice == "Day Before Yesterday's (Parso) Visits":
            pdf_source_df = df_dby_all_roster.rename(columns={"Baby Display": "Baby Name", "PHC": "Block/PHC", "Visit Status": "Outcome"})
            cols_to_print = ["Baby Name", "DOB", "Mother", "ASHA", "Block/PHC", "Subcentre", "Visit Stage", "Due Date", "Outcome"]
            pdf_title = f"HBNC Day Before Yesterday Visits ({day_before_yesterday_date.strftime('%d-%b-%Y')})"
        elif pdf_view_choice == "Yesterday's Visits & Outcomes":
            pdf_source_df = df_yest_all_roster.rename(columns={"Baby Display": "Baby Name", "PHC": "Block/PHC", "Visit Status": "Outcome"})
            cols_to_print = ["Baby Name", "DOB", "Mother", "ASHA", "Block/PHC", "Subcentre", "Visit Stage", "Due Date", "Outcome"]
            pdf_title = f"HBNC Yesterday's Visits & Outcomes ({yesterday_date.strftime('%d-%b-%Y')})"
        elif pdf_view_choice == "Today's Scheduled Visits":
            pdf_source_df = df_today_sched_roster.rename(columns={"Baby Display": "Baby Name", "PHC": "Block/PHC"})
            cols_to_print = ["Baby Name", "DOB", "Mother", "ASHA", "Block/PHC", "Subcentre", "Visit Stage", "Due Date"]
            pdf_title = f"HBNC Today's Scheduled Visits ({today_date.strftime('%d-%b-%Y')})"
        elif pdf_view_choice == "Tomorrow's Scheduled Visits":
            pdf_source_df = df_tom_sched_roster.rename(columns={"Baby Display": "Baby Name", "PHC": "Block/PHC"})
            cols_to_print = ["Baby Name", "DOB", "Mother", "ASHA", "Block/PHC", "Subcentre", "Visit Stage", "Due Date"]
            pdf_title = f"HBNC Tomorrow's Scheduled Visits ({tomorrow_date.strftime('%d-%b-%Y')})"
        elif pdf_view_choice == "Rejected Videos Actionable Report":
            pdf_source_df = df_rejected_report
            cols_to_print = ["ASHA", "Baby Name", "DOB", "Mother", "Subcentre", "Visit Date", "Visit_Id", "Rejected Feedback"]
            pdf_title = "HBNC Actionable Rejected Videos Quality Report (Visit_Id Deduplicated)"
        elif pdf_view_choice == "Complete Video Verification & Quality Feedback":
            pdf_source_df = df_fb_filtered
            cols_to_print = ["Visit_Id", "Annotation Verdict", "ASHA", "Baby Name", "DOB", "Mother", "Subcentre", "Visit Date"]
            pdf_title = "HBNC Video Verification & Feedback Audit (Visit_Id Single Visits)"
        elif pdf_view_choice == "Baby Registration vs Visit Audit":
            pdf_source_df = df_b_filtered.rename(
                columns={
                    "Baby Display": "Baby Name",
                    "Completed Visits Total": "Completed",
                    "Visits After Reg": "After Reg",
                    "Pending Visits Total": "Pending",
                }
            )
            cols_to_print = ["Baby Name", "DOB", "ASHA", "Registration Date", "Last Visit Date", "Completed", "After Reg", "Pending"]
            pdf_title = "HBNC Baby Registration vs. Visit Audit Report"
        else:
            pdf_source_df = df_b_filtered
            cols_to_print = ["Baby Display", "DOB", "Mother", "ASHA", "PHC", "Subcentre", "Missed Visits Count"]
            pdf_title = "HBNC Baby Master & Individual Missed Visits Ledger"

        pdf_bytes = generate_pdf_report(df_export=pdf_source_df, columns_to_show=cols_to_print, report_title=pdf_title, active_filters_desc=active_filters_str)
        st.download_button(label="📄 Download Landscape PDF (.pdf)", data=pdf_bytes, file_name=f"HBNC_{pdf_view_choice[:14].replace(' ', '_')}_{timestamp_slug}.pdf", mime="application/pdf", use_container_width=True)