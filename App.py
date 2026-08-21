"""
Form Processing + Notice & Hearing MIS Dashboard
=================================================

A single-file Streamlit application for monitoring Uttarakhand SIR (Special
Intensive Revision) electoral-roll workflow data across two source workbooks:

    Form_Processing.xlsx   -> AC x Form-Type snapshot of form processing status
    Notice_Hearing.xlsx    -> Part-level notice generation / hearing / DEO pendency data

--------------------------------------------------------------------------
HOW THIS DASHBOARD WAS DESIGNED (data-driven, not templated)
--------------------------------------------------------------------------
Both workbooks were inspected cell-by-cell before any KPI/chart/filter was
chosen. The findings that drive every design decision below:

FORM PROCESSING (Form_Processing.xlsx, sheet "Sheet1")
  * The real header row is row 5 of the sheet (0-indexed row 4). Rows 1-4
    are report metadata (state/district scope, the reporting date range,
    and gender/age filters used when the report was generated) -- not data.
  * Grain: one row per (AC No, Form Type). This is a SNAPSHOT/SUMMARY table,
    not one row per individual application.
  * The sheet contains subtotal rows ("AC Form Total") after every AC's
    form-type rows, and grand-total rows ("All AC Form-6 Total", "All AC
    Form-6A Total", "All AC Form-7 Total", "All AC Form-8 Total", "All AC
    Form Total") at the very bottom. These are DETECTED AND EXCLUDED from
    every calculation to avoid double counting -- confirmed by checking
    that AC No. is non-numeric on those rows.
  * Form Type values are FORM6 (new elector registration), FORM6A (NRI
    elector registration), FORM7 (objection to inclusion / deletion) and
    FORM8 (correction of entries / transposition) -- the standard Indian
    electoral-roll form categories that literally appear in the data.
  * The 18 status columns (Unprocessed, BLO Assigned, ... Eroll Inclusion)
    are MUTUALLY EXCLUSIVE and sum EXACTLY to "Total Form Received" for
    every row in the cleaned dataset (verified: 287,612 == 287,612 across
    the whole sheet). That means every form sits in exactly one current
    status bucket -- so this is a true status/stage distribution, and a
    donut/funnel of these columns is a meaningful, non-invented chart.
  * "Submitted Within AC", "RollBack Rejected" and "Eroll Pending" are zero
    for every row in the current data. They are kept in the underlying
    data/exports for completeness but are NOT promoted to KPI cards, since
    a KPI that is always zero tells an officer nothing.
  * There is NO District column in this workbook.

NOTICE & HEARING (Notice_Hearing.xlsx, sheet "sirNoticeGenerate")
  * Grain: one row per polling-station Part within an AC within a District
    -- genuine unit-level data (12,543 rows). Zero duplicate
    (District, AC, Part) combinations, zero blank cells anywhere.
  * "Notice Generated" == "Notice Delivered" + "Notice Pending Delivery"
    for every single row (verified exactly) -- so a Delivery Rate is a
    clean, non-invented derived KPI.
  * "Pending for Notice Generation" is 0 in every row of the current
    dataset (all notices already generated). It is not featured as a KPI
    for the same reason as above, but is shown as 0 if it appears.
  * "DEO-Status Pending GT 5 Days" is a genuine ageing/overdue indicator
    and is treated as a first-class KPI (backlog that is stale, not just
    backlog).
  * This workbook DOES carry District Number/Name and AC Number/Name for
    all 70 Uttarakhand Assembly Constituencies across 13 districts.

DISTRICT <-> AC MAPPING
  Form_Processing.xlsx has AC numbers/names but no District. Rather than
  inventing or hard-coding a mapping, the AC-to-District relationship is
  DERIVED AT RUNTIME from Notice_Hearing.xlsx (District Number, District
  Name, AC Number, Asmbly Name), which was verified to cover AC numbers
  1-70 with names that match Form_Processing's AC Name column exactly.
  AC No. 0 ("AC Missing") in Form_Processing has no district and is
  labelled "Unassigned / AC Not Mapped".

NO DATE COLUMN EXISTS AT ROW LEVEL IN EITHER FILE
  Form_Processing carries only a report-period string ("From ... To ...")
  in its metadata block, and Notice_Hearing has no date field at all. So
  this dashboard does NOT fabricate a date filter or a date-wise trend
  chart -- that would not be real data. The reporting period is instead
  surfaced as read-only context in the header.
--------------------------------------------------------------------------
"""

import io
import os
import re
import warnings
from datetime import datetime

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

warnings.filterwarnings("ignore")

# --------------------------------------------------------------------------
# Constants
# --------------------------------------------------------------------------
_APP_DIR = os.path.dirname(os.path.abspath(__file__))
FORM_PROCESSING_FILE = os.path.join(_APP_DIR, "Form_Processing.xlsx")
NOTICE_HEARING_FILE = os.path.join(_APP_DIR, "Notice_Hearing.xlsx")

BRAND_PRIMARY = "#0B3D91"
BRAND_PRIMARY_DARK = "#082B66"
BRAND_ACCENT = "#1F8A70"
BRAND_WARN = "#D97706"
BRAND_DANGER = "#C0392B"
BRAND_BG = "#F4F6FA"
BRAND_CARD = "#FFFFFF"
BRAND_TEXT = "#1B2430"
BRAND_MUTED = "#6B7685"

CHART_COLORWAY = ["#0B3D91", "#1F8A70", "#D97706", "#C0392B", "#6C5CE7",
                   "#0FA3B1", "#8D6A9F", "#3D5A80", "#B23A48", "#588157"]

FORM_TYPE_LABELS = {
    "FORM6": "Form 6 (New Registration)",
    "FORM6A": "Form 6A (NRI Elector)",
    "FORM7": "Form 7 (Objection / Deletion)",
    "FORM8": "Form 8 (Correction / Shifting)",
}

# Status columns as they appear in the raw Form_Processing.xlsx header
# (natural workflow order). These 18 columns are mutually exclusive and sum
# to Total Form Received (verified during data inspection).
FP_STATUS_COLS_RAW = [
    "Submitted Within AC", "Unprocessed", "BLO Assigned", "BLO DeAssigned",
    "Directly Forwarded to ERO", "Field Verified Submitted", "Forwarded to ERO",
    "Hearing Scheduled", "Re-Initiated to BLO", "Re-Initiated to AERO",
    "Wrong AC Marking", "Forwarded to Other DEO", "Rejected", "Accepted",
    "RollBack Requested", "RollBack Rejected", "RollBack Accepted",
    "Eroll Pending", "Eroll Inclusion",
]


def _clean_col(name: str) -> str:
    """Turn a raw Excel header into a safe underscore identifier."""
    return re.sub(r"[^0-9A-Za-z]+", "_", name).strip("_")


# Cleaned (underscore) column names used everywhere after load_form_processing().
FP_STATUS_COLS = [_clean_col(c) for c in FP_STATUS_COLS_RAW]
FP_RAW_TO_CLEAN = dict(zip(FP_STATUS_COLS_RAW, FP_STATUS_COLS))

# Columns that make up "finalized / disposed" outcomes for a form.
FP_FINAL_COLS = ["Accepted", "Rejected", "Eroll_Inclusion"]

# Everything else (besides Unprocessed and the final columns) is "in
# workflow" -- actively moving through the process but not yet finalized.
FP_INPROGRESS_COLS = [c for c in FP_STATUS_COLS
                       if c not in FP_FINAL_COLS + ["Unprocessed"]]


# --------------------------------------------------------------------------
# Small formatting helpers
# --------------------------------------------------------------------------
def fmt_int(n) -> str:
    try:
        return f"{int(round(float(n))):,}"
    except (TypeError, ValueError):
        return "0"


def fmt_pct(n, decimals=1) -> str:
    try:
        return f"{float(n):.{decimals}f}%"
    except (TypeError, ValueError):
        return "0.0%"


def safe_div(num, den):
    try:
        num = float(num)
        den = float(den)
        if den == 0:
            return 0.0
        return num / den * 100.0
    except (TypeError, ValueError):
        return 0.0


def clean_str(x) -> str:
    if pd.isna(x):
        return ""
    return re.sub(r"\s+", " ", str(x)).strip()


# --------------------------------------------------------------------------
# Data loading (cached)
# --------------------------------------------------------------------------
@st.cache_data(show_spinner=False)
def load_notice_hearing(path: str):
    """Load and clean Notice_Hearing.xlsx (sheet: sirNoticeGenerate)."""
    try:
        raw = pd.read_excel(path, sheet_name=0, header=0)
    except FileNotFoundError:
        return None, f"File not found: {path}"
    except Exception as exc:  # noqa: BLE001
        return None, f"Could not read {path}: {exc}"

    rename_map = {
        "State Id": "State_Id", "State Name": "State",
        "District Number": "District_No", "District Name": "District",
        "AC Number": "AC_No", "Asmbly Name": "AC_Name", "Part No": "Part_No",
        "Notice Generated": "Notice_Generated",
        "Pending for Notice Generation": "Pending_Notice_Generation",
        "Notice Delivered": "Notice_Delivered",
        "Notice Pending Delivery": "Notice_Pending_Delivery",
        "Hearings Held": "Hearings_Held",
        "DEO-Status Total Pending": "DEO_Total_Pending",
        "DEO-Status Pending GT 5 Days": "DEO_Pending_GT5",
        "DEO-Status Verified": "DEO_Verified",
        "DEO-Status Not Verified": "DEO_Not_Verified",
        "ERO/AERO Status Found Ineligible For Final w.r.t. Notice Generated": "Ineligible_Final",
        "ERO/AERO Status Parked For Final Publication": "Parked_Final_Publication",
        "ERO/AERO Parked For Final Publication w.r.t. Others": "Parked_Final_Publication_Others",
    }
    missing = [c for c in rename_map if c not in raw.columns]
    if missing:
        return None, ("Notice_Hearing.xlsx is missing expected column(s): "
                       + ", ".join(missing))

    df = raw.rename(columns=rename_map).copy()
    df["District"] = df["District"].apply(clean_str)
    df["AC_Name"] = df["AC_Name"].apply(clean_str)

    num_cols = ["District_No", "AC_No", "Part_No", "Notice_Generated",
                "Pending_Notice_Generation", "Notice_Delivered",
                "Notice_Pending_Delivery", "Hearings_Held", "DEO_Total_Pending",
                "DEO_Pending_GT5", "DEO_Verified", "DEO_Not_Verified",
                "Ineligible_Final", "Parked_Final_Publication",
                "Parked_Final_Publication_Others"]
    for c in num_cols:
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)

    df = df[(df["District"] != "") & (df["AC_Name"] != "")].copy()
    df = df.drop_duplicates(subset=["District_No", "AC_No", "Part_No"])

    df["Delivery_Rate"] = df.apply(
        lambda r: safe_div(r["Notice_Delivered"], r["Notice_Generated"]), axis=1)
    df["Hearing_Rate"] = df.apply(
        lambda r: safe_div(r["Hearings_Held"], r["Notice_Delivered"]), axis=1)

    return df, None


@st.cache_data(show_spinner=False)
def build_ac_district_map(nh_df: pd.DataFrame):
    """Derive AC -> District mapping from Notice_Hearing.xlsx (ground truth)."""
    if nh_df is None or nh_df.empty:
        return {}
    m = nh_df[["AC_No", "District"]].drop_duplicates()
    return dict(zip(m["AC_No"].astype(int), m["District"]))


@st.cache_data(show_spinner=False)
def load_form_processing(path: str, ac_district_map: dict):
    """Load and clean Form_Processing.xlsx (sheet: Sheet1)."""
    try:
        raw_head = pd.read_excel(path, sheet_name=0, header=None, nrows=6)
    except FileNotFoundError:
        return None, None, f"File not found: {path}"
    except Exception as exc:  # noqa: BLE001
        return None, None, f"Could not read {path}: {exc}"

    # Locate the real header row (the one starting with "AC No.").
    header_row_idx = None
    for i in range(min(10, len(raw_head))):
        first_cell = clean_str(raw_head.iloc[i, 0])
        if first_cell.lower().startswith("ac no"):
            header_row_idx = i
            break
    if header_row_idx is None:
        header_row_idx = 4  # fallback to the position observed in the source file

    meta_lines = [clean_str(raw_head.iloc[i, 0]) for i in range(header_row_idx)]
    period_match = re.search(r"From\s*-\s*([\d/]+)\s*To\s*-\s*([\d/]+)",
                              " ".join(meta_lines))
    report_period = (f"{period_match.group(1)} to {period_match.group(2)}"
                      if period_match else None)
    scope_line = next((m for m in meta_lines if ">>" in m), None)

    raw = pd.read_excel(path, sheet_name=0, header=header_row_idx)
    raw.columns = [clean_str(c) for c in raw.columns]

    required = ["AC No.", "AC Name", "Form Type", "Total Form Received"] + FP_STATUS_COLS_RAW
    missing = [c for c in required if c not in raw.columns]
    if missing:
        return None, None, ("Form_Processing.xlsx is missing expected column(s): "
                             + ", ".join(missing))

    # Keep only genuine data rows: AC No. must be numeric. This is exactly
    # how "AC Form Total" / "All AC Form-X Total" / "All AC Form Total"
    # subtotal and grand-total rows are detected and excluded.
    ac_no_numeric = pd.to_numeric(raw["AC No."], errors="coerce")
    df = raw[ac_no_numeric.notna()].copy()
    df["AC No."] = ac_no_numeric[ac_no_numeric.notna()].astype(int)
    df["AC Name"] = df["AC Name"].apply(clean_str)
    df["Form Type"] = df["Form Type"].apply(clean_str)
    df = df[(df["Form Type"] != "") & (df["Form Type"] != "nan")].copy()

    value_cols = ["Total Form Received"] + FP_STATUS_COLS_RAW
    for c in value_cols:
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)

    df = df.drop_duplicates(subset=["AC No.", "Form Type"])

    df["District"] = df["AC No."].map(ac_district_map).fillna("Unassigned / AC Not Mapped")
    df["Form_Type_Label"] = df["Form Type"].map(FORM_TYPE_LABELS).fillna(df["Form Type"])

    # Rename identifier + all 18 status columns to clean underscore names
    # BEFORE deriving Finalized/In_Progress, so every downstream reference
    # (module-level FP_STATUS_COLS / FP_FINAL_COLS / FP_INPROGRESS_COLS) is
    # consistent with the dataframe's actual column names.
    rename_map = {"AC No.": "AC_No", "AC Name": "AC_Name", "Form Type": "Form_Type",
                  "Total Form Received": "Total_Received"}
    rename_map.update(FP_RAW_TO_CLEAN)
    df = df.rename(columns=rename_map)

    df["Finalized"] = df[FP_FINAL_COLS].sum(axis=1)
    df["In_Progress"] = df[FP_INPROGRESS_COLS].sum(axis=1)

    meta = {"report_period": report_period, "scope": scope_line}
    return df, meta, None


# --------------------------------------------------------------------------
# Aggregation helpers
# --------------------------------------------------------------------------
def fp_district_report(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df
    g = df.groupby("District", as_index=False).agg(
        ACs_Reporting=("AC_No", "nunique"),
        Total_Received=("Total_Received", "sum"),
        Unprocessed=("Unprocessed", "sum"),
        In_Progress=("In_Progress", "sum"),
        Accepted=("Accepted", "sum"),
        Rejected=("Rejected", "sum"),
        Eroll_Inclusion=("Eroll_Inclusion", "sum"),
        Finalized=("Finalized", "sum"),
        Hearing_Scheduled=("Hearing_Scheduled", "sum"),
    )
    g["Disposal_Rate_%"] = g.apply(lambda r: safe_div(r["Finalized"], r["Total_Received"]), axis=1)
    g["Inclusion_Rate_%"] = g.apply(lambda r: safe_div(r["Eroll_Inclusion"], r["Total_Received"]), axis=1)
    g["Backlog_Rate_%"] = g.apply(lambda r: safe_div(r["Unprocessed"], r["Total_Received"]), axis=1)
    return g.sort_values("Total_Received", ascending=False)


def fp_ac_report(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df
    g = df.groupby(["District", "AC_No", "AC_Name"], as_index=False).agg(
        Total_Received=("Total_Received", "sum"),
        Unprocessed=("Unprocessed", "sum"),
        In_Progress=("In_Progress", "sum"),
        Accepted=("Accepted", "sum"),
        Rejected=("Rejected", "sum"),
        Eroll_Inclusion=("Eroll_Inclusion", "sum"),
        Finalized=("Finalized", "sum"),
        Hearing_Scheduled=("Hearing_Scheduled", "sum"),
    )
    g["Disposal_Rate_%"] = g.apply(lambda r: safe_div(r["Finalized"], r["Total_Received"]), axis=1)
    g["Inclusion_Rate_%"] = g.apply(lambda r: safe_div(r["Eroll_Inclusion"], r["Total_Received"]), axis=1)
    return g.sort_values("Total_Received", ascending=False)


def nh_district_report(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df
    g = df.groupby("District", as_index=False).agg(
        ACs_Reporting=("AC_No", "nunique"),
        Parts=("Part_No", "count"),
        Notice_Generated=("Notice_Generated", "sum"),
        Notice_Delivered=("Notice_Delivered", "sum"),
        Notice_Pending_Delivery=("Notice_Pending_Delivery", "sum"),
        Hearings_Held=("Hearings_Held", "sum"),
        DEO_Total_Pending=("DEO_Total_Pending", "sum"),
        DEO_Pending_GT5=("DEO_Pending_GT5", "sum"),
        Ineligible_Final=("Ineligible_Final", "sum"),
    )
    g["Delivery_Rate_%"] = g.apply(lambda r: safe_div(r["Notice_Delivered"], r["Notice_Generated"]), axis=1)
    g["Hearing_Rate_%"] = g.apply(lambda r: safe_div(r["Hearings_Held"], r["Notice_Delivered"]), axis=1)
    g["Pending_GT5_%"] = g.apply(lambda r: safe_div(r["DEO_Pending_GT5"], r["DEO_Total_Pending"]), axis=1)
    return g.sort_values("Notice_Generated", ascending=False)


def nh_ac_report(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df
    g = df.groupby(["District", "AC_No", "AC_Name"], as_index=False).agg(
        Parts=("Part_No", "count"),
        Notice_Generated=("Notice_Generated", "sum"),
        Notice_Delivered=("Notice_Delivered", "sum"),
        Notice_Pending_Delivery=("Notice_Pending_Delivery", "sum"),
        Hearings_Held=("Hearings_Held", "sum"),
        DEO_Total_Pending=("DEO_Total_Pending", "sum"),
        DEO_Pending_GT5=("DEO_Pending_GT5", "sum"),
        Ineligible_Final=("Ineligible_Final", "sum"),
    )
    g["Delivery_Rate_%"] = g.apply(lambda r: safe_div(r["Notice_Delivered"], r["Notice_Generated"]), axis=1)
    g["Hearing_Rate_%"] = g.apply(lambda r: safe_div(r["Hearings_Held"], r["Notice_Delivered"]), axis=1)
    return g.sort_values("Notice_Generated", ascending=False)


# --------------------------------------------------------------------------
# UI helpers
# --------------------------------------------------------------------------
def inject_css():
    st.markdown(f"""
    <style>
        .stApp {{ background-color: {BRAND_BG}; }}
        #MainMenu, footer {{ visibility: hidden; }}
        .block-container {{ padding-top: 1.2rem; padding-bottom: 2rem; max-width: 1400px; }}

        .mis-header {{
            background: linear-gradient(120deg, {BRAND_PRIMARY} 0%, {BRAND_PRIMARY_DARK} 100%);
            color: white; padding: 1.4rem 1.8rem; border-radius: 14px;
            margin-bottom: 1.2rem; box-shadow: 0 4px 18px rgba(11,61,145,0.25);
        }}
        .mis-header h1 {{ margin: 0; font-size: 1.55rem; font-weight: 700; letter-spacing: 0.2px; }}
        .mis-header p {{ margin: 0.3rem 0 0 0; opacity: 0.9; font-size: 0.92rem; }}
        .mis-badge {{
            display: inline-block; background: rgba(255,255,255,0.16);
            padding: 0.2rem 0.7rem; border-radius: 999px; font-size: 0.78rem;
            margin-top: 0.55rem; margin-right: 0.4rem; border: 1px solid rgba(255,255,255,0.25);
        }}

        .kpi-card {{
            background: {BRAND_CARD}; border-radius: 12px; padding: 1rem 1.1rem;
            border: 1px solid #E4E8F0; box-shadow: 0 2px 8px rgba(20,30,60,0.05);
            height: 100%;
        }}
        .kpi-label {{ color: {BRAND_MUTED}; font-size: 0.78rem; font-weight: 600;
                      text-transform: uppercase; letter-spacing: 0.4px; margin-bottom: 0.35rem; }}
        .kpi-value {{ color: {BRAND_TEXT}; font-size: 1.65rem; font-weight: 800; line-height: 1.1; }}
        .kpi-sub {{ font-size: 0.82rem; margin-top: 0.3rem; font-weight: 600; }}

        .section-title {{
            font-size: 1.05rem; font-weight: 700; color: {BRAND_TEXT};
            margin: 1.4rem 0 0.6rem 0; padding-bottom: 0.35rem;
            border-bottom: 2px solid {BRAND_PRIMARY};
        }}
        .note-box {{
            background: #EEF3FC; border-left: 4px solid {BRAND_PRIMARY};
            padding: 0.7rem 1rem; border-radius: 6px; font-size: 0.85rem;
            color: {BRAND_TEXT}; margin-bottom: 0.8rem;
        }}
        div[data-testid="stDataFrame"] {{ border-radius: 10px; overflow: hidden;
            border: 1px solid #E4E8F0; }}
        section[data-testid="stSidebar"] {{ background-color: #FFFFFF; border-right: 1px solid #E4E8F0; }}
    </style>
    """, unsafe_allow_html=True)


def kpi_card(col, label, value, sub=None, color=BRAND_PRIMARY):
    with col:
        sub_html = f'<div class="kpi-sub" style="color:{color}">{sub}</div>' if sub else ""
        st.markdown(f"""
        <div class="kpi-card">
            <div class="kpi-label">{label}</div>
            <div class="kpi-value">{value}</div>
            {sub_html}
        </div>
        """, unsafe_allow_html=True)


def section_title(text):
    st.markdown(f'<div class="section-title">{text}</div>', unsafe_allow_html=True)


def no_data_message():
    st.info("No data available for the selected filters.")


def apply_plotly_theme(fig, height=380):
    fig.update_layout(
        colorway=CHART_COLORWAY, height=height,
        margin=dict(l=10, r=10, t=45, b=10),
        paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
        font=dict(color=BRAND_TEXT, size=12.5),
        title_font=dict(size=15, color=BRAND_TEXT),
        legend=dict(orientation="h", yanchor="bottom", y=-0.25, x=0),
        hoverlabel=dict(bgcolor="white"),
    )
    fig.update_xaxes(gridcolor="#EAEDF3")
    fig.update_yaxes(gridcolor="#EAEDF3")
    return fig


# --------------------------------------------------------------------------
# Excel export
# --------------------------------------------------------------------------
def build_excel_download(sheets: dict) -> bytes:
    """sheets: {sheet_name: dataframe}. Returns formatted .xlsx bytes."""
    buf = io.BytesIO()
    try:
        engine = "xlsxwriter"
        writer = pd.ExcelWriter(buf, engine=engine)
    except Exception:  # noqa: BLE001
        engine = "openpyxl"
        writer = pd.ExcelWriter(buf, engine=engine)

    with writer:
        for name, df in sheets.items():
            safe_name = re.sub(r"[\[\]:*?/\\]", "_", name)[:31]
            if df is None or df.empty:
                pd.DataFrame({"Message": ["No data available for the selected filters."]}).to_excel(
                    writer, sheet_name=safe_name, index=False)
                continue
            df.to_excel(writer, sheet_name=safe_name, index=False)
            if engine == "xlsxwriter":
                wb = writer.book
                ws = writer.sheets[safe_name]
                header_fmt = wb.add_format({"bold": True, "bg_color": BRAND_PRIMARY,
                                             "font_color": "white", "border": 1})
                pct_fmt = wb.add_format({"num_format": "0.0", "border": 1})
                num_fmt = wb.add_format({"num_format": "#,##0", "border": 1})
                for c_idx, col in enumerate(df.columns):
                    header_fmt_local = header_fmt
                    ws.write(0, c_idx, col, header_fmt_local)
                    width = max(12, min(34, int(df[col].astype(str).str.len().max() or 10) + 3,
                                         len(str(col)) + 3))
                    if "%" in str(col):
                        ws.set_column(c_idx, c_idx, width, pct_fmt)
                    elif pd.api.types.is_numeric_dtype(df[col]):
                        ws.set_column(c_idx, c_idx, width, num_fmt)
                    else:
                        ws.set_column(c_idx, c_idx, width)
                ws.freeze_panes(1, 0)
    buf.seek(0)
    return buf.getvalue()


# --------------------------------------------------------------------------
# PDF export (ReportLab)
# --------------------------------------------------------------------------
def _fig_to_png(fig, width=1000, height=480):
    try:
        fig = go.Figure(fig)  # shallow copy so export-only tweaks don't affect the on-screen chart
        fig.update_layout(margin=dict(l=70, r=30, t=60, b=110), legend_title_text="")
        fig.update_xaxes(tickangle=-35)
        return fig.to_image(format="png", width=width, height=height, scale=2)
    except Exception:  # noqa: BLE001
        return None


def build_pdf_report(title, subtitle, filters_desc, kpis, district_df, ac_df,
                      charts, district_cols=None, ac_cols=None):
    """Builds a professional MIS PDF report and returns bytes."""
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4, landscape, portrait
    from reportlab.lib.units import mm
    from reportlab.platypus import (BaseDocTemplate, Frame, Image, PageTemplate,
                                     Paragraph, Spacer, Table, TableStyle,
                                     KeepTogether)
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.enums import TA_LEFT, TA_CENTER

    n_cols = max(
        len(district_cols) if district_cols else 0,
        len(ac_cols) if ac_cols else 0,
    )
    pagesize = landscape(A4) if n_cols >= 6 else portrait(A4)

    buf = io.BytesIO()
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle("MISTitle", parent=styles["Title"], fontSize=17,
                                  textColor=colors.HexColor(BRAND_PRIMARY_DARK), alignment=TA_LEFT)
    sub_style = ParagraphStyle("MISSub", parent=styles["Normal"], fontSize=10.5,
                                textColor=colors.HexColor(BRAND_MUTED), spaceAfter=4)
    h2_style = ParagraphStyle("MISH2", parent=styles["Heading2"], fontSize=12.5,
                               textColor=colors.HexColor(BRAND_PRIMARY_DARK), spaceBefore=10, spaceAfter=6)
    body_style = ParagraphStyle("MISBody", parent=styles["Normal"], fontSize=9, leading=13)
    filt_style = ParagraphStyle("MISFilt", parent=styles["Normal"], fontSize=9,
                                 textColor=colors.HexColor(BRAND_TEXT))

    def header_footer(canvas, doc):
        canvas.saveState()
        w, h = pagesize
        canvas.setFillColor(colors.HexColor(BRAND_PRIMARY))
        canvas.rect(0, h - 16 * mm, w, 16 * mm, fill=1, stroke=0)
        canvas.setFillColor(colors.white)
        canvas.setFont("Helvetica-Bold", 12)
        canvas.drawString(14 * mm, h - 10.5 * mm, "Form Processing & Notice-Hearing MIS")
        canvas.setFont("Helvetica", 8.5)
        canvas.drawRightString(w - 14 * mm, h - 10.5 * mm,
                                datetime.now().strftime("Generated: %d-%b-%Y %H:%M IST"))
        canvas.setFillColor(colors.HexColor(BRAND_MUTED))
        canvas.setFont("Helvetica", 8)
        canvas.drawString(14 * mm, 8 * mm, "Official MIS Report - Uttarakhand SIR")
        canvas.drawRightString(w - 14 * mm, 8 * mm, f"Page {doc.page}")
        canvas.restoreState()

    doc = BaseDocTemplate(buf, pagesize=pagesize,
                           leftMargin=14 * mm, rightMargin=14 * mm,
                           topMargin=20 * mm, bottomMargin=14 * mm)
    frame = Frame(doc.leftMargin, doc.bottomMargin, doc.width, doc.height, id="normal")
    doc.addPageTemplates([PageTemplate(id="mis", frames=frame, onPage=header_footer)])

    story = []
    story.append(Paragraph(title, title_style))
    story.append(Paragraph(subtitle, sub_style))
    story.append(Spacer(1, 4))

    story.append(Paragraph("<b>Filters Applied</b>", h2_style))
    story.append(Paragraph(filters_desc or "All records (no filters applied)", filt_style))
    story.append(Spacer(1, 6))

    if kpis:
        story.append(Paragraph("<b>Key Performance Summary</b>", h2_style))
        rows = [["Metric", "Value"]]
        for k, v in kpis.items():
            rows.append([k, v])
        t = Table(rows, colWidths=[doc.width * 0.6, doc.width * 0.4])
        t.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor(BRAND_PRIMARY)),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("FONTSIZE", (0, 0), (-1, -1), 9),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#EEF3FC")]),
            ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#D5DCE8")),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("TOPPADDING", (0, 0), (-1, -1), 4),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ]))
        story.append(t)
        story.append(Spacer(1, 8))

    NARROW_COLS = {"AC_No", "ACs_Reporting", "Parts", "District_No"}
    WIDE_COLS = {"District", "AC_Name"}
    cell_style = ParagraphStyle("MISCell", fontName="Helvetica", fontSize=7.4, leading=9)
    cell_style_r = ParagraphStyle("MISCellR", parent=cell_style, alignment=2)  # right-align
    header_cell_style = ParagraphStyle("MISCellH", fontName="Helvetica-Bold", fontSize=7.6,
                                        leading=9.5, textColor=colors.white)

    def _fmt_cell(val, col):
        if isinstance(val, float):
            if "%" in col:
                return f"{val:.1f}"
            return f"{val:,.0f}" if val == int(val) else f"{val:,.1f}"
        if isinstance(val, int):
            return f"{val:,}"
        return str(val)

    def df_to_table(df, cols, max_rows=40):
        cols = [c for c in cols if c in df.columns]
        show = df[cols].head(max_rows).copy()

        weights = []
        for c in cols:
            if c in NARROW_COLS:
                weights.append(0.55)
            elif c in WIDE_COLS:
                weights.append(1.7)
            else:
                weights.append(1.0)
        total_w = sum(weights)
        col_widths = [doc.width * w / total_w for w in weights]

        header = [Paragraph(c.replace("_", " "), header_cell_style) for c in cols]
        data_rows = []
        for _, row in show.iterrows():
            cells = []
            for c in cols:
                text = _fmt_cell(row[c], c)
                style = cell_style if c in WIDE_COLS or c == "District" else cell_style_r
                cells.append(Paragraph(text, style))
            data_rows.append(cells)
        data = [header] + data_rows

        t = Table(data, colWidths=col_widths, repeatRows=1)
        t.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor(BRAND_PRIMARY_DARK)),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F4F6FA")]),
            ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#D5DCE8")),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("TOPPADDING", (0, 0), (-1, -1), 3),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
            ("LEFTPADDING", (0, 0), (-1, -1), 4),
            ("RIGHTPADDING", (0, 0), (-1, -1), 4),
        ]))
        return t

    if district_df is not None and not district_df.empty and district_cols:
        story.append(Paragraph("<b>District-wise Report</b>", h2_style))
        story.append(df_to_table(district_df, district_cols))
        story.append(Spacer(1, 8))

    if ac_df is not None and not ac_df.empty and ac_cols:
        story.append(Paragraph("<b>AC-wise Report</b>", h2_style))
        story.append(df_to_table(ac_df, ac_cols))
        story.append(Spacer(1, 8))

    if charts:
        story.append(Paragraph("<b>Charts</b>", h2_style))
        for chart_title, fig in charts:
            png = _fig_to_png(fig)
            if png:
                img = Image(io.BytesIO(png), width=doc.width, height=doc.width * 0.45)
                story.append(Paragraph(chart_title, body_style))
                story.append(img)
                story.append(Spacer(1, 8))

    doc.build(story)
    buf.seek(0)
    return buf.getvalue()


# --------------------------------------------------------------------------
# App
# --------------------------------------------------------------------------
st.set_page_config(page_title="Form Processing & Notice-Hearing MIS",
                    page_icon="\U0001F5F3", layout="wide",
                    initial_sidebar_state="expanded")
inject_css()

st.markdown(f"""
<div class="mis-header">
    <h1>Form Processing &amp; Notice-Hearing MIS Dashboard</h1>
    <p>Special Intensive Revision (SIR) &mdash; Uttarakhand &nbsp;|&nbsp; District &amp; AC-level monitoring</p>
    <span class="mis-badge">Data-driven KPIs</span>
    <span class="mis-badge">Live filters</span>
    <span class="mis-badge">PDF &amp; Excel export</span>
</div>
""", unsafe_allow_html=True)

nh_df, nh_err = load_notice_hearing(NOTICE_HEARING_FILE)
ac_map = build_ac_district_map(nh_df) if nh_df is not None else {}
fp_df, fp_meta, fp_err = load_form_processing(FORM_PROCESSING_FILE, ac_map)

with st.sidebar:
    st.markdown("### \U0001F4CB Dashboard Controls")
    st.caption("Filters apply live to KPIs, charts, reports and exports.")

tab1, tab2 = st.tabs(["\U0001F4C4  Form Processing", "\U0001F4E8  Notice & Hearing"])

# ==========================================================================
# TAB 1: FORM PROCESSING
# ==========================================================================
with tab1:
    if fp_err:
        st.error(f"**Form Processing data could not be loaded.**\n\n{fp_err}")
    elif fp_df is None or fp_df.empty:
        st.warning("Form_Processing.xlsx loaded but contains no usable data rows.")
    else:
        if fp_meta.get("report_period"):
            st.markdown(f"""<div class="note-box">
                <b>Report period (from source file):</b> {fp_meta['report_period']}
                &nbsp;&nbsp;|&nbsp;&nbsp; <b>Scope:</b> {fp_meta.get('scope') or 'Uttarakhand - All Districts'}
                <br><span style="color:{BRAND_MUTED}">No per-record date field exists in this workbook,
                so no date-wise filter/trend is shown here (only this reporting-period range applies to the whole file).</span>
                </div>""", unsafe_allow_html=True)

        with st.sidebar:
            st.markdown("#### \U0001F4C4 Form Processing Filters")
            fp_districts = sorted(fp_df["District"].unique())
            fp_sel_districts = st.multiselect("District", fp_districts, default=[], key="fp_dist")
            fp_scope = fp_df[fp_df["District"].isin(fp_sel_districts)] if fp_sel_districts else fp_df
            fp_acs = sorted(fp_scope["AC_Name"].unique())
            fp_sel_acs = st.multiselect("Assembly Constituency (AC)", fp_acs, default=[], key="fp_ac")
            fp_form_types = sorted(fp_df["Form_Type"].unique())
            fp_sel_forms = st.multiselect(
                "Form Type", fp_form_types,
                format_func=lambda x: FORM_TYPE_LABELS.get(x, x), default=[], key="fp_form")
            fp_metric_options = {
                "Total Received": "Total_Received", "Unprocessed (Backlog)": "Unprocessed",
                "In Progress (Workflow)": "In_Progress", "Finalized (Disposed)": "Finalized",
                "Eroll Inclusion": "Eroll_Inclusion", "Rejected": "Rejected",
                "Hearing Scheduled": "Hearing_Scheduled",
            }
            fp_focus_label = st.selectbox("Chart focus metric", list(fp_metric_options.keys()),
                                           index=0, key="fp_metric")
            fp_focus_metric = fp_metric_options[fp_focus_label]
            st.caption("Status is captured as parallel stage-count columns per record "
                       "(not a single status field), so it drives the chart-focus metric "
                       "above and the status-mix chart on the page, rather than a row filter.")

        filtered = fp_df.copy()
        if fp_sel_districts:
            filtered = filtered[filtered["District"].isin(fp_sel_districts)]
        if fp_sel_acs:
            filtered = filtered[filtered["AC_Name"].isin(fp_sel_acs)]
        if fp_sel_forms:
            filtered = filtered[filtered["Form_Type"].isin(fp_sel_forms)]

        filt_parts = []
        if fp_sel_districts: filt_parts.append("District: " + ", ".join(fp_sel_districts))
        if fp_sel_acs: filt_parts.append("AC: " + ", ".join(fp_sel_acs))
        if fp_sel_forms: filt_parts.append("Form Type: " + ", ".join(fp_sel_forms))
        fp_filters_desc = " | ".join(filt_parts) if filt_parts else "All Districts, All ACs, All Form Types"

        if filtered.empty:
            no_data_message()
        else:
            total_received = filtered["Total_Received"].sum()
            unprocessed = filtered["Unprocessed"].sum()
            in_progress = filtered["In_Progress"].sum()
            finalized = filtered["Finalized"].sum()
            eroll_inclusion = filtered["Eroll_Inclusion"].sum()
            rejected = filtered["Rejected"].sum()
            accepted = filtered["Accepted"].sum()
            hearing_sched = filtered["Hearing_Scheduled"].sum()

            section_title("Key Performance Indicators")
            c1, c2, c3, c4 = st.columns(4)
            kpi_card(c1, "Total Forms Received", fmt_int(total_received),
                     f"{filtered['AC_No'].nunique()} ACs · {filtered['Form_Type'].nunique()} form types")
            kpi_card(c2, "Unprocessed (Backlog)", fmt_int(unprocessed),
                     f"{fmt_pct(safe_div(unprocessed, total_received))} of received", color=BRAND_DANGER)
            kpi_card(c3, "In Progress (Workflow)", fmt_int(in_progress),
                     f"{fmt_pct(safe_div(in_progress, total_received))} of received", color=BRAND_WARN)
            kpi_card(c4, "Finalized (Disposed)", fmt_int(finalized),
                     f"{fmt_pct(safe_div(finalized, total_received))} disposal rate", color=BRAND_ACCENT)

            c5, c6, c7, c8 = st.columns(4)
            kpi_card(c5, "Eroll Inclusion", fmt_int(eroll_inclusion),
                     f"{fmt_pct(safe_div(eroll_inclusion, total_received))} inclusion rate", color=BRAND_ACCENT)
            kpi_card(c6, "Rejected", fmt_int(rejected),
                     f"{fmt_pct(safe_div(rejected, total_received))} rejection rate", color=BRAND_DANGER)
            kpi_card(c7, "Accepted", fmt_int(accepted),
                     f"{fmt_pct(safe_div(accepted, total_received))} of received")
            kpi_card(c8, "Hearing Scheduled", fmt_int(hearing_sched),
                     f"{fmt_pct(safe_div(hearing_sched, total_received))} of received", color=BRAND_WARN)

            section_title("Visual Analysis")
            v1, v2 = st.columns(2)
            with v1:
                status_sums = filtered[FP_STATUS_COLS].sum()
                status_sums = status_sums[status_sums > 0].sort_values(ascending=False)
                status_labels = [s.replace("_", " ") for s in status_sums.index]
                if len(status_sums):
                    fig = px.pie(names=status_labels, values=status_sums.values, hole=0.5,
                                 title="Current Status Mix (all forms sum to exactly one stage each)")
                    fig.update_traces(textinfo="percent+label", textposition="inside")
                    st.plotly_chart(apply_plotly_theme(fig), use_container_width=True)
            with v2:
                ftype_sums = filtered.groupby("Form_Type_Label")["Total_Received"].sum().sort_values(ascending=False)
                fig = px.bar(x=ftype_sums.values, y=ftype_sums.index, orientation="h",
                             title="Total Forms Received by Form Type",
                             labels={"x": "Forms Received", "y": ""})
                fig.update_traces(marker_color=BRAND_PRIMARY)
                st.plotly_chart(apply_plotly_theme(fig), use_container_width=True)

            v3, v4 = st.columns(2)
            with v3:
                dist_cmp = filtered.groupby("District")[["Unprocessed", "In_Progress", "Finalized"]].sum()
                dist_cmp = dist_cmp.loc[dist_cmp.sum(axis=1).sort_values(ascending=False).index]
                fig = go.Figure()
                for col, color in zip(["Unprocessed", "In_Progress", "Finalized"],
                                       [BRAND_DANGER, BRAND_WARN, BRAND_ACCENT]):
                    fig.add_bar(name=col.replace("_", " "), x=dist_cmp.index, y=dist_cmp[col], marker_color=color)
                fig.update_layout(barmode="stack", title="District-wise Processing Status (stacked)")
                st.plotly_chart(apply_plotly_theme(fig), use_container_width=True)
            with v4:
                ac_focus = filtered.groupby("AC_Name")[fp_focus_metric].sum().sort_values(ascending=False).head(15)
                fig = px.bar(x=ac_focus.values, y=ac_focus.index, orientation="h",
                             title=f"Top 15 ACs by {fp_focus_label}",
                             labels={"x": fp_focus_label, "y": ""})
                fig.update_traces(marker_color=BRAND_PRIMARY)
                fig.update_yaxes(autorange="reversed")
                st.plotly_chart(apply_plotly_theme(fig), use_container_width=True)

            # ------------------ District-wise report ------------------
            section_title("District-wise Report")
            fp_dist_rep = fp_district_report(filtered)
            fp_dist_display_cols = ["District", "ACs_Reporting", "Total_Received", "Unprocessed",
                                     "In_Progress", "Finalized", "Eroll_Inclusion", "Rejected",
                                     "Disposal_Rate_%", "Inclusion_Rate_%", "Backlog_Rate_%"]
            sort_opt = st.selectbox("Sort district report by", fp_dist_display_cols[1:],
                                     index=fp_dist_display_cols.index("Total_Received") - 1, key="fp_dist_sort")
            fp_dist_view = fp_dist_rep[fp_dist_display_cols].sort_values(sort_opt, ascending=False)
            st.dataframe(
                fp_dist_view.style.format({
                    "Total_Received": "{:,.0f}", "Unprocessed": "{:,.0f}", "In_Progress": "{:,.0f}",
                    "Finalized": "{:,.0f}", "Eroll_Inclusion": "{:,.0f}", "Rejected": "{:,.0f}",
                    "Disposal_Rate_%": "{:.1f}%", "Inclusion_Rate_%": "{:.1f}%", "Backlog_Rate_%": "{:.1f}%",
                }), use_container_width=True, height=380)

            fig = px.bar(fp_dist_view, x="District", y=["Total_Received", "Finalized"],
                         barmode="group", title="District Comparison: Received vs Finalized",
                         color_discrete_sequence=CHART_COLORWAY, labels={"value": "Forms", "variable": ""})
            st.plotly_chart(apply_plotly_theme(fig, height=360), use_container_width=True)

            # ------------------ AC-wise report ------------------
            section_title("AC-wise Report")
            ac_dist_pick = st.selectbox("Select District to drill into its ACs (optional)",
                                         ["All Districts"] + sorted(filtered["District"].unique()),
                                         key="fp_ac_district_pick")
            ac_scope_df = filtered if ac_dist_pick == "All Districts" else filtered[filtered["District"] == ac_dist_pick]
            fp_ac_rep = fp_ac_report(ac_scope_df)
            fp_ac_display_cols = ["District", "AC_No", "AC_Name", "Total_Received", "Unprocessed",
                                   "In_Progress", "Finalized", "Eroll_Inclusion", "Rejected",
                                   "Disposal_Rate_%", "Inclusion_Rate_%"]
            st.dataframe(
                fp_ac_rep[fp_ac_display_cols].style.format({
                    "Total_Received": "{:,.0f}", "Unprocessed": "{:,.0f}", "In_Progress": "{:,.0f}",
                    "Finalized": "{:,.0f}", "Eroll_Inclusion": "{:,.0f}", "Rejected": "{:,.0f}",
                    "Disposal_Rate_%": "{:.1f}%", "Inclusion_Rate_%": "{:.1f}%",
                }), use_container_width=True, height=380)

            ac_chart_df = fp_ac_rep.sort_values("Total_Received", ascending=False).head(20)
            fig = px.bar(ac_chart_df, x="AC_Name", y=["Unprocessed", "In_Progress", "Finalized"],
                         barmode="stack", title=f"AC Comparison ({ac_dist_pick}) - Top 20 by Volume",
                         color_discrete_sequence=CHART_COLORWAY, labels={"value": "Forms", "variable": ""})
            fig.update_xaxes(tickangle=-40)
            st.plotly_chart(apply_plotly_theme(fig, height=380), use_container_width=True)

            # ------------------ Downloads ------------------
            section_title("Downloads")
            dcol1, dcol2, dcol3 = st.columns(3)
            with dcol1:
                excel_bytes = build_excel_download({
                    "Filtered Records": filtered.drop(columns=["Form_Type_Label"], errors="ignore"),
                    "District Report": fp_dist_rep,
                    "AC Report": fp_ac_rep,
                })
                st.download_button("\U0001F4E5 Download Filtered Report (Excel)", excel_bytes,
                                    file_name="Form_Processing_Filtered_Report.xlsx",
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                    use_container_width=True)
            with dcol2:
                kpis_for_pdf = {
                    "Total Forms Received": fmt_int(total_received),
                    "Unprocessed (Backlog)": f"{fmt_int(unprocessed)} ({fmt_pct(safe_div(unprocessed, total_received))})",
                    "In Progress (Workflow)": f"{fmt_int(in_progress)} ({fmt_pct(safe_div(in_progress, total_received))})",
                    "Finalized (Disposed)": f"{fmt_int(finalized)} ({fmt_pct(safe_div(finalized, total_received))})",
                    "Eroll Inclusion": f"{fmt_int(eroll_inclusion)} ({fmt_pct(safe_div(eroll_inclusion, total_received))})",
                    "Rejected": f"{fmt_int(rejected)} ({fmt_pct(safe_div(rejected, total_received))})",
                    "Hearing Scheduled": fmt_int(hearing_sched),
                }
                status_fig = apply_plotly_theme(px.pie(
                    names=status_labels, values=status_sums.values, hole=0.5,
                    title="Current Status Mix")) if len(status_sums) else None
                dist_fig = apply_plotly_theme(px.bar(
                    fp_dist_view, x="District", y=["Total_Received", "Finalized"],
                    barmode="group", title="District Comparison",
                    color_discrete_sequence=CHART_COLORWAY, labels={"value": "Forms", "variable": ""}))
                charts_for_pdf = [c for c in [
                    ("Current Status Mix", status_fig) if status_fig is not None else None,
                    ("District Comparison: Received vs Finalized", dist_fig),
                ] if c is not None]
                try:
                    pdf_bytes = build_pdf_report(
                        title="Form Processing MIS Report",
                        subtitle=f"Report period: {fp_meta.get('report_period') or 'N/A'}",
                        filters_desc=fp_filters_desc, kpis=kpis_for_pdf,
                        district_df=fp_dist_view, ac_df=fp_ac_rep[fp_ac_display_cols],
                        charts=charts_for_pdf,
                        district_cols=fp_dist_display_cols, ac_cols=fp_ac_display_cols,
                    )
                    st.download_button("\U0001F4C4 Download PDF Report", pdf_bytes,
                                        file_name="Form_Processing_MIS_Report.pdf",
                                        mime="application/pdf", use_container_width=True)
                except Exception as exc:  # noqa: BLE001
                    st.error(f"PDF generation failed: {exc}")
            with dcol3:
                dist_excel = build_excel_download({"District Report": fp_dist_rep, "AC Report": fp_ac_rep})
                st.download_button("\U0001F4E5 Download District + AC Report (Excel)", dist_excel,
                                    file_name="Form_Processing_District_AC_Report.xlsx",
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                    use_container_width=True)

# ==========================================================================
# TAB 2: NOTICE & HEARING
# ==========================================================================
with tab2:
    if nh_err:
        st.error(f"**Notice & Hearing data could not be loaded.**\n\n{nh_err}")
    elif nh_df is None or nh_df.empty:
        st.warning("Notice_Hearing.xlsx loaded but contains no usable data rows.")
    else:
        st.markdown(f"""<div class="note-box">
            <b>Granularity:</b> one row per polling-station Part within an AC within a District
            ({fmt_int(len(nh_df))} parts across {nh_df['District'].nunique()} districts and {nh_df['AC_No'].nunique()} ACs).
            <br><span style="color:{BRAND_MUTED}">No date field exists in this workbook, so no date filter/trend is shown here.</span>
            </div>""", unsafe_allow_html=True)

        with st.sidebar:
            st.markdown("#### \U0001F4E8 Notice & Hearing Filters")
            nh_districts = sorted(nh_df["District"].unique())
            nh_sel_districts = st.multiselect("District", nh_districts, default=[], key="nh_dist")
            nh_scope = nh_df[nh_df["District"].isin(nh_sel_districts)] if nh_sel_districts else nh_df
            nh_acs = sorted(nh_scope["AC_Name"].unique())
            nh_sel_acs = st.multiselect("Assembly Constituency (AC)", nh_acs, default=[], key="nh_ac")
            nh_pending_only = st.checkbox("Show only Parts with DEO-pending cases", value=False, key="nh_pending_only")

        nh_filtered = nh_df.copy()
        if nh_sel_districts:
            nh_filtered = nh_filtered[nh_filtered["District"].isin(nh_sel_districts)]
        if nh_sel_acs:
            nh_filtered = nh_filtered[nh_filtered["AC_Name"].isin(nh_sel_acs)]
        if nh_pending_only:
            nh_filtered = nh_filtered[nh_filtered["DEO_Total_Pending"] > 0]

        nh_filt_parts = []
        if nh_sel_districts: nh_filt_parts.append("District: " + ", ".join(nh_sel_districts))
        if nh_sel_acs: nh_filt_parts.append("AC: " + ", ".join(nh_sel_acs))
        if nh_pending_only: nh_filt_parts.append("Only Parts with DEO-pending cases")
        nh_filters_desc = " | ".join(nh_filt_parts) if nh_filt_parts else "All Districts, All ACs"

        if nh_filtered.empty:
            no_data_message()
        else:
            notice_gen = nh_filtered["Notice_Generated"].sum()
            notice_del = nh_filtered["Notice_Delivered"].sum()
            notice_pend_del = nh_filtered["Notice_Pending_Delivery"].sum()
            hearings_held = nh_filtered["Hearings_Held"].sum()
            deo_pending = nh_filtered["DEO_Total_Pending"].sum()
            deo_pending_gt5 = nh_filtered["DEO_Pending_GT5"].sum()
            ineligible = nh_filtered["Ineligible_Final"].sum()

            section_title("Key Performance Indicators")
            c1, c2, c3, c4 = st.columns(4)
            kpi_card(c1, "Notices Generated", fmt_int(notice_gen),
                     f"{nh_filtered['AC_No'].nunique()} ACs · {fmt_int(len(nh_filtered))} parts")
            kpi_card(c2, "Notices Delivered", fmt_int(notice_del),
                     f"{fmt_pct(safe_div(notice_del, notice_gen))} delivery rate", color=BRAND_ACCENT)
            kpi_card(c3, "Notice Pending Delivery", fmt_int(notice_pend_del),
                     f"{fmt_pct(safe_div(notice_pend_del, notice_gen))} of generated", color=BRAND_WARN)
            kpi_card(c4, "Hearings Held", fmt_int(hearings_held),
                     f"{fmt_pct(safe_div(hearings_held, notice_del))} of delivered notices")

            c5, c6, c7 = st.columns(3)
            kpi_card(c5, "DEO Pending (Backlog)", fmt_int(deo_pending),
                     f"{fmt_pct(safe_div(deo_pending, notice_gen))} of notices generated", color=BRAND_DANGER)
            kpi_card(c6, "DEO Pending > 5 Days", fmt_int(deo_pending_gt5),
                     f"{fmt_pct(safe_div(deo_pending_gt5, deo_pending))} of DEO backlog is overdue", color=BRAND_DANGER)
            kpi_card(c7, "Found Ineligible for Final", fmt_int(ineligible),
                     f"{fmt_pct(safe_div(ineligible, notice_gen))} of notices generated", color=BRAND_WARN)

            section_title("Visual Analysis")
            v1, v2 = st.columns(2)
            with v1:
                fig = px.pie(names=["Delivered", "Pending Delivery"],
                             values=[notice_del, notice_pend_del], hole=0.5,
                             title="Notice Delivery Split", color_discrete_sequence=[BRAND_ACCENT, BRAND_WARN])
                st.plotly_chart(apply_plotly_theme(fig), use_container_width=True)
            with v2:
                pend_df = nh_filtered.groupby("District")[["DEO_Total_Pending", "DEO_Pending_GT5"]].sum()
                pend_df = pend_df.loc[pend_df["DEO_Total_Pending"].sort_values(ascending=False).index]
                fig = go.Figure()
                fig.add_bar(name="Total Pending", x=pend_df.index, y=pend_df["DEO_Total_Pending"], marker_color=BRAND_WARN)
                fig.add_bar(name="Pending > 5 Days", x=pend_df.index, y=pend_df["DEO_Pending_GT5"], marker_color=BRAND_DANGER)
                fig.update_layout(barmode="overlay", title="DEO Pendency by District (overdue vs total)")
                st.plotly_chart(apply_plotly_theme(fig), use_container_width=True)

            v3, v4 = st.columns(2)
            with v3:
                cmp = nh_filtered.groupby("District")[["Notice_Generated", "Notice_Delivered", "Hearings_Held"]].sum()
                cmp = cmp.loc[cmp["Notice_Generated"].sort_values(ascending=False).index]
                fig = px.bar(cmp, x=cmp.index, y=["Notice_Generated", "Notice_Delivered", "Hearings_Held"],
                             barmode="group", title="District Comparison: Generated / Delivered / Hearings Held",
                             color_discrete_sequence=CHART_COLORWAY, labels={"value": "Count", "variable": ""})
                st.plotly_chart(apply_plotly_theme(fig), use_container_width=True)
            with v4:
                # Delivery rate is uniformly high (~97-100%) across ACs in this
                # data, so ranking by rate barely differentiates ACs. The
                # absolute "Pending Delivery" count is the actionable view for
                # an officer -- it shows exactly where the undelivered notices
                # are concentrated.
                ac_pend = nh_filtered.groupby("AC_Name")["Notice_Pending_Delivery"].sum().sort_values(ascending=False).head(15)
                fig = px.bar(x=ac_pend.values, y=ac_pend.index, orientation="h",
                             title="Top 15 ACs by Notice Pending Delivery (count)",
                             labels={"x": "Notices Pending Delivery", "y": ""})
                fig.update_traces(marker_color=BRAND_DANGER)
                fig.update_yaxes(autorange="reversed")
                st.plotly_chart(apply_plotly_theme(fig), use_container_width=True)

            # ------------------ District-wise report ------------------
            section_title("District-wise Report")
            nh_dist_rep = nh_district_report(nh_filtered)
            nh_dist_display_cols = ["District", "ACs_Reporting", "Parts", "Notice_Generated", "Notice_Delivered",
                                     "Delivery_Rate_%", "Hearings_Held", "Hearing_Rate_%", "DEO_Total_Pending",
                                     "DEO_Pending_GT5", "Pending_GT5_%", "Ineligible_Final"]
            nh_sort_opt = st.selectbox("Sort district report by", nh_dist_display_cols[1:],
                                        index=nh_dist_display_cols.index("Notice_Generated") - 1, key="nh_dist_sort")
            nh_dist_view = nh_dist_rep[nh_dist_display_cols].sort_values(nh_sort_opt, ascending=False)
            st.dataframe(
                nh_dist_view.style.format({
                    "Parts": "{:,.0f}", "Notice_Generated": "{:,.0f}", "Notice_Delivered": "{:,.0f}",
                    "Delivery_Rate_%": "{:.1f}%", "Hearings_Held": "{:,.0f}", "Hearing_Rate_%": "{:.1f}%",
                    "DEO_Total_Pending": "{:,.0f}", "DEO_Pending_GT5": "{:,.0f}", "Pending_GT5_%": "{:.1f}%",
                    "Ineligible_Final": "{:,.0f}",
                }), use_container_width=True, height=380)

            fig = px.bar(nh_dist_view, x="District", y=["Notice_Generated", "Notice_Delivered"],
                         barmode="group", title="District Comparison: Notices Generated vs Delivered",
                         color_discrete_sequence=CHART_COLORWAY, labels={"value": "Notices", "variable": ""})
            st.plotly_chart(apply_plotly_theme(fig, height=360), use_container_width=True)

            # ------------------ AC-wise report ------------------
            section_title("AC-wise Report")
            nh_ac_dist_pick = st.selectbox("Select District to drill into its ACs (optional)",
                                            ["All Districts"] + sorted(nh_filtered["District"].unique()),
                                            key="nh_ac_district_pick")
            nh_ac_scope = nh_filtered if nh_ac_dist_pick == "All Districts" else nh_filtered[nh_filtered["District"] == nh_ac_dist_pick]
            nh_ac_rep = nh_ac_report(nh_ac_scope)
            nh_ac_display_cols = ["District", "AC_No", "AC_Name", "Parts", "Notice_Generated", "Notice_Delivered",
                                   "Delivery_Rate_%", "Hearings_Held", "Hearing_Rate_%", "DEO_Total_Pending",
                                   "DEO_Pending_GT5"]
            st.dataframe(
                nh_ac_rep[nh_ac_display_cols].style.format({
                    "Parts": "{:,.0f}", "Notice_Generated": "{:,.0f}", "Notice_Delivered": "{:,.0f}",
                    "Delivery_Rate_%": "{:.1f}%", "Hearings_Held": "{:,.0f}", "Hearing_Rate_%": "{:.1f}%",
                    "DEO_Total_Pending": "{:,.0f}", "DEO_Pending_GT5": "{:,.0f}",
                }), use_container_width=True, height=380)

            ac_chart_df2 = nh_ac_rep.sort_values("Notice_Generated", ascending=False).head(20)
            fig = px.bar(ac_chart_df2, x="AC_Name", y=["Notice_Delivered", "Notice_Pending_Delivery"],
                         barmode="stack", title=f"AC Comparison ({nh_ac_dist_pick}) - Top 20 by Notices Generated",
                         color_discrete_sequence=CHART_COLORWAY, labels={"value": "Notices", "variable": ""})
            fig.update_xaxes(tickangle=-40)
            st.plotly_chart(apply_plotly_theme(fig, height=380), use_container_width=True)

            # ------------------ Downloads ------------------
            section_title("Downloads")
            dcol1, dcol2, dcol3 = st.columns(3)
            with dcol1:
                excel_bytes2 = build_excel_download({
                    "Filtered Records": nh_filtered,
                    "District Report": nh_dist_rep,
                    "AC Report": nh_ac_rep,
                })
                st.download_button("\U0001F4E5 Download Filtered Report (Excel)", excel_bytes2,
                                    file_name="Notice_Hearing_Filtered_Report.xlsx",
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                    use_container_width=True, key="nh_dl1")
            with dcol2:
                kpis_for_pdf2 = {
                    "Notices Generated": fmt_int(notice_gen),
                    "Notices Delivered": f"{fmt_int(notice_del)} ({fmt_pct(safe_div(notice_del, notice_gen))})",
                    "Notice Pending Delivery": fmt_int(notice_pend_del),
                    "Hearings Held": f"{fmt_int(hearings_held)} ({fmt_pct(safe_div(hearings_held, notice_del))} of delivered)",
                    "DEO Pending (Backlog)": fmt_int(deo_pending),
                    "DEO Pending > 5 Days": f"{fmt_int(deo_pending_gt5)} ({fmt_pct(safe_div(deo_pending_gt5, deo_pending))} of backlog)",
                    "Found Ineligible for Final": fmt_int(ineligible),
                }
                split_fig = apply_plotly_theme(px.pie(
                    names=["Delivered", "Pending Delivery"], values=[notice_del, notice_pend_del],
                    hole=0.5, title="Notice Delivery Split",
                    color_discrete_sequence=[BRAND_ACCENT, BRAND_WARN]))
                dist_fig2 = apply_plotly_theme(px.bar(
                    nh_dist_view, x="District", y=["Notice_Generated", "Notice_Delivered"],
                    barmode="group", title="District Comparison",
                    color_discrete_sequence=CHART_COLORWAY, labels={"value": "Notices", "variable": ""}))
                try:
                    pdf_bytes2 = build_pdf_report(
                        title="Notice & Hearing MIS Report",
                        subtitle=f"{fmt_int(len(nh_filtered))} polling-station parts in scope",
                        filters_desc=nh_filters_desc, kpis=kpis_for_pdf2,
                        district_df=nh_dist_view, ac_df=nh_ac_rep[nh_ac_display_cols],
                        charts=[("Notice Delivery Split", split_fig),
                                ("District Comparison: Generated vs Delivered", dist_fig2)],
                        district_cols=nh_dist_display_cols, ac_cols=nh_ac_display_cols,
                    )
                    st.download_button("\U0001F4C4 Download PDF Report", pdf_bytes2,
                                        file_name="Notice_Hearing_MIS_Report.pdf",
                                        mime="application/pdf", use_container_width=True, key="nh_dl2")
                except Exception as exc:  # noqa: BLE001
                    st.error(f"PDF generation failed: {exc}")
            with dcol3:
                dist_excel2 = build_excel_download({"District Report": nh_dist_rep, "AC Report": nh_ac_rep})
                st.download_button("\U0001F4E5 Download District + AC Report (Excel)", dist_excel2,
                                    file_name="Notice_Hearing_District_AC_Report.xlsx",
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                    use_container_width=True, key="nh_dl3")

st.markdown(f"""
<div style="text-align:center; color:{BRAND_MUTED}; font-size:0.78rem; margin-top:2rem; padding-top:1rem; border-top:1px solid #E4E8F0;">
    Form Processing &amp; Notice-Hearing MIS Dashboard &middot; Data refreshed from workbook files in this repository &middot;
    Generated {datetime.now().strftime('%d-%b-%Y %H:%M')} IST
</div>
""", unsafe_allow_html=True)