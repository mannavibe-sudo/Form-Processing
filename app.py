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
NOTICE_FILE = os.path.join(_APP_DIR, "Notice.xlsx")
ELECTORS_FILE = os.path.join(_APP_DIR, "Electors.xlsx")
# Comparison workbook for the Difference Report -- same Form Processing
# columns as Form_Processing.xlsx, but District-level (not AC-level) and
# covering an earlier report period, so the two can be diffed district by
# district to show what changed between the two cutoff dates.
FORM_PROCESSING_OLD_FILE = os.path.join(_APP_DIR, "1513.xlsx")

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

# Charts are intentionally switched off for now -- the dashboard shows only
# KPI cards + clean District-wise / AC-wise report tables + downloads. All
# chart code below is kept intact and working (it was fully tested); flip
# this back to True whenever charts are wanted again -- no other change needed.
SHOW_CHARTS = False

# --------------------------------------------------------------------------
# On-screen report tables: trimmed to the columns that matter for day-to-day
# monitoring (glanceable on a phone). Column NAMES are kept exactly as they
# are elsewhere in the app -- nothing is renamed. The fuller column sets
# (*_display_cols, defined next to each report below) are still used for the
# Excel/PDF exports, so no data is lost -- only the on-screen table is leaner.
# --------------------------------------------------------------------------
# Notice & Hearing report: exactly the column set the user specified (built
# from Notice.xlsx + Electors.xlsx), used as-is on screen, in Excel, and in
# PDF -- no separate "trimmed for screen" vs "full" column set here, unlike
# Form Processing below, since this set already IS the intended default.
NOTICE_DIST_COLS = ["District", "Electors", "Notice_Delivered", "Hearing_Held",
                     "Hearing_Date_Lapsed", "Lapsed_%", "DEO_Total_Pending",
                     "Ineligible_Final", "Parked_Notices_Generated", "Parked_Others",
                     "Parked_Final_%"]
NOTICE_AC_COLS = ["District", "AC_No", "AC_Name", "Electors", "Notice_Delivered",
                   "Hearing_Held", "Hearing_Date_Lapsed", "Lapsed_%", "DEO_Total_Pending",
                   "Ineligible_Final", "Parked_Notices_Generated", "Parked_Others",
                   "Parked_Final_%"]
# Header text for columns whose plain "replace underscore with space" label
# wouldn't match the exact wording the report is meant to use.
NOTICE_COL_LABELS = {
    "Notice_Delivered": "Notice Delivered",
    "Hearing_Held": "Hearing Held",
    "Hearing_Date_Lapsed": "Hearing Date Lapsed",
    "Lapsed_%": "% Lapsed",
    "DEO_Total_Pending": "Total pending Text (DEO)",
    "Ineligible_Final": "Found Ineligible for Final (ERO)",
    "Parked_Notices_Generated": "Parked for Final Publication w.r.t. Notices Generated",
    "Parked_Others": "Parked For Final Publication w.r.t. Others",
    "Parked_Final_%": "% Parked for Final",
}
NOTICE_COL_FORMATS = {
    "Electors": "{:,.0f}", "Notice_Delivered": "{:,.0f}", "Hearing_Held": "{:,.0f}",
    "Hearing_Date_Lapsed": "{:,.0f}", "Lapsed_%": "{:.2f}",
    "DEO_Total_Pending": "{:,.0f}", "Ineligible_Final": "{:,.0f}",
    "Parked_Notices_Generated": "{:,.0f}", "Parked_Others": "{:,.0f}",
    "Parked_Final_%": "{:.2f}",
}
# Form Processing's screen columns are defined further below, right after
# FP_STATUS_COLS -- they're built from that list (every raw status column
# from the sheet), so they can't be defined before it exists.

# Rate/percent columns get a colour-coded pill in the on-screen tables so an
# officer can spot problem districts/ACs at a glance without reading numbers.
RATE_COLS_HIGHER_BETTER = {"Inclusion_Rate_%"}
RATE_COLS_LOWER_BETTER = set()

FORM_TYPE_LABELS = {
    "FORM6": "Form 6 (New Registration)",
    "FORM6A": "Form 6A (NRI Elector)",
    "FORM7": "Form 7 (Objection / Deletion)",
    "FORM8": "Form 8 (Correction / Shifting)",
}
# Short form (no parenthetical) -- used in compact contexts like the PDF's
# filter-summary line, where "Form 7" reads better than the full label.
FORM_TYPE_SHORT = {"FORM6": "Form 6", "FORM6A": "Form 6A", "FORM7": "Form 7", "FORM8": "Form 8"}

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
# On-screen Form Processing report tables: default to literal, recognisable
# columns straight from Form_Processing.xlsx (identifiers + Total Form
# Received + a handful of the raw status columns) -- nothing invented. Every
# other raw status column, plus the derived In Progress / Inclusion Rate
# metrics, is available as an optional "add columns" picker so nothing from
# the sheet is hidden -- it's just not cluttering the default view.
#
# NOTE: "Finalized", "Disposal Rate %" and "Backlog Rate %" were removed
# entirely (not just hidden) at the user's request -- they were computed
# summaries (Finalized = Accepted+Rejected+Eroll Inclusion; the two rates
# were built from that and from Unprocessed) with no equivalent field in
# either source workbook, and caused confusion. Unprocessed, Accepted,
# Rejected and Eroll Inclusion are still shown -- those ARE real columns.
# --------------------------------------------------------------------------
FP_DIST_BASE_COLS = ["District", "Total_Received", "Hearing_Scheduled",
                      "Rejected", "Accepted", "Eroll_Inclusion"]
FP_AC_BASE_COLS = ["District", "AC_No", "AC_Name", "Total_Received",
                    "Hearing_Scheduled", "Rejected", "Accepted", "Eroll_Inclusion"]
FP_DIST_EXTRA_COLS = (["ACs_Reporting"]
                       + [c for c in FP_STATUS_COLS if c not in FP_DIST_BASE_COLS]
                       + ["In_Progress", "Inclusion_Rate_%"])
FP_AC_EXTRA_COLS = ([c for c in FP_STATUS_COLS if c not in FP_AC_BASE_COLS]
                     + ["In_Progress", "Inclusion_Rate_%"])

# Format spec covering every possible Form Processing report column (base +
# extra), used regardless of which optional columns the user adds on screen.
FP_COL_FORMATS = {c: "{:,.0f}" for c in FP_STATUS_COLS}
FP_COL_FORMATS.update({
    "Total_Received": "{:,.0f}", "In_Progress": "{:,.0f}",
    "ACs_Reporting": "{:,.0f}",
    "Inclusion_Rate_%": "{:.1f}%",
})

# --------------------------------------------------------------------------
# Difference Report: same columns as the Form Processing District-wise
# report, but the *values* are District-wise deltas between Form_Processing.
# xlsx (current period) and the comparison workbook (an earlier cutoff,
# District-level only -- no AC breakdown exists for it). District is the
# only granularity available for a diff, since the comparison file has no
# AC No./AC Name columns of its own.
# --------------------------------------------------------------------------
DIFF_DIST_BASE_COLS = ["District", "Total_Received", "Hearing_Scheduled",
                        "Rejected", "Accepted", "Eroll_Inclusion"]
DIFF_DIST_EXTRA_COLS = ([c for c in FP_STATUS_COLS if c not in DIFF_DIST_BASE_COLS]
                         + ["In_Progress", "Inclusion_Rate_%"])
DIFF_COL_LABELS = {"Inclusion_Rate_%": "Inclusion Rate % (change, points)",
                    "In_Progress": "In Progress (change)"}
# "{:+,.0f}" / "{:+.1f}%" (note the "+") so a positive change reads as
# "+1,234" and a negative one as "-1,234" -- this table is deltas, not
# absolute counts, so the sign needs to be visible at a glance.
DIFF_COL_FORMATS = {c: "{:+,.0f}" for c in FP_STATUS_COLS}
DIFF_COL_FORMATS.update({
    "Total_Received": "{:+,.0f}", "In_Progress": "{:+,.0f}",
    "Inclusion_Rate_%": "{:+.1f}%",
})


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


def fmt_diff(n) -> str:
    """Like fmt_int, but always shows a sign (+1,234 / -1,234) -- used for
    the Difference Report, where the sign is the whole point."""
    try:
        return f"{int(round(float(n))):+,}"
    except (TypeError, ValueError):
        return "0"


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
# Notice.xlsx has a two-row merged column header (row 2 = category like
# "DEO Status", row 3 = the actual sub-column like "Total Pending") over 20
# columns (A-T, the last blank) -- rather than fragile merged-cell text
# matching, the columns are given clean names here in their fixed left-to-
# right order, the same pragmatic approach FP_STATUS_COLS_RAW uses above.
NOTICE_RAW_COLS = [
    "S_No", "AC_Combo", "Notice_Generated", "Pending_Notice_Generation", "Notice_Delivered",
    "Notice_Pending_Delivery", "Hearing_Held", "Hearing_Date_Lapsed", "Reschedule_Date_Lapsed",
    "DEO_Total_Pending", "DEO_Pending_GT5", "DEO_Verified", "DEO_Not_Verified",
    "Ineligible_Final", "Parked_Notices_Generated", "Parked_Others",
    "Electors_DSE_UP", "Form7_Generated", "No_Action_Required", "Blank_Col",
]


@st.cache_data(show_spinner=False)
def load_notice_data(notice_path: str, electors_path: str):
    """Load Notice.xlsx (AC-level notice/hearing/DEO/ERO counts) and
    Electors.xlsx (AC-level elector totals + District mapping), and merge
    them on AC No. into one AC-level dataframe -- this replaces the old
    Part-level Notice_Hearing.xlsx as the data source for the Notice &
    Hearing tab.

    Both files have a 3-row header block (title / scope / column header --
    Notice.xlsx's header is split across two merged rows) followed by one
    data row per AC and a trailing "Total" row; the Total row is dropped
    here since every total in this app is recomputed from the row-level
    data rather than trusted from the sheet.
    """
    try:
        notice_raw = pd.read_excel(notice_path, sheet_name=0, header=None)
    except FileNotFoundError:
        return None, None, f"File not found: {notice_path}"
    except Exception as exc:  # noqa: BLE001
        return None, None, f"Could not read {notice_path}: {exc}"
    try:
        electors_raw = pd.read_excel(electors_path, sheet_name=0, header=None)
    except FileNotFoundError:
        return None, None, f"File not found: {electors_path}"
    except Exception as exc:  # noqa: BLE001
        return None, None, f"Could not read {electors_path}: {exc}"

    # The last column (T) is entirely blank in the source file, so pandas
    # drops it on read -- 19 columns is the normal/expected width; 20 is
    # also accepted in case a future export keeps that blank column.
    if notice_raw.shape[1] < 19 or notice_raw.shape[0] < 4:
        return None, None, "Notice.xlsx does not match the expected layout (19-20 columns, header in rows 1-3)."
    ncols = min(notice_raw.shape[1], len(NOTICE_RAW_COLS))
    n = notice_raw.iloc[3:, :ncols].copy()
    n.columns = NOTICE_RAW_COLS[:ncols]
    if "Blank_Col" not in n.columns:
        n["Blank_Col"] = None
    n = n[pd.to_numeric(n["S_No"], errors="coerce").notna()].copy()
    if n.empty:
        return None, None, "Notice.xlsx has no usable AC data rows."
    n["AC_Combo"] = n["AC_Combo"].apply(clean_str)
    n["AC_No"] = n["AC_Combo"].apply(lambda x: pd.to_numeric(x.split("-")[0], errors="coerce"))
    n = n[n["AC_No"].notna()].copy()
    n["AC_No"] = n["AC_No"].astype(int)
    n["AC_Name"] = n["AC_Combo"].apply(lambda x: clean_str(x.split("-", 1)[1]) if "-" in x else x)

    num_cols = [c for c in NOTICE_RAW_COLS if c not in ("S_No", "AC_Combo", "Blank_Col")]
    for c in num_cols:
        n[c] = pd.to_numeric(n[c], errors="coerce").fillna(0)
    n = n.drop(columns=["S_No", "AC_Combo", "Blank_Col"]).drop_duplicates(subset=["AC_No"])

    if electors_raw.shape[1] < 10 or electors_raw.shape[0] < 4:
        return None, None, "Electors.xlsx does not match the expected layout (10 columns, header in rows 1-3)."
    e = electors_raw.iloc[3:, :10].copy()
    e.columns = ["AC_No", "AC_Name_E", "District_No", "District", "PC_No", "PC_Name",
                 "Male", "Female", "Third_Gender", "Electors"]
    e["AC_No"] = pd.to_numeric(e["AC_No"], errors="coerce")
    e = e[e["AC_No"].notna()].copy()
    if e.empty:
        return None, None, "Electors.xlsx has no usable AC data rows."
    e["AC_No"] = e["AC_No"].astype(int)
    e["District"] = e["District"].apply(clean_str)
    e["Electors"] = pd.to_numeric(e["Electors"], errors="coerce").fillna(0)
    e = e.drop_duplicates(subset=["AC_No"])

    df = n.merge(e[["AC_No", "District", "Electors"]], on="AC_No", how="left")
    unmatched = int(df["District"].isna().sum())
    df["District"] = df["District"].fillna("Unassigned / AC Not Mapped")
    df["Electors"] = df["Electors"].fillna(0)

    meta = {"n_acs": len(df), "unmatched_acs": unmatched}
    return df, meta, None


@st.cache_data(show_spinner=False)
def build_ac_district_map(notice_df: pd.DataFrame):
    """Derive AC -> District mapping from the Notice & Hearing data (ground
    truth, sourced from Electors.xlsx) -- Form_Processing.xlsx has no
    District column of its own, so this is how it gets one."""
    if notice_df is None or notice_df.empty:
        return {}
    m = notice_df[["AC_No", "District"]].drop_duplicates()
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
    # BEFORE deriving In_Progress, so every downstream reference (module-level
    # FP_STATUS_COLS / FP_FINAL_COLS / FP_INPROGRESS_COLS) is consistent with
    # the dataframe's actual column names.
    rename_map = {"AC No.": "AC_No", "AC Name": "AC_Name", "Form Type": "Form_Type",
                  "Total Form Received": "Total_Received"}
    rename_map.update(FP_RAW_TO_CLEAN)
    df = df.rename(columns=rename_map)

    df["In_Progress"] = df[FP_INPROGRESS_COLS].sum(axis=1)

    meta = {"report_period": report_period, "scope": scope_line}
    return df, meta, None


@st.cache_data(show_spinner=False)
def load_form_processing_district(path: str):
    """Load a Form-Processing-style workbook that is keyed by District
    instead of AC (same header/column layout as Form_Processing.xlsx, just
    with 'District No.'/'District Name' identifier columns) -- this is the
    comparison workbook that backs the Difference Report."""
    try:
        raw_head = pd.read_excel(path, sheet_name=0, header=None, nrows=6)
    except FileNotFoundError:
        return None, None, f"File not found: {path}"
    except Exception as exc:  # noqa: BLE001
        return None, None, f"Could not read {path}: {exc}"

    header_row_idx = None
    for i in range(min(10, len(raw_head))):
        first_cell = clean_str(raw_head.iloc[i, 0])
        if first_cell.lower().startswith("district no"):
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

    required = ["District No.", "District Name", "Form Type", "Total Form Received"] + FP_STATUS_COLS_RAW
    missing = [c for c in required if c not in raw.columns]
    if missing:
        return None, None, ("Comparison workbook is missing expected column(s): "
                             + ", ".join(missing))

    # Keep only genuine data rows: District No. must be numeric. This is how
    # "District Form Total" / "All District Form-X Total" / "All District
    # Form Total" subtotal and grand-total rows are excluded.
    dist_no_numeric = pd.to_numeric(raw["District No."], errors="coerce")
    df = raw[dist_no_numeric.notna()].copy()
    df["District Name"] = df["District Name"].apply(clean_str)
    df["Form Type"] = df["Form Type"].apply(clean_str)
    df = df[(df["Form Type"] != "") & (df["Form Type"] != "nan")].copy()

    value_cols = ["Total Form Received"] + FP_STATUS_COLS_RAW
    for c in value_cols:
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)

    rename_map = {"District Name": "District", "Total Form Received": "Total_Received",
                  "Form Type": "Form_Type"}
    rename_map.update(FP_RAW_TO_CLEAN)
    df = df.rename(columns=rename_map)

    meta = {"report_period": report_period, "scope": scope_line}
    return df, meta, None


# --------------------------------------------------------------------------
# Aggregation helpers
# --------------------------------------------------------------------------
def fp_district_report(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df
    agg_kwargs = {"ACs_Reporting": ("AC_No", "nunique"), "Total_Received": ("Total_Received", "sum")}
    agg_kwargs.update({c: (c, "sum") for c in FP_STATUS_COLS})
    g = df.groupby("District", as_index=False).agg(**agg_kwargs)
    g["In_Progress"] = g[FP_INPROGRESS_COLS].sum(axis=1)
    g["Inclusion_Rate_%"] = g.apply(lambda r: safe_div(r["Eroll_Inclusion"], r["Total_Received"]), axis=1)
    return g.sort_values("Total_Received", ascending=False)


def fp_ac_report(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df
    agg_kwargs = {"Total_Received": ("Total_Received", "sum")}
    agg_kwargs.update({c: (c, "sum") for c in FP_STATUS_COLS})
    g = df.groupby(["District", "AC_No", "AC_Name"], as_index=False).agg(**agg_kwargs)
    g["In_Progress"] = g[FP_INPROGRESS_COLS].sum(axis=1)
    g["Inclusion_Rate_%"] = g.apply(lambda r: safe_div(r["Eroll_Inclusion"], r["Total_Received"]), axis=1)
    # Group rows by District (districts ordered by their own total volume,
    # highest first -- matching the District-wise report above), and within
    # each district rank ACs from top receiving to lowest.
    dist_total = g.groupby("District")["Total_Received"].transform("sum")
    return (g.assign(_dist_total=dist_total)
             .sort_values(["_dist_total", "District", "Total_Received"], ascending=[False, True, False])
             .drop(columns="_dist_total"))


def fp_district_report_flat(df: pd.DataFrame) -> pd.DataFrame:
    """Aggregate a District x Form_Type dataframe (already at district
    granularity -- no AC No./AC Name of its own, e.g. the comparison
    workbook behind the Difference Report) up to one row per district. Same
    shape as fp_district_report()'s output, minus ACs_Reporting (there's no
    AC-level data to count)."""
    if df.empty:
        return df
    agg_kwargs = {"Total_Received": ("Total_Received", "sum")}
    agg_kwargs.update({c: (c, "sum") for c in FP_STATUS_COLS})
    g = df.groupby("District", as_index=False).agg(**agg_kwargs)
    g["In_Progress"] = g[FP_INPROGRESS_COLS].sum(axis=1)
    g["Inclusion_Rate_%"] = g.apply(lambda r: safe_div(r["Eroll_Inclusion"], r["Total_Received"]), axis=1)
    return g.sort_values("Total_Received", ascending=False)


def fp_diff_report(new_dist: pd.DataFrame, old_dist: pd.DataFrame) -> pd.DataFrame:
    """District-wise difference: new period minus old/comparison period,
    same columns as the Form Processing District-wise report (Total_Received
    + every real status column, plus the derived In_Progress/Inclusion_Rate_%
    for reference). A district present in only one side is still included,
    treating its missing side as zero, so nothing silently drops out."""
    diff_cols = ["Total_Received"] + FP_STATUS_COLS + ["In_Progress", "Inclusion_Rate_%"]
    if new_dist is None or new_dist.empty:
        new_dist = pd.DataFrame(columns=["District"] + diff_cols)
    if old_dist is None or old_dist.empty:
        old_dist = pd.DataFrame(columns=["District"] + diff_cols)

    all_districts = sorted(set(new_dist["District"]) | set(old_dist["District"]))
    n = new_dist.set_index("District").reindex(all_districts, fill_value=0)
    o = old_dist.set_index("District").reindex(all_districts, fill_value=0)

    diff = pd.DataFrame(index=all_districts)
    for c in diff_cols:
        n_col = n[c] if c in n.columns else 0
        o_col = o[c] if c in o.columns else 0
        diff[c] = n_col - o_col
    diff = diff.reset_index().rename(columns={"index": "District"})
    return diff.sort_values("Total_Received", ascending=False)


def _add_notice_pct_cols(g: pd.DataFrame) -> pd.DataFrame:
    """% Lapsed and % Parked for Final -- the two derived columns in the
    Notice & Hearing report, computed the same way at every grouping level
    (district, AC, or the grand total row) so the numbers stay consistent:
      % Lapsed = Hearing Date Lapsed / Notice Delivered
      % Parked for Final = (Parked w.r.t. Notices Generated + Parked w.r.t.
                             Others) / Electors
    (safe_div already returns the ratio *100, e.g. 21.02 for 21.02%.)
    """
    g["Lapsed_%"] = g.apply(
        lambda r: safe_div(r["Hearing_Date_Lapsed"], r["Notice_Delivered"]), axis=1)
    g["Parked_Final_%"] = g.apply(
        lambda r: safe_div(r["Parked_Notices_Generated"] + r["Parked_Others"], r["Electors"]), axis=1)
    return g


def notice_district_report(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df
    g = df.groupby("District", as_index=False).agg(
        ACs_Reporting=("AC_No", "nunique"),
        Electors=("Electors", "sum"),
        Notice_Generated=("Notice_Generated", "sum"),
        Notice_Delivered=("Notice_Delivered", "sum"),
        Notice_Pending_Delivery=("Notice_Pending_Delivery", "sum"),
        Hearing_Held=("Hearing_Held", "sum"),
        Hearing_Date_Lapsed=("Hearing_Date_Lapsed", "sum"),
        DEO_Total_Pending=("DEO_Total_Pending", "sum"),
        DEO_Pending_GT5=("DEO_Pending_GT5", "sum"),
        Ineligible_Final=("Ineligible_Final", "sum"),
        Parked_Notices_Generated=("Parked_Notices_Generated", "sum"),
        Parked_Others=("Parked_Others", "sum"),
    )
    g = _add_notice_pct_cols(g)
    return g.sort_values("Electors", ascending=False)


def notice_ac_report(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df
    g = df.groupby(["District", "AC_No", "AC_Name"], as_index=False).agg(
        Electors=("Electors", "sum"),
        Notice_Generated=("Notice_Generated", "sum"),
        Notice_Delivered=("Notice_Delivered", "sum"),
        Notice_Pending_Delivery=("Notice_Pending_Delivery", "sum"),
        Hearing_Held=("Hearing_Held", "sum"),
        Hearing_Date_Lapsed=("Hearing_Date_Lapsed", "sum"),
        DEO_Total_Pending=("DEO_Total_Pending", "sum"),
        DEO_Pending_GT5=("DEO_Pending_GT5", "sum"),
        Ineligible_Final=("Ineligible_Final", "sum"),
        Parked_Notices_Generated=("Parked_Notices_Generated", "sum"),
        Parked_Others=("Parked_Others", "sum"),
    )
    g = _add_notice_pct_cols(g)
    # Group rows by District (districts ordered by their own total electors,
    # highest first -- matching the District-wise report above), and within
    # each district rank ACs from most electors to fewest.
    dist_total = g.groupby("District")["Electors"].transform("sum")
    return (g.assign(_dist_total=dist_total)
             .sort_values(["_dist_total", "District", "Electors"], ascending=[False, True, False])
             .drop(columns="_dist_total"))


def notice_total_row(dist_rep: pd.DataFrame) -> pd.DataFrame:
    """Grand-total row for the District-wise report -- summed from the
    per-district numbers (not read off a sheet total), with % Lapsed and
    % Parked for Final recomputed from those summed totals rather than
    averaged, so the Total row's percentages are internally consistent."""
    if dist_rep.empty:
        return dist_rep
    sum_cols = ["Electors", "Notice_Generated", "Notice_Delivered", "Notice_Pending_Delivery",
                "Hearing_Held", "Hearing_Date_Lapsed", "DEO_Total_Pending", "DEO_Pending_GT5",
                "Ineligible_Final", "Parked_Notices_Generated", "Parked_Others"]
    totals = {c: dist_rep[c].sum() for c in sum_cols if c in dist_rep.columns}
    totals["District"] = "Total"
    if "ACs_Reporting" in dist_rep.columns:
        totals["ACs_Reporting"] = dist_rep["ACs_Reporting"].sum()
    row = pd.DataFrame([totals])
    row = _add_notice_pct_cols(row)
    return row


# --------------------------------------------------------------------------
# UI helpers
# --------------------------------------------------------------------------
def inject_css():
    st.markdown(f"""
    <style>
        /* Force light native form controls (dropdowns, checkboxes) even when
           the viewer's OS/browser is set to dark mode -- a defense-in-depth
           layer independent of .streamlit/config.toml, whose theme block
           only takes effect from the exact path ".streamlit/config.toml". */
        html, body, .stApp, [data-testid="stAppViewContainer"], [data-testid="stHeader"] {{
            color-scheme: light !important;
        }}

        .stApp {{
            background: linear-gradient(165deg, #EAF1FC 0%, #F5F8FD 45%, #EDF2FA 100%) !important;
        }}
        [data-testid="stHeader"] {{ background: rgba(0,0,0,0) !important; }}
        #MainMenu, footer {{ visibility: hidden; }}
        .block-container {{ padding-top: 1.2rem; padding-bottom: 2rem; max-width: 1400px; }}

        /* Streamlit's own chrome (labels, widget text) forced to our light
           palette regardless of viewer theme. */
        section[data-testid="stSidebar"] {{
            background: rgba(255,255,255,0.72) !important;
            backdrop-filter: blur(14px) saturate(160%);
            -webkit-backdrop-filter: blur(14px) saturate(160%);
            border-right: 1px solid rgba(255,255,255,0.6);
        }}
        section[data-testid="stSidebar"] * {{ color: {BRAND_TEXT} !important; }}

        /* Multiselect selected-value "tags" -- restyle away from the
           default solid, saturated blue pill (which clashes with the rest
           of the light glassy design) to a soft tinted pill that matches
           the brand palette. */
        [data-testid="stMultiSelectTagsContainer"] span[data-tag] {{
            background-color: rgba(11,61,145,0.12) !important;
            border: 1px solid rgba(11,61,145,0.35) !important;
            border-radius: 999px !important;
            color: {BRAND_PRIMARY} !important;
        }}
        [data-testid="stMultiSelectTagsContainer"] span[data-tag] * {{
            color: {BRAND_PRIMARY} !important;
        }}
        [data-testid="stMultiSelectTagsContainer"] span[data-tag] svg {{
            stroke: {BRAND_PRIMARY} !important;
        }}
        [data-testid="stMultiSelectTagsContainer"] span[data-tag] button:hover {{
            background: rgba(11,61,145,0.22) !important; border-radius: 999px;
        }}

        /* Consistent gap between st.columns cards no matter how many columns
           are in the row (fixes KPI cards looking "stuck together"). */
        div[data-testid="stHorizontalBlock"] > div[data-testid="column"] {{
            padding: 0 0.5rem;
        }}
        div[data-testid="stHorizontalBlock"] > div[data-testid="column"]:first-child {{ padding-left: 0; }}
        div[data-testid="stHorizontalBlock"] > div[data-testid="column"]:last-child {{ padding-right: 0; }}
        div[data-testid="stHorizontalBlock"] {{ margin-bottom: 0.9rem; }}

        /* View switcher (replaces st.tabs() so only the active view's code,
           including its sidebar filters, runs each rerun) -- restyled from a
           plain radio group into a tab-like pill bar. Selectors verified
           against the actual rendered DOM (Streamlit 1.62's React-Aria
           radio markup: label[data-testid="stRadioOption"] with a
           data-selected="true" attribute on the active option). */
        div[data-testid="stRadio"] {{ margin-bottom: 0.4rem; }}
        div[data-testid="stRadioGroup"] {{
            gap: 0.5rem;
            background: rgba(255,255,255,0.55);
            padding: 0.35rem;
            border-radius: 12px;
            border: 1px solid rgba(255,255,255,0.7);
            width: fit-content;
        }}
        label[data-testid="stRadioOption"] {{
            background: transparent;
            border-radius: 8px;
            padding: 0.45rem 1.1rem !important;
            margin: 0 !important;
            transition: background 0.15s ease, color 0.15s ease;
            cursor: pointer;
        }}
        label[data-testid="stRadioOption"][data-selected="true"] {{
            background: {BRAND_PRIMARY} !important;
        }}
        label[data-testid="stRadioOption"][data-selected="true"] p {{
            color: #FFFFFF !important;
            font-weight: 600;
        }}
        /* Hide the small radio-dot circle, keep only the pill label text. */
        label[data-testid="stRadioOption"] > div > div > div:not([data-testid="stMarkdownContainer"]) {{
            display: none !important;
        }}
        .view-tab-spacer {{ margin-bottom: 0.6rem; }}
        /* NOTE: scoped to actual form-widget labels only (not a blanket
           p/span/li rule) -- an earlier, broader version of this rule also
           matched text inside .mis-header and turned the banner's own white
           text black. Keep this narrow. */
        label, .stSelectbox label, .stMultiSelect label, .stCheckbox label p {{
            color: {BRAND_TEXT} !important;
        }}

        .mis-header {{
            background: linear-gradient(120deg, {BRAND_PRIMARY} 0%, {BRAND_PRIMARY_DARK} 100%);
            color: #ffffff; padding: 1.4rem 1.8rem; border-radius: 16px;
            margin-bottom: 1.2rem; box-shadow: 0 8px 26px rgba(11,61,145,0.28);
            border: 1px solid rgba(255,255,255,0.18);
        }}
        .mis-header h1 {{
            margin: 0; font-size: 1.55rem; font-weight: 700; letter-spacing: 0.2px;
            color: #ffffff !important;
        }}
        .mis-header p {{
            margin: 0.3rem 0 0 0; opacity: 0.95; font-size: 0.92rem;
            color: #ffffff !important;
        }}
        .mis-badge {{
            display: inline-block; background: rgba(255,255,255,0.18);
            padding: 0.2rem 0.7rem; border-radius: 999px; font-size: 0.78rem;
            margin-top: 0.55rem; margin-right: 0.4rem; border: 1px solid rgba(255,255,255,0.3);
            color: #ffffff !important; font-weight: 600;
        }}

        /* Glassmorphism: translucent, blurred, soft-bordered cards. Each card
           also gets an inline colour accent (left border + tint wash, set
           per-card in kpi_card()) so the row reads as distinct, colourful
           tiles rather than one flat block. */
        .kpi-card {{
            background-color: rgba(255,255,255,0.6);
            backdrop-filter: blur(12px) saturate(160%);
            -webkit-backdrop-filter: blur(12px) saturate(160%);
            border-radius: 16px; padding: 1rem 1.1rem;
            border: 1px solid rgba(255,255,255,0.65);
            border-left-width: 5px; border-left-style: solid;
            box-shadow: 0 8px 22px rgba(20,30,60,0.08);
            height: 100%; transition: transform .15s ease, box-shadow .15s ease;
        }}
        .kpi-card:hover {{ transform: translateY(-2px); box-shadow: 0 12px 28px rgba(20,30,60,0.13); }}
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
            background: rgba(238,243,252,0.75); backdrop-filter: blur(6px);
            border-left: 4px solid {BRAND_PRIMARY};
            padding: 0.7rem 1rem; border-radius: 8px; font-size: 0.85rem;
            color: {BRAND_TEXT}; margin-bottom: 0.8rem;
        }}
        .table-caption {{
            font-size: 0.78rem; color: {BRAND_MUTED}; margin: -0.3rem 0 0.6rem 0;
        }}

        /* Custom glassy, theme-independent report tables (replaces the
           canvas-rendered st.dataframe grid, which cannot be styled with CSS
           and otherwise follows the viewer's dark/light mode regardless of
           our design). Horizontally scrollable so it stays usable on phones. */
        .glass-table-wrap {{
            overflow-x: auto; overflow-y: auto; max-height: 460px;
            border-radius: 14px; border: 1px solid rgba(255,255,255,0.65);
            background: rgba(255,255,255,0.55);
            backdrop-filter: blur(12px) saturate(160%);
            -webkit-backdrop-filter: blur(12px) saturate(160%);
            box-shadow: 0 8px 22px rgba(20,30,60,0.07);
            margin-bottom: 0.9rem;
        }}
        .glass-table {{ width: 100%; border-collapse: collapse; font-size: 0.86rem; min-width: 480px; }}
        .glass-table thead th {{
            position: sticky; top: 0; z-index: 2;
            background: {BRAND_PRIMARY}; color: #ffffff !important; text-align: left;
            padding: 0.6rem 0.85rem; font-weight: 600; font-size: 0.74rem;
            text-transform: uppercase; letter-spacing: 0.3px; white-space: nowrap;
        }}
        .glass-table tbody td {{
            padding: 0.55rem 0.85rem; border-bottom: 1px solid rgba(20,30,60,0.07);
            color: {BRAND_TEXT} !important; white-space: nowrap;
        }}
        .glass-table tbody tr:nth-child(even) {{ background: rgba(11,61,145,0.035); }}
        .glass-table tbody tr:hover {{ background: rgba(11,61,145,0.09); }}
        .rate-badge {{
            display: inline-block; padding: 0.15rem 0.6rem; border-radius: 999px;
            font-weight: 700; font-size: 0.78rem;
        }}
        .badge-good {{ background: rgba(31,138,112,0.16); color: #17715A; }}
        .badge-mid  {{ background: rgba(217,119,6,0.16); color: #A35004; }}
        .badge-bad  {{ background: rgba(192,57,43,0.16); color: #A5312A; }}

        /* Mobile: tighten padding and shrink type so cards/tables/header fit
           comfortably on a phone screen without horizontal overflow of the
           page itself (tables still scroll sideways inside their own box). */
        @media (max-width: 640px) {{
            .block-container {{ padding-left: 0.7rem; padding-right: 0.7rem; padding-top: 0.8rem; }}
            .mis-header {{ padding: 1rem 1.1rem; border-radius: 12px; }}
            .mis-header h1 {{ font-size: 1.15rem; }}
            .mis-header p {{ font-size: 0.82rem; }}
            .kpi-card {{ padding: 0.8rem 0.9rem; border-radius: 12px; }}
            .kpi-value {{ font-size: 1.28rem; }}
            .kpi-label {{ font-size: 0.68rem; }}
            .section-title {{ font-size: 0.95rem; margin: 1.1rem 0 0.5rem 0; }}
            .glass-table {{ font-size: 0.78rem; }}
            .glass-table thead th, .glass-table tbody td {{ padding: 0.45rem 0.6rem; }}
        }}
    </style>
    """, unsafe_allow_html=True)


def _hex_to_rgba(hex_color: str, alpha: float) -> str:
    hex_color = hex_color.lstrip("#")
    r, g, b = (int(hex_color[i:i + 2], 16) for i in (0, 2, 4))
    return f"rgba({r},{g},{b},{alpha})"


def kpi_card(col, label, value, sub=None, color=BRAND_PRIMARY):
    with col:
        tint = _hex_to_rgba(color, 0.12)
        sub_html = f'<div class="kpi-sub" style="color:{color}">{sub}</div>' if sub else ""
        st.markdown(f"""
        <div class="kpi-card" style="border-left-color:{color};
             background-image: linear-gradient(135deg, {tint} 0%, rgba(255,255,255,0) 70%);">
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


def _rate_badge_class(col: str, val) -> str:
    """Colour-code a rate/percent value for the on-screen tables so an
    officer can spot problem districts/ACs at a glance."""
    try:
        v = float(val)
    except (TypeError, ValueError):
        return ""
    if col in RATE_COLS_HIGHER_BETTER:
        if v >= 75:
            return "badge-good"
        if v >= 50:
            return "badge-mid"
        return "badge-bad"
    if col in RATE_COLS_LOWER_BETTER:
        if v <= 10:
            return "badge-good"
        if v <= 25:
            return "badge-mid"
        return "badge-bad"
    return ""


def render_html_table(df: pd.DataFrame, cols: list, formats: dict = None, caption: str = None,
                       labels: dict = None, total_row: dict = None):
    """Render a dataframe as a custom, theme-independent HTML table.

    Streamlit's st.dataframe() renders through a canvas-based grid component
    that ignores custom CSS and always follows the app's active theme -- that
    is the actual root cause of tables turning unreadable in dark mode no
    matter what CSS is added elsewhere. A plain HTML <table> rendered via
    st.markdown(unsafe_allow_html=True) is fully CSS-controllable and
    completely independent of the viewer's theme, and is wrapped here in a
    horizontally-scrollable "glass" container so it stays usable on a phone.

    labels: optional {column: header text} override for columns whose plain
    "replace underscore with space" name isn't the wording the table should
    show. total_row: optional {column: value} for a bold grand-total row
    appended after the regular rows (formatted with the same `formats`).
    """
    formats = formats or {}
    labels = labels or {}
    view = df[cols].copy()
    numeric_cols = set(view.select_dtypes(include="number").columns)

    def _label(c):
        return labels.get(c, c.replace("_", " "))

    thead = "".join(f"<th>{_label(c)}</th>" for c in cols)

    def _render_row(row_dict, bold=False):
        cells = []
        for c in cols:
            val = row_dict.get(c)
            if c in formats:
                try:
                    disp = formats[c].format(val)
                except (TypeError, ValueError):
                    disp = "" if pd.isna(val) else str(val)
            else:
                disp = "" if (val is None or (isinstance(val, float) and pd.isna(val))) else str(val)
            is_numeric = c in numeric_cols or isinstance(val, (int, float))
            align = "right" if is_numeric else "left"
            badge = _rate_badge_class(c, val) if (not bold and is_numeric) else ""
            style = f"text-align:{align}" + ("; font-weight:700" if bold else "")
            if badge:
                cells.append(f'<td style="{style}"><span class="rate-badge {badge}">{disp}</span></td>')
            else:
                cells.append(f'<td style="{style}">{disp}</td>')
        return "<tr>" + "".join(cells) + "</tr>"

    body_rows = [_render_row(row.to_dict()) for _, row in view.iterrows()]
    if total_row:
        body_rows.append(_render_row(total_row, bold=True))

    table_html = (
        '<div class="glass-table-wrap"><table class="glass-table">'
        f"<thead><tr>{thead}</tr></thead><tbody>{''.join(body_rows)}</tbody>"
        "</table></div>"
    )
    st.markdown(table_html, unsafe_allow_html=True)
    if caption:
        st.markdown(f'<div class="table-caption">{caption}</div>', unsafe_allow_html=True)


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
                      charts, district_cols=None, ac_cols=None, col_labels=None,
                      district_total_row=None):
    """Builds a professional MIS PDF report and returns bytes.

    col_labels: optional {column: header text} override (see render_html_table).
    district_total_row: optional {column: value} bold grand-total row appended
    to the District-wise table.
    """
    col_labels = col_labels or {}
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
    title_style = ParagraphStyle("MISTitle", parent=styles["Title"], fontSize=19,
                                  textColor=colors.HexColor(BRAND_PRIMARY_DARK), alignment=TA_LEFT)
    sub_style = ParagraphStyle("MISSub", parent=styles["Normal"], fontSize=11.5,
                                textColor=colors.HexColor(BRAND_MUTED), spaceAfter=4)
    h2_style = ParagraphStyle("MISH2", parent=styles["Heading2"], fontSize=14,
                               textColor=colors.HexColor(BRAND_PRIMARY_DARK), spaceBefore=8, spaceAfter=5)
    body_style = ParagraphStyle("MISBody", parent=styles["Normal"], fontSize=10, leading=14)
    filt_style = ParagraphStyle("MISFilt", parent=styles["Normal"], fontSize=10.5,
                                 textColor=colors.HexColor(BRAND_TEXT))

    def header_footer(canvas, doc):
        # Top brand band (title text + "Generated: ..." timestamp) and the
        # "Filters Applied" block are intentionally not drawn -- removed so
        # the District-wise Report starts higher up / fits on page 1.
        canvas.saveState()
        w, h = pagesize
        canvas.setFillColor(colors.HexColor(BRAND_MUTED))
        canvas.setFont("Helvetica", 8)
        canvas.drawString(12 * mm, 6 * mm, "Official Report - Uttarakhand SIR")
        canvas.drawRightString(w - 12 * mm, 6 * mm, f"Page {doc.page}")
        canvas.restoreState()

    doc = BaseDocTemplate(buf, pagesize=pagesize,
                           leftMargin=12 * mm, rightMargin=12 * mm,
                           topMargin=10 * mm, bottomMargin=10 * mm)
    frame = Frame(doc.leftMargin, doc.bottomMargin, doc.width, doc.height, id="normal")
    doc.addPageTemplates([PageTemplate(id="mis", frames=frame, onPage=header_footer)])

    story = []
    story.append(Paragraph(title, title_style))
    if filters_desc:
        story.append(Paragraph(f"<b>{filters_desc}</b>", filt_style))
    story.append(Paragraph(subtitle, sub_style))
    story.append(Spacer(1, 4))

    # A spacer is only inserted *between* sections below (never trailing
    # after the last one) -- a trailing Spacer that doesn't quite fit at the
    # bottom of a page gets pushed onto a fresh page by itself, which is what
    # was producing an almost-empty extra page even when everything actually
    # fit on page 1.
    section_added = False

    def _section_gap():
        if section_added:
            story.append(Spacer(1, 5))

    if kpis:
        _section_gap()
        story.append(Paragraph("<b>Key Performance Summary</b>", h2_style))
        rows = [["Metric", "Value"]]
        for k, v in kpis.items():
            rows.append([k, v])
        t = Table(rows, colWidths=[doc.width * 0.6, doc.width * 0.4])
        t.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor(BRAND_PRIMARY)),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("FONTSIZE", (0, 0), (-1, -1), 10.5),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#EEF3FC")]),
            ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#D5DCE8")),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("TOPPADDING", (0, 0), (-1, -1), 5),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ]))
        story.append(t)
        section_added = True

    NARROW_COLS = {"AC_No", "Parts", "District_No"}
    MEDIUM_COLS = {"ACs_Reporting"}  # short numbers, but a longer header ("ACs Reporting")
    WIDE_COLS = {"District", "AC_Name"}
    cell_style = ParagraphStyle("MISCell", fontName="Helvetica", fontSize=9.5, leading=12)
    cell_style_r = ParagraphStyle("MISCellR", parent=cell_style, alignment=2)  # right-align
    header_cell_style = ParagraphStyle("MISCellH", fontName="Helvetica-Bold", fontSize=9.7,
                                        leading=12, textColor=colors.white)

    def _fmt_cell(val, col):
        if isinstance(val, float):
            if "%" in col:
                return f"{val:.2f}"
            return f"{val:,.0f}" if val == int(val) else f"{val:,.1f}"
        if isinstance(val, int):
            return f"{val:,}"
        return str(val)

    total_row_style = ParagraphStyle("MISCellTotal", parent=cell_style, fontName="Helvetica-Bold")
    total_row_style_r = ParagraphStyle("MISCellTotalR", parent=cell_style_r, fontName="Helvetica-Bold")

    def df_to_table(df, cols, max_rows=None, total_row=None):
        cols = [c for c in cols if c in df.columns]
        show = df[cols].head(max_rows).copy() if max_rows else df[cols].copy()

        weights = []
        for c in cols:
            if c in NARROW_COLS:
                weights.append(0.55)
            elif c in MEDIUM_COLS:
                weights.append(0.95)
            elif c in WIDE_COLS:
                weights.append(1.7)
            else:
                weights.append(1.0)
        total_w = sum(weights)
        col_widths = [doc.width * w / total_w for w in weights]

        header = [Paragraph(col_labels.get(c, c.replace("_", " ")), header_cell_style) for c in cols]
        data_rows = []
        for _, row in show.iterrows():
            cells = []
            for c in cols:
                text = _fmt_cell(row[c], c)
                style = cell_style if c in WIDE_COLS or c == "District" else cell_style_r
                cells.append(Paragraph(text, style))
            data_rows.append(cells)
        data = [header] + data_rows

        style_cmds = [
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor(BRAND_PRIMARY_DARK)),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F4F6FA")]),
            ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#D5DCE8")),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("TOPPADDING", (0, 0), (-1, -1), 3.5),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 3.5),
            ("LEFTPADDING", (0, 0), (-1, -1), 5),
            ("RIGHTPADDING", (0, 0), (-1, -1), 5),
        ]

        if total_row:
            trow = []
            for c in cols:
                val = total_row.get(c)
                text = _fmt_cell(val, c) if val is not None else ""
                style = total_row_style if c in WIDE_COLS or c == "District" else total_row_style_r
                trow.append(Paragraph(text, style))
            data.append(trow)
            last = len(data) - 1
            style_cmds.append(("BACKGROUND", (0, last), (-1, last), colors.HexColor("#DCE6F5")))
            style_cmds.append(("LINEABOVE", (0, last), (-1, last), 1, colors.HexColor(BRAND_PRIMARY_DARK)))

        t = Table(data, colWidths=col_widths, repeatRows=1)
        t.setStyle(TableStyle(style_cmds))
        return t

    if district_df is not None and not district_df.empty and district_cols:
        _section_gap()
        story.append(Paragraph("<b>District-wise Report</b>", h2_style))
        story.append(df_to_table(district_df, district_cols, total_row=district_total_row))
        section_added = True

    if ac_df is not None and not ac_df.empty and ac_cols:
        _section_gap()
        story.append(Paragraph("<b>AC-wise Report</b>", h2_style))
        story.append(df_to_table(ac_df, ac_cols))
        section_added = True

    if charts:
        _section_gap()
        story.append(Paragraph("<b>Charts</b>", h2_style))
        for i, (chart_title, fig) in enumerate(charts):
            png = _fig_to_png(fig)
            if png:
                if i > 0:
                    story.append(Spacer(1, 8))
                img = Image(io.BytesIO(png), width=doc.width, height=doc.width * 0.45)
                story.append(Paragraph(chart_title, body_style))
                story.append(img)
        section_added = True

    doc.build(story)
    buf.seek(0)
    return buf.getvalue()


# --------------------------------------------------------------------------
# App
# --------------------------------------------------------------------------
st.set_page_config(page_title="Form Processing & Notice-Hearing Report",
                    page_icon="\U0001F5F3", layout="wide",
                    initial_sidebar_state="expanded")
inject_css()

st.markdown(f"""
<div class="mis-header">
    <h1>Form Processing &amp; Notice-Hearing Report Dashboard</h1>
    <p>Special Intensive Revision (SIR) &mdash; Uttarakhand &nbsp;|&nbsp; District &amp; AC-level monitoring</p>
    <span class="mis-badge">Data-driven KPIs</span>
    <span class="mis-badge">Live filters</span>
    <span class="mis-badge">PDF &amp; Excel export</span>
</div>
""", unsafe_allow_html=True)

nh_df, nh_meta, nh_err = load_notice_data(NOTICE_FILE, ELECTORS_FILE)
ac_map = build_ac_district_map(nh_df) if nh_df is not None else {}
fp_df, fp_meta, fp_err = load_form_processing(FORM_PROCESSING_FILE, ac_map)
fp_old_df, fp_old_meta, fp_old_err = load_form_processing_district(FORM_PROCESSING_OLD_FILE)

with st.sidebar:
    st.markdown("### \U0001F4CB Dashboard Controls")
    st.caption("Filters apply live to KPIs, charts, reports and exports.")

VIEW_LABELS = {"fp": "\U0001F4C4  Form Processing", "nh": "\U0001F4E8  Notice & Hearing",
               "diff": "\U0001F4CA  Difference Report"}
active_view = st.radio(
    "View", list(VIEW_LABELS.keys()), format_func=lambda k: VIEW_LABELS[k],
    horizontal=True, key="active_view", label_visibility="collapsed",
)
st.markdown('<div class="view-tab-spacer"></div>', unsafe_allow_html=True)

# Only the active view's branch runs each rerun (unlike st.tabs(), where both
# tab bodies execute every time regardless of which tab is visually shown) --
# this is what keeps the sidebar filters scoped to whichever view is active,
# instead of both "Form Processing Filters" and "Notice & Hearing Filters"
# always appearing together.

# ==========================================================================
# VIEW: FORM PROCESSING
# ==========================================================================
if active_view == "fp":
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
                "Total Received": "Total_Received", "Unprocessed": "Unprocessed",
                "In Progress": "In_Progress",
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
        if fp_sel_forms: filt_parts.append(", ".join(FORM_TYPE_SHORT.get(f, f) for f in fp_sel_forms))
        fp_filters_desc = " | ".join(filt_parts)

        if filtered.empty:
            no_data_message()
        else:
            total_received = filtered["Total_Received"].sum()
            unprocessed = filtered["Unprocessed"].sum()
            in_progress = filtered["In_Progress"].sum()
            eroll_inclusion = filtered["Eroll_Inclusion"].sum()
            rejected = filtered["Rejected"].sum()
            accepted = filtered["Accepted"].sum()
            hearing_sched = filtered["Hearing_Scheduled"].sum()

            section_title("Key Performance Indicators")
            c1, c2, c3, c4 = st.columns(4, gap="medium")
            kpi_card(c1, "Total Forms Received", fmt_int(total_received),
                     f"{filtered['AC_No'].nunique()} ACs · {filtered['Form_Type'].nunique()} form types")
            kpi_card(c2, "Unprocessed", fmt_int(unprocessed),
                     f"{fmt_pct(safe_div(unprocessed, total_received))} of received", color=BRAND_DANGER)
            kpi_card(c3, "In Progress", fmt_int(in_progress),
                     f"{fmt_pct(safe_div(in_progress, total_received))} of received", color=BRAND_WARN)
            kpi_card(c4, "Eroll Inclusion", fmt_int(eroll_inclusion),
                     f"{fmt_pct(safe_div(eroll_inclusion, total_received))} inclusion rate", color=BRAND_ACCENT)

            c5, c6, c7 = st.columns(3, gap="medium")
            kpi_card(c5, "Rejected", fmt_int(rejected),
                     f"{fmt_pct(safe_div(rejected, total_received))} rejection rate", color=BRAND_DANGER)
            kpi_card(c6, "Accepted", fmt_int(accepted),
                     f"{fmt_pct(safe_div(accepted, total_received))} of received")
            kpi_card(c7, "Hearing Scheduled", fmt_int(hearing_sched),
                     f"{fmt_pct(safe_div(hearing_sched, total_received))} of received", color=BRAND_WARN)

            st.caption(
                "ℹ️ **In Progress** is a calculated total, not a column in Form_Processing.xlsx -- "
                "it's every status column except Unprocessed/Accepted/Rejected/Eroll Inclusion, added "
                "together (forms still moving through the workflow). Every number that goes into it is "
                "visible as its own column in the reports below, or via “Add more columns”."
            )

            if SHOW_CHARTS:
                section_title("Visual Analysis")
                status_sums = filtered[FP_STATUS_COLS].sum()
                status_sums = status_sums[status_sums > 0].sort_values(ascending=False)
                status_labels = [s.replace("_", " ") for s in status_sums.index]
                v1, v2 = st.columns(2)
                with v1:
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
                    dist_cmp = filtered.groupby("District")[["Unprocessed", "In_Progress"]].sum()
                    dist_cmp = dist_cmp.loc[dist_cmp.sum(axis=1).sort_values(ascending=False).index]
                    fig = go.Figure()
                    for col, color in zip(["Unprocessed", "In_Progress"],
                                           [BRAND_DANGER, BRAND_WARN]):
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
            fp_dist_display_cols = FP_DIST_BASE_COLS + FP_DIST_EXTRA_COLS
            sort_opt = st.selectbox("Sort district report by", fp_dist_display_cols[1:],
                                     index=fp_dist_display_cols.index("Total_Received") - 1, key="fp_dist_sort")
            fp_dist_view = fp_dist_rep[fp_dist_display_cols].sort_values(sort_opt, ascending=False)
            fp_dist_extra_pick = st.multiselect(
                "Add more columns (optional) -- every column below is exactly as named in Form_Processing.xlsx",
                FP_DIST_EXTRA_COLS, default=[], key="fp_dist_extra_cols",
                format_func=lambda c: c.replace("_", " "),
            )
            render_html_table(
                fp_dist_view, FP_DIST_BASE_COLS + fp_dist_extra_pick,
                formats=FP_COL_FORMATS,
                caption="Columns match Form_Processing.xlsx exactly (In Progress/Inclusion Rate are optional, "
                        "derived totals -- add them above). Full column set is also in the Excel/PDF export below.",
            )

            if SHOW_CHARTS:
                fig = px.bar(fp_dist_view, x="District", y=["Total_Received", "Eroll_Inclusion"],
                             barmode="group", title="District Comparison: Received vs Eroll Inclusion",
                             color_discrete_sequence=CHART_COLORWAY, labels={"value": "Forms", "variable": ""})
                st.plotly_chart(apply_plotly_theme(fig, height=360), use_container_width=True)

            # ------------------ AC-wise report ------------------
            section_title("AC-wise Report")
            ac_dist_pick = st.selectbox("Select District to drill into its ACs (optional)",
                                         ["All Districts"] + sorted(filtered["District"].unique()),
                                         key="fp_ac_district_pick")
            ac_scope_df = filtered if ac_dist_pick == "All Districts" else filtered[filtered["District"] == ac_dist_pick]
            fp_ac_rep = fp_ac_report(ac_scope_df)
            fp_ac_display_cols = FP_AC_BASE_COLS + FP_AC_EXTRA_COLS
            fp_ac_extra_pick = st.multiselect(
                "Add more columns (optional) -- every column below is exactly as named in Form_Processing.xlsx",
                FP_AC_EXTRA_COLS, default=[], key="fp_ac_extra_cols",
                format_func=lambda c: c.replace("_", " "),
            )
            render_html_table(
                fp_ac_rep, FP_AC_BASE_COLS + fp_ac_extra_pick,
                formats=FP_COL_FORMATS,
                caption="Columns match Form_Processing.xlsx exactly (In Progress/Inclusion Rate are optional, "
                        "derived totals -- add them above). Full column set is also in the Excel/PDF export below.",
            )

            if SHOW_CHARTS:
                ac_chart_df = fp_ac_rep.sort_values("Total_Received", ascending=False).head(20)
                fig = px.bar(ac_chart_df, x="AC_Name", y=["Unprocessed", "In_Progress"],
                             barmode="stack", title=f"AC Comparison ({ac_dist_pick}) - Top 20 by Volume",
                             color_discrete_sequence=CHART_COLORWAY, labels={"value": "Forms", "variable": ""})
                fig.update_xaxes(tickangle=-40)
                st.plotly_chart(apply_plotly_theme(fig, height=380), use_container_width=True)

            # ------------------ Downloads ------------------
            section_title("Downloads")
            dcol1, dcol2 = st.columns(2)
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
                    "Unprocessed": f"{fmt_int(unprocessed)} ({fmt_pct(safe_div(unprocessed, total_received))})",
                    "In Progress": f"{fmt_int(in_progress)} ({fmt_pct(safe_div(in_progress, total_received))})",
                    "Eroll Inclusion": f"{fmt_int(eroll_inclusion)} ({fmt_pct(safe_div(eroll_inclusion, total_received))})",
                    "Rejected": f"{fmt_int(rejected)} ({fmt_pct(safe_div(rejected, total_received))})",
                    "Hearing Scheduled": fmt_int(hearing_sched),
                }
                try:
                    pdf_bytes = build_pdf_report(
                        title="Form Processing Report",
                        subtitle=f"Report period: {fp_meta.get('report_period') or 'N/A'}",
                        filters_desc=fp_filters_desc, kpis=kpis_for_pdf,
                        district_df=fp_dist_view, ac_df=fp_ac_rep,
                        charts=[],
                        # PDF stays print-friendly with the same base columns
                        # shown on screen by default (a PDF page can't fit all
                        # ~27 columns readably); the full set is in the Excel
                        # export below and on screen via "Add more columns".
                        district_cols=FP_DIST_BASE_COLS, ac_cols=FP_AC_BASE_COLS,
                    )
                    st.download_button("\U0001F4C4 Download PDF Report (full)", pdf_bytes,
                                        file_name="Form_Processing_Report.pdf",
                                        mime="application/pdf", use_container_width=True)
                except Exception as exc:  # noqa: BLE001
                    st.error(f"PDF generation failed: {exc}")

            st.caption("District-wise and AC-wise reports on their own, in Excel or PDF:")
            dcol3, dcol4, dcol5, dcol6 = st.columns(4)
            with dcol3:
                dist_excel = build_excel_download({"District Report": fp_dist_rep})
                st.download_button("\U0001F4E5 District-wise (Excel)", dist_excel,
                                    file_name="Form_Processing_District_Report.xlsx",
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                    use_container_width=True, key="fp_dl_dist_xlsx")
            with dcol4:
                try:
                    dist_pdf_bytes = build_pdf_report(
                        title="Form Processing Report - District-wise",
                        subtitle=f"Report period: {fp_meta.get('report_period') or 'N/A'}",
                        filters_desc=fp_filters_desc, kpis=None,
                        district_df=fp_dist_view, ac_df=None, charts=[],
                        district_cols=FP_DIST_BASE_COLS,
                    )
                    st.download_button("\U0001F4C4 District-wise (PDF)", dist_pdf_bytes,
                                        file_name="Form_Processing_District_Report.pdf",
                                        mime="application/pdf", use_container_width=True, key="fp_dl_dist_pdf")
                except Exception as exc:  # noqa: BLE001
                    st.error(f"PDF generation failed: {exc}")
            with dcol5:
                ac_excel = build_excel_download({"AC Report": fp_ac_rep})
                st.download_button("\U0001F4E5 AC-wise (Excel)", ac_excel,
                                    file_name="Form_Processing_AC_Report.xlsx",
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                    use_container_width=True, key="fp_dl_ac_xlsx")
            with dcol6:
                try:
                    ac_pdf_bytes = build_pdf_report(
                        title="Form Processing Report - AC-wise",
                        subtitle=f"Report period: {fp_meta.get('report_period') or 'N/A'}",
                        filters_desc=fp_filters_desc, kpis=None,
                        district_df=None, ac_df=fp_ac_rep, charts=[],
                        ac_cols=FP_AC_BASE_COLS,
                    )
                    st.download_button("\U0001F4C4 AC-wise (PDF)", ac_pdf_bytes,
                                        file_name="Form_Processing_AC_Report.pdf",
                                        mime="application/pdf", use_container_width=True, key="fp_dl_ac_pdf")
                except Exception as exc:  # noqa: BLE001
                    st.error(f"PDF generation failed: {exc}")

# ==========================================================================
# VIEW: NOTICE & HEARING
# ==========================================================================
elif active_view == "nh":
    if nh_err:
        st.error(f"**Notice & Hearing data could not be loaded.**\n\n{nh_err}")
    elif nh_df is None or nh_df.empty:
        st.warning("Notice.xlsx / Electors.xlsx loaded but contain no usable data rows.")
    else:
        unmatched_note = ""
        if nh_meta and nh_meta.get("unmatched_acs"):
            unmatched_note = (f" <span style=\"color:{BRAND_DANGER}\">"
                               f"({nh_meta['unmatched_acs']} AC(s) in Notice.xlsx could not be matched "
                               f"to a District in Electors.xlsx.)</span>")
        st.markdown(f"""<div class="note-box">
            <b>Granularity:</b> one row per Assembly Constituency (AC)
            ({fmt_int(len(nh_df))} ACs across {nh_df['District'].nunique()} districts).
            <br><span style="color:{BRAND_MUTED}">Built from Notice.xlsx (notice/hearing/DEO/ERO counts) and
            Electors.xlsx (elector totals + District mapping). No date field exists in these workbooks,
            so no date filter/trend is shown here.</span>{unmatched_note}
            </div>""", unsafe_allow_html=True)

        with st.sidebar:
            st.markdown("#### \U0001F4E8 Notice & Hearing Filters")
            nh_districts = sorted(nh_df["District"].unique())
            nh_sel_districts = st.multiselect("District", nh_districts, default=[], key="nh_dist")
            nh_scope = nh_df[nh_df["District"].isin(nh_sel_districts)] if nh_sel_districts else nh_df
            nh_acs = sorted(nh_scope["AC_Name"].unique())
            nh_sel_acs = st.multiselect("Assembly Constituency (AC)", nh_acs, default=[], key="nh_ac")
            nh_pending_only = st.checkbox("Show only ACs with DEO-pending cases", value=False, key="nh_pending_only")

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
        if nh_pending_only: nh_filt_parts.append("Only ACs with DEO-pending cases")
        nh_filters_desc = " | ".join(nh_filt_parts)

        if nh_filtered.empty:
            no_data_message()
        else:
            electors = nh_filtered["Electors"].sum()
            notice_gen = nh_filtered["Notice_Generated"].sum()
            notice_del = nh_filtered["Notice_Delivered"].sum()
            hearing_held = nh_filtered["Hearing_Held"].sum()
            hearing_lapsed = nh_filtered["Hearing_Date_Lapsed"].sum()
            deo_pending = nh_filtered["DEO_Total_Pending"].sum()
            ineligible = nh_filtered["Ineligible_Final"].sum()
            parked_notices = nh_filtered["Parked_Notices_Generated"].sum()
            parked_others = nh_filtered["Parked_Others"].sum()
            parked_total = parked_notices + parked_others
            lapsed_pct = safe_div(hearing_lapsed, notice_del)
            parked_pct = safe_div(parked_total, electors)

            section_title("Key Performance Indicators")
            c1, c2, c3, c4 = st.columns(4, gap="medium")
            kpi_card(c1, "Electors", fmt_int(electors),
                     f"{nh_filtered['AC_No'].nunique()} ACs · {nh_filtered['District'].nunique()} districts")
            kpi_card(c2, "Notice Delivered", fmt_int(notice_del),
                     f"{fmt_pct(safe_div(notice_del, notice_gen))} of notices generated", color=BRAND_ACCENT)
            kpi_card(c3, "Hearing Held", fmt_int(hearing_held),
                     f"{fmt_pct(safe_div(hearing_held, notice_del))} of delivered notices")
            kpi_card(c4, "Hearing Date Lapsed", fmt_int(hearing_lapsed),
                     f"{fmt_pct(lapsed_pct)} of delivered notices", color=BRAND_DANGER)

            c5, c6, c7 = st.columns(3, gap="medium")
            kpi_card(c5, "Total Pending Text (DEO)", fmt_int(deo_pending),
                     f"{fmt_pct(safe_div(deo_pending, notice_gen))} of notices generated", color=BRAND_WARN)
            kpi_card(c6, "Found Ineligible for Final (ERO)", fmt_int(ineligible),
                     f"{fmt_pct(safe_div(ineligible, notice_gen))} of notices generated", color=BRAND_WARN)
            kpi_card(c7, "Parked for Final Publication", fmt_int(parked_total),
                     f"{fmt_pct(parked_pct)} of electors", color=BRAND_ACCENT)

            if SHOW_CHARTS:
                section_title("Visual Analysis")
                v1, v2 = st.columns(2)
                with v1:
                    fig = px.pie(names=["Delivered", "Hearing Date Lapsed"],
                                 values=[notice_del, hearing_lapsed], hole=0.5,
                                 title="Hearing Lapse Split", color_discrete_sequence=[BRAND_ACCENT, BRAND_DANGER])
                    st.plotly_chart(apply_plotly_theme(fig), use_container_width=True)
                with v2:
                    pend_df = nh_filtered.groupby("District")[["DEO_Total_Pending", "Ineligible_Final"]].sum()
                    pend_df = pend_df.loc[pend_df["DEO_Total_Pending"].sort_values(ascending=False).index]
                    fig = go.Figure()
                    fig.add_bar(name="DEO Pending", x=pend_df.index, y=pend_df["DEO_Total_Pending"], marker_color=BRAND_WARN)
                    fig.add_bar(name="Ineligible for Final", x=pend_df.index, y=pend_df["Ineligible_Final"], marker_color=BRAND_DANGER)
                    fig.update_layout(barmode="overlay", title="DEO Pendency & Ineligibility by District")
                    st.plotly_chart(apply_plotly_theme(fig), use_container_width=True)

                v3, v4 = st.columns(2)
                with v3:
                    cmp = nh_filtered.groupby("District")[["Notice_Generated", "Notice_Delivered", "Hearing_Held"]].sum()
                    cmp = cmp.loc[cmp["Notice_Generated"].sort_values(ascending=False).index]
                    fig = px.bar(cmp, x=cmp.index, y=["Notice_Generated", "Notice_Delivered", "Hearing_Held"],
                                 barmode="group", title="District Comparison: Generated / Delivered / Hearing Held",
                                 color_discrete_sequence=CHART_COLORWAY, labels={"value": "Count", "variable": ""})
                    st.plotly_chart(apply_plotly_theme(fig), use_container_width=True)
                with v4:
                    ac_lapsed = nh_filtered.groupby("AC_Name")["Hearing_Date_Lapsed"].sum().sort_values(ascending=False).head(15)
                    fig = px.bar(x=ac_lapsed.values, y=ac_lapsed.index, orientation="h",
                                 title="Top 15 ACs by Hearing Date Lapsed (count)",
                                 labels={"x": "Hearing Date Lapsed", "y": ""})
                    fig.update_traces(marker_color=BRAND_DANGER)
                    fig.update_yaxes(autorange="reversed")
                    st.plotly_chart(apply_plotly_theme(fig), use_container_width=True)

            # ------------------ District-wise report ------------------
            section_title("District-wise Report")
            nh_dist_rep = notice_district_report(nh_filtered)
            nh_dist_view = nh_dist_rep.sort_values("Electors", ascending=False)
            nh_dist_total = notice_total_row(nh_dist_rep)
            render_html_table(
                nh_dist_view, NOTICE_DIST_COLS,
                formats=NOTICE_COL_FORMATS, labels=NOTICE_COL_LABELS,
                total_row=nh_dist_total.iloc[0].to_dict() if not nh_dist_total.empty else None,
                caption="Columns match the requested report format exactly.",
            )

            if SHOW_CHARTS:
                fig = px.bar(nh_dist_view, x="District", y=["Notice_Delivered", "Hearing_Held"],
                             barmode="group", title="District Comparison: Notice Delivered vs Hearing Held",
                             color_discrete_sequence=CHART_COLORWAY, labels={"value": "Count", "variable": ""})
                st.plotly_chart(apply_plotly_theme(fig, height=360), use_container_width=True)

            # ------------------ AC-wise report ------------------
            section_title("AC-wise Report")
            nh_ac_dist_pick = st.selectbox("Select District to drill into its ACs (optional)",
                                            ["All Districts"] + sorted(nh_filtered["District"].unique()),
                                            key="nh_ac_district_pick")
            nh_ac_scope = nh_filtered if nh_ac_dist_pick == "All Districts" else nh_filtered[nh_filtered["District"] == nh_ac_dist_pick]
            nh_ac_rep = notice_ac_report(nh_ac_scope)
            render_html_table(
                nh_ac_rep, NOTICE_AC_COLS,
                formats=NOTICE_COL_FORMATS, labels=NOTICE_COL_LABELS,
                caption="Columns match the requested report format exactly.",
            )

            if SHOW_CHARTS:
                ac_chart_df2 = nh_ac_rep.sort_values("Electors", ascending=False).head(20)
                fig = px.bar(ac_chart_df2, x="AC_Name", y=["Notice_Delivered", "Hearing_Held"],
                             barmode="stack", title=f"AC Comparison ({nh_ac_dist_pick}) - Top 20 by Electors",
                             color_discrete_sequence=CHART_COLORWAY, labels={"value": "Count", "variable": ""})
                fig.update_xaxes(tickangle=-40)
                st.plotly_chart(apply_plotly_theme(fig, height=380), use_container_width=True)

            # ------------------ Downloads ------------------
            section_title("Downloads")
            dcol1, dcol2 = st.columns(2)
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
                    "Electors": fmt_int(electors),
                    "Notice Delivered": f"{fmt_int(notice_del)} ({fmt_pct(safe_div(notice_del, notice_gen))} of generated)",
                    "Hearing Held": f"{fmt_int(hearing_held)} ({fmt_pct(safe_div(hearing_held, notice_del))} of delivered)",
                    "Hearing Date Lapsed": f"{fmt_int(hearing_lapsed)} ({fmt_pct(lapsed_pct)} of delivered)",
                    "Total Pending Text (DEO)": fmt_int(deo_pending),
                    "Found Ineligible for Final (ERO)": fmt_int(ineligible),
                    "Parked for Final Publication": f"{fmt_int(parked_total)} ({fmt_pct(parked_pct)} of electors)",
                }
                try:
                    pdf_bytes2 = build_pdf_report(
                        title="Notice & Hearing Report",
                        subtitle=f"{fmt_int(len(nh_filtered))} Assembly Constituencies in scope",
                        filters_desc=nh_filters_desc, kpis=kpis_for_pdf2,
                        district_df=nh_dist_view, ac_df=nh_ac_rep,
                        charts=[],
                        district_cols=NOTICE_DIST_COLS, ac_cols=NOTICE_AC_COLS,
                        col_labels=NOTICE_COL_LABELS,
                        district_total_row=nh_dist_total.iloc[0].to_dict() if not nh_dist_total.empty else None,
                    )
                    st.download_button("\U0001F4C4 Download PDF Report (full)", pdf_bytes2,
                                        file_name="Notice_Hearing_Report.pdf",
                                        mime="application/pdf", use_container_width=True, key="nh_dl2")
                except Exception as exc:  # noqa: BLE001
                    st.error(f"PDF generation failed: {exc}")

            st.caption("District-wise and AC-wise reports on their own, in Excel or PDF:")
            dcol3, dcol4, dcol5, dcol6 = st.columns(4)
            with dcol3:
                dist_excel2 = build_excel_download({"District Report": nh_dist_rep})
                st.download_button("\U0001F4E5 District-wise (Excel)", dist_excel2,
                                    file_name="Notice_Hearing_District_Report.xlsx",
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                    use_container_width=True, key="nh_dl_dist_xlsx")
            with dcol4:
                try:
                    nh_dist_pdf_bytes = build_pdf_report(
                        title="Notice & Hearing Report - District-wise",
                        subtitle=f"{fmt_int(len(nh_filtered))} Assembly Constituencies in scope",
                        filters_desc=nh_filters_desc, kpis=None,
                        district_df=nh_dist_view, ac_df=None, charts=[],
                        district_cols=NOTICE_DIST_COLS, col_labels=NOTICE_COL_LABELS,
                        district_total_row=nh_dist_total.iloc[0].to_dict() if not nh_dist_total.empty else None,
                    )
                    st.download_button("\U0001F4C4 District-wise (PDF)", nh_dist_pdf_bytes,
                                        file_name="Notice_Hearing_District_Report.pdf",
                                        mime="application/pdf", use_container_width=True, key="nh_dl_dist_pdf")
                except Exception as exc:  # noqa: BLE001
                    st.error(f"PDF generation failed: {exc}")
            with dcol5:
                ac_excel2 = build_excel_download({"AC Report": nh_ac_rep})
                st.download_button("\U0001F4E5 AC-wise (Excel)", ac_excel2,
                                    file_name="Notice_Hearing_AC_Report.xlsx",
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                    use_container_width=True, key="nh_dl_ac_xlsx")
            with dcol6:
                try:
                    nh_ac_pdf_bytes = build_pdf_report(
                        title="Notice & Hearing Report - AC-wise",
                        subtitle=f"{fmt_int(len(nh_filtered))} Assembly Constituencies in scope",
                        filters_desc=nh_filters_desc, kpis=None,
                        district_df=None, ac_df=nh_ac_rep, charts=[],
                        ac_cols=NOTICE_AC_COLS, col_labels=NOTICE_COL_LABELS,
                    )
                    st.download_button("\U0001F4C4 AC-wise (PDF)", nh_ac_pdf_bytes,
                                        file_name="Notice_Hearing_AC_Report.pdf",
                                        mime="application/pdf", use_container_width=True, key="nh_dl_ac_pdf")
                except Exception as exc:  # noqa: BLE001
                    st.error(f"PDF generation failed: {exc}")

# ==========================================================================
# VIEW: DIFFERENCE REPORT
# ==========================================================================
else:
    if fp_err:
        st.error(f"**Form Processing data could not be loaded.**\n\n{fp_err}")
    elif fp_old_err:
        st.error(f"**Comparison workbook (1513.xlsx) could not be loaded.**\n\n{fp_old_err}")
    elif fp_df is None or fp_df.empty:
        st.warning("Form_Processing.xlsx loaded but contains no usable data rows.")
    elif fp_old_df is None or fp_old_df.empty:
        st.warning("1513.xlsx loaded but contains no usable data rows.")
    else:
        new_dist_full = fp_district_report(fp_df)
        old_dist_full = fp_district_report_flat(fp_old_df)
        diff_dist_full = fp_diff_report(new_dist_full, old_dist_full)

        st.markdown(f"""<div class="note-box">
            <b>Comparing:</b> Form_Processing.xlsx (period: {fp_meta.get('report_period') or 'N/A'})
            <b>minus</b> comparison workbook 1513.xlsx (period: {fp_old_meta.get('report_period') or 'N/A'}).
            <br><span style="color:{BRAND_MUTED}">Same columns as the Form Processing District-wise report --
            every value here is the District-wise change between the two report periods. District-level
            only: the comparison workbook has no AC-level breakdown to diff against.</span>
            </div>""", unsafe_allow_html=True)

        with st.sidebar:
            st.markdown("#### \U0001F4CA Difference Report Filters")
            diff_districts = sorted(diff_dist_full["District"].unique())
            diff_sel_districts = st.multiselect("District", diff_districts, default=[], key="diff_dist")

        diff_view = (diff_dist_full[diff_dist_full["District"].isin(diff_sel_districts)]
                     if diff_sel_districts else diff_dist_full)

        diff_filt_parts = []
        if diff_sel_districts: diff_filt_parts.append("District: " + ", ".join(diff_sel_districts))
        diff_filters_desc = " | ".join(diff_filt_parts)

        if diff_view.empty:
            no_data_message()
        else:
            total_diff = diff_view["Total_Received"].sum()
            hearing_diff = diff_view["Hearing_Scheduled"].sum()
            rejected_diff = diff_view["Rejected"].sum()
            accepted_diff = diff_view["Accepted"].sum()
            eroll_diff = diff_view["Eroll_Inclusion"].sum()
            unprocessed_diff = diff_view["Unprocessed"].sum()
            inprogress_diff = diff_view["In_Progress"].sum()

            section_title("Key Performance Indicators")
            c1, c2, c3, c4 = st.columns(4, gap="medium")
            kpi_card(c1, "Total Forms Received (Δ)", fmt_diff(total_diff),
                     f"{diff_view['District'].nunique()} districts in scope")
            kpi_card(c2, "Hearing Scheduled (Δ)", fmt_diff(hearing_diff), color=BRAND_WARN)
            kpi_card(c3, "Rejected (Δ)", fmt_diff(rejected_diff), color=BRAND_DANGER)
            kpi_card(c4, "Accepted (Δ)", fmt_diff(accepted_diff))

            c5, c6, c7 = st.columns(3, gap="medium")
            kpi_card(c5, "Eroll Inclusion (Δ)", fmt_diff(eroll_diff), color=BRAND_ACCENT)
            kpi_card(c6, "Unprocessed (Δ)", fmt_diff(unprocessed_diff), color=BRAND_WARN)
            kpi_card(c7, "In Progress (Δ)", fmt_diff(inprogress_diff))

            st.caption(
                "ℹ️ Every KPI and column here is (Form_Processing.xlsx value) minus "
                "(1513.xlsx value) for the same District -- a positive number means that column grew "
                "between the two report periods, negative means it shrank."
            )

            section_title("District-wise Difference Report")
            diff_sort_options = DIFF_DIST_BASE_COLS[1:] + DIFF_DIST_EXTRA_COLS
            diff_sort_opt = st.selectbox(
                "Sort district report by", diff_sort_options, index=0, key="diff_dist_sort",
                format_func=lambda c: DIFF_COL_LABELS.get(c, c.replace("_", " ")),
            )
            diff_sorted = diff_view.sort_values(diff_sort_opt, ascending=False)
            diff_extra_pick = st.multiselect(
                "Add more columns (optional) -- every column below is exactly as named in Form_Processing.xlsx",
                DIFF_DIST_EXTRA_COLS, default=[], key="diff_dist_extra_cols",
                format_func=lambda c: DIFF_COL_LABELS.get(c, c.replace("_", " ")),
            )
            render_html_table(
                diff_sorted, DIFF_DIST_BASE_COLS + diff_extra_pick,
                formats=DIFF_COL_FORMATS, labels=DIFF_COL_LABELS,
                caption="Values are District-wise differences (Form_Processing.xlsx minus 1513.xlsx), "
                        "not absolute counts. Full column set is in the Excel/PDF export below.",
            )

            # ------------------ Downloads ------------------
            section_title("Downloads")
            dcol1, dcol2 = st.columns(2)
            with dcol1:
                diff_excel = build_excel_download({"Difference Report": diff_dist_full})
                st.download_button("\U0001F4E5 Download Difference Report (Excel)", diff_excel,
                                    file_name="Form_Processing_Difference_Report.xlsx",
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                    use_container_width=True, key="diff_dl_xlsx")
            with dcol2:
                try:
                    diff_pdf_bytes = build_pdf_report(
                        title="Form Processing - Difference Report",
                        subtitle=(f"Change: {fp_old_meta.get('report_period') or 'N/A'}  "
                                  f"vs  {fp_meta.get('report_period') or 'N/A'}"),
                        filters_desc=diff_filters_desc, kpis=None,
                        district_df=diff_sorted, ac_df=None, charts=[],
                        district_cols=DIFF_DIST_BASE_COLS, col_labels=DIFF_COL_LABELS,
                    )
                    st.download_button("\U0001F4C4 Download Difference Report (PDF)", diff_pdf_bytes,
                                        file_name="Form_Processing_Difference_Report.pdf",
                                        mime="application/pdf", use_container_width=True, key="diff_dl_pdf")
                except Exception as exc:  # noqa: BLE001
                    st.error(f"PDF generation failed: {exc}")

st.markdown(f"""
<div style="text-align:center; color:{BRAND_MUTED}; font-size:0.78rem; margin-top:2rem; padding-top:1rem; border-top:1px solid #E4E8F0;">
    Form Processing &amp; Notice-Hearing Report Dashboard &middot; Data refreshed from workbook files in this repository &middot;
    Generated {datetime.now().strftime('%d-%b-%Y %H:%M')} IST
</div>
""", unsafe_allow_html=True)
