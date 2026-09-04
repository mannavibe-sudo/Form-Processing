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
# Part-level Notice/Hearing counts (one row per polling Part within an AC) --
# used only by the Part-wise Report tab's "bottom N parts per AC" drill-down.
PARTWISE_FILE = os.path.join(_APP_DIR, "Partwise.xlsx")

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
                     "Hearing_Held_%", "Hearing_Date_Lapsed", "Lapsed_%", "DEO_Total_Pending",
                     "Ineligible_Final", "Parked_Notices_Generated", "Parked_Others",
                     "Parked_Final_%"]
NOTICE_AC_COLS = ["District", "AC_No", "AC_Name", "Electors", "Notice_Delivered",
                   "Hearing_Held", "Hearing_Held_%", "Hearing_Date_Lapsed", "Lapsed_%", "DEO_Total_Pending",
                   "Ineligible_Final", "Parked_Notices_Generated", "Parked_Others",
                   "Parked_Final_%"]
# Header text for columns whose plain "replace underscore with space" label
# wouldn't match the exact wording the report is meant to use.
NOTICE_COL_LABELS = {
    "Notice_Delivered": "Notice Delivered",
    "Hearing_Held": "Hearing Held",
    "Hearing_Held_%": "% Hearing Held",
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
    "Hearing_Held_%": "{:.2f}",
    "Hearing_Date_Lapsed": "{:,.0f}", "Lapsed_%": "{:.2f}",
    "DEO_Total_Pending": "{:,.0f}", "Ineligible_Final": "{:,.0f}",
    "Parked_Notices_Generated": "{:,.0f}", "Parked_Others": "{:,.0f}",
    "Parked_Final_%": "{:.2f}",
}
# PDF-only: % Parked for Final figures under 80% are flagged in red text in
# the District-wise/AC-wise tables (on-screen tables are unaffected).
NOTICE_RED_BELOW = {"Parked_Final_%": 90}

# Part-wise Report: same column-naming convention/wording as the Notice &
# Hearing report above, but at Part granularity (one row per polling Part
# within an AC), for the "bottom N parts per AC" drill-down. Partwise.xlsx
# has no elector count per Part (only Electors.xlsx has electors, and only
# at AC level), so "Electors" and "% Parked for Final" -- which is defined
# as a share of Electors -- are not meaningful at this granularity and are
# left out here; everything else mirrors NOTICE_AC_COLS.
PARTWISE_COLS = ["Rank", "District", "AC_No", "AC_Name", "Part_No", "Notice_Delivered",
                 "Hearing_Held", "Hearing_Held_%", "Hearing_Date_Lapsed", "Lapsed_%",
                 "DEO_Total_Pending", "Ineligible_Final", "Parked_Notices_Generated",
                 "Parked_Others"]
PARTWISE_COL_LABELS = {
    "Rank": "Rank (1 = worst)",
    "Part_No": "Part No.",
    "Notice_Delivered": "Notice Delivered",
    "Hearing_Held": "Hearing Held",
    "Hearing_Held_%": "% Hearing Held",
    "Hearing_Date_Lapsed": "Hearing Date Lapsed",
    "Lapsed_%": "% Lapsed",
    "DEO_Total_Pending": "Total pending Text (DEO)",
    "Ineligible_Final": "Found Ineligible for Final (ERO)",
    "Parked_Notices_Generated": "Parked for Final Publication w.r.t. Notices Generated",
    "Parked_Others": "Parked For Final Publication w.r.t. Others",
}
PARTWISE_COL_FORMATS = {
    "Rank": "{:,.0f}", "Part_No": "{:,.0f}",
    "Notice_Delivered": "{:,.0f}", "Hearing_Held": "{:,.0f}", "Hearing_Held_%": "{:.2f}",
    "Hearing_Date_Lapsed": "{:,.0f}", "Lapsed_%": "{:.2f}",
    "DEO_Total_Pending": "{:,.0f}", "Ineligible_Final": "{:,.0f}",
    "Parked_Notices_Generated": "{:,.0f}", "Parked_Others": "{:,.0f}",
}
# "Bottom N parts per AC" ranking options -- label -> (column, ascending).
# ascending=True means the *lowest* values are "worst" (e.g. % Hearing Held);
# ascending=False means the *highest* values are "worst" (e.g. % Lapsed).
PARTWISE_RANK_OPTIONS = {
    "Found Ineligible for Final -- ERO (highest first)": ("Ineligible_Final", False),
    "% Hearing Held (lowest first)": ("Hearing_Held_%", True),
    "% Lapsed (highest first)": ("Lapsed_%", False),
    "Total pending Text -- DEO (highest first)": ("DEO_Total_Pending", False),
    "Notice Pending Delivery (highest first)": ("Notice_Pending_Delivery", False),
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
# Natural workflow order for the Form Type-wise breakdown report (6, 6A, 7, 8
# -- NOT alphabetical, which would wrongly put "FORM7" before "FORM6A").
FORM_TYPE_ORDER = {"FORM6": 0, "FORM6A": 1, "FORM7": 2, "FORM8": 3}

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
# Form Processing District-wise/AC-wise report -- Form Type-wise breakdown.
# Matches the official "Form Processed Current Status - All Forms" report
# layout exactly (per the user-supplied reference report): one row per
# District (or AC) x Form Type combination, a "District/AC Form Total"
# subtotal row after each group, and a final grand-Total row. This is the
# report shown by default -- it replaced the older single-row-per-district/
# AC summary (still available by selecting a single Form Type in the sidebar
# filter, which collapses each group to one form type's row + its subtotal).
# --------------------------------------------------------------------------
FP_FORMTYPE_DIST_COLS = ["District_No", "District", "Form_Type", "Total_Received",
                          "Hearing_Scheduled", "Rejected", "Accepted", "Eroll_Inclusion"]
FP_FORMTYPE_AC_COLS = ["District_No", "District", "AC_No", "AC_Name", "Form_Type",
                        "Total_Received", "Hearing_Scheduled", "Rejected", "Accepted", "Eroll_Inclusion"]
FP_FORMTYPE_DIST_EXTRA_COLS = (["ACs_Reporting"]
                                + [c for c in FP_STATUS_COLS if c not in FP_FORMTYPE_DIST_COLS]
                                + ["In_Progress", "Inclusion_Rate_%"])
FP_FORMTYPE_AC_EXTRA_COLS = ([c for c in FP_STATUS_COLS if c not in FP_FORMTYPE_AC_COLS]
                              + ["In_Progress", "Inclusion_Rate_%"])
# Header text matching the reference report's exact wording (plain
# "replace underscore with space" would give "District No"/"Total Received").
FP_FORMTYPE_COL_LABELS = {
    "District_No": "District No.",
    "District": "District Name",
    "Form_Type": "Form Type",
    "Total_Received": "Total Form Received",
}

# --------------------------------------------------------------------------
# Difference Report: same columns as the Form Processing District-wise/
# AC-wise reports, but the *values* are deltas between Form_Processing.xlsx
# (current period) and the comparison workbook (an earlier cutoff) -- both
# now AC-level workbooks with identical layout, so the diff is available at
# both District-wise and AC-wise granularity.
# --------------------------------------------------------------------------
DIFF_DIST_BASE_COLS = ["District", "Total_Received", "Hearing_Scheduled",
                        "Rejected", "Accepted", "Eroll_Inclusion"]
DIFF_DIST_EXTRA_COLS = ([c for c in FP_STATUS_COLS if c not in DIFF_DIST_BASE_COLS]
                         + ["In_Progress", "Inclusion_Rate_%"])
DIFF_AC_BASE_COLS = ["District", "AC_No", "AC_Name", "Total_Received",
                      "Hearing_Scheduled", "Rejected", "Accepted", "Eroll_Inclusion"]
DIFF_AC_EXTRA_COLS = ([c for c in FP_STATUS_COLS if c not in DIFF_AC_BASE_COLS]
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
def fmt_indian(n, decimals: int = 0) -> str:
    """Format a number with Indian digit grouping (last 3 digits, then
    groups of 2: 71,33,785 rather than the western 7,133,785) -- every
    report in this app is for an Indian (Uttarakhand SIR) audience, so all
    grouped numbers use this instead of Python's built-in ',' format spec,
    which always groups in 3s."""
    try:
        n = float(n)
    except (TypeError, ValueError):
        return "0"
    negative = n < 0
    formatted = f"{abs(n):.{decimals}f}"
    int_str, _, frac_str = formatted.partition(".")
    if len(int_str) <= 3:
        grouped = int_str
    else:
        last3 = int_str[-3:]
        rest = int_str[:-3]
        parts = []
        while len(rest) > 2:
            parts.insert(0, rest[-2:])
            rest = rest[:-2]
        if rest:
            parts.insert(0, rest)
        grouped = ",".join(parts) + "," + last3
    result = f"{grouped}.{frac_str}" if frac_str else grouped
    return ("-" if negative else "") + result


def fmt_int(n) -> str:
    return fmt_indian(n, 0)


def fmt_pct(n, decimals=1) -> str:
    try:
        return f"{float(n):.{decimals}f}%"
    except (TypeError, ValueError):
        return "0.0%"


def fmt_diff(n) -> str:
    """Like fmt_int, but always shows a sign (+1,234 / -1,234) -- used for
    the Difference Report, where the sign is the whole point."""
    try:
        v = float(n)
    except (TypeError, ValueError):
        return "0"
    body = fmt_indian(v, 0)  # already carries its own "-" when negative
    return body if v < 0 else f"+{body}"


def _apply_col_format(val, spec: str) -> str:
    """Render a value using one of the app's column format-spec strings.
    ",.0f"-style specs (plain or signed grouped integer) are special-cased
    to route through fmt_indian()/fmt_diff() for Indian digit grouping --
    Python's own str.format() only groups in 3s (western style). Every
    other spec (percentages, etc -- always under 100, so grouping doesn't
    apply) falls through to plain str.format()."""
    if spec == "{:,.0f}":
        return fmt_indian(val)
    if spec == "{:+,.0f}":
        return fmt_diff(val)
    return spec.format(val)


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

    # Row 1 is the sheet title, and now sometimes has a "Last Updated On :
    # <timestamp>" suffix appended to it (added directly in the source
    # workbook) -- pull that out here so it can be shown on the dashboard
    # and in the PDF export, same as the report-period line already is for
    # Form_Processing.xlsx.
    last_updated = None
    title_cell = clean_str(notice_raw.iat[0, 0]) if notice_raw.shape[0] > 0 else ""
    last_updated_match = re.search(r"Last Updated On\s*:\s*(.+)$", title_cell, re.IGNORECASE)
    if last_updated_match:
        last_updated = last_updated_match.group(1).strip()

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
    e["District_No"] = pd.to_numeric(e["District_No"], errors="coerce").fillna(0).astype(int)
    e["Electors"] = pd.to_numeric(e["Electors"], errors="coerce").fillna(0)
    e = e.drop_duplicates(subset=["AC_No"])

    df = n.merge(e[["AC_No", "District", "District_No", "Electors"]], on="AC_No", how="left")
    unmatched = int(df["District"].isna().sum())
    df["District"] = df["District"].fillna("Unassigned / AC Not Mapped")
    df["District_No"] = df["District_No"].fillna(0).astype(int)
    df["Electors"] = df["Electors"].fillna(0)

    meta = {"n_acs": len(df), "unmatched_acs": unmatched, "last_updated": last_updated}
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
def build_ac_districtno_map(notice_df: pd.DataFrame):
    """Derive AC -> District No. mapping (the fixed government district
    numbering from Electors.xlsx) -- used to order the Form Type-wise
    breakdown report exactly the way the official 'Form Processed Current
    Status' report is ordered (by District No., not by volume)."""
    if notice_df is None or notice_df.empty or "District_No" not in notice_df.columns:
        return {}
    m = notice_df[["AC_No", "District_No"]].drop_duplicates()
    return dict(zip(m["AC_No"].astype(int), m["District_No"].astype(int)))


@st.cache_data(show_spinner=False)
def load_form_processing(path: str, ac_district_map: dict, ac_districtno_map: dict = None):
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
    if ac_districtno_map:
        df["District_No"] = df["AC No."].map(ac_districtno_map).fillna(0).astype(int)
    else:
        df["District_No"] = 0
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


def fp_total_row(rep: pd.DataFrame) -> pd.DataFrame:
    """Grand-total row for the Form Processing District-wise/AC-wise report --
    summed from the per-row numbers (not read off a sheet total), with
    Inclusion_Rate_% recomputed from those summed totals rather than
    averaged, so the Total row's percentage is internally consistent. Works
    for both the District-wise report (has ACs_Reporting) and the AC-wise
    report (has AC_No/AC_Name instead -- left blank on the Total row)."""
    if rep is None or rep.empty:
        return pd.DataFrame()
    sum_cols = ["Total_Received"] + FP_STATUS_COLS + ["In_Progress"]
    totals = {c: rep[c].sum() for c in sum_cols if c in rep.columns}
    totals["District"] = "Total"
    if "ACs_Reporting" in rep.columns:
        totals["ACs_Reporting"] = rep["ACs_Reporting"].sum()
    if "AC_No" in rep.columns:
        totals["AC_No"] = ""
    if "AC_Name" in rep.columns:
        totals["AC_Name"] = ""
    row = pd.DataFrame([totals])
    row["Inclusion_Rate_%"] = row.apply(
        lambda r: safe_div(r["Eroll_Inclusion"], r["Total_Received"]), axis=1)
    return row


def fp_district_formtype_report(df: pd.DataFrame) -> pd.DataFrame:
    """District + Form Type breakdown, matching the official 'Form Processed
    Current Status - All Forms' report layout exactly: districts ordered by
    District No. (the fixed government numbering, NOT by volume), each
    district's own Form Type rows in natural order (6, 6A, 7, 8), followed by
    a 'District Form Total' subtotal row summed across that district's form
    types. The boolean '_is_subtotal' column (not meant for display) marks
    the subtotal rows so the caller can render them bold."""
    if df.empty:
        return df
    agg_kwargs = {"ACs_Reporting": ("AC_No", "nunique"), "Total_Received": ("Total_Received", "sum")}
    agg_kwargs.update({c: (c, "sum") for c in FP_STATUS_COLS})
    g = df.groupby(["District_No", "District", "Form_Type"], as_index=False).agg(**agg_kwargs)
    g["In_Progress"] = g[FP_INPROGRESS_COLS].sum(axis=1)
    g["Inclusion_Rate_%"] = g.apply(lambda r: safe_div(r["Eroll_Inclusion"], r["Total_Received"]), axis=1)
    g["_form_order"] = g["Form_Type"].map(FORM_TYPE_ORDER).fillna(99)
    g = g.sort_values(["District_No", "_form_order"]).drop(columns="_form_order")

    sum_cols = ["Total_Received"] + FP_STATUS_COLS + ["In_Progress"]
    chunks = []
    for (dno, dist), chunk in g.groupby(["District_No", "District"], sort=False):
        chunk = chunk.copy()
        chunk["_is_subtotal"] = False
        chunks.append(chunk)
        if len(chunk) <= 1:
            continue  # only one Form Type row -- a subtotal identical to it would be redundant
        subtotal = {c: chunk[c].sum() for c in sum_cols}
        subtotal["ACs_Reporting"] = chunk["ACs_Reporting"].sum()
        subtotal["District_No"] = dno
        subtotal["District"] = dist
        subtotal["Form_Type"] = "District Form Total"
        subtotal["Inclusion_Rate_%"] = safe_div(subtotal["Eroll_Inclusion"], subtotal["Total_Received"])
        subtotal["_is_subtotal"] = True
        chunks.append(pd.DataFrame([subtotal]))
    return pd.concat(chunks, ignore_index=True)


def fp_ac_formtype_report(df: pd.DataFrame) -> pd.DataFrame:
    """AC + Form Type breakdown -- same idea as fp_district_formtype_report()
    one level deeper: districts ordered by District No., ACs within a
    district ordered by AC No., each AC's Form Type rows in natural order (6,
    6A, 7, 8), followed by an 'AC Form Total' subtotal row. Same
    '_is_subtotal' marker convention."""
    if df.empty:
        return df
    agg_kwargs = {"Total_Received": ("Total_Received", "sum")}
    agg_kwargs.update({c: (c, "sum") for c in FP_STATUS_COLS})
    g = df.groupby(["District_No", "District", "AC_No", "AC_Name", "Form_Type"], as_index=False).agg(**agg_kwargs)
    g["In_Progress"] = g[FP_INPROGRESS_COLS].sum(axis=1)
    g["Inclusion_Rate_%"] = g.apply(lambda r: safe_div(r["Eroll_Inclusion"], r["Total_Received"]), axis=1)
    g["_form_order"] = g["Form_Type"].map(FORM_TYPE_ORDER).fillna(99)
    g = g.sort_values(["District_No", "AC_No", "_form_order"]).drop(columns="_form_order")

    sum_cols = ["Total_Received"] + FP_STATUS_COLS + ["In_Progress"]
    chunks = []
    for (dno, dist, acno, acname), chunk in g.groupby(["District_No", "District", "AC_No", "AC_Name"], sort=False):
        chunk = chunk.copy()
        chunk["_is_subtotal"] = False
        chunks.append(chunk)
        if len(chunk) <= 1:
            continue  # only one Form Type row -- a subtotal identical to it would be redundant
        subtotal = {c: chunk[c].sum() for c in sum_cols}
        subtotal["District_No"] = dno
        subtotal["District"] = dist
        subtotal["AC_No"] = acno
        subtotal["AC_Name"] = acname
        subtotal["Form_Type"] = "AC Form Total"
        subtotal["Inclusion_Rate_%"] = safe_div(subtotal["Eroll_Inclusion"], subtotal["Total_Received"])
        subtotal["_is_subtotal"] = True
        chunks.append(pd.DataFrame([subtotal]))
    return pd.concat(chunks, ignore_index=True)


def fp_diff_report(new_rep: pd.DataFrame, old_rep: pd.DataFrame, key_cols: list) -> pd.DataFrame:
    """Difference report: new_rep minus old_rep, aligned on key_cols (e.g.
    ["District"] for a District-wise diff, ["District", "AC_No", "AC_Name"]
    for an AC-wise diff). new_rep/old_rep are expected to already be the
    output of fp_district_report()/fp_ac_report() (or empty/None). Same
    value columns as those reports (Total_Received + every real status
    column, plus the derived In_Progress/Inclusion_Rate_% for reference;
    ACs_Reporting is excluded -- a diff of "how many distinct ACs reported"
    isn't a meaningful number). A row (district or AC) present on only one
    side is still included, with the missing side treated as zero, so
    nothing silently drops out of the diff."""
    diff_cols = ["Total_Received"] + FP_STATUS_COLS + ["In_Progress", "Inclusion_Rate_%"]
    if new_rep is None or new_rep.empty:
        new_rep = pd.DataFrame(columns=key_cols + diff_cols)
    if old_rep is None or old_rep.empty:
        old_rep = pd.DataFrame(columns=key_cols + diff_cols)

    n = new_rep.set_index(key_cols)
    o = old_rep.set_index(key_cols)
    all_keys = n.index.union(o.index)
    n = n.reindex(all_keys, fill_value=0)
    o = o.reindex(all_keys, fill_value=0)

    diff = pd.DataFrame(index=all_keys)
    for c in diff_cols:
        n_col = n[c] if c in n.columns else 0
        o_col = o[c] if c in o.columns else 0
        diff[c] = n_col - o_col
    diff = diff.reset_index()
    return diff.sort_values("Total_Received", ascending=False)


def _add_notice_pct_cols(g: pd.DataFrame) -> pd.DataFrame:
    """% Hearing Held, % Lapsed and % Parked for Final -- the derived columns
    in the Notice & Hearing report, computed the same way at every grouping
    level (district, AC, or the grand total row) so the numbers stay
    consistent:
      % Hearing Held = Hearing Held / Notice Delivered
      % Lapsed = Hearing Date Lapsed / Notice Delivered
      % Parked for Final = (Parked w.r.t. Notices Generated + Parked w.r.t.
                             Others) / Electors
    (safe_div already returns the ratio *100, e.g. 21.02 for 21.02%.)
    """
    g["Hearing_Held_%"] = g.apply(
        lambda r: safe_div(r["Hearing_Held"], r["Notice_Delivered"]), axis=1)
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


# Partwise.xlsx source-column -> internal-column mapping. Unlike Notice.xlsx
# / Electors.xlsx / Form_Processing.xlsx, this file has one plain header row
# (no title/scope rows, no merged header) so it can be read with pandas'
# normal header=0 and just renamed -- no positional-column workaround needed.
PARTWISE_RAW_RENAME = {
    "District Number": "District_No", "District Name": "District",
    "AC Number": "AC_No", "Asmbly Name": "AC_Name", "Part No": "Part_No",
    "Notice Generated": "Notice_Generated",
    "Pending for Notice Generation": "Pending_Notice_Generation",
    "Notice Delivered": "Notice_Delivered",
    "Notice Pending Delivery": "Notice_Pending_Delivery",
    "Hearings Held": "Hearing_Held",
    "Hearing Date Lapsed": "Hearing_Date_Lapsed",
    "Reschedule Date Lapsed": "Reschedule_Date_Lapsed",
    "DEO-Status Total Pending": "DEO_Total_Pending",
    "DEO-Status Pending GT 5 Days": "DEO_Pending_GT5",
    "DEO-Status Verified": "DEO_Verified",
    "DEO-Status Not Verified": "DEO_Not_Verified",
    "ERO/AERO Status Found Ineligible For Final w.r.t. Notice Generated": "Ineligible_Final",
    "ERO/AERO Status Parked For Final Publication": "Parked_Notices_Generated",
    "ERO/AERO Parked For Final Publication w.r.t. Others": "Parked_Others",
}


@st.cache_data(show_spinner=False)
def load_partwise_data(path: str):
    """Load Partwise.xlsx -- one row per polling Part within an AC, used only
    by the Part-wise Report tab's "bottom N parts per AC" drill-down.

    This workbook carries its own District No./District Name/AC No./AC Name
    per row (no separate Electors.xlsx merge needed here), but has no
    elector count per Part, so % Parked for Final (a share of Electors)
    cannot be computed at this granularity -- see PARTWISE_COLS.
    """
    try:
        raw = pd.read_excel(path, sheet_name=0)
    except FileNotFoundError:
        return None, None, f"File not found: {path}"
    except Exception as exc:  # noqa: BLE001
        return None, None, f"Could not read {path}: {exc}"

    missing = [c for c in PARTWISE_RAW_RENAME if c not in raw.columns]
    if missing:
        return None, None, ("Partwise.xlsx does not match the expected layout "
                             f"(missing column(s): {', '.join(missing)}).")

    df = raw.rename(columns=PARTWISE_RAW_RENAME).copy()
    df["District"] = df["District"].apply(clean_str)
    df["AC_Name"] = df["AC_Name"].apply(clean_str)
    id_cols = {"District", "AC_Name"}
    num_cols = ["District_No", "AC_No", "Part_No"] + [
        c for c in PARTWISE_RAW_RENAME.values() if c not in id_cols
        and c not in ("District_No", "AC_No", "Part_No")
    ]
    for c in num_cols:
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)
    df["District_No"] = df["District_No"].astype(int)
    df["AC_No"] = df["AC_No"].astype(int)
    df["Part_No"] = df["Part_No"].astype(int)
    df = df[df["Notice_Delivered"] > 0].copy()  # avoid divide-by-zero in the % columns below
    if df.empty:
        return None, None, "Partwise.xlsx has no usable Part-level rows (Notice Delivered is 0 everywhere)."

    df["Hearing_Held_%"] = df.apply(
        lambda r: safe_div(r["Hearing_Held"], r["Notice_Delivered"]), axis=1)
    df["Lapsed_%"] = df.apply(
        lambda r: safe_div(r["Hearing_Date_Lapsed"], r["Notice_Delivered"]), axis=1)

    meta = {"n_parts": len(df), "n_acs": df["AC_No"].nunique(), "n_districts": df["District"].nunique()}
    return df, meta, None


def partwise_bottom_n_report(df: pd.DataFrame, rank_col: str = "Hearing_Held_%",
                              ascending: bool = True, n: int = 10) -> pd.DataFrame:
    """The N worst-performing Parts within each AC, ranked by `rank_col`
    (ascending=True -- lowest values are worst, e.g. % Hearing Held;
    ascending=False -- highest values are worst, e.g. % Lapsed). Adds a
    "Rank" column (1 = worst) scoped to each AC, and orders the output by
    District No. then AC No. then Rank so it reads like the other AC-wise
    reports (grouped, in official District order)."""
    if df.empty:
        return df
    g = df.sort_values(["AC_No", rank_col], ascending=[True, ascending])
    out = g.groupby("AC_No", group_keys=False).head(n).copy()
    out["Rank"] = out.groupby("AC_No").cumcount() + 1
    return out.sort_values(["District_No", "AC_No", "Rank"])


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

        /* Compact variant: used for dense multi-row KPI grids (e.g. Form
           Processing's Total/Form6/Form7/Form8 breakdown) where many small
           cards need to fit without overwhelming the page. */
        .kpi-card-compact {{
            padding: 0.55rem 0.65rem; border-radius: 10px; border-left-width: 4px;
            box-shadow: 0 4px 12px rgba(20,30,60,0.06);
        }}
        .kpi-card-compact .kpi-label {{ font-size: 0.62rem; margin-bottom: 0.15rem; }}
        .kpi-card-compact .kpi-value {{ font-size: 1.05rem; line-height: 1.05; }}
        .kpi-row-title {{
            color: {BRAND_TEXT}; font-size: 0.85rem; font-weight: 700;
            margin: 0.65rem 0 0.3rem 0.1rem; opacity: 0.85;
        }}
        .kpi-row-title:first-of-type {{ margin-top: 0.1rem; }}

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
            .kpi-card-compact {{ padding: 0.45rem 0.5rem; border-radius: 8px; }}
            .kpi-card-compact .kpi-value {{ font-size: 0.88rem; }}
            .kpi-card-compact .kpi-label {{ font-size: 0.56rem; }}
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


def kpi_card(col, label, value, sub=None, color=BRAND_PRIMARY, compact=False):
    with col:
        tint = _hex_to_rgba(color, 0.12)
        sub_html = f'<div class="kpi-sub" style="color:{color}">{sub}</div>' if sub else ""
        card_class = "kpi-card kpi-card-compact" if compact else "kpi-card"
        st.markdown(f"""
        <div class="{card_class}" style="border-left-color:{color};
             background-image: linear-gradient(135deg, {tint} 0%, rgba(255,255,255,0) 70%);">
            <div class="kpi-label">{label}</div>
            <div class="kpi-value">{value}</div>
            {sub_html}
        </div>
        """, unsafe_allow_html=True)


def kpi_row(title, values, color=BRAND_PRIMARY):
    """Render one compact KPI row: a small row title followed by 5 metric
    cards (Receiving / Hearing / Reject / Accept / Inclusion)."""
    st.markdown(f'<div class="kpi-row-title">{title}</div>', unsafe_allow_html=True)
    cols = st.columns(5, gap="small")
    labels = ["Receiving", "Hearing", "Reject", "Accept", "Inclusion"]
    for c, lbl, val in zip(cols, labels, values):
        kpi_card(c, lbl, fmt_int(val), color=color, compact=True)


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
    If `df` itself has a boolean "_is_subtotal" column (not included in
    `cols`, so never displayed), those rows are rendered bold in place --
    used by the Form Type-wise breakdown report for its per-district/per-AC
    subtotal rows, interspersed among the regular rows rather than only at
    the very end.
    """
    formats = formats or {}
    labels = labels or {}
    view = df[cols].copy()
    subtotal_flags = df["_is_subtotal"] if "_is_subtotal" in df.columns else None
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
                    disp = _apply_col_format(val, formats[c])
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

    body_rows = [
        _render_row(row.to_dict(), bold=bool(subtotal_flags.loc[idx]) if subtotal_flags is not None else False)
        for idx, row in view.iterrows()
    ]
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
            # Internal-only helper columns (e.g. "_is_subtotal", used to mark
            # bold subtotal rows on screen/PDF) are never meant to reach the
            # exported spreadsheet.
            df = df.drop(columns=[c for c in df.columns if c.startswith("_")], errors="ignore")
            df.to_excel(writer, sheet_name=safe_name, index=False)
            if engine == "xlsxwriter":
                wb = writer.book
                ws = writer.sheets[safe_name]
                header_fmt = wb.add_format({"bold": True, "bg_color": BRAND_PRIMARY,
                                             "font_color": "white", "border": 1})
                pct_fmt = wb.add_format({"num_format": "0.0", "border": 1})
                # "#,##,##0" (not the western "#,##0") -- Excel's custom
                # number-format codes support arbitrary comma placement, and
                # this is the standard way to get Indian digit grouping
                # (71,33,785) natively in the spreadsheet, matching the
                # on-screen/PDF formatting elsewhere in this app.
                num_fmt = wb.add_format({"num_format": "#,##,##0", "border": 1})
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
                      district_total_row=None, ac_total_row=None, red_below=None,
                      kpi_groups=None, district_title="District-wise Report",
                      ac_title="AC-wise Report"):
    """Builds a professional MIS PDF report and returns bytes.

    col_labels: optional {column: header text} override (see render_html_table).
    district_total_row / ac_total_row: optional {column: value} bold
    grand-total row appended to the District-wise / AC-wise table respectively.
    red_below: optional {column: threshold} -- any District-wise/AC-wise table
    cell in that column (including the Total row) whose numeric value is below
    the threshold is rendered in red text, e.g. {"Parked_Final_%": 80} flags
    every % Parked for Final figure under 80%.
    kpi_groups: optional list of (group_title, {label: value}) tuples -- when
    given, this renders the Key Performance Summary as one small metrics row
    per group (e.g. Total / Form 6 / Form 7 / Form 8), mirroring the on-screen
    multi-row KPI grid, instead of the flat `kpis` grid. Takes precedence over
    `kpis` when both are supplied.
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
                               textColor=colors.HexColor(BRAND_PRIMARY_DARK), spaceBefore=6, spaceAfter=4)
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
        canvas.drawString(12 * mm, 5 * mm, "Official Report - Uttarakhand SIR")
        canvas.drawRightString(w - 12 * mm, 5 * mm, f"Page {doc.page}")
        canvas.restoreState()

    doc = BaseDocTemplate(buf, pagesize=pagesize,
                           leftMargin=12 * mm, rightMargin=12 * mm,
                           topMargin=8 * mm, bottomMargin=8 * mm)
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

    if kpi_groups:
        _section_gap()
        story.append(Paragraph("<b>Key Performance Summary</b>", h2_style))
        # One small metrics strip per group (Total / Form 6 / Form 7 / Form 8)
        # -- mirrors the on-screen multi-row KPI grid instead of one flat list.
        kpi_group_title_style = ParagraphStyle("MISKpiGroupTitle", fontName="Helvetica-Bold",
                                                fontSize=9.3, leading=11,
                                                textColor=colors.HexColor(BRAND_PRIMARY_DARK),
                                                spaceBefore=2, spaceAfter=1)
        kpi_label_style = ParagraphStyle("MISKpiLabel", fontName="Helvetica-Bold", fontSize=8.3,
                                          leading=10, textColor=colors.white)
        kpi_value_style = ParagraphStyle("MISKpiValue", fontName="Helvetica-Bold", fontSize=10.3,
                                          leading=12.5, textColor=colors.HexColor(BRAND_PRIMARY_DARK))
        for gi, (group_title, metrics) in enumerate(kpi_groups):
            items = list(metrics.items())
            if not items:
                continue
            col_w = doc.width / len(items)
            story.append(Paragraph(group_title, kpi_group_title_style))
            label_row = [Paragraph(str(k), kpi_label_style) for k, v in items]
            value_row = [Paragraph(str(v), kpi_value_style) for k, v in items]
            t = Table([label_row, value_row], colWidths=[col_w] * len(items))
            t.setStyle(TableStyle([
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor(BRAND_PRIMARY)),
                ("BACKGROUND", (0, 1), (-1, 1), colors.HexColor("#EEF3FC")),
                ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#D5DCE8")),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("TOPPADDING", (0, 0), (-1, 0), 3),
                ("BOTTOMPADDING", (0, 0), (-1, 0), 3),
                ("TOPPADDING", (0, 1), (-1, 1), 4),
                ("BOTTOMPADDING", (0, 1), (-1, 1), 5),
                ("LEFTPADDING", (0, 0), (-1, -1), 6),
                ("RIGHTPADDING", (0, 0), (-1, -1), 6),
            ]))
            story.append(t)
            if gi < len(kpi_groups) - 1:
                story.append(Spacer(1, 3))
        section_added = True
    elif kpis:
        _section_gap()
        story.append(Paragraph("<b>Key Performance Summary</b>", h2_style))
        # A wide grid (up to 4 metrics per row: label row + value row) instead
        # of a tall single-column list -- this uses far less page height, so
        # the District-wise table (with its Total row) has room to fit on
        # page 1 alongside it instead of spilling onto page 2.
        kpi_items = list(kpis.items())
        KPI_COLS_PER_ROW = 4
        kpi_label_style = ParagraphStyle("MISKpiLabel", fontName="Helvetica-Bold", fontSize=8.3,
                                          leading=10, textColor=colors.white)
        kpi_value_style = ParagraphStyle("MISKpiValue", fontName="Helvetica-Bold", fontSize=10.7,
                                          leading=13, textColor=colors.HexColor(BRAND_PRIMARY_DARK))
        col_w = doc.width / KPI_COLS_PER_ROW
        for i in range(0, len(kpi_items), KPI_COLS_PER_ROW):
            chunk = kpi_items[i:i + KPI_COLS_PER_ROW]
            label_row = [Paragraph(str(k), kpi_label_style) for k, v in chunk]
            value_row = [Paragraph(str(v), kpi_value_style) for k, v in chunk]
            while len(label_row) < KPI_COLS_PER_ROW:
                label_row.append("")
                value_row.append("")
            t = Table([label_row, value_row], colWidths=[col_w] * KPI_COLS_PER_ROW)
            t.setStyle(TableStyle([
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor(BRAND_PRIMARY)),
                ("BACKGROUND", (0, 1), (-1, 1), colors.HexColor("#EEF3FC")),
                ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#D5DCE8")),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("TOPPADDING", (0, 0), (-1, 0), 4),
                ("BOTTOMPADDING", (0, 0), (-1, 0), 4),
                ("TOPPADDING", (0, 1), (-1, 1), 5),
                ("BOTTOMPADDING", (0, 1), (-1, 1), 6),
                ("LEFTPADDING", (0, 0), (-1, -1), 6),
                ("RIGHTPADDING", (0, 0), (-1, -1), 6),
            ]))
            story.append(t)
            if i + KPI_COLS_PER_ROW < len(kpi_items):
                story.append(Spacer(1, 2))
        section_added = True

    NARROW_COLS = {"AC_No", "Parts", "District_No", "Part_No", "Rank"}
    MEDIUM_COLS = {"ACs_Reporting"}  # short numbers, but a longer header ("ACs Reporting")
    WIDE_COLS = {"District", "AC_Name", "Form_Type"}
    cell_style = ParagraphStyle("MISCell", fontName="Helvetica", fontSize=9.5, leading=12)
    cell_style_r = ParagraphStyle("MISCellR", parent=cell_style, alignment=2)  # right-align
    # Red variants -- used for red_below flagged cells (e.g. % Parked for
    # Final under 80%) so the low figure stands out in the printed report.
    cell_style_red = ParagraphStyle("MISCellRed", parent=cell_style, textColor=colors.HexColor("#C0392B"))
    cell_style_red_r = ParagraphStyle("MISCellRedR", parent=cell_style_r, textColor=colors.HexColor("#C0392B"))
    header_cell_style = ParagraphStyle("MISCellH", fontName="Helvetica-Bold", fontSize=9.7,
                                        leading=12, textColor=colors.white)

    def _fmt_cell(val, col):
        if isinstance(val, float):
            if "%" in col:
                return f"{val:.2f}"
            return fmt_indian(val, 0) if val == int(val) else fmt_indian(val, 1)
        if isinstance(val, int):
            return fmt_indian(val, 0)
        return str(val)

    total_row_style = ParagraphStyle("MISCellTotal", parent=cell_style, fontName="Helvetica-Bold")
    total_row_style_r = ParagraphStyle("MISCellTotalR", parent=cell_style_r, fontName="Helvetica-Bold")
    total_row_style_red_r = ParagraphStyle("MISCellTotalRedR", parent=total_row_style_r,
                                            textColor=colors.HexColor("#C0392B"))

    def _below_threshold(val, col, red_below):
        if not red_below or col not in red_below:
            return False
        try:
            return float(val) < red_below[col]
        except (TypeError, ValueError):
            return False

    def df_to_table(df, cols, max_rows=None, total_row=None, red_below=None):
        # "_is_subtotal" (if present) marks rows to render bold + tinted --
        # the Form Type-wise breakdown report's per-district/per-AC subtotal
        # rows, interspersed among the regular rows rather than only at the
        # very end (unlike total_row, which is always the final row).
        subtotal_flags = df["_is_subtotal"] if "_is_subtotal" in df.columns else None
        cols = [c for c in cols if c in df.columns]
        show = df[cols].head(max_rows).copy() if max_rows else df[cols].copy()
        if subtotal_flags is not None:
            subtotal_flags = subtotal_flags.loc[show.index]

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
        subtotal_row_nums = []
        for i, (_, row) in enumerate(show.iterrows()):
            is_subtotal = bool(subtotal_flags.iloc[i]) if subtotal_flags is not None else False
            if is_subtotal:
                subtotal_row_nums.append(i + 1)  # +1 for the header row
            cells = []
            for c in cols:
                text = _fmt_cell(row[c], c)
                is_wide = c in WIDE_COLS or c == "District"
                if _below_threshold(row[c], c, red_below):
                    style = cell_style_red if is_wide else cell_style_red_r
                elif is_subtotal:
                    style = total_row_style if is_wide else total_row_style_r
                else:
                    style = cell_style if is_wide else cell_style_r
                cells.append(Paragraph(text, style))
            data_rows.append(cells)
        data = [header] + data_rows

        style_cmds = [
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor(BRAND_PRIMARY_DARK)),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F4F6FA")]),
            ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#D5DCE8")),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("TOPPADDING", (0, 0), (-1, -1), 2.8),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 2.8),
            ("LEFTPADDING", (0, 0), (-1, -1), 5),
            ("RIGHTPADDING", (0, 0), (-1, -1), 5),
        ]
        for r in subtotal_row_nums:
            style_cmds.append(("BACKGROUND", (0, r), (-1, r), colors.HexColor("#EEF3FC")))

        if total_row:
            trow = []
            for c in cols:
                val = total_row.get(c)
                text = _fmt_cell(val, c) if val is not None else ""
                is_wide = c in WIDE_COLS or c == "District"
                if val is not None and _below_threshold(val, c, red_below):
                    style = total_row_style_red_r if not is_wide else total_row_style
                else:
                    style = total_row_style if is_wide else total_row_style_r
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
        story.append(Paragraph(f"<b>{district_title}</b>", h2_style))
        story.append(df_to_table(district_df, district_cols, total_row=district_total_row,
                                  red_below=red_below))
        section_added = True

    if ac_df is not None and not ac_df.empty and ac_cols:
        _section_gap()
        story.append(Paragraph(f"<b>{ac_title}</b>", h2_style))
        story.append(df_to_table(ac_df, ac_cols, total_row=ac_total_row, red_below=red_below))
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
ac_districtno_map = build_ac_districtno_map(nh_df) if nh_df is not None else {}
fp_df, fp_meta, fp_err = load_form_processing(FORM_PROCESSING_FILE, ac_map, ac_districtno_map)
fp_old_df, fp_old_meta, fp_old_err = load_form_processing(FORM_PROCESSING_OLD_FILE, ac_map, ac_districtno_map)
pw_df, pw_meta, pw_err = load_partwise_data(PARTWISE_FILE)

with st.sidebar:
    st.markdown("### \U0001F4CB Dashboard Controls")
    st.caption("Filters apply live to KPIs, charts, reports and exports.")

VIEW_LABELS = {"fp": "\U0001F4C4  Form Processing", "nh": "\U0001F4E8  Notice & Hearing",
               "pw": "\U0001F4CD  Part-wise Report", "diff": "\U0001F4CA  Difference Report"}
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

            def _fp_row_values(df_row_scope):
                return [
                    df_row_scope["Total_Received"].sum(),
                    df_row_scope["Hearing_Scheduled"].sum(),
                    df_row_scope["Rejected"].sum(),
                    df_row_scope["Accepted"].sum(),
                    df_row_scope["Eroll_Inclusion"].sum(),
                ]

            kpi_row("Total (All Forms)", _fp_row_values(filtered), color=BRAND_PRIMARY)
            kpi_row("Form 6", _fp_row_values(filtered[filtered["Form_Type"] == "FORM6"]), color=BRAND_ACCENT)
            kpi_row("Form 7", _fp_row_values(filtered[filtered["Form_Type"] == "FORM7"]), color=BRAND_WARN)
            kpi_row("Form 8", _fp_row_values(filtered[filtered["Form_Type"] == "FORM8"]), color=BRAND_DANGER)

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
            # Form Type-wise breakdown by default (matches the official "Form
            # Processed Current Status" report layout) -- fp_dist_rep (flat,
            # one row per district, all form types summed) is kept only for
            # the grand-Total row and the chart below, not for display.
            section_title("District-wise Report")
            fp_dist_rep = fp_district_report(filtered)
            fp_dist_formtype_rep = fp_district_formtype_report(filtered)
            fp_dist_extra_pick = st.multiselect(
                "Add more columns (optional) -- every column below is exactly as named in Form_Processing.xlsx",
                FP_FORMTYPE_DIST_EXTRA_COLS, default=[], key="fp_dist_extra_cols",
                format_func=lambda c: c.replace("_", " "),
            )
            fp_dist_total = fp_total_row(fp_dist_rep)
            fp_dist_total_dict = fp_dist_total.iloc[0].to_dict() if not fp_dist_total.empty else None
            if fp_dist_total_dict:
                fp_dist_total_dict["Form_Type"] = "All District Form Total"
            render_html_table(
                fp_dist_formtype_rep, FP_FORMTYPE_DIST_COLS + fp_dist_extra_pick,
                formats=FP_COL_FORMATS, labels=FP_FORMTYPE_COL_LABELS,
                total_row=fp_dist_total_dict,
                caption="Matches the official 'Form Processed Current Status' report: each district's Form "
                        "6/6A/7/8 rows, its District Form Total, then the overall Total at the very end. Pick a "
                        "single Form Type in the sidebar filter to collapse this back to one row per district.",
            )

            if SHOW_CHARTS:
                fp_dist_chart_view = fp_dist_rep.sort_values("Total_Received", ascending=False)
                fig = px.bar(fp_dist_chart_view, x="District", y=["Total_Received", "Eroll_Inclusion"],
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
            fp_ac_formtype_rep = fp_ac_formtype_report(ac_scope_df)
            fp_ac_extra_pick = st.multiselect(
                "Add more columns (optional) -- every column below is exactly as named in Form_Processing.xlsx",
                FP_FORMTYPE_AC_EXTRA_COLS, default=[], key="fp_ac_extra_cols",
                format_func=lambda c: c.replace("_", " "),
            )
            fp_ac_total = fp_total_row(fp_ac_rep)
            fp_ac_total_dict = fp_ac_total.iloc[0].to_dict() if not fp_ac_total.empty else None
            if fp_ac_total_dict:
                fp_ac_total_dict["Form_Type"] = "All AC Form Total"
            render_html_table(
                fp_ac_formtype_rep, FP_FORMTYPE_AC_COLS + fp_ac_extra_pick,
                formats=FP_COL_FORMATS, labels=FP_FORMTYPE_COL_LABELS,
                total_row=fp_ac_total_dict,
                caption="Matches the official 'Form Processed Current Status' report: each AC's Form 6/6A/7/8 "
                        "rows, its AC Form Total, then the overall Total at the very end. Pick a single Form "
                        "Type in the sidebar filter to collapse this back to one row per AC.",
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
                    "District Report": fp_dist_formtype_rep,
                    "AC Report": fp_ac_formtype_rep,
                })
                st.download_button("\U0001F4E5 Download Filtered Report (Excel)", excel_bytes,
                                    file_name="Form_Processing_Filtered_Report.xlsx",
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                    use_container_width=True)
            with dcol2:
                def _fp_kpi_group(row_scope):
                    vals = _fp_row_values(row_scope)
                    return {
                        "Receiving": fmt_int(vals[0]),
                        "Hearing": fmt_int(vals[1]),
                        "Reject": fmt_int(vals[2]),
                        "Accept": fmt_int(vals[3]),
                        "Inclusion": fmt_int(vals[4]),
                    }

                fp_kpi_groups = [("Total (All Forms)", _fp_kpi_group(filtered))]
                for _grp_label, _grp_code in [("Form 6", "FORM6"), ("Form 7", "FORM7"), ("Form 8", "FORM8")]:
                    fp_kpi_groups.append(
                        (_grp_label, _fp_kpi_group(filtered[filtered["Form_Type"] == _grp_code])))

                try:
                    pdf_bytes = build_pdf_report(
                        title="Form Processing Report",
                        subtitle=f"Report period: {fp_meta.get('report_period') or 'N/A'}",
                        filters_desc=fp_filters_desc, kpis=None, kpi_groups=fp_kpi_groups,
                        district_df=fp_dist_formtype_rep, ac_df=fp_ac_formtype_rep,
                        charts=[],
                        district_cols=FP_FORMTYPE_DIST_COLS, ac_cols=FP_FORMTYPE_AC_COLS,
                        col_labels=FP_FORMTYPE_COL_LABELS,
                        district_total_row=fp_dist_total_dict, ac_total_row=fp_ac_total_dict,
                    )
                    st.download_button("\U0001F4C4 Download PDF Report (full)", pdf_bytes,
                                        file_name="Form_Processing_Report.pdf",
                                        mime="application/pdf", use_container_width=True)
                except Exception as exc:  # noqa: BLE001
                    st.error(f"PDF generation failed: {exc}")

            st.caption("District-wise and AC-wise reports on their own, in Excel or PDF:")
            dcol3, dcol4, dcol5, dcol6 = st.columns(4)
            with dcol3:
                dist_excel = build_excel_download({"District Report": fp_dist_formtype_rep})
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
                        district_df=fp_dist_formtype_rep, ac_df=None, charts=[],
                        district_cols=FP_FORMTYPE_DIST_COLS, col_labels=FP_FORMTYPE_COL_LABELS,
                        district_total_row=fp_dist_total_dict,
                    )
                    st.download_button("\U0001F4C4 District-wise (PDF)", dist_pdf_bytes,
                                        file_name="Form_Processing_District_Report.pdf",
                                        mime="application/pdf", use_container_width=True, key="fp_dl_dist_pdf")
                except Exception as exc:  # noqa: BLE001
                    st.error(f"PDF generation failed: {exc}")
            with dcol5:
                ac_excel = build_excel_download({"AC Report": fp_ac_formtype_rep})
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
                        district_df=None, ac_df=fp_ac_formtype_rep, charts=[],
                        ac_cols=FP_FORMTYPE_AC_COLS, col_labels=FP_FORMTYPE_COL_LABELS,
                        ac_total_row=fp_ac_total_dict,
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
        nh_last_updated = nh_meta.get("last_updated") if nh_meta else None
        last_updated_line = (f"<br><b>Last Updated On:</b> {nh_last_updated}"
                              if nh_last_updated else "")
        st.markdown(f"""<div class="note-box">
            <b>Granularity:</b> one row per Assembly Constituency (AC)
            ({fmt_int(len(nh_df))} ACs across {nh_df['District'].nunique()} districts).{last_updated_line}
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
                nh_subtitle_suffix = f"  |  Last Updated On: {nh_last_updated}" if nh_last_updated else ""
                try:
                    pdf_bytes2 = build_pdf_report(
                        title="Notice & Hearing Report",
                        subtitle=f"{fmt_int(len(nh_filtered))} Assembly Constituencies in scope{nh_subtitle_suffix}",
                        filters_desc=nh_filters_desc, kpis=kpis_for_pdf2,
                        district_df=nh_dist_view, ac_df=nh_ac_rep,
                        charts=[],
                        district_cols=NOTICE_DIST_COLS, ac_cols=NOTICE_AC_COLS,
                        col_labels=NOTICE_COL_LABELS,
                        district_total_row=nh_dist_total.iloc[0].to_dict() if not nh_dist_total.empty else None,
                        red_below=NOTICE_RED_BELOW,
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
                        subtitle=f"{fmt_int(len(nh_filtered))} Assembly Constituencies in scope{nh_subtitle_suffix}",
                        filters_desc=nh_filters_desc, kpis=None,
                        district_df=nh_dist_view, ac_df=None, charts=[],
                        district_cols=NOTICE_DIST_COLS, col_labels=NOTICE_COL_LABELS,
                        district_total_row=nh_dist_total.iloc[0].to_dict() if not nh_dist_total.empty else None,
                        red_below=NOTICE_RED_BELOW,
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
                        subtitle=f"{fmt_int(len(nh_filtered))} Assembly Constituencies in scope{nh_subtitle_suffix}",
                        filters_desc=nh_filters_desc, kpis=None,
                        district_df=None, ac_df=nh_ac_rep, charts=[],
                        ac_cols=NOTICE_AC_COLS, col_labels=NOTICE_COL_LABELS,
                        red_below=NOTICE_RED_BELOW,
                    )
                    st.download_button("\U0001F4C4 AC-wise (PDF)", nh_ac_pdf_bytes,
                                        file_name="Notice_Hearing_AC_Report.pdf",
                                        mime="application/pdf", use_container_width=True, key="nh_dl_ac_pdf")
                except Exception as exc:  # noqa: BLE001
                    st.error(f"PDF generation failed: {exc}")

# ==========================================================================
# VIEW: PART-WISE REPORT
# ==========================================================================
elif active_view == "pw":
    if pw_err:
        st.error(f"**Part-wise data could not be loaded.**\n\n{pw_err}")
    elif pw_df is None or pw_df.empty:
        st.warning("Partwise.xlsx loaded but contains no usable data rows.")
    else:
        # Partwise.xlsx itself carries no title/timestamp row -- reuse the
        # "Last Updated On" timestamp already parsed from Notice.xlsx (same
        # underlying SIR data refresh) so this tab shows it too, on-screen
        # and in both downloads, per the user's request.
        pw_last_updated = nh_meta.get("last_updated") if nh_meta else None
        pw_last_updated_line = (f"<br><b>Last Updated On (from Notice.xlsx):</b> {pw_last_updated}"
                                 if pw_last_updated else "")
        st.markdown(f"""<div class="note-box">
            <b>Granularity:</b> one row per polling Part
            ({fmt_int(len(pw_df))} Parts across {pw_df['AC_No'].nunique()} ACs,
            {pw_df['District'].nunique()} districts).{pw_last_updated_line}
            <br><span style="color:{BRAND_MUTED}">Built from Partwise.xlsx. Same column set/wording as the
            Notice &amp; Hearing report, except <b>Electors</b> and <b>% Parked for Final</b> -- Partwise.xlsx
            has no elector count per Part (only Electors.xlsx does, at AC level), and % Parked for Final is
            defined as a share of Electors, so it isn't meaningful at this granularity.</span>
            </div>""", unsafe_allow_html=True)

        with st.sidebar:
            st.markdown("#### \U0001F4CD Part-wise Filters")
            pw_districts = sorted(pw_df["District"].unique())
            pw_sel_districts = st.multiselect("District", pw_districts, default=[], key="pw_dist")
            pw_scope = pw_df[pw_df["District"].isin(pw_sel_districts)] if pw_sel_districts else pw_df
            pw_acs = sorted(pw_scope["AC_Name"].unique())
            pw_sel_acs = st.multiselect("Assembly Constituency (AC)", pw_acs, default=[], key="pw_ac")
            pw_rank_label = st.selectbox("Rank parts by", list(PARTWISE_RANK_OPTIONS.keys()),
                                          index=0, key="pw_rank")
            pw_n = st.selectbox("Parts per AC to show", [5, 10, 15, 20], index=1, key="pw_n")

        pw_filtered = pw_df.copy()
        if pw_sel_districts:
            pw_filtered = pw_filtered[pw_filtered["District"].isin(pw_sel_districts)]
        if pw_sel_acs:
            pw_filtered = pw_filtered[pw_filtered["AC_Name"].isin(pw_sel_acs)]

        pw_rank_col, pw_ascending = PARTWISE_RANK_OPTIONS[pw_rank_label]
        # "Bottom" when the lowest values are worst (e.g. % Hearing Held),
        # "Top" when the highest values are worst (e.g. Found Ineligible for
        # Final -- ERO) -- so the section wording always reads naturally
        # regardless of which metric is selected.
        pw_direction_word = "Bottom" if pw_ascending else "Top"

        pw_filt_parts = []
        if pw_sel_districts: pw_filt_parts.append("District: " + ", ".join(pw_sel_districts))
        if pw_sel_acs: pw_filt_parts.append("AC: " + ", ".join(pw_sel_acs))
        pw_filt_parts.append(f"{pw_direction_word} {pw_n} per AC by {pw_rank_label}")
        pw_filters_desc = " | ".join(pw_filt_parts)

        if pw_filtered.empty:
            no_data_message()
        else:
            pw_bottom_rep = partwise_bottom_n_report(pw_filtered, pw_rank_col, pw_ascending, pw_n)

            section_title("Key Performance Indicators")
            c1, c2, c3, c4 = st.columns(4, gap="medium")
            kpi_card(c1, "Parts in Scope", fmt_int(len(pw_filtered)),
                     f"{pw_filtered['AC_No'].nunique()} ACs · {pw_filtered['District'].nunique()} districts")
            kpi_card(c2, "Avg % Hearing Held", fmt_pct(pw_filtered["Hearing_Held_%"].mean()),
                     "across all parts in scope", color=BRAND_ACCENT)
            kpi_card(c3, "Total Found Ineligible for Final (ERO)", fmt_int(pw_filtered["Ineligible_Final"].sum()),
                     "across all parts in scope", color=BRAND_DANGER)
            kpi_card(c4, "Parts Shown Below", fmt_int(len(pw_bottom_rep)),
                     f"{pw_direction_word.lower()} {pw_n} per AC, {pw_bottom_rep['AC_No'].nunique()} ACs", color=BRAND_WARN)

            section_title(f"{pw_direction_word} {pw_n} Parts per AC (by {pw_rank_label})")
            rank_note = "Rank 1 = highest value in that AC." if not pw_ascending else "Rank 1 = worst part in that AC."
            render_html_table(
                pw_bottom_rep, PARTWISE_COLS,
                formats=PARTWISE_COL_FORMATS, labels=PARTWISE_COL_LABELS,
                caption="Columns match the Notice & Hearing report format (Electors / % Parked for Final "
                        f"excluded -- not available at Part level). {rank_note}",
            )

            # ------------------ Downloads ------------------
            section_title("Downloads")
            dcol1, dcol2 = st.columns(2)
            with dcol1:
                pw_excel_sheets = {}
                if pw_last_updated:
                    pw_excel_sheets["Report Info"] = pd.DataFrame(
                        {"Note": [f"Last Updated On (from Notice.xlsx): {pw_last_updated}"]})
                pw_excel_sheets[f"{pw_direction_word} N Parts per AC"] = pw_bottom_rep
                pw_excel_sheets["All Parts (Filtered)"] = pw_filtered
                pw_excel_bytes = build_excel_download(pw_excel_sheets)
                st.download_button("\U0001F4E5 Download Report (Excel)", pw_excel_bytes,
                                    file_name="Partwise_Report.xlsx",
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                    use_container_width=True, key="pw_dl_xlsx")
            with dcol2:
                pw_subtitle_suffix = f"  |  Last Updated On: {pw_last_updated}" if pw_last_updated else ""
                try:
                    pw_pdf_bytes = build_pdf_report(
                        title="Part-wise Notice & Hearing Report",
                        subtitle=f"{pw_direction_word} {pw_n} parts per AC, ranked by "
                                 f"{pw_rank_label}{pw_subtitle_suffix}",
                        filters_desc=pw_filters_desc, kpis=None,
                        district_df=pw_bottom_rep, ac_df=None, charts=[],
                        district_cols=PARTWISE_COLS, col_labels=PARTWISE_COL_LABELS,
                        district_title=f"{pw_direction_word} {pw_n} Parts per AC",
                    )
                    st.download_button("\U0001F4C4 Download PDF Report", pw_pdf_bytes,
                                        file_name="Partwise_Report.pdf",
                                        mime="application/pdf", use_container_width=True, key="pw_dl_pdf")
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
        st.markdown(f"""<div class="note-box">
            <b>Comparing:</b> Form_Processing.xlsx (period: {fp_meta.get('report_period') or 'N/A'})
            <b>minus</b> comparison workbook 1513.xlsx (period: {fp_old_meta.get('report_period') or 'N/A'}).
            <br><span style="color:{BRAND_MUTED}">Same columns as the Form Processing District-wise/AC-wise
            reports -- every value here is the change between the two report periods (District or AC).</span>
            </div>""", unsafe_allow_html=True)

        with st.sidebar:
            st.markdown("#### \U0001F4CA Difference Report Filters")
            diff_districts = sorted(set(fp_df["District"].unique()) | set(fp_old_df["District"].unique()))
            diff_sel_districts = st.multiselect("District", diff_districts, default=[], key="diff_dist")
            diff_ac_scope_src = pd.concat([fp_df, fp_old_df])
            if diff_sel_districts:
                diff_ac_scope_src = diff_ac_scope_src[diff_ac_scope_src["District"].isin(diff_sel_districts)]
            diff_acs = sorted(diff_ac_scope_src["AC_Name"].unique())
            diff_sel_acs = st.multiselect("Assembly Constituency (AC)", diff_acs, default=[], key="diff_ac")
            diff_form_types = sorted(set(fp_df["Form_Type"].unique()) | set(fp_old_df["Form_Type"].unique()))
            diff_sel_forms = st.multiselect(
                "Form Type", diff_form_types,
                format_func=lambda x: FORM_TYPE_LABELS.get(x, x), default=[], key="diff_form")

        # Apply District/AC/Form Type filters to the row-level data BEFORE
        # aggregating -- the diff reports sum across every matching row, so
        # filtering has to happen here, not after fp_district_report()/
        # fp_ac_report() have already summed Form Type away.
        fp_scope = fp_df.copy()
        fp_old_scope = fp_old_df.copy()
        if diff_sel_districts:
            fp_scope = fp_scope[fp_scope["District"].isin(diff_sel_districts)]
            fp_old_scope = fp_old_scope[fp_old_scope["District"].isin(diff_sel_districts)]
        if diff_sel_acs:
            fp_scope = fp_scope[fp_scope["AC_Name"].isin(diff_sel_acs)]
            fp_old_scope = fp_old_scope[fp_old_scope["AC_Name"].isin(diff_sel_acs)]
        if diff_sel_forms:
            fp_scope = fp_scope[fp_scope["Form_Type"].isin(diff_sel_forms)]
            fp_old_scope = fp_old_scope[fp_old_scope["Form_Type"].isin(diff_sel_forms)]

        diff_dist_view = fp_diff_report(fp_district_report(fp_scope), fp_district_report(fp_old_scope),
                                         ["District"])
        diff_ac_view = fp_diff_report(fp_ac_report(fp_scope), fp_ac_report(fp_old_scope),
                                       ["District", "AC_No", "AC_Name"])

        diff_filt_parts = []
        if diff_sel_districts: diff_filt_parts.append("District: " + ", ".join(diff_sel_districts))
        if diff_sel_acs: diff_filt_parts.append("AC: " + ", ".join(diff_sel_acs))
        if diff_sel_forms: diff_filt_parts.append(", ".join(FORM_TYPE_SHORT.get(f, f) for f in diff_sel_forms))
        diff_filters_desc = " | ".join(diff_filt_parts)

        if diff_dist_view.empty:
            no_data_message()
        else:
            total_diff = diff_dist_view["Total_Received"].sum()
            hearing_diff = diff_dist_view["Hearing_Scheduled"].sum()
            rejected_diff = diff_dist_view["Rejected"].sum()
            accepted_diff = diff_dist_view["Accepted"].sum()
            eroll_diff = diff_dist_view["Eroll_Inclusion"].sum()
            unprocessed_diff = diff_dist_view["Unprocessed"].sum()
            inprogress_diff = diff_dist_view["In_Progress"].sum()

            section_title("Key Performance Indicators")
            c1, c2, c3, c4 = st.columns(4, gap="medium")
            kpi_card(c1, "Total Forms Received (Δ)", fmt_diff(total_diff),
                     f"{diff_dist_view['District'].nunique()} districts in scope")
            kpi_card(c2, "Hearing Scheduled (Δ)", fmt_diff(hearing_diff), color=BRAND_WARN)
            kpi_card(c3, "Rejected (Δ)", fmt_diff(rejected_diff), color=BRAND_DANGER)
            kpi_card(c4, "Accepted (Δ)", fmt_diff(accepted_diff))

            c5, c6, c7 = st.columns(3, gap="medium")
            kpi_card(c5, "Eroll Inclusion (Δ)", fmt_diff(eroll_diff), color=BRAND_ACCENT)
            kpi_card(c6, "Unprocessed (Δ)", fmt_diff(unprocessed_diff), color=BRAND_WARN)
            kpi_card(c7, "In Progress (Δ)", fmt_diff(inprogress_diff))

            st.caption(
                "ℹ️ Every KPI and column here is (Form_Processing.xlsx value) minus "
                "(1513.xlsx value) for the same District/AC -- a positive number means that column grew "
                "between the two report periods, negative means it shrank."
            )

            # ------------------ District-wise diff report ------------------
            section_title("District-wise Difference Report")
            diff_dist_sort_options = DIFF_DIST_BASE_COLS[1:] + DIFF_DIST_EXTRA_COLS
            diff_dist_sort_opt = st.selectbox(
                "Sort district report by", diff_dist_sort_options, index=0, key="diff_dist_sort",
                format_func=lambda c: DIFF_COL_LABELS.get(c, c.replace("_", " ")),
            )
            diff_dist_sorted = diff_dist_view.sort_values(diff_dist_sort_opt, ascending=False)
            diff_dist_extra_pick = st.multiselect(
                "Add more columns (optional) -- every column below is exactly as named in Form_Processing.xlsx",
                DIFF_DIST_EXTRA_COLS, default=[], key="diff_dist_extra_cols",
                format_func=lambda c: DIFF_COL_LABELS.get(c, c.replace("_", " ")),
            )
            render_html_table(
                diff_dist_sorted, DIFF_DIST_BASE_COLS + diff_dist_extra_pick,
                formats=DIFF_COL_FORMATS, labels=DIFF_COL_LABELS,
                caption="Values are District-wise differences (Form_Processing.xlsx minus 1513.xlsx), "
                        "not absolute counts. Full column set is in the Excel/PDF export below.",
            )

            # ------------------ AC-wise diff report ------------------
            section_title("AC-wise Difference Report")
            diff_ac_dist_pick = st.selectbox(
                "Select District to drill into its ACs (optional)",
                ["All Districts"] + sorted(diff_ac_view["District"].unique()),
                key="diff_ac_district_pick",
            )
            diff_ac_scope = (diff_ac_view if diff_ac_dist_pick == "All Districts"
                              else diff_ac_view[diff_ac_view["District"] == diff_ac_dist_pick])
            diff_ac_extra_pick = st.multiselect(
                "Add more columns (optional) -- every column below is exactly as named in Form_Processing.xlsx",
                DIFF_AC_EXTRA_COLS, default=[], key="diff_ac_extra_cols",
                format_func=lambda c: DIFF_COL_LABELS.get(c, c.replace("_", " ")),
            )
            render_html_table(
                diff_ac_scope, DIFF_AC_BASE_COLS + diff_ac_extra_pick,
                formats=DIFF_COL_FORMATS, labels=DIFF_COL_LABELS,
                caption="Values are AC-wise differences (Form_Processing.xlsx minus 1513.xlsx), "
                        "not absolute counts. Full column set is in the Excel/PDF export below.",
            )

            # ------------------ Downloads ------------------
            section_title("Downloads")
            dcol1, dcol2 = st.columns(2)
            with dcol1:
                diff_excel = build_excel_download({
                    "District Report": diff_dist_view, "AC Report": diff_ac_view,
                })
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
                        district_df=diff_dist_sorted, ac_df=diff_ac_scope, charts=[],
                        district_cols=DIFF_DIST_BASE_COLS, ac_cols=DIFF_AC_BASE_COLS,
                        col_labels=DIFF_COL_LABELS,
                    )
                    st.download_button("\U0001F4C4 Download PDF Report (full)", diff_pdf_bytes,
                                        file_name="Form_Processing_Difference_Report.pdf",
                                        mime="application/pdf", use_container_width=True, key="diff_dl_pdf")
                except Exception as exc:  # noqa: BLE001
                    st.error(f"PDF generation failed: {exc}")

            st.caption("District-wise and AC-wise reports on their own, in Excel or PDF:")
            dcol3, dcol4, dcol5, dcol6 = st.columns(4)
            with dcol3:
                diff_dist_excel = build_excel_download({"District Report": diff_dist_view})
                st.download_button("\U0001F4E5 District-wise (Excel)", diff_dist_excel,
                                    file_name="Form_Processing_Difference_District_Report.xlsx",
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                    use_container_width=True, key="diff_dl_dist_xlsx")
            with dcol4:
                try:
                    diff_dist_pdf_bytes = build_pdf_report(
                        title="Form Processing - Difference Report (District-wise)",
                        subtitle=(f"Change: {fp_old_meta.get('report_period') or 'N/A'}  "
                                  f"vs  {fp_meta.get('report_period') or 'N/A'}"),
                        filters_desc=diff_filters_desc, kpis=None,
                        district_df=diff_dist_sorted, ac_df=None, charts=[],
                        district_cols=DIFF_DIST_BASE_COLS, col_labels=DIFF_COL_LABELS,
                    )
                    st.download_button("\U0001F4C4 District-wise (PDF)", diff_dist_pdf_bytes,
                                        file_name="Form_Processing_Difference_District_Report.pdf",
                                        mime="application/pdf", use_container_width=True, key="diff_dl_dist_pdf")
                except Exception as exc:  # noqa: BLE001
                    st.error(f"PDF generation failed: {exc}")
            with dcol5:
                diff_ac_excel = build_excel_download({"AC Report": diff_ac_view})
                st.download_button("\U0001F4E5 AC-wise (Excel)", diff_ac_excel,
                                    file_name="Form_Processing_Difference_AC_Report.xlsx",
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                    use_container_width=True, key="diff_dl_ac_xlsx")
            with dcol6:
                try:
                    diff_ac_pdf_bytes = build_pdf_report(
                        title="Form Processing - Difference Report (AC-wise)",
                        subtitle=(f"Change: {fp_old_meta.get('report_period') or 'N/A'}  "
                                  f"vs  {fp_meta.get('report_period') or 'N/A'}"),
                        filters_desc=diff_filters_desc, kpis=None,
                        district_df=None, ac_df=diff_ac_scope, charts=[],
                        ac_cols=DIFF_AC_BASE_COLS, col_labels=DIFF_COL_LABELS,
                    )
                    st.download_button("\U0001F4C4 AC-wise (PDF)", diff_ac_pdf_bytes,
                                        file_name="Form_Processing_Difference_AC_Report.pdf",
                                        mime="application/pdf", use_container_width=True, key="diff_dl_ac_pdf")
                except Exception as exc:  # noqa: BLE001
                    st.error(f"PDF generation failed: {exc}")

st.markdown(f"""
<div style="text-align:center; color:{BRAND_MUTED}; font-size:0.78rem; margin-top:2rem; padding-top:1rem; border-top:1px solid #E4E8F0;">
    Form Processing &amp; Notice-Hearing Report Dashboard &middot; Data refreshed from workbook files in this repository &middot;
    Generated {datetime.now().strftime('%d-%b-%Y %H:%M')} IST
</div>
""", unsafe_allow_html=True)
