# Form Processing & Notice-Hearing MIS Dashboard

A single-file Streamlit MIS (Management Information System) dashboard for monitoring
**Special Intensive Revision (SIR)** electoral-roll workflow data for Uttarakhand:
form processing status (Form 6 / 6A / 7 / 8) and notice-generation / hearing / DEO
pendency data, both broken down by District and Assembly Constituency (AC).

Every KPI card, filter, chart and report column in this dashboard was derived by
directly inspecting the two source workbooks (`Form_Processing.xlsx` and
`Notice_Hearing.xlsx`) cell-by-cell — headers, subtotal rows, duplicate/blank
checks, and the real relationship between columns — rather than from a generic
template. The reasoning is documented in detail at the top of `app.py`.

## What this dashboard actually shows

### Tab 1 — Form Processing
Source: `Form_Processing.xlsx` (one row per AC x Form Type; subtotal/grand-total
rows such as "AC Form Total" and "All AC Form Total" are detected by their
non-numeric AC number and excluded so nothing is double-counted).

- **KPIs**: Total Forms Received, Unprocessed (Backlog), In Progress (Workflow),
  Finalized (Disposed), Eroll Inclusion, Rejected, Accepted, Hearing Scheduled —
  each derived from the fact that the 18 status columns in the sheet are
  mutually exclusive and sum exactly to Total Form Received (verified during
  inspection), i.e. every form sits in exactly one current stage.
- **Filters**: District, Assembly Constituency, Form Type, plus a chart-focus
  metric selector (the sheet has no single "status" field — status is a set of
  parallel stage-count columns — so this selector drives which stage the
  comparison charts rank ACs by).
- **Charts**: current status mix (donut), forms received by form type, district
  processing status (stacked), top ACs by the selected focus metric, plus
  comparison charts inside the District-wise and AC-wise reports.
- **District-wise report** and **AC-wise report**: sortable tables with
  Total Received, Unprocessed, In Progress, Finalized, Eroll Inclusion,
  Rejected, Disposal Rate %, Inclusion Rate % (and Backlog Rate % at district
  level). Selecting a District in the AC-wise report scopes the table and
  chart to that district's ACs only.

### Tab 2 — Notice & Hearing
Source: `Notice_Hearing.xlsx` (one row per polling-station Part within an AC
within a District — genuine unit-level data, no subtotal rows, no duplicates).

- **KPIs**: Notices Generated, Notices Delivered (delivery rate), Notice
  Pending Delivery, Hearings Held (rate vs. delivered), DEO Pending (backlog),
  DEO Pending > 5 Days (aged/overdue backlog), Found Ineligible for Final.
- **Filters**: District, Assembly Constituency, and a "show only Parts with
  DEO-pending cases" toggle.
- **Charts**: notice delivery split (donut), DEO pendency by district (total
  vs. overdue), district comparison (generated/delivered/hearings held), and
  top ACs by notices pending delivery (chosen over a "delivery rate" ranking
  because delivery rates cluster at 97-100% across ACs and a rate chart barely
  differentiates them — the absolute pending count is what an officer can act
  on).
- **District-wise report** and **AC-wise report**: sortable tables with Parts,
  Notice Generated/Delivered, Delivery Rate %, Hearings Held, Hearing Rate %,
  DEO Total Pending, DEO Pending > 5 Days, Pending > 5 Days %, and Found
  Ineligible. District -> AC drill-down works the same way as Tab 1.

### What's deliberately *not* in this dashboard
Neither workbook has a genuine per-record date field (Form_Processing only has
a report-period string in its title block; Notice_Hearing has no date column
at all), so no date filter or date-wise trend chart was built — that would
have been fabricated, not derived from data. The report period from
Form_Processing is shown as read-only context instead.

## Project structure

```
Form-Processing/
├── app.py                # Everything: data loading, KPIs, charts, reports, PDF/Excel export, UI
├── Form_Processing.xlsx  # Source workbook (Tab 1)
├── Notice_Hearing.xlsx   # Source workbook (Tab 2)
├── requirements.txt
└── README.md
```

## Installation

```bash
pip install -r requirements.txt
```

Requires Python 3.9+.

## Running the dashboard

From the repository root (the app locates the two `.xlsx` files next to
`app.py`, so it works regardless of your current working directory):

```bash
streamlit run app.py
```

Then open the URL Streamlit prints (typically `http://localhost:8501`).

## Updating the data

Replace `Form_Processing.xlsx` and/or `Notice_Hearing.xlsx` with updated
exports that keep the same column headers. The app re-reads the files (cached
per session via `st.cache_data`) and every KPI, chart, report and export
recalculates automatically — nothing is hard-coded.

## PDF reports

Each tab has a **Download PDF Report** button that generates a report (via
ReportLab, built directly in `app.py`, no screenshots involved) reflecting
whatever filters are currently applied: title, report type, generated
timestamp, selected filters, a KPI summary table, the District-wise and
AC-wise report tables (auto-switches to landscape A4 when a table has 6+
columns, with repeating headers across pages), and two supporting charts
rendered as embedded images. Every page has a branded header/footer with page
numbers.

## Excel downloads

Each tab offers:
- **Filtered Report** — the currently filtered detail rows, District Report,
  and AC Report as separate sheets.
- **District + AC Report** — just the two summary reports, for a lighter file.

Exports use `xlsxwriter` for a formatted header row, sensible column widths,
number formatting, and frozen header row (falls back to `openpyxl` if
`xlsxwriter` isn't available).

## Error handling

- Missing workbook file, missing sheet/columns, or a filter combination with
  no matching rows all show a clear inline message instead of crashing the
  app.
- Subtotal/grand-total rows in `Form_Processing.xlsx` are detected and
  excluded automatically so KPIs are never double-counted.

## Notes on the District/AC relationship

`Form_Processing.xlsx` has AC numbers and names but no District column.
Rather than hard-coding a mapping, the app derives AC -> District at runtime
from `Notice_Hearing.xlsx` (District Number/Name, AC Number/Name), which was
verified to cover all 70 Uttarakhand ACs with names matching
`Form_Processing.xlsx` exactly. AC No. 0 ("AC Missing") has no district and is
labelled "Unassigned / AC Not Mapped" rather than being dropped.
