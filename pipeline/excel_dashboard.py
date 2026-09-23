"""Builds Job_Portal_Performance_Tracker.xlsx from the full accumulated Weekly Input / Recruiter
Activity history (CSV rows as produced by build.py). Structure mirrors the original hand-built
tracker: Instructions, Platform Settings, Weekly Input, Recruiter Activity, Recruiter Summary,
Dashboard — but every tab is now generated from real accumulated history instead of a single
starter snapshot, and the Dashboard includes a genuine period-over-period trend chart once more
than one period exists.
"""
import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.chart import BarChart, LineChart, Reference

FONT_NAME = "Arial"
HEADER_FONT = Font(name=FONT_NAME, size=10, bold=True, color="FFFFFF")
LABEL_FONT = Font(name=FONT_NAME, size=10, bold=True)
NORMAL_FONT = Font(name=FONT_NAME, size=10)
INPUT_FONT = Font(name=FONT_NAME, size=10, color="0000FF")
KPI_VALUE_FONT = Font(name=FONT_NAME, size=18, bold=True, color="1F4E78")
KPI_LABEL_FONT = Font(name=FONT_NAME, size=9, bold=True, color="595959")

HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
INPUT_FILL = PatternFill("solid", fgColor="FFF2CC")
FORMULA_FILL = PatternFill("solid", fgColor="F2F2F2")
TITLE_FILL = PatternFill("solid", fgColor="1F4E78")
KPI_FILL = PatternFill("solid", fgColor="EAF1FB")
SUBHEADER_FILL = PatternFill("solid", fgColor="D9E1F2")

thin = Side(style="thin", color="BFBFBF")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)

PLATFORM_SETTINGS = [
    ("Vivian", "PREMIUM ELITE", None, "", "175/week (proposals)",
     "Cap is 6 hires/month (free max) + 25 proposals/hour, 175/week — not a raw view limit."),
    ("Monster", "ENTERPRISE", 270000, "Year", 35, "Views cap = 270,000/year."),
    ("Indeed", "BUDGET", None, "", "Unlimited", "Budget-based spend, not a fixed view cap."),
    ("Job Diva", "ATS+DATABASE", None, "", "Internal DB (no limit)", "Own database; not a paid job board subscription."),
    ("LinkedIn", "RECRUITER PRO", None, "", "111 total (85 US)",
     "Posting cap is a contract-term allotment, not weekly — track cumulative postings used against it."),
    ("CareerBuilder", "STANDARD", None, "", 8, "Merged with Monster account; no separate view cap."),
    ("Resume-Library", "STANDARD", 30000, "Year", "Unlimited", "Cap unit = unlocks/year."),
    ("Dice", "STANDARD", 50000, "Year", 35, "Views cap = 50,000/year."),
    ("Facebook", "Paid Groups", None, "", "Unlimited", "10-recruiter seat pack; no view cap."),
    ("DocCafe", "STANDARD", 375, "Month", "Unlimited", "375 views/month; +5,000 mails/month tracked separately."),
    ("Zoominfo", "", None, "", "", "No usage data received."),
    ("SignalHire", "", None, "", "", "Recruiter click activity tracked in Recruiter Activity tab."),
    ("Referred", "BONUS BASED", None, "", "Commission basis", "No subscription cost; pay-per-hire referral bonus."),
]
PLATFORM_LIST = [p[0] for p in PLATFORM_SETTINGS]


def _f(v):
    """CSV string -> python number or None."""
    if v is None or v == "" or v == "None":
        return None
    try:
        return float(v)
    except (TypeError, ValueError):
        return v  # leave text as-is


def _date(v):
    """CSV date string ('YYYY-MM-DD') -> datetime.date, so Excel date formulas (MAX, SUMIFS) work."""
    if v is None or v == "":
        return None
    if isinstance(v, datetime.date):
        return v
    for fmt in ("%Y-%m-%d", "%m/%d/%Y", "%Y/%m/%d"):
        try:
            return datetime.datetime.strptime(str(v).strip(), fmt).date()
        except ValueError:
            continue
    return v  # give up, leave as text (formulas involving it will just skip it)


def _style_header_row(ws, row, ncols, start_col=1):
    for c in range(start_col, start_col + ncols):
        cell = ws.cell(row=row, column=c)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = CENTER
        cell.border = BORDER


def _autosize(ws, widths):
    for col, w in widths.items():
        ws.column_dimensions[col].width = w


def build_excel(weekly_history, recruiter_history, output_path):
    wb = Workbook()

    # ================= Instructions =================
    ws = wb.active
    ws.title = "Instructions"
    ws.sheet_view.showGridLines = False
    ws["A1"] = "HonorVet Job Portal Performance Tracker — Automated Pipeline"
    ws["A1"].font = Font(name=FONT_NAME, size=16, bold=True, color="1F4E78")
    ws.merge_cells("A1:F1")
    ws.row_dimensions[1].height = 28
    notes = [
        ("This file is auto-generated", "Every tab below is regenerated from data/history/*.csv each time the "
         "pipeline runs. Do not hand-edit this workbook — edits will be overwritten on the next run. To correct "
         "or add data, edit the raw exports / manual CSV inputs and re-run the pipeline (or push to GitHub, if "
         "Actions is set up)."),
        ("Weekly Input & Recruiter Activity", "Reflect the full accumulated history across every period the "
         "pipeline has processed, not just the latest one."),
        ("Dashboard", "Always shows the most recent period, plus a trend chart across all periods once more "
         "than one exists."),
        ("Cumulative vs. period figures", "Recruiter Activity values are this period's actual activity — for "
         "platforms whose raw export reports running totals (LinkedIn, Indeed, etc.), the pipeline has already "
         "subtracted the prior snapshot. Nothing here is a raw cumulative total."),
    ]
    r = 3
    for a, b in notes:
        ws.cell(row=r, column=1, value=a).font = LABEL_FONT
        c = ws.cell(row=r, column=2, value=b)
        c.font = NORMAL_FONT
        c.alignment = Alignment(wrap_text=True, vertical="top")
        ws.cell(row=r, column=1).alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[r].height = 40
        r += 2
    _autosize(ws, {"A": 26, "B": 100})

    # ================= Platform Settings =================
    ws = wb.create_sheet("Platform Settings")
    ws.sheet_view.showGridLines = False
    ws["A1"] = "Platform Settings — reference data"
    ws["A1"].font = Font(name=FONT_NAME, size=13, bold=True, color="1F4E78")
    ws.merge_cells("A1:G1")
    headers = ["Platform", "Subscription Type", "Views/Unit Cap (value)", "Cap Period (Year/Month/Week)",
               "Daily Views/Unit Cap (auto)", "Postings Cap (concurrent, not prorated)", "Notes"]
    hr = 3
    for i, h in enumerate(headers, start=1):
        ws.cell(row=hr, column=i, value=h)
    _style_header_row(ws, hr, len(headers))
    ws.row_dimensions[hr].height = 32
    row = hr + 1
    for platform, sub_type, cap_val, cap_period, post_cap, notes_txt in PLATFORM_SETTINGS:
        ws.cell(row=row, column=1, value=platform).font = LABEL_FONT
        c = ws.cell(row=row, column=2, value=sub_type); c.font = INPUT_FONT; c.fill = INPUT_FILL
        c = ws.cell(row=row, column=3, value=cap_val); c.font = INPUT_FONT; c.fill = INPUT_FILL
        c = ws.cell(row=row, column=4, value=cap_period); c.font = INPUT_FONT; c.fill = INPUT_FILL
        f_cell = ws.cell(row=row, column=5)
        f_cell.value = (f'=IF(C{row}="","N/A",IF(D{row}="Year",C{row}/365,'
                        f'IF(D{row}="Month",C{row}*12/365,IF(D{row}="Week",C{row}/7,"N/A"))))')
        f_cell.font = NORMAL_FONT; f_cell.fill = FORMULA_FILL
        c = ws.cell(row=row, column=6, value=post_cap); c.font = INPUT_FONT; c.fill = INPUT_FILL
        c = ws.cell(row=row, column=7, value=notes_txt); c.font = Font(name=FONT_NAME, size=9, italic=True)
        c.alignment = Alignment(wrap_text=True, vertical="top")
        for col in range(1, 8):
            ws.cell(row=row, column=col).border = BORDER
            if col != 7:
                ws.cell(row=row, column=col).alignment = CENTER
        ws.row_dimensions[row].height = 30
        row += 1
    SETTINGS_LAST_ROW = row - 1
    _autosize(ws, {"A": 16, "B": 16, "C": 16, "D": 16, "E": 18, "F": 22, "G": 46})
    ws.freeze_panes = "A4"

    # ================= Weekly Input =================
    ws = wb.create_sheet("Weekly Input")
    ws.sheet_view.showGridLines = False
    ws["A1"] = "Weekly Input — auto-populated from the pipeline; do not hand-edit"
    ws["A1"].font = Font(name=FONT_NAME, size=13, bold=True, color="1F4E78")
    ws.merge_cells("A1:X1")
    wi_headers = [
        "Period Ending", "Platform", "Cost ($)", "Recruiters Assigned", "Active Recruiters",
        "Active Recruiter %", "Views Used (this period)", "Period Views Cap", "Views Utilization %",
        "Active Postings (current)", "Postings Cap", "Postings Utilization %",
        "Applications", "Contacts Made", "Responses Received",
        "Submissions", "Interviews Scheduled", "Offers Made",
        "Cost / Submission", "Cost / Interview", "Cost / Offer",
        "Submission->Interview Rate", "Interview->Offer Rate", "Period Length (Days)", "Notes"
    ]
    WI_HR = 3
    for i, h in enumerate(wi_headers, start=1):
        ws.cell(row=WI_HR, column=i, value=h)
    _style_header_row(ws, WI_HR, len(wi_headers))
    ws.row_dimensions[WI_HR].height = 44

    row = WI_HR + 1
    WI_FIRST_DATA_ROW = row
    for r_ in weekly_history:
        ws.cell(row=row, column=1, value=_date(r_["period_ending"]))
        ws.cell(row=row, column=2, value=r_["platform"])
        ws.cell(row=row, column=3, value=_f(r_["cost"]))
        ws.cell(row=row, column=4, value=_f(r_["recruiters_assigned"]))
        ws.cell(row=row, column=5, value=_f(r_["active_recruiters"]))
        ws.cell(row=row, column=6, value=f'=IFERROR(E{row}/D{row},"")')
        ws.cell(row=row, column=7, value=_f(r_["views_used"]))
        ws.cell(row=row, column=8, value=(f'=IFERROR(INDEX(\'Platform Settings\'!$E:$E,MATCH(B{row},'
                                           f'\'Platform Settings\'!$A:$A,0))*X{row},"N/A")'))
        ws.cell(row=row, column=9, value=f'=IF(ISNUMBER(H{row}),IFERROR(G{row}/H{row},""),"N/A")')
        ws.cell(row=row, column=10, value=_f(r_["active_postings"]))
        ws.cell(row=row, column=11, value=(f'=IFERROR(INDEX(\'Platform Settings\'!$F:$F,MATCH(B{row},'
                                            f'\'Platform Settings\'!$A:$A,0)),"N/A")'))
        ws.cell(row=row, column=12, value=f'=IF(ISNUMBER(K{row}),IFERROR(J{row}/K{row},""),"N/A")')
        ws.cell(row=row, column=13, value=_f(r_["applications"]))
        ws.cell(row=row, column=14, value=_f(r_["contacts_made"]))
        ws.cell(row=row, column=15, value=_f(r_["responses_received"]))
        ws.cell(row=row, column=16, value=_f(r_["submissions"]))
        ws.cell(row=row, column=17, value=_f(r_["interviews"]))
        ws.cell(row=row, column=18, value=_f(r_["offers"]))
        ws.cell(row=row, column=19, value=f'=IFERROR(C{row}/P{row},"")')
        ws.cell(row=row, column=20, value=f'=IFERROR(C{row}/Q{row},"")')
        ws.cell(row=row, column=21, value=f'=IFERROR(C{row}/R{row},"")')
        ws.cell(row=row, column=22, value=f'=IFERROR(Q{row}/P{row},"")')
        ws.cell(row=row, column=23, value=f'=IFERROR(R{row}/Q{row},"")')
        ws.cell(row=row, column=24, value=_f(r_["period_length_days"]) or 7)
        ws.cell(row=row, column=25, value=r_.get("notes", ""))
        for col in range(1, 26):
            cell = ws.cell(row=row, column=col)
            cell.border = BORDER
            cell.font = NORMAL_FONT
            cell.fill = FORMULA_FILL
            if col == 1:
                cell.number_format = "mm/dd/yyyy"
            if col in (3, 19, 20, 21):
                cell.number_format = '$#,##0.00;($#,##0.00);"-"'
            elif col in (6, 9, 12, 22, 23):
                cell.number_format = '0.0%;(0.0%);"-"'
            elif col == 25:
                cell.alignment = Alignment(wrap_text=True, vertical="top")
            else:
                cell.alignment = CENTER
        row += 1
    WI_LAST_DATA_ROW = max(row - 1, WI_FIRST_DATA_ROW)
    col_widths = {"A": 12, "B": 15, "C": 11, "D": 10, "E": 10, "F": 9, "G": 12, "H": 11, "I": 9, "J": 11, "K": 9,
                  "L": 9, "M": 11, "N": 10, "O": 10, "P": 10, "Q": 10, "R": 9, "S": 11, "T": 11, "U": 10, "V": 10,
                  "W": 10, "X": 11, "Y": 30}
    _autosize(ws, col_widths)
    ws.freeze_panes = "C4"

    # ================= Recruiter Activity =================
    ws = wb.create_sheet("Recruiter Activity")
    ws.sheet_view.showGridLines = False
    ws["A1"] = "Recruiter Activity — auto-populated from the pipeline; do not hand-edit"
    ws["A1"].font = Font(name=FONT_NAME, size=13, bold=True, color="1F4E78")
    ws.merge_cells("A1:J1")
    ra_headers = ["Period Ending", "Platform", "Recruiter", "Active This Period?",
                  "Searches", "Profiles Viewed", "Outreach Sent", "Responses Received", "Data Scope", "Notes"]
    RA_HR = 3
    for i, h in enumerate(ra_headers, start=1):
        ws.cell(row=RA_HR, column=i, value=h)
    _style_header_row(ws, RA_HR, len(ra_headers))
    ws.row_dimensions[RA_HR].height = 30
    row = RA_HR + 1
    RA_FIRST_DATA_ROW = row
    for r_ in recruiter_history:
        ws.cell(row=row, column=1, value=_date(r_["period_ending"]))
        ws.cell(row=row, column=1).number_format = "mm/dd/yyyy"
        ws.cell(row=row, column=2, value=r_["platform"])
        ws.cell(row=row, column=3, value=r_["recruiter"])
        ws.cell(row=row, column=4, value=r_["active"])
        ws.cell(row=row, column=5, value=_f(r_["searches"]))
        ws.cell(row=row, column=6, value=_f(r_["profiles_viewed"]))
        ws.cell(row=row, column=7, value=_f(r_["outreach_sent"]))
        ws.cell(row=row, column=8, value=_f(r_["responses_received"]))
        ws.cell(row=row, column=9, value=r_.get("data_scope", ""))
        ws.cell(row=row, column=10, value=r_.get("notes", ""))
        for col in range(1, 11):
            cell = ws.cell(row=row, column=col)
            cell.border = BORDER
            cell.font = NORMAL_FONT
            cell.fill = FORMULA_FILL
            if col in (9, 10):
                cell.alignment = Alignment(wrap_text=True, vertical="top")
            else:
                cell.alignment = CENTER
        row += 1
    RA_LAST_DATA_ROW = max(row - 1, RA_FIRST_DATA_ROW)
    _autosize(ws, {"A": 12, "B": 14, "C": 22, "D": 12, "E": 10, "F": 12, "G": 11, "H": 12, "I": 26, "J": 40})
    ws.freeze_panes = "D4"

    # ================= Recruiter Summary =================
    ws = wb.create_sheet("Recruiter Summary")
    ws.sheet_view.showGridLines = False
    ws["A1"] = "Recruiter Summary — auto-calculated across all periods (no manual entry)"
    ws["A1"].font = Font(name=FONT_NAME, size=13, bold=True, color="1F4E78")
    ws.merge_cells("A1:H1")
    rs_headers = ["Platform", "Recruiter", "Periods Logged", "Total Searches", "Total Profiles Viewed",
                  "Total Outreach Sent", "Total Responses", "Response Rate"]
    RS_HR = 3
    for i, h in enumerate(rs_headers, start=1):
        ws.cell(row=RS_HR, column=i, value=h)
    _style_header_row(ws, RS_HR, len(rs_headers))
    ws.row_dimensions[RS_HR].height = 30
    seen = set()
    unique_pairs = []
    for r_ in recruiter_history:
        key = (r_["platform"], r_["recruiter"])
        if key not in seen:
            seen.add(key)
            unique_pairs.append(key)
    row = RS_HR + 1
    RANGE_END = RA_FIRST_DATA_ROW + max(800, len(recruiter_history) + 200)
    P = f"'Recruiter Activity'!$B${RA_FIRST_DATA_ROW}:$B${RANGE_END}"
    RCOL = f"'Recruiter Activity'!$C${RA_FIRST_DATA_ROW}:$C${RANGE_END}"
    SCOL = f"'Recruiter Activity'!$E${RA_FIRST_DATA_ROW}:$E${RANGE_END}"
    VCOL = f"'Recruiter Activity'!$F${RA_FIRST_DATA_ROW}:$F${RANGE_END}"
    OCOL = f"'Recruiter Activity'!$G${RA_FIRST_DATA_ROW}:$G${RANGE_END}"
    RECOL = f"'Recruiter Activity'!$H${RA_FIRST_DATA_ROW}:$H${RANGE_END}"
    for platform, recruiter in unique_pairs:
        ws.cell(row=row, column=1, value=platform)
        ws.cell(row=row, column=2, value=recruiter)
        ws.cell(row=row, column=3, value=f'=COUNTIFS({P},A{row},{RCOL},B{row})')
        ws.cell(row=row, column=4, value=f'=SUMIFS({SCOL},{P},A{row},{RCOL},B{row})')
        ws.cell(row=row, column=5, value=f'=SUMIFS({VCOL},{P},A{row},{RCOL},B{row})')
        ws.cell(row=row, column=6, value=f'=SUMIFS({OCOL},{P},A{row},{RCOL},B{row})')
        ws.cell(row=row, column=7, value=f'=SUMIFS({RECOL},{P},A{row},{RCOL},B{row})')
        ws.cell(row=row, column=8, value=f'=IFERROR(G{row}/F{row},"")')
        for col in range(1, 9):
            cell = ws.cell(row=row, column=col)
            cell.border = BORDER
            cell.font = NORMAL_FONT if col > 2 else LABEL_FONT
            cell.fill = FORMULA_FILL
            cell.alignment = CENTER
            if col == 8:
                cell.number_format = '0.0%;(0.0%);"-"'
        row += 1
    _autosize(ws, {"A": 14, "B": 22, "C": 12, "D": 12, "E": 14, "F": 13, "G": 12, "H": 11})
    ws.freeze_panes = "C4"

    # ================= Dashboard =================
    ws = wb.create_sheet("Dashboard")
    ws.sheet_view.showGridLines = False
    ws.merge_cells("A1:L1")
    ws["A1"] = "HonorVet — Job Portal Performance Dashboard"
    ws["A1"].font = Font(name=FONT_NAME, size=18, bold=True, color="FFFFFF")
    ws["A1"].fill = TITLE_FILL
    ws["A1"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[1].height = 32
    ws.merge_cells("A2:L2")
    ws["A2"] = "Auto-updates from the pipeline. Always shows the most recent period logged."
    ws["A2"].font = Font(name=FONT_NAME, size=10, italic=True, color="595959")
    ws["A3"] = "Period shown:"
    ws["A3"].font = LABEL_FONT
    ws.merge_cells("B3:D3")
    ws["B3"] = "=MAX('Weekly Input'!$A:$A)"
    ws["B3"].number_format = "mmmm d, yyyy"
    ws["B3"].font = Font(name=FONT_NAME, size=12, bold=True, color="1F4E78")
    PERIOD_CELL = "$B$3"

    def kpi_card(label, formula, start_row, start_col, width, number_format):
        end_col = start_col + width - 1
        ws.merge_cells(start_row=start_row, start_column=start_col, end_row=start_row, end_column=end_col)
        ws.merge_cells(start_row=start_row + 1, start_column=start_col, end_row=start_row + 2, end_column=end_col)
        lc = ws.cell(row=start_row, column=start_col, value=label)
        lc.font = KPI_LABEL_FONT
        lc.alignment = Alignment(horizontal="center", vertical="center")
        vc = ws.cell(row=start_row + 1, column=start_col, value=formula)
        vc.font = KPI_VALUE_FONT
        vc.alignment = Alignment(horizontal="center", vertical="center")
        vc.number_format = number_format
        for rr in range(start_row, start_row + 3):
            for cc in range(start_col, end_col + 1):
                ws.cell(row=rr, column=cc).fill = KPI_FILL
                ws.cell(row=rr, column=cc).border = BORDER

    TABLE_HEADER_ROW = 13
    TABLE_FIRST_ROW = TABLE_HEADER_ROW + 1
    N_PLATFORMS = len(PLATFORM_LIST)
    TABLE_LAST_ROW = TABLE_FIRST_ROW + N_PLATFORMS - 1
    TOTAL_ROW = TABLE_LAST_ROW + 1

    kpi_card("TOTAL SPEND", f"=B{TOTAL_ROW}", 5, 1, 3, '$#,##0;($#,##0);"-"')
    kpi_card("TOTAL SUBMISSIONS", f"=H{TOTAL_ROW}", 5, 4, 3, '#,##0;(#,##0);"-"')
    kpi_card("TOTAL INTERVIEWS", f"=I{TOTAL_ROW}", 5, 7, 3, '#,##0;(#,##0);"-"')
    kpi_card("TOTAL OFFERS", f"=J{TOTAL_ROW}", 5, 10, 3, '#,##0;(#,##0);"-"')
    kpi_card("OVERALL OFFER RATE", f'=IFERROR(J{TOTAL_ROW}/H{TOTAL_ROW},"-")', 9, 1, 3, '0.0%;(0.0%);"-"')
    kpi_card("AVG COST PER OFFER", f"=K{TOTAL_ROW}", 9, 4, 3, '$#,##0;($#,##0);"-"')
    kpi_card("AVG RECRUITER UTILIZATION", f'=IFERROR(AVERAGE(E{TABLE_FIRST_ROW}:E{TABLE_LAST_ROW}),"-")', 9, 7, 3, '0.0%;(0.0%);"-"')
    kpi_card("AVG VIEWS UTILIZATION", f'=IFERROR(AVERAGE(F{TABLE_FIRST_ROW}:F{TABLE_LAST_ROW}),"-")', 9, 10, 3, '0.0%;(0.0%);"-"')

    ws.merge_cells("A12:L12")
    ws["A12"] = "Platform Performance — selected period"
    ws["A12"].font = Font(name=FONT_NAME, size=12, bold=True, color="1F4E78")
    pt_headers = ["Platform", "Cost ($)", "Active Recruiters", "Recruiters Assigned", "Active Recruiter %",
                  "Views Utilization %", "Postings Utilization %", "Submissions", "Interviews", "Offers",
                  "Cost / Offer", "Share of Total Offers"]
    for i, h in enumerate(pt_headers, start=1):
        ws.cell(row=TABLE_HEADER_ROW, column=i, value=h)
    _style_header_row(ws, TABLE_HEADER_ROW, len(pt_headers))
    ws.row_dimensions[TABLE_HEADER_ROW].height = 36

    r = TABLE_FIRST_ROW
    for i, platform in enumerate(PLATFORM_LIST):
        settings_row = 4 + i
        ws.cell(row=r, column=1, value=f"='Platform Settings'!A{settings_row}")
        ws.cell(row=r, column=2, value=(f"=SUMIFS('Weekly Input'!$C:$C,'Weekly Input'!$B:$B,A{r},"
                                         f"'Weekly Input'!$A:$A,{PERIOD_CELL})"))
        ws.cell(row=r, column=3, value=(f"=SUMIFS('Weekly Input'!$E:$E,'Weekly Input'!$B:$B,A{r},"
                                         f"'Weekly Input'!$A:$A,{PERIOD_CELL})"))
        ws.cell(row=r, column=4, value=(f"=SUMIFS('Weekly Input'!$D:$D,'Weekly Input'!$B:$B,A{r},"
                                         f"'Weekly Input'!$A:$A,{PERIOD_CELL})"))
        ws.cell(row=r, column=5, value=f'=IFERROR(C{r}/D{r},"")')
        ws.cell(row=r, column=6, value=(f"=IFERROR(SUMIFS('Weekly Input'!$G:$G,'Weekly Input'!$B:$B,A{r},"
                                         f"'Weekly Input'!$A:$A,{PERIOD_CELL})/SUMIFS('Weekly Input'!$H:$H,"
                                         f"'Weekly Input'!$B:$B,A{r},'Weekly Input'!$A:$A,{PERIOD_CELL}),\"N/A\")"))
        ws.cell(row=r, column=7, value=(f"=IFERROR(SUMIFS('Weekly Input'!$J:$J,'Weekly Input'!$B:$B,A{r},"
                                         f"'Weekly Input'!$A:$A,{PERIOD_CELL})/SUMIFS('Weekly Input'!$K:$K,"
                                         f"'Weekly Input'!$B:$B,A{r},'Weekly Input'!$A:$A,{PERIOD_CELL}),\"N/A\")"))
        ws.cell(row=r, column=8, value=(f"=SUMIFS('Weekly Input'!$P:$P,'Weekly Input'!$B:$B,A{r},"
                                         f"'Weekly Input'!$A:$A,{PERIOD_CELL})"))
        ws.cell(row=r, column=9, value=(f"=SUMIFS('Weekly Input'!$Q:$Q,'Weekly Input'!$B:$B,A{r},"
                                         f"'Weekly Input'!$A:$A,{PERIOD_CELL})"))
        ws.cell(row=r, column=10, value=(f"=SUMIFS('Weekly Input'!$R:$R,'Weekly Input'!$B:$B,A{r},"
                                          f"'Weekly Input'!$A:$A,{PERIOD_CELL})"))
        ws.cell(row=r, column=11, value=f'=IFERROR(B{r}/J{r},"")')
        ws.cell(row=r, column=12, value=f'=IFERROR(J{r}/J${TOTAL_ROW},"")')
        for col in range(1, 13):
            cell = ws.cell(row=r, column=col)
            cell.border = BORDER
            cell.font = NORMAL_FONT
            cell.alignment = CENTER
            cell.fill = FORMULA_FILL
            if col in (2, 11):
                cell.number_format = '$#,##0;($#,##0);"-"'
            elif col in (5, 6, 7, 12):
                cell.number_format = '0.0%;(0.0%);"-"'
        r += 1

    ws.cell(row=TOTAL_ROW, column=1, value="TOTAL")
    ws.cell(row=TOTAL_ROW, column=2, value=f"=SUM(B{TABLE_FIRST_ROW}:B{TABLE_LAST_ROW})")
    ws.cell(row=TOTAL_ROW, column=3, value=f"=SUM(C{TABLE_FIRST_ROW}:C{TABLE_LAST_ROW})")
    ws.cell(row=TOTAL_ROW, column=4, value=f"=SUM(D{TABLE_FIRST_ROW}:D{TABLE_LAST_ROW})")
    ws.cell(row=TOTAL_ROW, column=5, value=f'=IFERROR(C{TOTAL_ROW}/D{TOTAL_ROW},"")')
    ws.cell(row=TOTAL_ROW, column=8, value=f"=SUM(H{TABLE_FIRST_ROW}:H{TABLE_LAST_ROW})")
    ws.cell(row=TOTAL_ROW, column=9, value=f"=SUM(I{TABLE_FIRST_ROW}:I{TABLE_LAST_ROW})")
    ws.cell(row=TOTAL_ROW, column=10, value=f"=SUM(J{TABLE_FIRST_ROW}:J{TABLE_LAST_ROW})")
    ws.cell(row=TOTAL_ROW, column=11, value=f'=IFERROR(B{TOTAL_ROW}/J{TOTAL_ROW},"")')
    ws.cell(row=TOTAL_ROW, column=12, value=f'=IFERROR(J{TOTAL_ROW}/J{TOTAL_ROW},"")')
    for col in range(1, 13):
        cell = ws.cell(row=TOTAL_ROW, column=col)
        cell.border = BORDER
        cell.font = Font(name=FONT_NAME, size=10, bold=True)
        cell.fill = SUBHEADER_FILL
        cell.alignment = CENTER
        if col in (2, 11):
            cell.number_format = '$#,##0;($#,##0);"-"'
        elif col in (5, 6, 7, 12):
            cell.number_format = '0.0%;(0.0%);"-"'

    ws.conditional_formatting.add(f"E{TABLE_FIRST_ROW}:E{TABLE_LAST_ROW}",
        ColorScaleRule(start_type="min", start_color="F8696B", mid_type="percentile", mid_value=50,
                       mid_color="FFEB84", end_type="max", end_color="63BE7B"))
    ws.conditional_formatting.add(f"F{TABLE_FIRST_ROW}:F{TABLE_LAST_ROW}",
        ColorScaleRule(start_type="min", start_color="F8696B", mid_type="percentile", mid_value=50,
                       mid_color="FFEB84", end_type="max", end_color="63BE7B"))
    ws.conditional_formatting.add(f"J{TABLE_FIRST_ROW}:J{TABLE_LAST_ROW}",
        ColorScaleRule(start_type="min", start_color="F8696B", mid_type="percentile", mid_value=50,
                       mid_color="FFEB84", end_type="max", end_color="63BE7B"))
    ws.conditional_formatting.add(f"K{TABLE_FIRST_ROW}:K{TABLE_LAST_ROW}",
        ColorScaleRule(start_type="min", start_color="63BE7B", mid_type="percentile", mid_value=50,
                       mid_color="FFEB84", end_type="max", end_color="F8696B"))

    col_widths = {"A": 16, "B": 12, "C": 13, "D": 13, "E": 13, "F": 13, "G": 13, "H": 11, "I": 10, "J": 9,
                  "K": 11, "L": 13}
    _autosize(ws, col_widths)

    # ---- Trend-over-time table + chart (real once >1 period exists) ----
    unique_periods = sorted({r_["period_ending"] for r_ in weekly_history})
    trend_row0 = TOTAL_ROW + 2
    ws.cell(row=trend_row0, column=1, value="Trend by Period").font = Font(name=FONT_NAME, size=12, bold=True, color="1F4E78")
    trend_hdr = trend_row0 + 1
    for i, h in enumerate(["Period Ending", "Total Cost", "Total Submissions", "Total Interviews", "Total Offers"], start=1):
        ws.cell(row=trend_hdr, column=i, value=h)
    _style_header_row(ws, trend_hdr, 5)
    tr = trend_hdr + 1
    for period in unique_periods:
        ws.cell(row=tr, column=1, value=_date(period))
        ws.cell(row=tr, column=1).number_format = "mm/dd/yyyy"
        ws.cell(row=tr, column=2, value=f"=SUMIFS('Weekly Input'!$C:$C,'Weekly Input'!$A:$A,A{tr})")
        ws.cell(row=tr, column=3, value=f"=SUMIFS('Weekly Input'!$P:$P,'Weekly Input'!$A:$A,A{tr})")
        ws.cell(row=tr, column=4, value=f"=SUMIFS('Weekly Input'!$Q:$Q,'Weekly Input'!$A:$A,A{tr})")
        ws.cell(row=tr, column=5, value=f"=SUMIFS('Weekly Input'!$R:$R,'Weekly Input'!$A:$A,A{tr})")
        for col in range(1, 6):
            cell = ws.cell(row=tr, column=col)
            cell.border = BORDER; cell.font = NORMAL_FONT; cell.fill = FORMULA_FILL; cell.alignment = CENTER
            if col == 2:
                cell.number_format = '$#,##0;($#,##0);"-"'
        tr += 1
    trend_last = tr - 1

    note_row = tr + 1
    ws.merge_cells(f"A{note_row}:L{note_row}")
    ws.cell(row=note_row, column=1,
            value=("Note: with only one period logged so far, the trend chart below will show a single point — "
                   "it fills in automatically as the pipeline processes more weeks."))
    ws.cell(row=note_row, column=1).font = Font(name=FONT_NAME, size=9, italic=True, color="808080")
    ws.row_dimensions[note_row].height = 18

    chart_row = note_row + 2
    bar1 = BarChart(); bar1.title = "Cost by Platform"; bar1.y_axis.title = "Cost ($)"
    bar1.height, bar1.width = 8, 14
    data = Reference(ws, min_col=2, min_row=TABLE_HEADER_ROW, max_row=TABLE_LAST_ROW)
    cats = Reference(ws, min_col=1, min_row=TABLE_FIRST_ROW, max_row=TABLE_LAST_ROW)
    bar1.add_data(data, titles_from_data=True); bar1.set_categories(cats); bar1.legend = None
    ws.add_chart(bar1, f"A{chart_row}")

    bar2 = BarChart(); bar2.title = "Offers by Platform"; bar2.y_axis.title = "Offers"
    bar2.height, bar2.width = 8, 14
    data2 = Reference(ws, min_col=10, min_row=TABLE_HEADER_ROW, max_row=TABLE_LAST_ROW)
    bar2.add_data(data2, titles_from_data=True); bar2.set_categories(cats); bar2.legend = None
    ws.add_chart(bar2, f"E{chart_row}")

    bar3 = BarChart(); bar3.title = "Active Recruiter % by Platform"; bar3.y_axis.title = "Active %"
    bar3.height, bar3.width = 8, 14
    data3 = Reference(ws, min_col=5, min_row=TABLE_HEADER_ROW, max_row=TABLE_LAST_ROW)
    bar3.add_data(data3, titles_from_data=True); bar3.set_categories(cats); bar3.legend = None
    ws.add_chart(bar3, f"I{chart_row}")

    if len(unique_periods) > 1:
        line = LineChart(); line.title = "Spend & Offers Trend"; line.height, line.width = 8, 20
        cats2 = Reference(ws, min_col=1, min_row=trend_hdr + 1, max_row=trend_last)
        cost_data = Reference(ws, min_col=2, min_row=trend_hdr, max_row=trend_last)
        offers_data = Reference(ws, min_col=5, min_row=trend_hdr, max_row=trend_last)
        line.add_data(cost_data, titles_from_data=True)
        line.add_data(offers_data, titles_from_data=True)
        line.set_categories(cats2)
        ws.add_chart(line, f"A{chart_row + 17}")

    ws.freeze_panes = "A13"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_area = f"A1:L{chart_row+17}"

    wb.move_sheet("Dashboard", offset=-4)
    wb.save(output_path)
