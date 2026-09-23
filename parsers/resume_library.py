"""Parses a Resume-Library_Reports.xlsx export. Picks the most recent month block present."""
from .utils import load_ws, to_num, Recruiter, ParseResult

MONTHS = ["January", "February", "March", "April", "May", "June", "July", "August",
          "September", "October", "November", "December"]


def parse(filepath):
    rows = load_ws(filepath, "Monthly Report")
    recruiters = []
    month_used = None
    if rows:
        # find every row whose col A is a month name; each starts a fixed-size block of recruiter rows
        month_start_rows = [i for i, r in enumerate(rows) if r and r[0] in MONTHS]
        if month_start_rows:
            last_start = month_start_rows[-1]
            month_used = rows[last_start][0]
            # block runs until the next blank row or 'Period Total' row
            i = last_start
            while i < len(rows):
                r = rows[i]
                if r is None or (r[0] and str(r[0]).strip().lower().startswith("period total")):
                    break
                name = r[1] if len(r) > 1 else None
                if name:
                    searches, views, unlocks, downloads = to_num(r[2]), to_num(r[3]), to_num(r[4]), to_num(r[5])
                    recruiters.append(Recruiter(
                        name=name, searches=searches, profiles_viewed=views, outreach_sent=unlocks,
                        responses_received=0, notes=f"Downloads: {downloads}",
                    ))
                i += 1
    return ParseResult(
        platform="Resume-Library",
        recruiters=recruiters,
        platform_totals={},
        cumulative_fields=set(),  # each month's figures are already period-scoped, not cumulative
        source_note=f"Monthly Report tab, most recent month block found: {month_used}.",
    )
