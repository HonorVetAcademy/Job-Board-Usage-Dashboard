"""Parses a SignalHire_Reports.xlsx export. Picks the most recent named month-block per account."""
from .utils import load_ws, sheet_names, to_num, Recruiter, ParseResult


def _find_latest_month_block(header_row):
    """header_row[i] is a month label (e.g. 'September 2026') at the start of each 15-col block,
    None everywhere else. Returns (start_col_idx, label) for the last one found, or None."""
    blocks = [(i, v) for i, v in enumerate(header_row) if v and v not in ("Total",)]
    return blocks[-1] if blocks else None


def parse(filepath):
    accounts = [s for s in sheet_names(filepath) if s.startswith("Account")]
    recruiters = []
    month_used = None
    for sheet in accounts:
        rows = load_ws(filepath, sheet)
        if not rows or len(rows) < 10:
            continue
        header_row = rows[6] if len(rows) > 6 else None
        if not header_row:
            continue
        block = _find_latest_month_block(header_row)
        if not block:
            continue
        start_col, month_used = block
        for r in rows[9:]:
            if not r or not r[0] or r[0] == "All users":
                continue
            clicks = to_num(r[start_col]) if len(r) > start_col else 0
            successful = to_num(r[start_col + 3]) if len(r) > start_col + 3 else 0
            with_contact = to_num(r[start_col + 6]) if len(r) > start_col + 6 else 0
            if clicks > 0:
                recruiters.append(Recruiter(
                    name=r[0], searches=clicks, profiles_viewed=0,
                    outreach_sent=successful, responses_received=with_contact,
                    notes=f"{month_used} clicks; account: {sheet}; email: {r[1] if len(r)>1 else ''}",
                ))
    return ParseResult(
        platform="SignalHire",
        recruiters=recruiters,
        platform_totals={},
        cumulative_fields=set(),  # the selected month-column is already period-scoped
        source_note=f"Most recent monthly column found across account sheets: {month_used}.",
    )
