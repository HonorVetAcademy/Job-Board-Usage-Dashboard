"""Parses a Vivian.xlsx export."""
import re
from .utils import load_ws, to_num, Recruiter, ParseResult


def parse(filepath):
    rows = load_ws(filepath, "Active Recruiters in Vivian")
    recruiters = []
    if rows:
        for r in rows[1:]:
            if not r or not r[0]:
                continue
            name, inbound, weekly_prop, resp_rate, resp_time = (list(r) + [None]*5)[:5]
            m = re.search(r'(\d+)\s*/\s*(\d+)', weekly_prop or "")
            used = to_num(m.group(1)) if m else 0
            cap = to_num(m.group(2)) if m else None
            status = "Enabled" if "Enabled" in (weekly_prop or "") or "Limit reached" in (weekly_prop or "") else "Disabled"
            recruiters.append(Recruiter(
                name=name, searches=0, profiles_viewed=0, outreach_sent=used, responses_received=0,
                notes=f"Proposal cap: {cap}/wk; Status: {status}; Resp rate: {resp_rate}; Resp time: {resp_time}",
            ))

    active_jobs_rows = load_ws(filepath, "Active Jobs ") or load_ws(filepath, "Active Jobs") or []
    active_postings = max(0, len(active_jobs_rows) - 1) if active_jobs_rows else 0

    return ParseResult(
        platform="Vivian",
        recruiters=recruiters,
        platform_totals={"active_postings": active_postings},
        cumulative_fields=set(),  # weekly proposal cap resets each week; active postings is a current count
        source_note="Active Recruiters in Vivian tab (weekly proposal usage) + Active Jobs tab (current posting count).",
    )
