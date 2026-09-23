"""Parses a DocCafe_Reports.xlsx export."""
from .utils import load_ws, to_num, Recruiter, ParseResult


def parse(filepath):
    rows = load_ws(filepath, "Sheet5")  # Company User, Role, Status, Last Login, Active Days, Searches, Views, Downloads, Saved
    recruiters = []
    if rows:
        for r in rows[1:]:
            if not r[0] or r[0] == "Total":
                continue
            name, role, status, last_login, active_days, searches, profile_views, downloads, saved = (list(r) + [None]*9)[:9]
            recruiters.append(Recruiter(
                name=name, searches=to_num(searches), profiles_viewed=to_num(profile_views),
                outreach_sent=to_num(downloads), responses_received=0,
                notes=f"Active days (cumulative): {active_days}",
            ))
    return ParseResult(
        platform="DocCafe",
        recruiters=recruiters,
        platform_totals={},
        cumulative_fields={"searches", "profiles_viewed", "outreach_sent"},
        source_note="Sheet5 (per-recruiter activity). No cost/submissions/interviews/offers in this export.",
    )
