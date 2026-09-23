"""Parses an Indeed_Reports.xlsx export."""
from .utils import load_ws, to_num, Recruiter, ParseResult


def parse(filepath):
    rows = load_ws(filepath, "Sheet9")  # per-recruiter sourcing usage
    recruiters = []
    total_applications = 0
    if rows:
        for r in rows[1:]:
            if not r[0]:
                continue
            rec_email, name, sub, searches, resumes_viewed, shared_c, free_c, contacts, cas, resp_rate, \
                responses, prr, prrsa, pos_resp, pos_resp_sa, started_apps, applications = (list(r) + [None]*17)[:17]
            label = name or rec_email
            recruiters.append(Recruiter(
                name=label, searches=to_num(searches), profiles_viewed=to_num(resumes_viewed),
                outreach_sent=to_num(contacts), responses_received=to_num(responses),
                notes=f"Subscription: {sub}",
            ))
            total_applications += to_num(applications)
    return ParseResult(
        platform="Indeed",
        recruiters=recruiters,
        platform_totals={"applications": total_applications},
        cumulative_fields={"searches", "profiles_viewed", "outreach_sent", "responses_received", "applications"},
        source_note="Sheet9 (per-recruiter sourcing usage). Cost/submissions/interviews/offers not in this export.",
    )
