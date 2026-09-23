"""Parses a Monster_Reports.xlsx export."""
from .utils import load_ws, to_num, ParseResult


def parse(filepath):
    usage_rows = load_ws(filepath, "Usage")
    credit_used = 0
    if usage_rows:
        for r in usage_rows:
            if r and r[0] == "Credit Used":
                credit_used = to_num(r[1])
                break
    return ParseResult(
        platform="Monster",
        recruiters=[],  # Monster's export has a user roster but no per-recruiter activity breakdown
        platform_totals={"credits_used": credit_used},
        cumulative_fields={"credits_used"},
        source_note="No per-recruiter activity in this export (roster only). Credit usage is cumulative "
                     "for the current subscription term.",
    )
