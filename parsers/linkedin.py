"""Parses a LinkedIn.xlsx (Recruiter) export."""
from .utils import load_ws, to_num, Recruiter, ParseResult


def parse(filepath, top_n=None):
    rows = load_ws(filepath, "Usage Details")
    recruiters = []
    if rows:
        for r in rows[1:]:
            if not r[0]:
                continue
            seat_id = r[1]
            active_days = to_num(r[10])
            profiles = to_num(r[15])
            searches = to_num(r[28])
            inmails_sent = to_num(r[34])
            inmails_accepted = to_num(r[37])
            if active_days > 0 or profiles > 0 or searches > 0 or inmails_sent > 0:
                # LinkedIn exports can have two seats sharing a display name (e.g. two accounts for
                # one person, or two different people with the same name) — key on name+seat id so
                # their history/deltas never collide.
                label = f"{r[0]} ({seat_id})"
                recruiters.append(Recruiter(
                    name=label, searches=searches, profiles_viewed=profiles,
                    outreach_sent=inmails_sent, responses_received=inmails_accepted,
                    notes=f"Active days (cumulative): {active_days}",
                ))
    recruiters.sort(key=lambda x: -x.outreach_sent)
    if top_n:
        recruiters = recruiters[:top_n]
    return ParseResult(
        platform="LinkedIn",
        recruiters=recruiters,
        platform_totals={},
        cumulative_fields={"searches", "profiles_viewed", "outreach_sent", "responses_received"},
        source_note="Usage Details tab. Figures are cumulative since contract start; pipeline computes period deltas.",
    )
