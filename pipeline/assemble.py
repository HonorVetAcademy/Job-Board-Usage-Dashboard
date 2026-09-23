"""Turns raw-export parse results + the small manual cost/ROI CSV into this period's
Weekly Input row and Recruiter Activity rows per platform.
"""
import csv
import os
from . import history as hist

MANUAL_FIELDS = ["cost", "submissions", "interviews", "offers", "active_postings"]


def load_manual_csv(path):
    """platform -> {cost, submissions, interviews, offers, active_postings}, blanks -> None."""
    out = {}
    if not os.path.exists(path):
        return out
    with open(path, newline="") as f:
        for row in csv.DictReader(f):
            platform = row.get("platform", "").strip()
            if not platform:
                continue
            rec = {}
            for field in MANUAL_FIELDS:
                val = (row.get(field) or "").strip()
                rec[field] = float(val) if val not in ("", None) else None
            out[platform] = rec
    return out


def _safe_sum(values):
    nums = [v for v in values if v is not None]
    return sum(nums) if nums else None


def assemble_period(parse_results, manual_data, period_ending, period_days, snapshots):
    """parse_results: dict platform -> ParseResult.
    Returns (weekly_input_rows: list[dict], recruiter_activity_rows: list[dict], warnings: list[str])
    """
    weekly_rows = []
    recruiter_rows = []
    warnings = []
    as_of = period_ending

    for platform, pr in parse_results.items():
        period_recruiters = hist.period_recruiter_rows(pr, snapshots, as_of)
        period_totals_raw = hist.period_platform_totals(pr, snapshots, as_of)

        first_obs = [r["name"] for r in period_recruiters if r["first_observation"]]
        if first_obs:
            warnings.append(
                f"{platform}: {len(first_obs)} recruiter(s) seen for the first time — no delta yet this "
                f"period (their cumulative counters are only being captured for next time)."
            )

        views_used = _safe_sum(r["profiles_viewed"] for r in period_recruiters)
        contacts_made = _safe_sum(r["outreach_sent"] for r in period_recruiters)
        responses_received = _safe_sum(r["responses_received"] for r in period_recruiters)
        applications = period_totals_raw.get("applications")

        recruiters_assigned = len(pr.recruiters)
        active_recruiters = sum(
            1 for r in period_recruiters
            if any((r.get(f) or 0) > 0 for f in ("searches", "profiles_viewed", "outreach_sent", "responses_received"))
        )

        manual = manual_data.get(platform, {})
        active_postings = manual.get("active_postings")
        if active_postings is None:
            active_postings = period_totals_raw.get("active_postings")

        weekly_rows.append({
            "period_ending": period_ending,
            "platform": platform,
            "cost": manual.get("cost"),
            "recruiters_assigned": recruiters_assigned,
            "active_recruiters": active_recruiters,
            "views_used": views_used,
            "active_postings": active_postings,
            "applications": applications,
            "contacts_made": contacts_made,
            "responses_received": responses_received,
            "submissions": manual.get("submissions"),
            "interviews": manual.get("interviews"),
            "offers": manual.get("offers"),
            "period_length_days": period_days,
            "notes": pr.source_note,
        })

        for r in period_recruiters:
            recruiter_rows.append({
                "period_ending": period_ending,
                "platform": platform,
                "recruiter": r["name"],
                "active": "Yes" if any((r.get(f) or 0) > 0 for f in
                                        ("searches", "profiles_viewed", "outreach_sent", "responses_received")) else "No",
                "searches": r["searches"],
                "profiles_viewed": r["profiles_viewed"],
                "outreach_sent": r["outreach_sent"],
                "responses_received": r["responses_received"],
                "data_scope": "This period only (first observation, no prior data to diff)" if r["first_observation"]
                              else "This period only",
                "notes": r["notes"],
            })

    # platforms present in manual CSV but with no raw export this period (e.g. CareerBuilder, Dice, ...)
    for platform, manual in manual_data.items():
        if platform in parse_results:
            continue
        weekly_rows.append({
            "period_ending": period_ending, "platform": platform, "cost": manual.get("cost"),
            "recruiters_assigned": None, "active_recruiters": None, "views_used": None,
            "active_postings": manual.get("active_postings"), "applications": None, "contacts_made": None,
            "responses_received": None, "submissions": manual.get("submissions"),
            "interviews": manual.get("interviews"), "offers": manual.get("offers"),
            "period_length_days": period_days, "notes": "No raw export parsed for this platform this period.",
        })

    return weekly_rows, recruiter_rows, warnings
