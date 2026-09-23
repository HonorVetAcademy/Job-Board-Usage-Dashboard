"""Tracks cumulative snapshots per platform/recruiter so period-over-period deltas can be computed
automatically for platforms whose raw exports report running totals rather than clean per-period figures.
"""
import json
import os

SNAPSHOT_FIELDS = ("searches", "profiles_viewed", "outreach_sent", "responses_received")


def load_snapshots(path):
    if os.path.exists(path):
        with open(path) as f:
            return json.load(f)
    return {}


def save_snapshots(path, data):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w") as f:
        json.dump(data, f, indent=2, sort_keys=True)


def period_recruiter_rows(parse_result, snapshots, as_of):
    """Returns a list of dicts, one per recruiter, with values already converted to THIS PERIOD's
    activity: a straight passthrough for period-scoped platforms, or (current - last snapshot) for
    cumulative platforms. Also returns the updated snapshot block to persist.
    """
    platform = parse_result.platform
    prev = snapshots.get(platform, {}).get("recruiters", {})
    new_snapshot_recruiters = {}
    out_rows = []

    for rec in parse_result.recruiters:
        current = {f: getattr(rec, f) for f in SNAPSHOT_FIELDS}
        new_snapshot_recruiters[rec.name] = current

        row = {"name": rec.name, "notes": rec.notes, "first_observation": False}
        prior = prev.get(rec.name)
        for field in SNAPSHOT_FIELDS:
            if field in parse_result.cumulative_fields:
                if prior is None:
                    row[field] = None  # can't diff yet
                    row["first_observation"] = True
                else:
                    row[field] = max(0, current[field] - prior.get(field, 0))
            else:
                row[field] = current[field]
        out_rows.append(row)

    snapshots.setdefault(platform, {})["recruiters"] = new_snapshot_recruiters
    snapshots[platform]["as_of"] = as_of
    return out_rows


def period_platform_totals(parse_result, snapshots, as_of):
    platform = parse_result.platform
    prev = snapshots.get(platform, {}).get("platform_totals", {})
    new_totals = dict(parse_result.platform_totals)
    out = {}
    for key, current_val in parse_result.platform_totals.items():
        if key in parse_result.cumulative_fields:
            if key not in prev:
                out[key] = None
            else:
                out[key] = max(0, current_val - prev.get(key, 0))
        else:
            out[key] = current_val
    snapshots.setdefault(platform, {})["platform_totals"] = new_totals
    snapshots[platform]["as_of"] = as_of
    return out
