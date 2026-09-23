#!/usr/bin/env python3
"""
Job Portal Performance Tracker — pipeline entrypoint.

Usage:
    python build.py --raw-dir data/raw/2026-09-22 --period-ending 2026-09-22 [--period-days 7]

What it does:
  1. Parses every recognized raw export in --raw-dir (DocCafe, Indeed, LinkedIn, Monster,
     Resume-Library, SignalHire, Vivian).
  2. Diffs cumulative-counter exports against the last snapshot to get this period's real usage
     (data/history/snapshots.json), and uses period-scoped exports directly.
  3. Merges in data/manual/<period-ending>.csv for cost/submissions/interviews/offers (the figures
     no raw export contains).
  4. Appends this period's rows to data/history/weekly_input.csv and recruiter_activity.csv
     (the permanent, ever-growing history — never overwritten).
  5. Regenerates output/Job_Portal_Performance_Tracker.xlsx and docs/data.json + docs/index.html
     from the FULL accumulated history.
"""
import argparse
import csv
import os
import sys

sys.path.insert(0, os.path.dirname(__file__))
from parsers import parser_for_filename
from pipeline import history as hist
from pipeline.assemble import assemble_period, load_manual_csv
from pipeline.excel_dashboard import build_excel
from pipeline.web_dashboard import build_web

ROOT = os.path.dirname(os.path.abspath(__file__))
SNAPSHOTS_PATH = os.path.join(ROOT, "data", "history", "snapshots.json")
WEEKLY_CSV = os.path.join(ROOT, "data", "history", "weekly_input.csv")
RECRUITER_CSV = os.path.join(ROOT, "data", "history", "recruiter_activity.csv")
MANUAL_DIR = os.path.join(ROOT, "data", "manual")
OUTPUT_XLSX = os.path.join(ROOT, "output", "Job_Portal_Performance_Tracker.xlsx")
DOCS_DIR = os.path.join(ROOT, "docs")

WEEKLY_FIELDS = ["period_ending", "platform", "cost", "recruiters_assigned", "active_recruiters",
                  "views_used", "active_postings", "applications", "contacts_made", "responses_received",
                  "submissions", "interviews", "offers", "period_length_days", "notes"]
RECRUITER_FIELDS = ["period_ending", "platform", "recruiter", "active", "searches", "profiles_viewed",
                     "outreach_sent", "responses_received", "data_scope", "notes"]


def append_csv(path, fieldnames, rows):
    file_exists = os.path.exists(path)
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "a", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        if not file_exists:
            writer.writeheader()
        for row in rows:
            writer.writerow(row)


def read_csv_history(path, fieldnames):
    if not os.path.exists(path):
        return []
    with open(path, newline="") as f:
        return list(csv.DictReader(f))


def already_processed(period_ending, weekly_history):
    return any(r["period_ending"] == period_ending for r in weekly_history)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--raw-dir", required=True, help="Folder containing this period's raw exports")
    ap.add_argument("--period-ending", required=True, help="YYYY-MM-DD")
    ap.add_argument("--period-days", type=int, default=7)
    ap.add_argument("--force", action="store_true", help="Reprocess even if this period_ending already exists")
    args = ap.parse_args()

    weekly_history = read_csv_history(WEEKLY_CSV, WEEKLY_FIELDS)
    if already_processed(args.period_ending, weekly_history) and not args.force:
        print(f"Period {args.period_ending} already in history — skipping (use --force to redo).")
    else:
        # 1. Parse every raw file we recognize
        parse_results = {}
        skipped = []
        for fname in sorted(os.listdir(args.raw_dir)):
            fpath = os.path.join(args.raw_dir, fname)
            if not os.path.isfile(fpath) or fname.startswith("~$"):
                continue
            module = parser_for_filename(fname)
            if module is None:
                skipped.append(fname)
                continue
            result = module.parse(fpath)
            parse_results[result.platform] = result
        if skipped:
            print("Skipped (no matching parser):", skipped)

        # 2. Manual cost/ROI CSV for this period
        manual_path = os.path.join(MANUAL_DIR, f"{args.period_ending}.csv")
        manual_data = load_manual_csv(manual_path)
        if not manual_data:
            print(f"WARNING: no manual data file found at {manual_path} — Cost/Submissions/Interviews/"
                  f"Offers will be blank for platforms not covered by a raw export this period.")

        # 3. Load snapshots, assemble this period, save snapshots
        snapshots = hist.load_snapshots(SNAPSHOTS_PATH)
        weekly_rows, recruiter_rows, warnings = assemble_period(
            parse_results, manual_data, args.period_ending, args.period_days, snapshots)
        hist.save_snapshots(SNAPSHOTS_PATH, snapshots)

        for w in warnings:
            print("NOTE:", w)

        # 4. Append to permanent history
        append_csv(WEEKLY_CSV, WEEKLY_FIELDS, weekly_rows)
        append_csv(RECRUITER_CSV, RECRUITER_FIELDS, recruiter_rows)
        print(f"Appended {len(weekly_rows)} Weekly Input rows and {len(recruiter_rows)} Recruiter Activity "
              f"rows for period {args.period_ending}.")

    # 5. Regenerate outputs from the FULL accumulated history every run
    weekly_history = read_csv_history(WEEKLY_CSV, WEEKLY_FIELDS)
    recruiter_history = read_csv_history(RECRUITER_CSV, RECRUITER_FIELDS)

    os.makedirs(os.path.dirname(OUTPUT_XLSX), exist_ok=True)
    build_excel(weekly_history, recruiter_history, OUTPUT_XLSX)
    print("Wrote", OUTPUT_XLSX)

    os.makedirs(DOCS_DIR, exist_ok=True)
    build_web(weekly_history, recruiter_history, DOCS_DIR)
    print("Wrote", os.path.join(DOCS_DIR, "data.json"), "and index.html")


if __name__ == "__main__":
    main()
