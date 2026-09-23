# Job Portal Performance Tracker — Automated Pipeline

Drop in the raw weekly exports from each job board, and this regenerates:
- `output/Job_Portal_Performance_Tracker.xlsx` — the full Excel dashboard (KPIs, platform table, charts)
- `docs/index.html` — a live web dashboard (deployable to GitHub Pages)

Recruiter/usage data is parsed automatically from the raw exports. **Cost, Submissions, Interviews,
and Offers are not in any raw export** — no platform's dashboard exposes them in a downloadable
file — so those five numbers per platform still need a small manual CSV each period (~2 minutes,
see below).

---

## ⚠️ Before you push this anywhere: privacy

The raw exports in `data/raw/` contain candidate names, emails, and recruiter performance data.

- **This repo must be private.** Do not create it as a public repository.
- If you ever want a *public* dashboard, publish only `docs/` (aggregated platform-level stats,
  no candidate PII) separately — don't make the repo itself public, since `data/raw/` would go
  with it.
- GitHub Pages built from a private repo is only private-by-default on GitHub Enterprise/certain
  paid plans — check your plan's Pages visibility settings before enabling it, or restrict Pages
  access accordingly.

---

## One-time setup

```bash
# 1. Create a new PRIVATE repo on GitHub first (via github.com/new), then:
cd job-portal-dashboard
git init
git add .
git commit -m "Initial commit"
git branch -M main
git remote add origin https://github.com/<your-username>/<your-repo>.git
git push -u origin main

# 2. In the repo on GitHub: Settings -> Pages -> Source -> "GitHub Actions"
#    (This lets the workflow deploy docs/ automatically. Skip this step if you only want the
#    Excel file and don't need a hosted web dashboard.)
```

That's it — the next push that touches `data/raw/` or `data/manual/` will trigger the workflow.

---

## Weekly workflow

1. Download this period's export from each job board (same files you already pull: DocCafe,
   Indeed, LinkedIn, Monster, Resume-Library, SignalHire, Vivian — whatever you have).
2. Create a new dated folder and drop them in:
   ```
   data/raw/2026-09-29/
     DocCafe_Reports.xlsx
     Indeed_Reports.xlsx
     LinkedIn.xlsx
     Monster_Reports.xlsx
     Resume-Library_Reports.xlsx
     SignalHire_Reports.xlsx
     Vivian.xlsx
   ```
   (Filenames just need to *contain* the platform name, case-insensitive — exact export
   filenames from each vendor work as-is.)
3. Fill in the tiny manual CSV for that same date — copy `data/manual/2026-09-15.csv` as a
   template. One row per platform:
   ```
   platform,cost,submissions,interviews,offers,active_postings
   Vivian,1875.00,19,1,3,
   Monster,666.41,20,2,1,35
   ...
   ```
   Leave `active_postings` blank for any platform the pipeline can already derive it for
   (currently: Vivian, from its live job count).
4. Commit and push:
   ```bash
   git add data/raw/2026-09-29 data/manual/2026-09-29.csv
   git commit -m "Week of 2026-09-29"
   git push
   ```
5. GitHub Actions runs automatically: parses the raw files, updates the history, regenerates
   the Excel file and web dashboard, commits the results back, and (if Pages is enabled) deploys
   the live site.

No local Python needed for the weekly cadence — only for the one-time setup test or if you'd
rather run it locally instead of via Actions (see below).

---

## Running it locally instead of / in addition to Actions

```bash
pip install -r requirements.txt
python3 build.py --raw-dir data/raw/2026-09-29 --period-ending 2026-09-29 --period-days 7
```

`--period-days` should match how many days that export actually covers (7 for a normal week —
the very first period you ever load was a 15-day bi-weekly snapshot, hence `--period-days 15`
in the seed data).

---

## How the "cumulative vs. period" problem is handled

Some platforms' raw exports report **running totals since your contract started** (LinkedIn,
Indeed, DocCafe), not a clean per-period figure. Others already report a clean period-scoped
number (Resume-Library's monthly breakdown, Vivian's weekly proposal cap, SignalHire's monthly
columns).

The pipeline doesn't need you to know which is which. It stores every raw snapshot, dated, in
`data/history/snapshots.json`, and for platforms flagged as cumulative it computes
`this_period = current_snapshot − last_snapshot`. Practical effect:

- **The very first time you run it**, cumulative-platform recruiters will show blank/zero
  activity for that period — there's no prior snapshot yet to diff against. This is expected,
  not a bug. From the second period onward, real deltas appear automatically.
- If a brand-new recruiter shows up on a cumulative platform, their first appearance is also a
  "first observation" (no delta yet) for the same reason.

---

## Known gaps (things no raw export currently gives us)

- **Cost, Submissions, Interviews, Offers** — manual CSV, every platform, every period. If you
  ever get a consistent export for any of these (e.g. an invoice API, an ATS report), tell
  whoever maintains this pipeline — it's a small addition to wire in.
- **Monster** — the export has a user roster but no per-recruiter activity breakdown, so it
  never appears in Recruiter Activity. Check whether Monster's dashboard offers a per-user
  export; if so, a parser can be added for it.
- **CareerBuilder, Dice, Facebook, Job Diva, Zoominfo, Referred** — no raw per-platform export
  has ever been provided for these; they exist only via the manual CSV. If you start receiving
  raw exports for any of them, a parser can be added the same way as the other seven.

---

## Repo structure

```
build.py                       # entrypoint — run this each period
parsers/                       # one module per platform, each returns a standard ParseResult
pipeline/
  history.py                   # cumulative-snapshot delta engine
  assemble.py                  # merges parsed data + manual CSV into this period's rows
  excel_dashboard.py           # builds the .xlsx from full accumulated history
  web_dashboard.py             # builds docs/data.json + docs/index.html
data/
  raw/<period>/                # you drop platform exports here, one dated folder per period
  manual/<period>.csv          # you fill this in, one row per platform
  history/                     # pipeline-owned — weekly_input.csv, recruiter_activity.csv,
                                # snapshots.json — never hand-edit these
output/Job_Portal_Performance_Tracker.xlsx   # regenerated every run
docs/                          # regenerated every run; served by GitHub Pages if enabled
.github/workflows/update-dashboard.yml       # the automation
```

## Testing changes before you trust them

Before relying on a change to a parser, run it against a real export and sanity-check the
numbers against what the platform's own dashboard shows, the same way you'd check any new report.
`python3 build.py --raw-dir ... --period-ending ... --force` reprocesses a period you've already
run, useful when iterating.
