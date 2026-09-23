"""Computes the same metrics as the Excel dashboard in plain Python and writes docs/data.json +
a static docs/index.html for GitHub Pages (no backend, just Chart.js reading the JSON).
"""
import json
import os
from .excel_dashboard import PLATFORM_SETTINGS

CAP_LOOKUP = {p[0]: (p[2], p[3]) for p in PLATFORM_SETTINGS}  # platform -> (cap_value, cap_period)
POST_CAP_LOOKUP = {p[0]: p[4] for p in PLATFORM_SETTINGS}


def _num(v):
    if v is None or v == "" or v == "None":
        return None
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def _daily_cap(platform):
    val, period = CAP_LOOKUP.get(platform, (None, ""))
    if val is None:
        return None
    if period == "Year":
        return val / 365
    if period == "Month":
        return val * 12 / 365
    if period == "Week":
        return val / 7
    return None


def _safe_div(a, b):
    if a is None or b in (None, 0):
        return None
    return a / b


def compute_dashboard(weekly_history):
    if not weekly_history:
        return {"period": None, "kpis": {}, "platforms": [], "trend": []}

    periods = sorted({r["period_ending"] for r in weekly_history})
    latest = periods[-1]
    rows = [r for r in weekly_history if r["period_ending"] == latest]

    platforms = []
    total_cost = total_sub = total_int = total_off = 0
    active_pcts, views_pcts = [], []
    for r in rows:
        cost = _num(r["cost"]) or 0
        active = _num(r["active_recruiters"])
        assigned = _num(r["recruiters_assigned"])
        active_pct = _safe_div(active, assigned)
        views_used = _num(r["views_used"])
        days = _num(r["period_length_days"]) or 7
        daily_cap = _daily_cap(r["platform"])
        period_cap = daily_cap * days if daily_cap is not None else None
        views_pct = _safe_div(views_used, period_cap)
        sub = _num(r["submissions"]) or 0
        interviews = _num(r["interviews"]) or 0
        offers = _num(r["offers"]) or 0
        cost_per_offer = _safe_div(cost, offers)

        total_cost += cost
        total_sub += sub
        total_int += interviews
        total_off += offers
        if active_pct is not None:
            active_pcts.append(active_pct)
        if views_pct is not None:
            views_pcts.append(views_pct)

        platforms.append({
            "platform": r["platform"], "cost": cost, "active_recruiters": active, "recruiters_assigned": assigned,
            "active_pct": active_pct, "views_pct": views_pct, "submissions": sub, "interviews": interviews,
            "offers": offers, "cost_per_offer": cost_per_offer,
        })

    for p in platforms:
        p["share_of_offers"] = _safe_div(p["offers"], total_off)

    kpis = {
        "total_spend": total_cost, "total_submissions": total_sub, "total_interviews": total_int,
        "total_offers": total_off, "offer_rate": _safe_div(total_off, total_sub),
        "avg_cost_per_offer": _safe_div(total_cost, total_off),
        "avg_recruiter_utilization": (sum(active_pcts) / len(active_pcts)) if active_pcts else None,
        "avg_views_utilization": (sum(views_pcts) / len(views_pcts)) if views_pcts else None,
    }

    trend = []
    for period in periods:
        prows = [r for r in weekly_history if r["period_ending"] == period]
        trend.append({
            "period": period,
            "cost": sum((_num(r["cost"]) or 0) for r in prows),
            "submissions": sum((_num(r["submissions"]) or 0) for r in prows),
            "interviews": sum((_num(r["interviews"]) or 0) for r in prows),
            "offers": sum((_num(r["offers"]) or 0) for r in prows),
        })

    return {"period": latest, "kpis": kpis, "platforms": platforms, "trend": trend}


HTML_TEMPLATE = """<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">
<title>Job Portal Performance Dashboard</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.min.js"></script>
<style>
  :root {
    --navy: #1F4E78; --bg: #f5f7fa; --card: #ffffff; --border: #e2e8f0; --text: #1a202c; --muted: #718096;
    --good: #63BE7B; --mid: #FFEB84; --bad: #F8696B;
  }
  * { box-sizing: border-box; }
  body { font-family: Arial, Helvetica, sans-serif; margin: 0; background: var(--bg); color: var(--text);
         padding-top: env(safe-area-inset-top, 0px); padding-bottom: env(safe-area-inset-bottom, 0px); }
  header { background: var(--navy); color: #fff; padding: 20px 24px; }
  header h1 { margin: 0; font-size: 1.4rem; }
  header p { margin: 6px 0 0; opacity: 0.85; font-size: 0.85rem; }
  main { max-width: 1200px; margin: 0 auto; padding: 20px; }
  .kpis { display: grid; grid-template-columns: repeat(auto-fit, minmax(180px, 1fr)); gap: 12px; margin-bottom: 24px; }
  .kpi { background: var(--card); border: 1px solid var(--border); border-radius: 8px; padding: 14px; text-align: center; }
  .kpi .label { font-size: 0.72rem; font-weight: bold; color: var(--muted); text-transform: uppercase; }
  .kpi .value { font-size: 1.6rem; font-weight: bold; color: var(--navy); margin-top: 4px; }
  h2 { font-size: 1.1rem; color: var(--navy); margin: 28px 0 10px; }
  table { width: 100%; border-collapse: collapse; background: var(--card); border-radius: 8px; overflow: hidden;
          box-shadow: 0 1px 3px rgba(0,0,0,0.08); }
  th, td { padding: 8px 10px; text-align: center; font-size: 0.85rem; border-bottom: 1px solid var(--border); }
  th { background: var(--navy); color: #fff; font-size: 0.75rem; }
  td:first-child, th:first-child { text-align: left; }
  .charts { display: grid; grid-template-columns: repeat(auto-fit, minmax(320px, 1fr)); gap: 16px; margin-top: 20px; }
  .chart-box { background: var(--card); border: 1px solid var(--border); border-radius: 8px; padding: 12px; }
  .overflow { overflow-x: auto; }
  footer { text-align: center; color: var(--muted); font-size: 0.75rem; padding: 20px; }
</style>
</head>
<body>
<header>
  <h1>HonorVet — Job Portal Performance Dashboard</h1>
  <p id="period-label">Loading…</p>
</header>
<main>
  <div class="kpis" id="kpi-cards"></div>
  <h2>Platform Performance</h2>
  <div class="overflow">
    <table id="platform-table">
      <thead><tr>
        <th>Platform</th><th>Cost</th><th>Active / Assigned</th><th>Active %</th><th>Views Util %</th>
        <th>Submissions</th><th>Interviews</th><th>Offers</th><th>Cost / Offer</th><th>Share of Offers</th>
      </tr></thead>
      <tbody></tbody>
    </table>
  </div>
  <div class="charts">
    <div class="chart-box"><canvas id="costChart"></canvas></div>
    <div class="chart-box"><canvas id="offersChart"></canvas></div>
    <div class="chart-box"><canvas id="activeChart"></canvas></div>
    <div class="chart-box"><canvas id="trendChart"></canvas></div>
  </div>
</main>
<footer>Auto-generated by the pipeline on every push. Not hand-edited.</footer>
<script>
async function load() {
  const res = await fetch('data.json');
  const d = await res.json();
  if (!d.period) { document.getElementById('period-label').textContent = 'No data yet.'; return; }
  document.getElementById('period-label').textContent = 'Period: ' + d.period;

  const pct = v => v === null || v === undefined ? '–' : (v*100).toFixed(1) + '%';
  const money = v => v === null || v === undefined ? '–' : '$' + Math.round(v).toLocaleString();
  const num = v => v === null || v === undefined ? '–' : Math.round(v).toLocaleString();

  const kpis = [
    ['TOTAL SPEND', money(d.kpis.total_spend)],
    ['TOTAL SUBMISSIONS', num(d.kpis.total_submissions)],
    ['TOTAL INTERVIEWS', num(d.kpis.total_interviews)],
    ['TOTAL OFFERS', num(d.kpis.total_offers)],
    ['OFFER RATE', pct(d.kpis.offer_rate)],
    ['AVG COST / OFFER', money(d.kpis.avg_cost_per_offer)],
    ['AVG RECRUITER UTIL', pct(d.kpis.avg_recruiter_utilization)],
    ['AVG VIEWS UTIL', pct(d.kpis.avg_views_utilization)],
  ];
  document.getElementById('kpi-cards').innerHTML = kpis.map(([l,v]) =>
    `<div class="kpi"><div class="label">${l}</div><div class="value">${v}</div></div>`).join('');

  const tbody = document.querySelector('#platform-table tbody');
  tbody.innerHTML = d.platforms.map(p => `<tr>
      <td>${p.platform}</td><td>${money(p.cost)}</td>
      <td>${p.active_recruiters ?? '–'} / ${p.recruiters_assigned ?? '–'}</td>
      <td>${pct(p.active_pct)}</td><td>${pct(p.views_pct)}</td>
      <td>${num(p.submissions)}</td><td>${num(p.interviews)}</td><td>${num(p.offers)}</td>
      <td>${money(p.cost_per_offer)}</td><td>${pct(p.share_of_offers)}</td>
    </tr>`).join('');

  const labels = d.platforms.map(p => p.platform);
  new Chart(document.getElementById('costChart'), { type: 'bar',
    data: { labels, datasets: [{ label: 'Cost ($)', data: d.platforms.map(p => p.cost), backgroundColor: '#1F4E78' }] },
    options: { plugins: { title: { display: true, text: 'Cost by Platform' } } } });
  new Chart(document.getElementById('offersChart'), { type: 'bar',
    data: { labels, datasets: [{ label: 'Offers', data: d.platforms.map(p => p.offers), backgroundColor: '#63BE7B' }] },
    options: { plugins: { title: { display: true, text: 'Offers by Platform' } } } });
  new Chart(document.getElementById('activeChart'), { type: 'bar',
    data: { labels, datasets: [{ label: 'Active %', data: d.platforms.map(p => (p.active_pct||0)*100), backgroundColor: '#F8A652' }] },
    options: { plugins: { title: { display: true, text: 'Active Recruiter % by Platform' } },
               scales: { y: { max: 100 } } } });
  new Chart(document.getElementById('trendChart'), { type: 'line',
    data: { labels: d.trend.map(t => t.period),
            datasets: [
              { label: 'Total Cost', data: d.trend.map(t => t.cost), borderColor: '#1F4E78', yAxisID: 'y' },
              { label: 'Total Offers', data: d.trend.map(t => t.offers), borderColor: '#63BE7B', yAxisID: 'y1' },
            ] },
    options: { plugins: { title: { display: true, text: 'Spend & Offers Trend' } },
               scales: { y: { position: 'left' }, y1: { position: 'right', grid: { drawOnChartArea: false } } } } });
}
load();
</script>
</body>
</html>
"""


def build_web(weekly_history, recruiter_history, docs_dir):
    data = compute_dashboard(weekly_history)
    os.makedirs(docs_dir, exist_ok=True)
    with open(os.path.join(docs_dir, "data.json"), "w") as f:
        json.dump(data, f, indent=2)
    with open(os.path.join(docs_dir, "index.html"), "w") as f:
        f.write(HTML_TEMPLATE)
