"""Computes the same metrics as the Excel dashboard in plain Python and writes docs/data.json +
a static docs/index.html for GitHub Pages (no backend, just Chart.js reading the JSON).
"""
import json
import os
from .excel_dashboard import PLATFORM_SETTINGS

CAP_LOOKUP = {p[0]: (p[2], p[3]) for p in PLATFORM_SETTINGS}  # platform -> (cap_value, cap_period)
POST_CAP_LOOKUP = {p[0]: p[4] for p in PLATFORM_SETTINGS}      # platform -> raw postings cap (num or text)


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


def compute_dashboard(weekly_history, recruiter_history):
    if not weekly_history:
        return {"period": None, "kpis": {}, "platforms": [], "trend": []}

    periods = sorted({r["period_ending"] for r in weekly_history})
    latest = periods[-1]
    rows = [r for r in weekly_history if r["period_ending"] == latest]
    rec_rows = [r for r in recruiter_history if r["period_ending"] == latest]

    platforms = []
    total_cost = total_sub = total_int = total_off = 0
    active_pcts, views_pcts = [], []
    for r in rows:
        platform = r["platform"]
        cost = _num(r["cost"]) or 0
        active = _num(r["active_recruiters"])
        assigned = _num(r["recruiters_assigned"])
        active_pct = _safe_div(active, assigned)

        views_used = _num(r["views_used"])
        days = _num(r["period_length_days"]) or 7
        daily_cap = _daily_cap(platform)
        views_cap = daily_cap * days if daily_cap is not None else None
        views_pct = _safe_div(views_used, views_cap)

        postings_used = _num(r["active_postings"])
        raw_post_cap = POST_CAP_LOOKUP.get(platform)
        postings_cap = raw_post_cap if isinstance(raw_post_cap, (int, float)) else None
        postings_pct = _safe_div(postings_used, postings_cap)

        sub = _num(r["submissions"]) or 0
        interviews = _num(r["interviews"]) or 0
        offers = _num(r["offers"]) or 0
        cost_per_offer = _safe_div(cost, offers)

        # Recruiter activity/usage aggregated up to platform level for this period
        platform_recs = [rr for rr in rec_rows if rr["platform"] == platform]
        total_searches = sum((_num(rr["searches"]) or 0) for rr in platform_recs) if platform_recs else None
        total_profiles = sum((_num(rr["profiles_viewed"]) or 0) for rr in platform_recs) if platform_recs else None
        total_outreach = sum((_num(rr["outreach_sent"]) or 0) for rr in platform_recs) if platform_recs else None
        total_responses = sum((_num(rr["responses_received"]) or 0) for rr in platform_recs) if platform_recs else None

        total_cost += cost
        total_sub += sub
        total_int += interviews
        total_off += offers
        if active_pct is not None:
            active_pcts.append(active_pct)
        if views_pct is not None:
            views_pcts.append(views_pct)

        platforms.append({
            "platform": platform, "cost": cost, "active_recruiters": active, "recruiters_assigned": assigned,
            "active_pct": active_pct,
            "views_used": views_used, "views_cap": views_cap, "views_pct": views_pct,
            "postings_used": postings_used, "postings_cap_raw": raw_post_cap, "postings_pct": postings_pct,
            "submissions": sub, "interviews": interviews, "offers": offers, "cost_per_offer": cost_per_offer,
            "total_searches": total_searches, "total_profiles_viewed": total_profiles,
            "total_outreach_sent": total_outreach, "total_responses_received": total_responses,
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
<title>HonorVet — Recruiting Spend Performance</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Source+Serif+4:opsz,wght@8..60,500;8..60,600;8..60,700&family=IBM+Plex+Sans:wght@400;500;600&display=swap" rel="stylesheet">
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.min.js"></script>
<style>
  :root {
    --ink: #0F1B2D;
    --panel: #16233A;
    --panel-alt: #1C2B44;
    --paper: #EDEFF3;
    --muted: #8B97AC;
    --brass: #C9A15A;
    --brass-dim: #8C7343;
    --sage: #6FA98A;
    --clay: #D08957;
    --hairline: #263752;
  }
  * { box-sizing: border-box; }
  html { scroll-padding-top: env(safe-area-inset-top, 0px); }
  body {
    margin: 0; background: var(--ink); color: var(--paper);
    font-family: "IBM Plex Sans", Arial, sans-serif;
    font-variant-numeric: tabular-nums;
    padding-top: env(safe-area-inset-top, 0px);
    padding-bottom: env(safe-area-inset-bottom, 0px);
  }
  .serif { font-family: "Source Serif 4", Georgia, serif; }
  main { max-width: 980px; margin: 0 auto; padding: 40px 24px 64px; }

  header.masthead {
    display: flex; justify-content: space-between; align-items: baseline;
    padding-bottom: 18px; border-bottom: 1px solid var(--hairline);
  }
  header.masthead .brand { font-size: 0.95rem; font-weight: 600; letter-spacing: 0.02em; }
  header.masthead .period { font-size: 0.85rem; color: var(--muted); }

  section.hero { padding: 40px 0 36px; border-bottom: 1px solid var(--hairline); }
  .hero-figure { display: flex; align-items: baseline; gap: 14px; flex-wrap: wrap; }
  .hero-figure .value { font-size: 3.6rem; font-weight: 600; line-height: 1; }
  .hero-figure .label { font-size: 1rem; color: var(--muted); }
  .hero-stats { display: flex; gap: 40px; margin-top: 26px; flex-wrap: wrap; }
  .hero-stats .stat .value { font-size: 1.5rem; font-weight: 600; }
  .hero-stats .stat .label { font-size: 0.8rem; color: var(--muted); margin-top: 2px; }

  section.funnel { padding: 36px 0; border-bottom: 1px solid var(--hairline); }
  h2.section-title { font-size: 0.95rem; font-weight: 600; color: var(--paper); margin: 0 0 22px; }
  .funnel-row { display: flex; align-items: stretch; gap: 0; }
  .funnel-stage { flex: 1; position: relative; padding-right: 28px; }
  .funnel-stage:last-child { padding-right: 0; }
  .funnel-stage .bar { height: 6px; background: var(--panel-alt); border-radius: 3px; overflow: hidden; margin-bottom: 12px; }
  .funnel-stage .bar-fill { height: 100%; background: var(--brass); border-radius: 3px; }
  .funnel-stage .value { font-size: 1.6rem; font-weight: 600; font-family: "Source Serif 4", serif; }
  .funnel-stage .label { font-size: 0.8rem; color: var(--muted); margin-top: 2px; }
  .funnel-stage .arrow { position: absolute; right: 6px; top: 0; color: var(--hairline); font-size: 1.2rem; }

  section.usage, section.capacity, section.roi { padding: 36px 0; border-bottom: 1px solid var(--hairline); }
  h2.section-title { font-size: 0.95rem; font-weight: 600; color: var(--paper); margin: 0 0 4px; }
  .section-sub { font-size: 0.8rem; color: var(--muted); margin: 0 0 20px; }

  .row-head { display: grid; gap: 12px; font-size: 0.75rem; color: var(--muted);
              padding: 0 0 10px; border-bottom: 1px solid var(--hairline); }
  .row-body { display: grid; gap: 12px; align-items: center; padding: 13px 0; border-bottom: 1px solid var(--hairline); }
  .row-body:last-child { border-bottom: none; }
  .row-body .platform { font-weight: 500; }
  .row-body .sub { font-size: 0.78rem; color: var(--muted); margin-top: 2px; }
  .row-body .figure { text-align: right; font-weight: 500; }
  .row-body .figure .sub { text-align: right; }

  .usage .row-head, .usage .row-body { grid-template-columns: 1.3fr 1fr 0.85fr 0.85fr 0.9fr 0.9fr; }
  .roi .row-head, .roi .row-body { grid-template-columns: 1.3fr 0.9fr 0.85fr 0.85fr 0.85fr 0.9fr 0.9fr; }
  .capacity .row-head, .capacity .row-body { grid-template-columns: 1.3fr 1.7fr 1.7fr; }

  .track { height: 5px; background: var(--panel-alt); border-radius: 3px; overflow: hidden; }
  .fill.cost { background: var(--brass); }
  .fill.offers { background: var(--sage); }
  .fill.views { background: var(--brass); }
  .fill.postings { background: var(--sage); }
  .bar-cell .num { font-size: 0.82rem; color: var(--muted); margin-top: 5px; }
  .na { color: var(--muted); }
  .flag { color: var(--clay); }

  section.trend { padding: 36px 0 0; }
  .trend-chart-box { height: 220px; }

  footer { color: var(--muted); font-size: 0.75rem; padding-top: 28px; }

  .overflow-x { overflow-x: auto; }
  @media (max-width: 760px) {
    .hero-figure .value { font-size: 2.6rem; }
    .funnel-row { flex-direction: column; gap: 20px; }
    .funnel-stage { padding-right: 0; }
    .funnel-stage .arrow { display: none; }
    .row-head { display: none; }
    .row-body { grid-template-columns: 1fr !important; row-gap: 8px; }
    .row-body .figure { text-align: left; }
    .row-body .figure .sub { text-align: left; }
    .row-body .figure::before { content: attr(data-label) ": "; color: var(--muted); font-weight: 400; }
  }
  :focus-visible { outline: 2px solid var(--brass); outline-offset: 2px; }
</style>
</head>
<body>
<main>
  <header class="masthead">
    <div class="brand">HonorVet</div>
    <div class="period" id="period-label">Loading…</div>
  </header>

  <section class="hero">
    <div class="hero-figure">
      <span class="value serif" id="hero-spend">–</span>
      <span class="label">total recruiting spend this period</span>
    </div>
    <div class="hero-stats">
      <div class="stat"><div class="value serif" id="hero-offers">–</div><div class="label">offers made</div></div>
      <div class="stat"><div class="value serif" id="hero-rate">–</div><div class="label">offer rate</div></div>
      <div class="stat"><div class="value serif" id="hero-cpo">–</div><div class="label">cost per offer</div></div>
      <div class="stat"><div class="value serif" id="hero-util">–</div><div class="label">avg recruiter utilization</div></div>
    </div>
  </section>

  <section class="funnel">
    <h2 class="section-title">The funnel</h2>
    <div class="funnel-row" id="funnel-row"></div>
  </section>

  <section class="usage">
    <h2 class="section-title">Recruiter activity &amp; usage per platform</h2>
    <p class="section-sub">Recruiters assigned vs. actually active, and what that activity looked like this period.</p>
    <div class="overflow-x">
      <div style="min-width:640px;">
        <div class="row-head">
          <div>Platform</div><div>Active / assigned</div><div>Searches</div><div>Profiles viewed</div>
          <div>Outreach sent</div><div>Responses</div>
        </div>
        <div id="usage-body"></div>
      </div>
    </div>
  </section>

  <section class="capacity">
    <h2 class="section-title">Views &amp; postings — utilized vs. available</h2>
    <p class="section-sub">This period's usage against each platform's subscription cap.</p>
    <div class="overflow-x">
      <div style="min-width:520px;">
        <div class="row-head">
          <div>Platform</div><div>Views used / cap</div><div>Postings used / cap</div>
        </div>
        <div id="capacity-body"></div>
      </div>
    </div>
  </section>

  <section class="roi">
    <h2 class="section-title">Cost &amp; ROI per platform</h2>
    <p class="section-sub">Ranked by share of total offers.</p>
    <div class="overflow-x">
      <div style="min-width:680px;">
        <div class="row-head">
          <div>Platform</div><div>Cost</div><div>Submissions</div><div>Interviews</div><div>Offers</div>
          <div>Cost / offer</div><div>Share of offers</div>
        </div>
        <div id="roi-body"></div>
      </div>
    </div>
  </section>

  <section class="trend">
    <h2 class="section-title">Spend &amp; offers over time</h2>
    <div class="trend-chart-box"><canvas id="trendChart"></canvas></div>
  </section>

  <footer>Auto-generated by the pipeline on every push. Not hand-edited.</footer>
</main>
<script>
async function load() {
  const res = await fetch('data.json');
  const d = await res.json();
  if (!d.period) { document.getElementById('period-label').textContent = 'No data yet.'; return; }

  const money = v => v === null || v === undefined ? '–' : '$' + Math.round(v).toLocaleString();
  const num = v => v === null || v === undefined ? '–' : Math.round(v).toLocaleString();
  const pct = v => v === null || v === undefined ? '–' : (v*100).toFixed(1) + '%';

  document.getElementById('period-label').textContent =
    new Date(d.period + 'T00:00:00').toLocaleDateString('en-US', { year: 'numeric', month: 'long', day: 'numeric' });

  document.getElementById('hero-spend').textContent = money(d.kpis.total_spend);
  document.getElementById('hero-offers').textContent = num(d.kpis.total_offers);
  document.getElementById('hero-rate').textContent = pct(d.kpis.offer_rate);
  document.getElementById('hero-cpo').textContent = money(d.kpis.avg_cost_per_offer);
  document.getElementById('hero-util').textContent = pct(d.kpis.avg_recruiter_utilization);

  // ---- Funnel: Spend -> Submissions -> Interviews -> Offers ----
  const stages = [
    { label: 'Spend', value: d.kpis.total_spend, display: money(d.kpis.total_spend) },
    { label: 'Submissions', value: d.kpis.total_submissions, display: num(d.kpis.total_submissions) },
    { label: 'Interviews', value: d.kpis.total_interviews, display: num(d.kpis.total_interviews) },
    { label: 'Offers', value: d.kpis.total_offers, display: num(d.kpis.total_offers) },
  ];
  const maxima = stages.slice(1).map(s => s.value || 0);
  const maxNonSpend = Math.max(1, ...maxima);
  const funnelHtml = stages.map((s, i) => {
    const pctWidth = i === 0 ? 100 : Math.max(4, ((s.value || 0) / maxNonSpend) * 100);
    const arrow = i < stages.length - 1 ? '<span class="arrow">→</span>' : '';
    return `<div class="funnel-stage">
        <div class="bar"><div class="bar-fill" style="width:${pctWidth}%"></div></div>
        <div class="value serif">${s.display}</div>
        <div class="label">${s.label}</div>
        ${arrow}
      </div>`;
  }).join('');
  document.getElementById('funnel-row').innerHTML = funnelHtml;

  // ---- shared platform ordering across all three sections ----
  const platforms = [...d.platforms].sort((a, b) => (b.offers||0) - (a.offers||0) || (b.cost||0) - (a.cost||0));

  // ---- Section: Recruiter activity & usage ----
  document.getElementById('usage-body').innerHTML = platforms.map(p => {
    const util = p.active_pct;
    const utilFlag = (util !== null && util !== undefined && util < 0.3);
    return `<div class="row-body">
        <div><div class="platform">${p.platform}</div></div>
        <div class="figure ${utilFlag ? 'flag' : ''}" data-label="Active / assigned">${p.active_recruiters ?? '–'} / ${p.recruiters_assigned ?? '–'}
          <div class="sub ${utilFlag ? 'flag' : ''}">${pct(util)} active</div></div>
        <div class="figure" data-label="Searches">${num(p.total_searches)}</div>
        <div class="figure" data-label="Profiles viewed">${num(p.total_profiles_viewed)}</div>
        <div class="figure" data-label="Outreach sent">${num(p.total_outreach_sent)}</div>
        <div class="figure" data-label="Responses">${num(p.total_responses_received)}</div>
      </div>`;
  }).join('');

  // ---- Section: Views & postings utilized vs. available ----
  document.getElementById('capacity-body').innerHTML = platforms.map(p => {
    const viewsCell = (p.views_cap !== null && p.views_cap !== undefined)
      ? `<div class="track"><div class="fill views" style="width:${Math.min(100, (p.views_pct||0)*100).toFixed(0)}%;height:100%"></div></div>
         <div class="num">${num(p.views_used)} / ${num(Math.round(p.views_cap))} (${pct(p.views_pct)})</div>`
      : `<div class="num na">${p.views_used !== null && p.views_used !== undefined ? num(p.views_used) + ' used — no fixed cap' : 'not tracked for this platform'}</div>`;
    const rawPostCap = p.postings_cap_raw;
    const postingsIsNumeric = typeof rawPostCap === 'number';
    const postingsCell = postingsIsNumeric
      ? `<div class="track"><div class="fill postings" style="width:${Math.min(100, (p.postings_pct||0)*100).toFixed(0)}%;height:100%"></div></div>
         <div class="num">${num(p.postings_used)} / ${num(rawPostCap)} (${pct(p.postings_pct)})</div>`
      : `<div class="num na">${(rawPostCap && String(rawPostCap).trim()) ? String(rawPostCap) : '–'}${(p.postings_used !== null && p.postings_used !== undefined) ? ' — ' + num(p.postings_used) + ' active' : ''}</div>`;
    return `<div class="row-body">
        <div><div class="platform">${p.platform}</div></div>
        <div class="bar-cell" data-label="Views">${viewsCell}</div>
        <div class="bar-cell" data-label="Postings">${postingsCell}</div>
      </div>`;
  }).join('');

  // ---- Section: Cost & ROI ----
  const maxCost = Math.max(1, ...platforms.map(p => p.cost || 0));
  const maxOffers = Math.max(1, ...platforms.map(p => p.offers || 0));
  document.getElementById('roi-body').innerHTML = platforms.map(p => {
    return `<div class="row-body">
        <div><div class="platform">${p.platform}</div></div>
        <div class="bar-cell" data-label="Cost">
          <div class="track"><div class="fill cost" style="width:${((p.cost||0)/maxCost*100).toFixed(0)}%;height:100%"></div></div>
          <div class="num">${money(p.cost)}</div>
        </div>
        <div class="figure" data-label="Submissions">${num(p.submissions)}</div>
        <div class="figure" data-label="Interviews">${num(p.interviews)}</div>
        <div class="bar-cell" data-label="Offers">
          <div class="track"><div class="fill offers" style="width:${((p.offers||0)/maxOffers*100).toFixed(0)}%;height:100%"></div></div>
          <div class="num">${num(p.offers)}</div>
        </div>
        <div class="figure" data-label="Cost / offer">${money(p.cost_per_offer)}</div>
        <div class="figure" data-label="Share of offers">${pct(p.share_of_offers)}</div>
      </div>`;
  }).join('');

  // ---- Trend chart ----
  Chart.defaults.font.family = "'IBM Plex Sans', Arial, sans-serif";
  Chart.defaults.color = '#8B97AC';
  new Chart(document.getElementById('trendChart'), {
    type: 'line',
    data: {
      labels: d.trend.map(t => t.period),
      datasets: [
        { label: 'Total cost', data: d.trend.map(t => t.cost), borderColor: '#C9A15A',
          backgroundColor: 'transparent', borderWidth: 2, tension: 0.25, yAxisID: 'y', pointRadius: 3 },
        { label: 'Total offers', data: d.trend.map(t => t.offers), borderColor: '#6FA98A',
          backgroundColor: 'transparent', borderWidth: 2, tension: 0.25, yAxisID: 'y1', pointRadius: 3 },
      ],
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: { legend: { labels: { boxWidth: 12, usePointStyle: true } } },
      scales: {
        x: { grid: { color: '#263752' } },
        y: { position: 'left', grid: { color: '#263752' }, title: { display: true, text: 'Cost ($)' } },
        y1: { position: 'right', grid: { display: false }, title: { display: true, text: 'Offers' } },
      },
    },
  });
}
load();
</script>
</body>
</html>
"""


def build_web(weekly_history, recruiter_history, docs_dir):
    data = compute_dashboard(weekly_history, recruiter_history)
    os.makedirs(docs_dir, exist_ok=True)
    with open(os.path.join(docs_dir, "data.json"), "w") as f:
        json.dump(data, f, indent=2)
    with open(os.path.join(docs_dir, "index.html"), "w") as f:
        f.write(HTML_TEMPLATE)
