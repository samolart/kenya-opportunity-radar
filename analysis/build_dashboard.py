"""Single-file interactive dashboard (data embedded). Weights are adjustable and rescore live in the browser."""
import os, json, pandas as pd, numpy as np
import model as M
ROOT = M.ROOT; S = M.S.reset_index()
gdp = M.gdp
# quarterly year-on-year growth by group
qs = [(y, q) for y in range(2021, 2027) for q in ("Q1", "Q2", "Q3", "Q4") if not (y == 2026 and q != "Q1")]
ser = {}
for g, (acts, _) in M.GROUPS.items():
    lv = {(y, q): gdp[(gdp.year == y) & (gdp.period == q) & gdp.activity.isin(acts)].value.sum() for y, q in qs}
    ser[g] = [round(100 * (lv[(y, q)] / lv[(y - 1, q)] - 1), 2) for y, q in qs if y >= 2022]
labels = [f"{y} {q}" for y, q in qs if y >= 2022]
rows = []
for _, r in S.iterrows():
    rows.append(dict(g=r.group, va=r.value_added_2025_ksh_m, share=r.share_of_gdp_pct, g25=r.growth_2025_pct, cagr=r.cagr_2021_25_pct, q1=r.growth_q1_2026_pct,
                     npl=None if pd.isna(r.npl_ratio_pct) else r.npl_ratio_pct, dem=None if pd.isna(r.credit_demand_net_pct) else r.credit_demand_net_pct,
                     loans=None if pd.isna(r.gross_loans_ksh_m) else r.gross_loans_ksh_m, banks=r.bank_sectors, acts=r.activities,
                     n=[None if pd.isna(r[f"norm_{p}"]) else r[f"norm_{p}"] for p in M.PILLARS], series=ser[r.group]))
sv = pd.read_csv(os.path.join(ROOT, "data", "cbk_credit_officer_survey_2026.csv"))
sv = sv[sv.period == "2026-06"].pivot(index="bank_sector", columns="measure", values="up_pct").reset_index()
sv2 = pd.read_csv(os.path.join(ROOT, "data", "cbk_credit_officer_survey_2026.csv")); sv2 = sv2[(sv2.period == "2026-06") & (sv2.measure == "demand")][["bank_sector", "up_pct", "down_pct"]]
D = dict(rows=rows, scen=M.SCENARIOS, pillars=["Scale", "Recent growth", "Track record", "Credit demand", "Credit health"], labels=labels, neutral=M.NEUTRAL,
         survey=sv2.to_dict("records"), loans=pd.read_csv(os.path.join(ROOT, "data", "cbk_sector_loans_npl_dec2025.csv")).to_dict("records"))
html = open(os.path.join(ROOT, "analysis", "dashboard_template.html")).read().replace("/*DATA*/null", json.dumps(D))
# standalone document for GitHub / browsers (the template itself is a fragment, as the Claude artifact format requires)
i = html.index('<div class="w">'); head, body = html[:i], html[i:]
head = head.replace('<meta name="claude:viewport-fit-sides" content="true">', '')
doc = '<!doctype html>\n<html lang="en">\n<head>\n<meta charset="utf-8">\n<meta name="viewport" content="width=device-width, initial-scale=1">\n' + head + '</head>\n<body>\n' + body + '\n</body>\n</html>\n'
for out in (os.path.join(ROOT, "dashboard", "opportunity_radar_dashboard.html"), os.path.join(ROOT, "index.html")):
    open(out, "w").write(doc)
print(len(doc))
