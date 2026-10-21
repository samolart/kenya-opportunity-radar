"""Industry Opportunity Radar: map real KNBS and CBK data onto ten industry groups and score them."""
import os, json, pandas as pd, numpy as np
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
D = lambda f: os.path.join(ROOT, "data", f)
gdp = pd.read_csv(D("gdp_constant_2016_ksh_m.csv"))
loans = pd.read_csv(D("cbk_sector_loans_npl_dec2025.csv")).set_index("bank_sector")
cos = pd.read_csv(D("cbk_credit_officer_survey_2026.csv"))

GROUPS = {  # industry group -> (GDP activities, CBK bank sectors)  [analyst mapping, documented in README]
 "Financial Services": (["Financial & Insurance"], ["Financial Services"]),
 "Energy & Resources": (["Electricity & Water Supply", "Mining & Quarrying"], ["Energy and Water", "Mining and Quarrying"]),
 "Technology, Media & Telecom": (["Information & Communication"], ["Transport and Communication"]),
 "Transport & Logistics": (["Transportation & Storage"], ["Transport and Communication"]),
 "Public Sector & Development": (["Public Administration", "Education"], []),
 "Health": (["Health"], []),
 "Consumer & Retail": (["Wholesale & Retail Trade", "Accommodation & Food Services"], ["Trade", "Tourism, Restaurant and Hotels"]),
 "Agribusiness": (["Agriculture"], ["Agriculture"]),
 "Real Estate & Construction": (["Real Estate", "Construction"], ["Real Estate", "Building and Construction"]),
 "Manufacturing": (["Manufacturing"], ["Manufacturing"]),
}
NEUTRAL = 50.0
PILLARS = ["scale", "growth_recent", "growth_track", "credit_demand", "credit_health"]
SCENARIOS = {  # weights in % per pillar
 "Balanced":   [20, 25, 15, 20, 20],
 "Growth-led": [10, 35, 25, 20, 10],
 "Risk-aware": [20, 15, 10, 15, 40],
 "Scale-led":  [45, 15, 10, 15, 15],
 "Equal":      [20, 20, 20, 20, 20],
}
def level(acts, y, p): return gdp[(gdp.year == y) & (gdp.period == p) & gdp.activity.isin(acts)].value.sum()
tot25 = level(["All Industries at Basic Prices"], 2025, "FY")
rows = []
for g, (acts, banks) in GROUPS.items():
    v25, v24, v21 = level(acts, 2025, "FY"), level(acts, 2024, "FY"), level(acts, 2021, "FY")
    q1_26, q1_25 = level(acts, 2026, "Q1"), level(acts, 2025, "Q1")
    r = dict(group=g, activities="; ".join(acts), bank_sectors="; ".join(banks) or "n/a",
             value_added_2025_ksh_m=v25, share_of_gdp_pct=100 * v25 / tot25,
             growth_2025_pct=100 * (v25 / v24 - 1), cagr_2021_25_pct=100 * ((v25 / v21) ** 0.25 - 1), growth_q1_2026_pct=100 * (q1_26 / q1_25 - 1))
    if banks:
        L = loans.loc[banks]; r["gross_loans_ksh_m"] = L.gross_loans_ksh_m.sum(); r["npl_ratio_pct"] = 100 * L.gross_npl_ksh_m.sum() / L.gross_loans_ksh_m.sum()
        r["loan_share_pct"] = 100 * L.gross_loans_ksh_m.sum() / loans.gross_loans_ksh_m.sum()
        dm = cos[(cos.measure == "demand") & (cos.period == "2026-06")].set_index("bank_sector").loc[banks]
        w = L.gross_loans_ksh_m / L.gross_loans_ksh_m.sum()
        r["credit_demand_net_pct"] = float(((dm.up_pct - dm.down_pct) * w).sum())
        r["credit_demand_up_pct"] = float((dm.up_pct * w).sum())
    rows.append(r)
S = pd.DataFrame(rows).set_index("group")
def mm(x, inv=False):
    z = (x - x.min()) / (x.max() - x.min()) * 100
    return 100 - z if inv else z
N = pd.DataFrame({"scale": mm(S.value_added_2025_ksh_m), "growth_recent": mm(S.growth_q1_2026_pct), "growth_track": mm(S.cagr_2021_25_pct),
                  "credit_demand": mm(S.credit_demand_net_pct), "credit_health": mm(S.npl_ratio_pct, inv=True)})
def score(w):
    w = np.array(w, float)  # missing credit pillars (no matching bank sector) are set to a neutral 50
    return pd.Series({g: float((N.loc[g].fillna(NEUTRAL).values * w).sum() / w.sum()) for g in N.index})
sc = pd.DataFrame({k: score(v) for k, v in SCENARIOS.items()})
rk = sc.rank(ascending=False, method="min").astype(int)
S = S.join(N.add_prefix("norm_")).join(sc.add_prefix("score_")).join(rk.add_prefix("rank_"))
S["rank_min"] = rk.min(axis=1); S["rank_max"] = rk.max(axis=1)
S["pillars_used"] = N.notna().sum(axis=1)  # 5 = full evidence; 3 = credit pillars neutral-imputed
S["tier"] = np.where(S.rank_Balanced <= 3, "Lead", np.where(S.rank_Balanced <= 7, "Grow", "Selective"))
S["robust"] = np.where((S.rank_max - S.rank_min) <= 3, "Stable", "Sensitive")
S = S.sort_values("rank_Balanced")
os.makedirs(os.path.join(ROOT, "outputs"), exist_ok=True)
S.reset_index().to_csv(os.path.join(ROOT, "outputs", "industry_scorecard.csv"), index=False)
json.dump(dict(scenarios=SCENARIOS, pillars=PILLARS, groups=GROUPS, total_va_2025=tot25), open(os.path.join(ROOT, "outputs", "model_meta.json"), "w"), indent=1)
if __name__ == "__main__":
    pd.set_option("display.width", 250)
    print(S[["share_of_gdp_pct","growth_2025_pct","cagr_2021_25_pct","growth_q1_2026_pct","npl_ratio_pct","credit_demand_net_pct","score_Balanced","rank_Balanced","rank_min","rank_max","tier","robust","pillars_used"]].round(1).to_string())
