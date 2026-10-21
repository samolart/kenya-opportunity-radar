"""Parse real CBK tables from the text extracted from the uploaded PDFs (raw/*.txt) into tidy CSVs."""
import re, os, pandas as pd
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
R = lambda f: open(os.path.join(ROOT, "raw", f), encoding="utf-8", errors="ignore").read()
D = os.path.join(ROOT, "data"); os.makedirs(D, exist_ok=True)
num = lambda s: float(s.replace(",", ""))

# ---- Table 18, CBK Bank Supervision Annual Report 2025: gross loans and NPLs by sector, Dec 2025
t = R("annual_report_2025.txt")
blk = t.split("Table 18: Sectoral Distribution")[1].split("Source: CBK")[0]
blk = blk.replace("Transport and \nCommunication", "Transport and Communication").replace("Tourism, Restaurant and \nHotels", "Tourism, Restaurant and Hotels")
rows = []
for line in blk.splitlines():
    m = re.match(r"^([A-Za-z ,&]+?)\s+([\d,]+)\s+([\d.]+)\s+([\d,.]+)\s+([\d.]+)\s+([\d,.]+)\s+([\d.]+)$", line.strip())
    if m:
        rows.append(dict(bank_sector=m[1].strip(), loan_accounts=num(m[2]), gross_loans_ksh_m=num(m[4]), gross_npl_ksh_m=num(m[6])))
c = pd.DataFrame(rows)
tot = c[c.bank_sector == "Total"].iloc[0]; c = c[c.bank_sector != "Total"].copy()
assert len(c) == 11, c
assert abs(c.gross_loans_ksh_m.sum() - tot.gross_loans_ksh_m) < 5 and abs(c.gross_npl_ksh_m.sum() - tot.gross_npl_ksh_m) < 5
c["npl_ratio"] = c.gross_npl_ksh_m / c.gross_loans_ksh_m
c["loan_share"] = c.gross_loans_ksh_m / c.gross_loans_ksh_m.sum()
c.to_csv(os.path.join(D, "cbk_sector_loans_npl_dec2025.csv"), index=False)
print(c.round(3).to_string()); print("TOTAL", tot.to_dict())

# ---- Credit Officer Survey, quarter ended June 2026 (Tables 1, 3, 5; each has March and June columns)
s = R("cos_jun2026.txt")
def table(title_start, end_marker, names):
    seg = s.split(title_start)[1].split(end_marker)[0]
    out = []
    for line in seg.splitlines():
        for n in names:
            if line.strip().startswith(n):
                v = [int(x) for x in re.findall(r"\b(\d{1,3})\b", line[line.index(n) + len(n):])]
                if len(v) == 6: out.append((n, v))
    return out
NAMES = ["Agriculture","Manufacturing","Building and Construction","Mining and Quarrying","Energy and Water","Trade","Tourism, Restaurant and Hotels","Transport and Communication","Real Estate","Financial Services","Personal and Household"]
rec = []
for tname, start, end in [("demand", "Table 1: Change in Demand for Credit", "Chart 1: Demand"), ("standards", "Table 3: Credit Standards", "2.4."), ("npl_expect", "Table 5: Non Performing Loans Trend", "Chart 6")]:
    for n, v in table(start, end, NAMES):
        for per, k in (("2026-03", 0), ("2026-06", 3)):
            up, same, down = v[k:k+3]
            rec.append(dict(measure=tname, bank_sector=n, period=per, up_pct=up, same_pct=same, down_pct=down))
q = pd.DataFrame(rec)
assert len(q) == 3 * 11 * 2, len(q)
assert (q[["up_pct","same_pct","down_pct"]].sum(axis=1).between(98, 102)).all()
q.to_csv(os.path.join(D, "cbk_credit_officer_survey_2026.csv"), index=False)
print(q[q.period == "2026-06"].pivot(index="bank_sector", columns="measure", values="up_pct"))

# ---- KNBS Economic Survey 2026, Annex 1 Table 0: selected indicators 2021-2025
e = R("economic_survey_2026.txt").split("ANNEX 1")[1].split("Index Numbers")[0]
want = {"Manufacturing output":"KSh Mn","Construction output":"KSh Mn","Cement Consumption":"'000 tonnes","Electricity consumption":"GWh","Tea-marketed production":"'000 tonnes","Fresh Horticultural Produce exports":"'000 tonnes","Milk sold centrally":"Mn litres","Rail freight":"'000 tonnes","Air passengers handled":"'000 No.","Mobile Subscriptions":"'000 No.","Total mobile money transfer":"KSh Bn","Wage employment":"'000 No.","Total domestic credit":"KSh Mn","Ordinary Revenue and Grants":"KSh Mn","Total National Government Expenditure1":"KSh Mn","Licensed doctors and dentists":"Number","Education-University enrolment":"'000 No.","GDP (Current) at Market Prices":"KSh Mn","Money Supply (M3)":"KSh Mn"}
ind = []
for line in e.splitlines():
    for k in want:
        if line.strip().startswith(k):
            v = re.findall(r"-?[\d,]+\.?\d*", line[line.index(k) + len(k):].replace(want[k], ""))
            v = [num(x) for x in v][-5:]
            if len(v) == 5: ind.append(dict(indicator=k.rstrip("1"), unit=want[k], **{str(y): x for y, x in zip(range(2021, 2026), v)}))
ind = pd.DataFrame(ind).drop_duplicates("indicator")
ind["growth_2025_pct"] = (ind["2025"] / ind["2024"] - 1) * 100
ind.to_csv(os.path.join(D, "knbs_economic_survey_indicators_2026.csv"), index=False)
print(ind[["indicator","2024","2025","growth_2025_pct"]].round(1).to_string())
