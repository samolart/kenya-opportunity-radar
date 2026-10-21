"""Extract KNBS Q1 2026 GDP tables (rotated PDF) into tidy CSVs."""
import re, subprocess, sys, os, pandas as pd
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PDF = sys.argv[1]
COLS = ["Agriculture","Mining & Quarrying","Manufacturing","Electricity & Water Supply","Construction","Wholesale & Retail Trade","Accommodation & Food Services","Transportation & Storage","Information & Communication","Financial & Insurance","Public Administration","Professional, Admin & Support Services","Real Estate","Education","Health","Other Services","FISIM","All Industries at Basic Prices","Taxes on Products","GDP at Market Prices","GDP, Seasonally Adjusted"]
def nums(page):
    t = subprocess.run(["pdftotext","-f",str(page),"-l",str(page),"-raw",PDF,"-"],capture_output=True,text=True).stdout
    toks = t.split("\n")
    i = max(j for j,x in enumerate(toks) if x.strip()=="Adjusted")
    out=[]
    for x in toks[i+1:]:
        x=x.strip()
        if re.fullmatch(r"\(?-?[\d,]+(\.\d+)?\)?",x):
            neg = x.startswith("(") or x.startswith("-")
            v=float(x.strip("()-").replace(",",""))
            out.append(-v if neg else v)
    return out
for page,name in [(11,"gdp_current"),(12,"gdp_constant"),(13,"gdp_growth")]:
    n=nums(page); print(name,len(n))
    pd.Series(n).to_csv(f"/tmp/{name}_raw.csv",index=False)

def table(page):
    n = nums(page)[:-11]
    n = [v for v in n if v not in (1.0,2.0,3.0,4.0)]
    return n
def build(page):
    n = table(page); rows=[]; k=0
    for y in range(2021,2026): rows.append((y,"FY",n[k:k+20])); k+=20
    for y in range(2021,2027):
        for q in range(1,5):
            if y==2026 and q>1: break
            w = 21
            rows.append((y,f"Q{q}",n[k:k+w])); k+=w
    assert k==len(n),(k,len(n))
    recs=[]
    for y,p,v in rows:
        for c,x in zip(COLS,v): recs.append({"year":y,"period":p,"activity":c,"value":x})
    return pd.DataFrame(recs)
os.makedirs(os.path.join(ROOT,"data"),exist_ok=True)
for page,name in [(11,"gdp_current_prices_ksh_m"),(12,"gdp_constant_2016_ksh_m")]:
    d=build(page); d.to_csv(os.path.join(ROOT,"data",name+".csv"),index=False); print(name,d.shape)
