# Methods appendix

**Extraction.** KNBS GDP tables are rotated in the PDF, so `analysis/extract_gdp.py` reads the numbers by position (annual FY2021-FY2025 then quarterly 2021 Q1 to 2026 Q1, 21 columns). Check: growth rates recomputed from the extracted levels equal KNBS Table 3 for Q1 2026 (GDP 5.3%, agriculture 4.9%, mining 9.1%, manufacturing 4.4%, construction 6.6%, ICT 5.0%, finance 6.3%, real estate 5.0%, health 4.3%, education 6.2%, accommodation 14.7%). CBK Table 18 sector rows sum exactly to the published totals (KSh 4,347,164.1m loans; KSh 696,931.0m NPLs).

**Mapping.** GDP activities and CBK bank sectors to ten groups: see `GDP` and `Credit_Map` sheets (editable) or `analysis/model.py`. Excluded: professional/admin services, other services, FISIM, taxes on products.

**Pillars.** Scale = 2025 real value added. Recent growth = Q1 2026 vs Q1 2025. Track record = 2021-25 CAGR. Credit demand = loan-weighted (banks reporting increased demand minus decreased), June 2026. Credit health = 100 minus the scaled NPL ratio (loans and NPLs summed over mapped bank sectors). Each pillar min-max scaled over the ten groups.

**Scenarios (weights %: scale, recent, track, demand, health).** Balanced 20/25/15/20/20; Growth-led 10/35/25/20/10; Risk-aware 20/15/10/15/40; Scale-led 45/15/10/15/15; Equal 20 each. Tiers on Balanced rank: 1-3 Lead, 4-7 Grow, 8-10 Selective. "Stable" means rank varies by 3 places or fewer across scenarios.

**Missing evidence.** No matching bank sector: credit pillars set to a neutral 50 (before this treatment Public Sector ranked 2nd; neutral treatment is the conservative choice).

**Verification.** The workbook recomputes everything with formulas (SUMIFS, SUMPRODUCT, INDEX-MATCH, RANK) and the Checks sheet compares it to the Python model: all pass, zero formula errors.

**Reproduce.**
```
pip install -r requirements.txt
python analysis/extract_gdp.py <path to KNBS Q1 2026 GDP PDF>
python analysis/extract_credit.py
cd analysis && python model.py && python build_workbook.py && python build_dashboard.py && python build_comms.py
```
(Recalculate the workbook in Excel or LibreOffice after building.)
