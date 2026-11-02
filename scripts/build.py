"""Build US business counts by industry, state and occupation.

Inputs (public data, no API key):
  US Census Bureau, County Business Patterns 2023: cbp23us.txt, cbp23st.txt, cbp23co.txt
    https://www2.census.gov/programs-surveys/cbp/datasets/2023/
    plus georef22.txt (county names) from
    https://www2.census.gov/programs-surveys/cbp/technical-documentation/reference/state-county-geography-reference/
  US Bureau of Labor Statistics, OEWS May 2024 (national, nat3d/nat4d/natsector, state files)
    https://www.bls.gov/oes/tables.htm
Usage: python build.py <cbp_dir> <oews_dir> <out.json>    (requires openpyxl)
Values that cannot be read stay null; nothing is imputed.
"""
import csv, json, sys, os
import openpyxl

CBP, OEWS, OUT = sys.argv[1], sys.argv[2], sys.argv[3]

STATES = {"01":"Alabama","02":"Alaska","04":"Arizona","05":"Arkansas","06":"California","08":"Colorado","09":"Connecticut","10":"Delaware","11":"District of Columbia","12":"Florida","13":"Georgia","15":"Hawaii","16":"Idaho","17":"Illinois","18":"Indiana","19":"Iowa","20":"Kansas","21":"Kentucky","22":"Louisiana","23":"Maine","24":"Maryland","25":"Massachusetts","26":"Michigan","27":"Minnesota","28":"Mississippi","29":"Missouri","30":"Montana","31":"Nebraska","32":"Nevada","33":"New Hampshire","34":"New Jersey","35":"New Mexico","36":"New York","37":"North Carolina","38":"North Dakota","39":"Ohio","40":"Oklahoma","41":"Oregon","42":"Pennsylvania","44":"Rhode Island","45":"South Carolina","46":"South Dakota","47":"Tennessee","48":"Texas","49":"Utah","50":"Vermont","51":"Virginia","53":"Washington","54":"West Virginia","55":"Wisconsin","56":"Wyoming"}

# slug -> (CBP NAICS code, Census label, OEWS industry code)
INDUSTRIES = {
 "software-companies": ("511210", "Software publishers", "513200"),
 "it-services-companies": ("5415//", "Computer systems design and related services", "541500"),
 "marketing-agencies": ("5418//", "Advertising, public relations and related services", "541800"),
 "law-firms": ("541110", "Offices of lawyers", "541100"),
 "accounting-firms": ("5412//", "Accounting, tax preparation, bookkeeping and payroll services", "541200"),
 "dental-practices": ("621210", "Offices of dentists", "621200"),
 "real-estate-agencies": ("531210", "Offices of real estate agents and brokers", None),
 "insurance-agencies": ("524210", "Insurance agencies and brokerages", "524200"),
 "construction-companies": ("236///", "Construction of buildings", None),
 "hvac-plumbing-contractors": ("238220", "Plumbing, heating and air-conditioning contractors", "238200"),
 "restaurants": ("7225//", "Restaurants and other eating places", "722500"),
 "manufacturing-companies": ("31----", "Manufacturing (all subsectors)", None),
 "management-consulting-firms": ("541611", "Administrative and general management consulting services", "541600"),
 "staffing-agencies": ("5613//", "Employment services", "561300"),
 "architecture-firms": ("541310", "Architectural services", "541300"),
 "trucking-companies": ("484///", "Truck transportation", "484000"),
 "hotels": ("721110", "Hotels (except casino hotels) and motels", "721100"),
 "medical-practices": ("6211//", "Offices of physicians", "621100"),
 "property-management-companies": ("53131/", "Real estate property managers", None),
 "fitness-centers": ("713940", "Fitness and recreational sports centers", "713900"),
}
# OEWS 3-digit (nat3d) codes for industries missing from the 4-digit file
OEWS3 = {"real-estate-agencies": "531000", "construction-companies": "236000", "property-management-companies": "531000"}
OEWS_SECTOR = {"manufacturing-companies": "31-33"}

def short_ind(t):
    t = str(t or "")
    if t.startswith("Federal, State, and Local Government"):
        return "Government (federal, state and local)"
    return t

TITLES = {
 "ceo": "11-1011", "operations-manager": "11-1021", "marketing-manager": "11-2021", "sales-manager": "11-2022",
 "hr-manager": "11-3121", "it-manager": "11-3021", "finance-manager": "11-3031", "purchasing-manager": "11-3061",
 "construction-manager": "11-9021", "medical-practice-manager": "11-9111", "property-manager": "11-9141",
 "recruiter": "13-1071", "software-engineer": "15-1252", "accountant": "13-2011", "financial-advisor": "13-2052",
}

SIZE_KEYS = [("n<5","1-4"),("n5_9","5-9"),("n10_19","10-19"),("n20_49","20-49"),("n50_99","50-99"),("n100_249","100-249"),("n250_499","250-499"),("n500_999","500-999"),("n1000","1,000+")]

def num(v):
    try:
        return int(v)
    except Exception:
        return None

def cbp_rows(path, keycol):
    out = {}
    with open(path, encoding="latin-1") as f:
        for r in csv.DictReader(f):
            if r["lfo"] != "-":
                continue
            out[(r[keycol], r["naics"])] = r
    return out

us = cbp_rows(os.path.join(CBP, "cbp23us.txt"), "uscode")
st = cbp_rows(os.path.join(CBP, "cbp23st.txt"), "fipstate")

def cbp_block(r):
    est = num(r["est"]); emp = num(r["emp"]); ap = num(r["ap"])
    # Census withholds some size classes ("N"); those cells stay None.
    sizes = [{"band": lab, "n": num(r[k])} for k, lab in SIZE_KEYS]
    if sum(1 for x in sizes if x["n"] is None) > 3:
        sizes = None
    # emp_nf / ap_nf noise flags: G/H = low noise (used); J = high noise and D/S = withheld (not used)
    emp_ok = r.get("emp_nf") in ("G", "H") and emp
    ap_ok = r.get("ap_nf") in ("G", "H") and ap
    return {
        "establishments": est,
        "employees": emp if emp_ok else None,
        "annual_payroll_usd": ap * 1000 if ap_ok else None,
        "sizes": sizes,
    }

data = {"meta": {
    "cbp": {"name": "County Business Patterns 2023", "publisher": "U.S. Census Bureau", "url": "https://www.census.gov/programs-surveys/cbp.html", "files": "https://www2.census.gov/programs-surveys/cbp/datasets/2023/", "released": "2025"},
    "oews": {"name": "Occupational Employment and Wage Statistics, May 2024", "publisher": "U.S. Bureau of Labor Statistics", "url": "https://www.bls.gov/oes/tables.htm", "released": "2025"},
    "generated": "2026-11-02"}, "industries": {}, "titles": {}}

# ---- read OEWS
def xl(path):
    wb = openpyxl.load_workbook(path, read_only=True)
    ws = wb.active
    it = ws.iter_rows(values_only=True)
    h = next(it)
    for r in it:
        yield dict(zip(h, r))

def onum(v):
    try:
        return int(float(v))
    except Exception:
        return None

nat = {d["OCC_CODE"]: d for d in xl(os.path.join(OEWS, "nat/oesm24nat/national_M2024_dl.xlsx")) if d["O_GROUP"] in ("detailed",)}
ind4 = {}
for d in xl(os.path.join(OEWS, "in4/oesm24in4/nat4d_M2024_dl.xlsx")):
    if d["O_GROUP"] == "detailed" and str(d["OCC_CODE"]).startswith("11-"):
        ind4.setdefault(d["NAICS"], []).append(d)
ind3 = {}
occ_by_ind3 = {}
for d in xl(os.path.join(OEWS, "in4/oesm24in4/nat3d_M2024_dl.xlsx")):
    if d["O_GROUP"] != "detailed":
        continue
    if str(d["OCC_CODE"]).startswith("11-"):
        ind3.setdefault(d["NAICS"], []).append(d)
    occ_by_ind3.setdefault(d["OCC_CODE"], []).append(d)
sector = {}
for d in xl(os.path.join(OEWS, "in4/oesm24in4/natsector_M2024_dl.xlsx")):
    if d["O_GROUP"] == "detailed" and str(d["OCC_CODE"]).startswith("11-"):
        sector.setdefault(str(d["NAICS"]), []).append(d)
state_occ = {}
for d in xl(os.path.join(OEWS, "st/oesm24st/state_M2024_dl.xlsx")):
    if d["O_GROUP"] == "detailed" and d["OCC_CODE"] in TITLES.values():
        state_occ.setdefault(d["OCC_CODE"], []).append(d)

def mgmt_roles(slug, code4):
    rows = None
    if code4 and code4 in ind4:
        rows = ind4[code4]
    elif OEWS3.get(slug) and OEWS3[slug] in ind3:
        rows = ind3[OEWS3[slug]]
    elif OEWS_SECTOR.get(slug):
        rows = sector.get(OEWS_SECTOR[slug]) or sector.get(OEWS_SECTOR[slug].replace("-", ""))
    if not rows:
        return None, None
    rr = [(r["OCC_TITLE"], onum(r["TOT_EMP"]), onum(r["A_MEDIAN"])) for r in rows]
    rr = [x for x in rr if x[1]]
    rr.sort(key=lambda x: -x[1])
    return [{"occupation": t, "employment": e, "median_wage": w} for t, e, w in rr[:6]], rows[0]["NAICS_TITLE"]

for slug, (code, label, o4) in INDUSTRIES.items():
    r = us.get(("98", code))
    if not r:
        raise SystemExit(f"CBP satiri yok: {slug} {code}")
    nb = cbp_block(r)
    states = []
    for (fips, c), sr in st.items():
        if c != code or fips not in STATES:
            continue
        b = cbp_block(sr)
        b["state"] = STATES[fips]
        b["fips"] = fips
        states.append(b)
    states.sort(key=lambda b: -(b["establishments"] or 0))
    for i, b in enumerate(states):
        b["rank"] = i + 1
    roles, oews_label = mgmt_roles(slug, o4)
    data["industries"][slug] = {"naics_raw": code, "naics": code.strip("/-"), "census_label": label, "national": nb,
                                "states": states, "states_with_data": len(states),
                                "mgmt_roles": roles, "oews_industry": oews_label}

for slug, soc in TITLES.items():
    n = nat.get(soc)
    if not n:
        raise SystemExit(f"OEWS satiri yok: {slug} {soc}")
    inds = [(short_ind(d["NAICS_TITLE"]), onum(d["TOT_EMP"])) for d in occ_by_ind3.get(soc, []) if onum(d["TOT_EMP"])]
    inds.sort(key=lambda x: -x[1])
    sts = [(d["AREA_TITLE"], onum(d["TOT_EMP"]), d["JOBS_1000"], onum(d["A_MEDIAN"])) for d in state_occ.get(soc, []) if onum(d["TOT_EMP"]) and d["AREA_TITLE"] in STATES.values()]
    sts.sort(key=lambda x: -x[1])
    conc = sorted([s for s in sts if isinstance(s[2], (int, float))], key=lambda s: -s[2])
    data["titles"][slug] = {
        "soc": soc, "occupation": n["OCC_TITLE"], "employment": onum(n["TOT_EMP"]), "median_wage": onum(n["A_MEDIAN"]),
        "top_industries": [{"industry": a, "employment": b} for a, b in inds[:6]],
        "top_states": [{"state": a, "employment": b, "median_wage": w} for a, b, _, w in sts[:6]],
        "highest_concentration": [{"state": a, "per_1000_jobs": round(float(c), 2)} for a, _, c, _ in conc[:3]],
    }

# ---- counties: top 5 counties for the 3 states with the most establishments per industry
names = {}
with open(os.path.join(CBP, "georef22.txt"), encoding="latin-1") as f:
    for r in csv.DictReader(f):
        names[(r["fipstate"], r["fipscty"])] = r["county_name"].split(",")[0].strip()
want = {(v["naics_raw"], b["fips"]) for v in data["industries"].values() for b in v["states"][:3]}
cty = {}
with open(os.path.join(CBP, "cbp23co.txt"), encoding="latin-1") as f:
    for r in csv.DictReader(f):
        k = (r["naics"], r["fipstate"])
        if k in want and r["fipscty"] != "999" and num(r["est"]):
            cty.setdefault(k, []).append((names.get((r["fipstate"], r["fipscty"])), num(r["est"])))
for v in data["industries"].values():
    for b in v["states"][:3]:
        rows = sorted([x for x in cty.get((v["naics_raw"], b["fips"]), []) if x[0]], key=lambda x: -x[1])
        b["top_counties"] = [{"county": c, "establishments": e} for c, e in rows[:5]]
        b["counties_with_data"] = len(rows)
for v in data["industries"].values():
    v.pop("naics_raw", None)

json.dump(data, open(OUT, "w"), indent=1)
print("industries", len(data["industries"]), "roles", len(data["titles"]))
for s, v in data["industries"].items():
    print(s, v["national"]["establishments"], v["national"]["sizes"] is not None, [ (x["state"], x["establishments"]) for x in v["states"][:3]], (v["mgmt_roles"] or [{}])[0].get("occupation"))
for s, v in data["titles"].items():
    print(s, v["employment"], v["median_wage"], [x["industry"] for x in v["top_industries"][:2]], [x["state"] for x in v["top_states"][:3]])
