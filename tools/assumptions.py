"""All inputs for the Granite Peak Materials Group five-year plan (fictional company).

Shared by build_workbook.py (writes them into the Inputs/Macro/Projects sheets) and
reference_model.py (independent Python calculation used to verify the workbook).
Money in $ millions; volumes in millions of tons (cubic yards for ready-mix).
"""
import pandas as pd
from pathlib import Path

COMPANY = "Granite Peak Materials Group"
BASE_YEAR = 2025                      # last full actual year
YEARS = list(range(2025, 2032))       # 2025A, 2026F (9+3), 2027 budget, 2027-2031 plan
PLAN_YEARS = YEARS[1:]
FY_BUDGET = 2027
DATA = Path(__file__).resolve().parents[1] / "data"

# ---- Macro drivers (source: FRED, St. Louis Fed) -------------------------------------
MACRO_SERIES = {
    "HOUST": "Housing starts (thousand units, SAAR)",
    "TTLCONS": "Total construction spending ($M, SAAR)",
    "TLHWYCONS": "Highway & street construction spending ($M, SAAR)",
    "MORTGAGE30US": "30-year fixed mortgage rate (%)",
    "WPUSI012011": "PPI: construction materials (index 1982=100)",
}
DRIVERS = ["HOUST", "TTLCONS", "TLHWYCONS"]

# Plan growth for each volume driver. 2026 is replaced by the actual year-to-date change
# from FRED (see macro_ytd_growth); 2027-2031 are planning assumptions.
MACRO_PLAN = {
    "HOUST":     {2027: 0.010, 2028: 0.030, 2029: 0.030, 2030: 0.020, 2031: 0.020},
    "TTLCONS":   {2027: 0.000, 2028: 0.020, 2029: 0.030, 2030: 0.030, 2031: 0.030},
    "TLHWYCONS": {2027: 0.030, 2028: 0.030, 2029: 0.020, 2030: 0.020, 2031: 0.020},
}

# ---- Operating companies ----------------------------------------------------------------
# driver: macro series that moves volume; elasticity: volume % change per 1% driver change
OPCOS = [
    dict(name="Aggregates", unit="tons", driver="TLHWYCONS", elasticity=0.6, base_growth=0.005,
         volume=95.0, price=16.50, var_cost=8.20, fixed=260.0, var_infl=0.042, fixed_infl=0.035,
         maint_pct=0.060, capex_2025=118.0, ic_pct=0.10),
    dict(name="Cement", unit="tons", driver="TTLCONS", elasticity=0.8, base_growth=0.000,
         volume=9.5, price=155.0, var_cost=88.0, fixed=180.0, var_infl=0.036, fixed_infl=0.035,
         maint_pct=0.070, capex_2025=121.0, ic_pct=0.18),
    dict(name="Ready-Mix Concrete", unit="cubic yards", driver="HOUST", elasticity=0.7, base_growth=0.000,
         volume=14.0, price=165.0, var_cost=128.0, fixed=290.0, var_infl=0.036, fixed_infl=0.035,
         maint_pct=0.030, capex_2025=64.0, ic_pct=0.0),
    dict(name="Asphalt & Paving", unit="tons", driver="TLHWYCONS", elasticity=0.9, base_growth=0.000,
         volume=11.0, price=82.0, var_cost=63.0, fixed=95.0, var_infl=0.036, fixed_infl=0.035,
         maint_pct=0.040, capex_2025=33.0, ic_pct=0.0),
]
# Price growth by year (2026 = latest forecast; 2027+ = plan)
PRICE_GROWTH = {
    "Aggregates":         {2026: 0.060, 2027: 0.055, 2028: 0.050, 2029: 0.045, 2030: 0.040, 2031: 0.040},
    "Cement":             {2026: 0.020, 2027: 0.030, 2028: 0.035, 2029: 0.035, 2030: 0.030, 2031: 0.030},
    "Ready-Mix Concrete": {2026: 0.030, 2027: 0.035, 2028: 0.035, 2029: 0.035, 2030: 0.030, 2031: 0.030},
    "Asphalt & Paving":   {2026: 0.030, 2027: 0.035, 2028: 0.035, 2029: 0.030, 2030: 0.030, 2031: 0.030},
}

# ---- Scenarios (applied to plan years 2027-2031 only) -----------------------------------
SCENARIOS = {1: "Base", 2: "Downside", 3: "Upside"}
SCEN_VOL_ADJ = {   # added to every opco's volume growth
    1: {y: 0.0 for y in range(2027, 2032)},
    2: {2027: -0.040, 2028: -0.030, 2029: -0.010, 2030: 0.0, 2031: 0.0},
    3: {2027: 0.015, 2028: 0.020, 2029: 0.010, 2030: 0.010, 2031: 0.010},
}
SCEN_PRICE_ADJ = {1: 0.0, 2: -0.010, 3: 0.005}
SCENARIO = 1   # active scenario written to the workbook

# ---- Group-level inputs -------------------------------------------------------------------
CORP_COST_2025 = 120.0
CORP_INFL = 0.035
TAX_RATE = 0.24
NWC_PCT = 0.14           # net working capital as % of consolidated revenue
PPE_2025 = 7800.0        # closing net PP&E
DA_2025 = 420.0
ASSET_LIFE = 18          # years; D&A = opening PP&E / life
WACC = 0.085
VAR_PCT, VAR_AMT = 0.10, 5.0   # capex variance flag: >=10% AND >=$5M

# FY2026 maintenance capex: original budget and actual spend through September
MAINT_2026 = {
    "Aggregates": dict(budget=104.0, ytd=71.5),
    "Cement": dict(budget=112.0, ytd=76.0),
    "Ready-Mix Concrete": dict(budget=72.0, ytd=55.2),
    "Asphalt & Paving": dict(budget=40.0, ytd=27.1),
}

# ---- Growth capex projects ----------------------------------------------------------------
# spend = latest forecast schedule by year; fy26_budget = original FY2026 budget
PROJECTS = [
    dict(id="G1", opco="Aggregates", name="Southeast greenfield quarry", approved=160.0,
         spend={2026: 31.0, 2027: 94.0, 2028: 35.0}, fy26_budget=40.0, fy26_ytd=19.8,
         in_service=2029, uplift=32.0, status="Delayed: permitting; $9M shifted to 2027"),
    dict(id="G2", opco="Aggregates", name="Rail-served distribution yard", approved=40.0,
         spend={2026: 18.0, 2027: 22.0}, fy26_budget=18.0, fy26_ytd=13.1,
         in_service=2028, uplift=9.0, status="On track"),
    dict(id="G3", opco="Cement", name="Kiln efficiency and alternative fuels", approved=75.0,
         spend={2026: 36.0, 2027: 45.0}, fy26_budget=30.0, fy26_ytd=27.4,
         in_service=2028, uplift=14.0, status="Over budget: equipment cost +$6M"),
    dict(id="G4", opco="Cement", name="Import terminal expansion", approved=55.0,
         spend={2027: 25.0, 2028: 30.0}, fy26_budget=0.0, fy26_ytd=0.0,
         in_service=2029, uplift=8.0, status="Approved; starts 2027"),
    dict(id="G5", opco="Ready-Mix Concrete", name="Plant automation (8 plants)", approved=26.0,
         spend={2026: 12.0, 2027: 14.0}, fy26_budget=12.0, fy26_ytd=8.6,
         in_service=2028, uplift=6.0, status="On track"),
    dict(id="G6", opco="Ready-Mix Concrete", name="Two new Texas plants", approved=46.0,
         spend={2027: 20.0, 2028: 26.0}, fy26_budget=0.0, fy26_ytd=0.0,
         in_service=2029, uplift=7.0, status="Approved; starts 2027"),
    dict(id="G7", opco="Asphalt & Paving", name="Recycled asphalt (RAP) upgrade", approved=22.0,
         spend={2026: 7.0, 2027: 15.0}, fy26_budget=10.0, fy26_ytd=4.2,
         in_service=2028, uplift=5.0, status="Delayed: vendor lead time; $3M shifted to 2027"),
    dict(id="G8", opco="Asphalt & Paving", name="Hot-mix plant bolt-on acquisition", approved=45.0,
         spend={2028: 45.0}, fy26_budget=0.0, fy26_ytd=0.0,
         in_service=2029, uplift=9.5, status="Pipeline; board approval 2027"),
]
RETURN_YEARS = list(range(2026, 2046))   # 20-year project cash flow horizon


def macro_annual():
    """Annual averages 2019-2026 (2026 = year to date) and the 2026 YTD change vs. 2025 YTD."""
    rows, ytd_growth, latest = {}, {}, {}
    for sid in MACRO_SERIES:
        d = pd.read_csv(DATA / f"{sid}.csv")
        d.columns = ["date", "value"]
        d["date"] = pd.to_datetime(d["date"])
        d["value"] = pd.to_numeric(d["value"], errors="coerce")
        d = d.dropna()
        d = d[d["date"].dt.year >= 2019]
        rows[sid] = d.groupby(d["date"].dt.year)["value"].mean()
        last = d["date"].max()
        ytd = d[d["date"].dt.month <= last.month]
        y = ytd.groupby(ytd["date"].dt.year)["value"].mean()
        ytd_growth[sid] = y[2026] / y[2025] - 1
        latest[sid] = (last.date().isoformat(), float(d.loc[d["date"] == last, "value"].iloc[0]))
    return pd.DataFrame(rows), ytd_growth, latest


def driver_growth():
    """Volume-driver growth by plan year: 2026 actual YTD from FRED, 2027+ plan assumptions."""
    _, ytd, _ = macro_annual()
    return {sid: {2026: round(ytd[sid], 4), **MACRO_PLAN[sid]} for sid in DRIVERS}
