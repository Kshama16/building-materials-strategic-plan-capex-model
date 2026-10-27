"""Independent Python implementation of the plan, consolidation and capex model.

verify_workbook.py compares the recalculated Excel workbook against these results.
Run directly to print the plan summary for all three scenarios.
"""
import numpy as np
import pandas as pd

import assumptions as A


def growth_projects():
    rows = []
    for p in A.PROJECTS:
        spend = {y: p["spend"].get(y, 0.0) for y in A.RETURN_YEARS}
        cf = np.array([-spend[y] + (p["uplift"] * (1 - A.TAX_RATE) if y >= p["in_service"] else 0.0)
                       for y in A.RETURN_YEARS])
        npv = cf[0] + sum(c / (1 + A.WACC) ** i for i, c in enumerate(cf[1:], start=1))
        cum = np.cumsum(cf)
        rows.append({**p, "forecast_total": sum(p["spend"].values()),
                     "fy26_forecast": p["spend"].get(2026, 0.0),
                     "npv": npv, "irr": irr(cf), "payback_years": int((cum < 0).sum())})
    return pd.DataFrame(rows).set_index("id")


def irr(cf, lo=-0.99, hi=1.0):
    f = lambda r: sum(c / (1 + r) ** i for i, c in enumerate(cf))
    for _ in range(200):
        mid = (lo + hi) / 2
        lo, hi = (mid, hi) if f(mid) > 0 else (lo, mid)
    return (lo + hi) / 2


def opco_plan(scenario=A.SCENARIO):
    """Returns {opco name: DataFrame indexed by year}."""
    g = A.driver_growth()
    proj = A.PROJECTS
    out = {}
    for o in A.OPCOS:
        n = o["name"]
        rows = {A.BASE_YEAR: dict(volume=o["volume"], price=o["price"], var_cost=o["var_cost"], fixed=o["fixed"])}
        for y in A.PLAN_YEARS:
            p = rows[y - 1]
            vol_adj = A.SCEN_VOL_ADJ[scenario].get(y, 0.0) if y >= 2027 else 0.0
            price_adj = A.SCEN_PRICE_ADJ[scenario] if y >= 2027 else 0.0
            vg = o["base_growth"] + o["elasticity"] * g[o["driver"]][y] + vol_adj
            pg = A.PRICE_GROWTH[n][y] + price_adj
            rows[y] = dict(driver_growth=g[o["driver"]][y], vol_growth=vg, price_growth=pg,
                           volume=p["volume"] * (1 + vg), price=p["price"] * (1 + pg),
                           var_cost=p["var_cost"] * (1 + o["var_infl"]), fixed=p["fixed"] * (1 + o["fixed_infl"]))
        df = pd.DataFrame(rows).T
        df["revenue"] = df.volume * df.price
        df["var_costs"] = df.volume * df.var_cost
        df["growth_ebitda"] = [sum(p["uplift"] for p in proj if p["opco"] == n and p["in_service"] <= y) for y in df.index]
        df["ebitda"] = df.revenue - df.var_costs - df.fixed + df.growth_ebitda
        df["margin"] = df.ebitda / df.revenue
        df["maint_capex"] = o["maint_pct"] * df.revenue
        df["growth_capex"] = [sum(p["spend"].get(y, 0.0) for p in proj if p["opco"] == n) for y in df.index]
        df["capex"] = df.maint_capex + df.growth_capex
        df.loc[A.BASE_YEAR, ["maint_capex", "growth_capex"]] = np.nan
        df.loc[A.BASE_YEAR, "capex"] = o["capex_2025"]
        out[n] = df
    return out


def consolidation(scenario=A.SCENARIO):
    plans = opco_plan(scenario)
    names = [o["name"] for o in A.OPCOS]
    c = pd.DataFrame(index=A.YEARS, dtype=float)
    c["opco_revenue"] = sum(plans[n].revenue for n in names)
    c["ic_elim"] = -sum(plans[o["name"]].revenue * o["ic_pct"] for o in A.OPCOS)
    c["revenue"] = c.opco_revenue + c.ic_elim
    c["opco_ebitda"] = sum(plans[n].ebitda for n in names)
    c["corporate"] = [-A.CORP_COST_2025 * (1 + A.CORP_INFL) ** (y - A.BASE_YEAR) for y in A.YEARS]
    c["ebitda"] = c.opco_ebitda + c.corporate
    c["margin"] = c.ebitda / c.revenue
    c["capex"] = sum(plans[n].capex for n in names)
    c["maint_capex"] = sum(plans[n].maint_capex for n in names)
    c["growth_capex"] = sum(plans[n].growth_capex for n in names)
    da, ppe = [A.DA_2025], [A.PPE_2025]
    for y in A.PLAN_YEARS:
        d = ppe[-1] / A.ASSET_LIFE
        da.append(d)
        ppe.append(ppe[-1] + c.loc[y, "capex"] - d)
    c["da"], c["ppe"] = da, ppe
    c["ebit"] = c.ebitda - c.da
    c["tax"] = c.ebit * A.TAX_RATE
    c["nwc"] = c.revenue * A.NWC_PCT
    c["d_nwc"] = c.nwc.diff()
    c["fcf"] = c.ebitda - c.tax - c.capex - c.d_nwc
    c["fcf_conversion"] = c.fcf / c.ebitda
    c.loc[A.BASE_YEAR, ["tax", "d_nwc", "fcf", "fcf_conversion"]] = np.nan
    return c, plans


def ebitda_bridge(scenario=A.SCENARIO, start=2026, end=2031):
    """Splits the EBITDA change into volume, price, cost inflation, growth projects, corporate."""
    c, plans = consolidation(scenario)
    b = dict(volume=0.0, price=0.0, var_cost=0.0, fixed=0.0, growth=0.0)
    for df in plans.values():
        s, e = df.loc[start], df.loc[end]
        b["volume"] += (e.volume - s.volume) * (s.price - s.var_cost)
        b["price"] += e.volume * (e.price - s.price)
        b["var_cost"] -= e.volume * (e.var_cost - s.var_cost)
        b["fixed"] -= e.fixed - s.fixed
        b["growth"] += e.growth_ebitda - s.growth_ebitda
    b["corporate"] = c.loc[end, "corporate"] - c.loc[start, "corporate"]
    b["start"], b["end"] = c.loc[start, "ebitda"], c.loc[end, "ebitda"]
    return b


def capex_2026():
    plans = opco_plan()
    proj = growth_projects()
    rows = []
    for o in A.OPCOS:
        n = o["name"]
        m = A.MAINT_2026[n]
        p = proj[proj.opco == n]
        rows.append(dict(opco=n, maint_budget=m["budget"], maint_ytd=m["ytd"], maint_forecast=plans[n].loc[2026, "maint_capex"],
                         growth_budget=p.fy26_budget.sum(), growth_ytd=p.fy26_ytd.sum(), growth_forecast=p.fy26_forecast.sum()))
    df = pd.DataFrame(rows).set_index("opco")
    df["budget"] = df.maint_budget + df.growth_budget
    df["ytd"] = df.maint_ytd + df.growth_ytd
    df["forecast"] = df.maint_forecast + df.growth_forecast
    df["variance"] = df.budget - df.forecast          # + = under budget
    df["var_pct"] = df.variance / df.budget
    df["flag"] = np.where((df.var_pct.abs() >= A.VAR_PCT) & (df.variance.abs() >= A.VAR_AMT),
                          np.where(df.variance > 0, "UNDER", "OVER"), "")
    return df


def summary(scenario):
    c, _ = consolidation(scenario)
    s, e = c.loc[2026], c.loc[2031]
    return dict(
        scenario=A.SCENARIOS[scenario],
        revenue_2026=s.revenue, revenue_2031=e.revenue, ebitda_2026=s.ebitda, ebitda_2031=e.ebitda,
        margin_2031=e.margin,
        revenue_cagr=(e.revenue / s.revenue) ** (1 / 5) - 1, ebitda_cagr=(e.ebitda / s.ebitda) ** (1 / 5) - 1,
        capex_5yr=c.loc[2027:2031, "capex"].sum(), growth_capex_5yr=c.loc[2027:2031, "growth_capex"].sum(),
        fcf_5yr=c.loc[2027:2031, "fcf"].sum(),
        fcf_conversion_5yr=c.loc[2027:2031, "fcf"].sum() / c.loc[2027:2031, "ebitda"].sum(),
        budget_2027_revenue=c.loc[2027, "revenue"], budget_2027_ebitda=c.loc[2027, "ebitda"],
    )


if __name__ == "__main__":
    pd.set_option("display.width", 200, "display.float_format", "{:,.1f}".format)
    c, plans = consolidation()
    print(c[["revenue", "ebitda", "margin", "capex", "da", "fcf"]].T)
    for n, df in plans.items():
        print(n, df[["volume", "price", "revenue", "ebitda", "margin", "capex"]].round(2).T.to_string(), sep="\n")
    print(pd.DataFrame([summary(s) for s in A.SCENARIOS]).set_index("scenario").T)
    print(ebitda_bridge())
    print(growth_projects()[["opco", "approved", "forecast_total", "npv", "irr", "payback_years"]])
    print(capex_2026())
