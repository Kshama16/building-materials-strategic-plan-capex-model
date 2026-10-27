"""Compares a recalculated workbook against the Python reference model.

Usage (from tools/):  ../.venv/bin/python verify_workbook.py <recalculated.xlsx> <scenario 1|2|3>
"""
import math
import sys

from openpyxl import load_workbook

import assumptions as A
import build_workbook as B
import reference_model as R

TOL = 1e-6


def main(path, scenario):
    wb = load_workbook(path, data_only=True)
    c, plans = R.consolidation(scenario)
    fails = checked = 0

    def check(label, got, want, tol=TOL):
        nonlocal fails, checked
        checked += 1
        if want is None or (isinstance(want, float) and math.isnan(want)):
            ok = got in (None, "")
        else:
            ok = isinstance(got, (int, float)) and abs(got - want) <= tol * max(1.0, abs(want))
        if not ok:
            fails += 1
            print(f"MISMATCH {label}: workbook {got!r} reference {want!r}")

    # Build the row maps without writing a file
    B.build(B.OUT.parent / "_tmp_layout.xlsx", scenario)
    (B.OUT.parent / "_tmp_layout.xlsx").unlink()

    op = wb["OpCo_Plan"]
    keys = {"vol": "volume", "price": "price", "rev": "revenue", "var": "var_costs", "fixed": "fixed",
            "grow": "growth_ebitda", "ebitda": "ebitda", "margin": "margin", "maint": "maint_capex",
            "gcap": "growth_capex", "capex": "capex"}
    for n, df in plans.items():
        for k, col in keys.items():
            for y in A.YEARS:
                check(f"{n} {col} {y}", op.cell(B.OP[n][k], B.YEAR_COL[y]).value, float(df.loc[y, col]))

    con = wb["Consolidation"]
    for k in ["rev", "ic", "ebitda", "corp", "margin", "da", "ebit", "tax", "maint", "gcap", "capex", "ppe",
              "nwc", "dnwc", "fcf", "fcf_conv"]:
        col = {"rev": "revenue", "ic": "ic_elim", "corp": "corporate", "maint": "maint_capex", "gcap": "growth_capex",
               "dnwc": "d_nwc", "fcf_conv": "fcf_conversion"}.get(k, k)
        for y in A.YEARS:
            check(f"Consolidation {col} {y}", con.cell(B.CON[k], B.YEAR_COL[y]).value, float(c.loc[y, col]))

    pj = wb["Projects"]
    proj = R.growth_projects()
    for i, pid in enumerate(proj.index):
        r = B.PRJ["first"] + i
        check(f"{pid} NPV", pj.cell(r, 20).value, proj.loc[pid, "npv"])
        check(f"{pid} IRR", pj.cell(r, 21).value, proj.loc[pid, "irr"], 1e-6)
        check(f"{pid} payback", pj.cell(r, 22).value, float(proj.loc[pid, "payback_years"]))

    cr = wb["Capex_Report"]
    cap = R.capex_2026()
    for i, n in enumerate(cap.index):
        r = B.CAP["first"] + i
        for j, col in [(8, "budget"), (9, "ytd"), (10, "forecast"), (12, "variance")]:
            check(f"Capex report {n} {col}", cr.cell(r, j).value, float(cap.loc[n, col]))
        if (cr.cell(r, 14).value or "") != cap.loc[n, "flag"]:
            fails += 1
            print(f"MISMATCH capex flag {n}: {cr.cell(r, 14).value!r} vs {cap.loc[n, 'flag']!r}")

    bs = wb["Board_Summary"]
    b = R.ebitda_bridge(scenario)
    for k in ["start", "volume", "price", "var_cost", "fixed", "growth", "corporate", "end"]:
        check(f"Bridge {k}", bs.cell(B.BRD["bridge_" + k], 2).value, float(b[k]))
    s = R.summary(scenario)
    check("Revenue CAGR", bs.cell(B.BRD["rev"], 6).value, s["revenue_cagr"])
    check("EBITDA CAGR", bs.cell(B.BRD["ebitda"], 6).value, s["ebitda_cagr"])
    check("5-year FCF", bs.cell(B.BRD["5-year free cash flow FY27-31"], 2).value, s["fcf_5yr"])

    ck = wb["Checks"]
    for r in range(4, ck.max_row + 1):
        if ck.cell(r, 2).value is False:
            fails += 1
            print(f"CHECK FAILED: {ck.cell(r, 1).value}")
    all_pass = ck.cell(ck.max_row, 2).value
    print(f"Scenario {A.SCENARIOS[scenario]}: {checked} values compared, Checks sheet ALL PASS = {all_pass} -> "
          + ("ALL MATCH" if fails == 0 and all_pass is True else f"{fails} MISMATCHES"))
    return fails + (all_pass is not True)


if __name__ == "__main__":
    sys.exit(1 if main(sys.argv[1], int(sys.argv[2])) else 0)
