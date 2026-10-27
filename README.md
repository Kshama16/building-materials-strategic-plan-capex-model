# Five-Year Strategic Plan, Consolidation & CAPEX Model

**Portfolio project | FP&A / corporate finance**

The annual planning cycle for a multi-business building materials group, **Granite Peak Materials Group** (fictional). Four operating companies submit volume × price plans driven by real US construction data from FRED. The model then:

- consolidates them with intercompany eliminations into the **FY2027 annual budget** and a **FY2027–FY2031 strategic plan**;
- tracks **FY2026 CAPEX** budget vs. year to date vs. forecast;
- evaluates growth projects (NPV, IRR, payback);
- builds free cash flow;
- produces a **board-style presentation**.

Everything is a live Excel formula, with a one-cell scenario switch (Base / Downside / Upside). A separate Python model recalculates every output and confirms the workbook matches it in all three scenarios.

> **Illustrative company.** Granite Peak and all company figures are fictional. The macro data (housing starts, construction spending, highway spending, mortgage rates, materials prices) is real, from [FRED](https://fred.stlouisfed.org/). Figures are in $ millions.

📥 [Excel model](model/Strategic_Plan_Consolidation_Capex_Model.xlsx) · 📊 [Board deck (PowerPoint)](deck/Board_Strategic_Plan_FY2027.pptx)

---

## Results (base case)

| $M | FY2026F (9+3) | FY2027 Budget | FY2031 Plan | CAGR FY26–31 |
|---|---|---|---|---|
| Revenue | 6,032 | 6,348 | 7,864 | 5.4% |
| EBITDA | 1,242 | 1,333 | 1,800 | 7.7% |
| EBITDA margin | 20.6% | 21.0% | 22.9% | |
| Capex | 417 | 564 | 410 | |
| Free cash flow | 603 | 509 | 1,010 | |

Five-year (FY2027–31) capex is **$2,220M**, of which $371M is growth capex. Five-year free cash flow is **$4,117M**, 51.6% of EBITDA.

**EBITDA bridge FY2026F → FY2031:**

| Start | Volume | Price | Variable cost inflation | Fixed cost inflation | Growth projects | Corporate costs | End |
|---|---|---|---|---|---|---|---|
| 1,242 | +209 | +1,371 | −928 | −160 | +90 | −23 | 1,800 |

Pricing more than covers cost inflation; volume adds a modest +$209M.

**Scenarios:**

| | Downside | Base | Upside |
|---|---|---|---|
| FY2027 budget EBITDA | 1,178 | 1,333 | 1,401 |
| FY2031 EBITDA | 1,211 | 1,800 | 2,207 |
| Five-year free cash flow | 2,948 | 4,117 | 4,806 |

**FY2026 CAPEX:** budget $438M, $303M spent through September, forecast $417M. That's $21M under budget, mainly because the Southeast quarry permitting delay moved $9M into FY2027. Asphalt & Paving is flagged at 10.3% under budget. The kiln project is $6M over its approved budget.

**Growth projects:** six of eight clear the 8.5% hurdle rate. The cement terminal expansion (IRR 7.6%) and the Texas ready-mix plants (8.3%) don't, so the deck recommends re-scoping both before funds are released.

## Workbook

| Sheet | What it does |
|---|---|
| **Inputs** | Scenario selector (`Scenario`), group assumptions as named ranges, drivers for each operating company, price growth by year, scenario adjustments |
| **Macro** | FRED annual averages 2019–2026 YTD and year-to-date change; plan growth for each volume driver (FY2026 links to the actual YTD change) |
| **Projects** | Growth CAPEX register: approved vs. forecast, FY2026 budget vs. YTD vs. forecast, status, 20-year after-tax cash flows, NPV, IRR, payback |
| **OpCo_Plan** | For each operating company: driver growth × elasticity → volume, price, revenue, variable and fixed costs, growth-project EBITDA, EBITDA, maintenance and growth capex |
| **Consolidation** | Revenue and EBITDA by operating company, intercompany eliminations, corporate costs, D&A from a PP&E roll-forward, taxes, working capital, free cash flow |
| **Capex_Report** | FY2026 budget vs. YTD vs. forecast by operating company with automatic flags (≥10% and ≥$5M); five-year capex plan, capex/D&A, capex % of revenue |
| **Board_Summary** | Headline KPIs and CAGRs for the active scenario; EBITDA bridge by driver |
| **Checks** | 10 integrity checks (consolidation ties, capex = maintenance + growth, register ties to plan, PP&E roll-forward, FCF build, bridge, capex report) rolled up to ALL CHECKS PASS |

## Method

- **Volume** = prior year × (1 + base growth + elasticity × driver growth + scenario adjustment). Each operating company follows one end market: aggregates and asphalt follow highway spending, cement follows total construction spending, and ready-mix follows housing starts.
- **Price and costs:** price grows by year-specific assumptions. Variable cost per unit and fixed costs grow with inflation.
- **Growth projects** add their EBITDA uplift from the in-service year. Project spend flows into growth capex by year.
- **Intercompany:** aggregates and cement sold to other group companies are eliminated from revenue at cost, so EBITDA is unchanged.
- **Cash flow:** D&A = opening PP&E ÷ 18-year average life. Taxes = 24% of EBIT. Net working capital is 14% of revenue. FCF = EBITDA − taxes − capex − increase in NWC.

## Verification

`tools/recalc_and_verify.sh` builds the workbook under each scenario, recalculates it headless in LibreOffice, and compares 471 values per scenario against `tools/reference_model.py`:

```
Scenario Base: 471 values compared, Checks sheet ALL PASS = True -> ALL MATCH
Scenario Downside: 471 values compared, Checks sheet ALL PASS = True -> ALL MATCH
Scenario Upside: 471 values compared, Checks sheet ALL PASS = True -> ALL MATCH
```

## Reproduce

```bash
python3 -m venv .venv && .venv/bin/pip install -r requirements.txt
(cd tools && ../.venv/bin/python build_workbook.py)   # Excel model
(cd tools && ../.venv/bin/python build_deck.py)       # board deck
tools/recalc_and_verify.sh                            # needs LibreOffice
```

The raw FRED files are in `data/`. All other inputs are in `tools/assumptions.py`.

---

Built by [Kshama Upadhyay](https://github.com/Kshama16). Related project: [FP&A Budget vs. Actual & Rolling Forecast Model](https://github.com/Kshama16/fpa-budget-variance-forecast-model) (monthly close, variance flags, 9+3 forecast).
