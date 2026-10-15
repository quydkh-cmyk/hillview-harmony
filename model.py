"""
Hillview Harmony: Business Analytics model.
Author: Quang, N. (2026), MQBS3010 Capstone, Group 1I, Macquarie University.

Sources (Harvard):
  Drake, J. (2026a) Hillview cost structure and break-even (FINAL) [Excel workbook]. MQBS3010 Group 1I. Unpublished.
  Drake, J. (2026b) Hillview performance and returns (FINAL) [Excel workbook]. MQBS3010 Group 1I. Unpublished.
  Drake, J. (2026c) Hillview capital allocation (FINAL) [Excel workbook]. MQBS3010 Group 1I. Unpublished.
Run order: analysis.py -> tradeoff2.py -> charts.py (all from this folder; Python 3, numpy, matplotlib).
"""
import numpy as np

# ---------------- Inputs taken directly from Finance's workbooks ----------------
YEARS = [1, 2, 3, 4, 5]
# School student-days in Finance's plan (Y1..Y5) and the mix between day / residential
PLAN_SD   = np.array([0, 9720, 10692, 11745, 12897], float)
DAY_SD    = np.array([0, 3240, 3321, 3402, 3474], float)
RES_SD    = np.array([0, 6480, 7371, 8343, 9423], float)
RES_NIGHT = np.array([0, 4320, 4914, 5562, 6282], float)
PROG_DAYS = np.array([0, 72, 81.9, 92.7, 104.7], float)       # 'cycles' run
DRY_HIRE  = np.array([0, 12, 27, 0, 0], float)                # room-days (Low reliability)
WET_HIRE  = np.array([0, 0, 0, 36, 45], float)                # room-days (Low reliability)

DAY_PRICE = 98.602898     # $/student-day  (cost-plus, Pricing tab, FINAL)
RES_PRICE = 172.142983      # $/student-night
DRY_RATE  = 700.0         # $/room-day
WET_PRICE = 11306.6667    # $/room-day
WET_COST  = 9476.0        # bought-in services per wet-hire room-day

CONTRACT_PER_CYCLE = 3955.0       # SRAC 2640 + ceramics 987 + horticulturist 328
CATERING_PER_NIGHT = 21.33
SUPPLIES_PER_SD    = 1.65414 + 5.0  # workshop + market-garden supplies
PAY_FEE            = 0.02         # % of total revenue
CORE_STAFF  = np.array([102492.613, 310902.953, 310902.953, 310902.953, 310902.953])  # Site Manager + 3 General Supervisors (FINAL)
OTHER_FIXED = np.array([164443, 229443, 234147, 256121, 257171], float)
UTILITIES   = 80000.0             # inside OTHER_FIXED (Low reliability)
MAINTENANCE = 20000.0             # inside OTHER_FIXED (Low reliability)
MARKETING   = np.array([0, 65000, 65000, 65000, 65000], float)  # inside OTHER_FIXED
DEPREC      = np.array([57880, 57880, 65025, 51099, 51099], float)

CAPEX, SETUP = 452583.19, 89500.0
OP_FUNDING_Y1, WORKING_CAP = 266935.613, 55664.972
ENDOWMENT = 1_000_000.0
AVG_INVEST = 281956.567           # 2.2 ARR Base B8 (capex/2 + working capital); 2.5 adds reinvestment/2
HURDLE, TARGET = 0.03, 0.035
# ---- Section 2.5 layer (Performance and Returns FINAL): reinvestment funded from operating surplus ----
WET_EQUIP_Y3 = 35727.0            # capex, already depreciated inside the 2.4 DEPREC line
RESTORATION  = 100000.0           # Year 3 heritage restoration, 25-yr life, NOT in 2.4 depreciation
CMP_Y3       = 57500.0            # Year 3 conservation management plan, expensed
POLICY_SHARE, POLICY_START, POLICY_LIFE = 0.5, 3, 25   # 50% of surplus reinvested in conservation from Year 3 (capex)
CAPACITY_SD = 270 * 250           # 67,500 student-days (Assumptions C26*C27)
SPC_PLAN = np.where(PROG_DAYS > 0, PLAN_SD / np.where(PROG_DAYS > 0, PROG_DAYS, 1), 0)  # students per cycle

DAY_SHARE = np.where(PLAN_SD > 0, DAY_SD / np.where(PLAN_SD > 0, PLAN_SD, 1), 0)
NIGHT_PER_RES_SD = 2.0 / 3.0      # 3 days / 2 nights camp

def run(sd=None, price=1.0, contractor=1.0, wet=1.0, utilities=1.0, staff=1.0,
        catering=1.0, maint=1.0, marketing=1.0, extra_fixed=0.0, day_price=None, policy=True, pol_share=None):
    """Return dict of yearly arrays. `sd` = school student-days by year (Y1..Y5); default = Finance's plan."""
    sd = PLAN_SD.copy() if sd is None else np.asarray(sd, float)
    day = sd * DAY_SHARE
    res = sd - day
    nights = res * NIGHT_PER_RES_SD
    cycles = np.where(sd > 0, sd / np.where(SPC_PLAN > 0, SPC_PLAN, 1), 0)
    wet_rd = WET_HIRE * wet
    dp = DAY_PRICE if day_price is None else day_price
    rev_school = (day * dp + nights * RES_PRICE) * price
    rev_dry = DRY_HIRE * DRY_RATE
    rev_wet = wet_rd * WET_PRICE
    revenue = rev_school + rev_dry + rev_wet
    fixed = (CORE_STAFF * staff
             + OTHER_FIXED
             + UTILITIES * (utilities - 1) * (YEARS_ARR > 0)
             + MAINTENANCE * (maint - 1) * (YEARS_ARR > 0)
             + MARKETING * (marketing - 1)
             + extra_fixed * (YEARS_ARR > 1))
    var = (cycles * CONTRACT_PER_CYCLE * contractor
           + nights * CATERING_PER_NIGHT * catering
           + sd * SUPPLIES_PER_SD
           + revenue * PAY_FEE
           + wet_rd * WET_COST)
    opcost = fixed + var
    surplus_pre = revenue - opcost                       # cash surplus before depreciation (2.4 row 46)
    surplus24 = surplus_pre - DEPREC                     # 2.4 'operating surplus after depreciation'
    # --- Section 2.5 layer ---
    dep_rest = np.where(YEARS_ARR >= 3, RESTORATION / 25.0, 0.0)
    cmp_ = np.where(YEARS_ARR == 3, CMP_Y3, 0.0)
    pshare = POLICY_SHARE if pol_share is None else pol_share
    pol_capex = np.zeros(5); pol_dep_cum = 0.0; surplus = np.zeros(5)
    for t in range(5):
        base = surplus24[t] - dep_rest[t] - cmp_[t] - pol_dep_cum
        if policy and (t + 1) >= POLICY_START and base > 0:
            pol_capex[t] = pshare * base
        pol_dep_cum += pol_capex[t] / POLICY_LIFE
        surplus[t] = base - pol_capex[t] / POLICY_LIFE
    reinvest = np.zeros(5); reinvest[2] += WET_EQUIP_Y3 + RESTORATION
    reinvest = reinvest + pol_capex
    return dict(sd=sd, revenue=revenue, rev_school=rev_school, rev_wet=rev_wet, fixed=fixed, var=var,
                opcost=opcost, surplus_pre=surplus_pre, surplus24=surplus24, surplus=surplus,
                cmp=cmp_, reinvest=reinvest, pol_capex=pol_capex, cycles=cycles, contrib_sd=None)
YEARS_ARR = np.array(YEARS)

def cash_flows(r, horizon=10):
    """Cash flows t0..t(horizon). t0 = capex + set-up. Y1 = set-up-year loss + working capital.
    Y2-Y5 = cash surplus (before depreciation) less the Year 3 conservation plan and reinvestment capex (2.5 layer).
    Y6+ (used only to date payback) = Year 5 cash surplus held flat, no further reinvestment.
    Working capital recovered at the end of the horizon."""
    sp = r['surplus_pre'] - r['cmp'] - r['reinvest']
    cf = [-(CAPEX + SETUP), sp[0] - WORKING_CAP] + list(sp[1:])
    while len(cf) < horizon + 1:
        cf.append(r['surplus_pre'][-1])
    cf[-1] += WORKING_CAP
    return np.array(cf[:horizon + 1])

def npv(rate, cf):
    return float(sum(c / (1 + rate) ** i for i, c in enumerate(cf)))

def irr(cf):
    lo, hi = -0.99, 2.0
    if npv(lo, cf) * npv(hi, cf) > 0:
        return float('nan')
    for _ in range(200):
        m = (lo + hi) / 2
        if npv(m, cf) > 0: lo = m
        else: hi = m
    return m

def den5(r):
    """2.5 ARR denominator at Year 5: 2.2 average investment + (cost)/2 of each surplus-funded capex item."""
    return AVG_INVEST + (WET_EQUIP_Y3 + RESTORATION + r['pol_capex'].sum()) / 2.0

def metrics(r):
    s = r['surplus']
    cf10 = cash_flows(r, 10); cf5 = cf10[:6]   # 5-year view: no terminal value, no working-capital recovery
    cum = np.cumsum(cf10)
    pay = next((i for i, c in enumerate(cum) if c >= 0), None)
    return dict(
        y5_surplus=float(s[4]), y5_revenue=float(r['revenue'][4]),
        arr_headline=float(s[4] / den5(r)),               # 2.5 definition: Y5 operating surplus / (average investment + reinvestment/2)
        arr_conservative=float(s[4] / ENDOWMENT),         # Y5 surplus / $1m
        arr_5yr_avg_endow=float(s.mean() / ENDOWMENT),    # average surplus Y1-5 incl. set-up loss / $1m
        arr_5yr_avg_inv=float(s.mean() / den5(r)),
        arr_24_headline=float(r['surplus24'][4] / AVG_INVEST),   # 2.2 ARR Base (before 2.5 reinvestment layer)
        npv5_3=npv(0.03, cf5), npv10_3=npv(0.03, cf10), npv10_35=npv(0.035, cf10),
        irr10=irr(cf10), payback_year=pay,
        cost_ratio_y5=float((r['opcost'][4] + DEPREC[4]) / r['revenue'][4]),
        den5=float(den5(r)), min_cash=float(min_cash(r)))

def min_cash(r):
    """2.5 cash test: endowment less launch spending, plus cash surplus less reinvestment; lowest year-end balance."""
    bal = ENDOWMENT - (CAPEX + SETUP); lo = 1e18
    for t in range(5):
        bal += r['surplus_pre'][t] - r['cmp'][t] - r['reinvest'][t]; lo = min(lo, bal)
    return lo

def breakeven_sd(r_year5_kwargs=None, target_surplus=0.0):
    """Student-days needed in Y5 for Y5 surplus (after dep, wet-hire included) = target. Bisection."""
    kw = r_year5_kwargs or {}
    lo, hi = 0.0, 67500.0
    for _ in range(80):
        m = (lo + hi) / 2
        sd = PLAN_SD.copy(); sd[4] = m
        if run(sd, **kw)['surplus'][4] > target_surplus: hi = m
        else: lo = m
    return hi
