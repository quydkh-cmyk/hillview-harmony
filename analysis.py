"""
Hillview Harmony: Business Analytics model.
Author: Quang, N. (2026), MQBS3010 Capstone, Group 1I, Macquarie University.

Sources (Harvard):
  Drake, J. (2026a) Hillview cost structure and break-even (FINAL) [Excel workbook]. MQBS3010 Group 1I. Unpublished.
  Drake, J. (2026b) Hillview performance and returns (FINAL) [Excel workbook]. MQBS3010 Group 1I. Unpublished.
  Drake, J. (2026c) Hillview capital allocation (FINAL) [Excel workbook]. MQBS3010 Group 1I. Unpublished.
Run order: analysis.py -> tradeoff2.py -> charts.py (all from this folder; Python 3, numpy, matplotlib).
"""
import sys, json
sys.path.insert(0, '.')
import numpy as np
import model as m

OUT = {}
# ---------------- 1. Benchmark growth (from Hillview_Comparable_Charity_Benchmarking.xlsx, Growth Trend tab) -------------
GROWTH_OBS = {"Stewart House CY23-24": 3316064/3162237-1, "Crusader Union FY23/24-24/25 (excl. capital campaign)": 18249398/16813900-1,
              "YMCA Camping FY23-24": 19825707/18921663-1, "YMCA Camping FY24-25": 21166872/19825707-1}
g_low, g_high = min(GROWTH_OBS.values()), max(GROWTH_OBS.values())
g_mean, g_med = float(np.mean(list(GROWTH_OBS.values()))), float(np.median(list(GROWTH_OBS.values())))
OUT['growth_obs'] = GROWTH_OBS
OUT['growth'] = dict(low=g_low, mean=g_mean, median=g_med, high=g_high)

# Finance's implied growth in school student-days
OUT['finance_growth'] = [float(m.PLAN_SD[i]/m.PLAN_SD[i-1]-1) for i in (2, 3, 4)]

# ---------------- 2. Predictive curve + scenarios ----------------
LAUNCH = dict(Low=50/72, Base=1.0, High=90/72)   # Low 50, Base 72, High 90 = Finance maximum students
GROW = dict(Low=g_low, Base=g_mean, High=g_high)
def path(launch, g):
    sd = np.zeros(5); sd[1] = m.PLAN_SD[1]*launch
    for t in (2, 3, 4): sd[t] = sd[t-1]*(1+g)
    return sd
SCEN = {k: path(LAUNCH[k], GROW[k]) for k in LAUNCH}
SCEN['Finance plan'] = m.PLAN_SD.copy()
BE_ACC = [None, 9687, 10035, 10318, 10486]      # Finance's FINAL Break-Even tab, block B (schools only, accounting)
res = {}
for k, sd in SCEN.items():
    r = m.run(sd); mt = m.metrics(r)
    mt.update(sd=[float(x) for x in sd], occupancy=[float(x/m.CAPACITY_SD) for x in sd],
              revenue=[float(x) for x in r['revenue']], surplus=[float(x) for x in r['surplus']],
              surplus_pre=[float(x) for x in r['surplus_pre']], cycles=[float(x) for x in r['cycles']])
    res[k] = mt
OUT['scenarios'] = res
OUT['launch'] = LAUNCH; OUT['grow'] = GROW

# ---------------- 3. One-at-a-time sensitivity (Finance plan = centre) ----------------
base = m.run(); base_y5 = float(base['surplus'][4])
DEN_PLAN = m.den5(m.run())
H_HEAD, H_CONS, TARGET = 0.03*DEN_PLAN, 0.03*m.ENDOWMENT, 0.035*m.ENDOWMENT
def f_vol(x):   return m.run(m.PLAN_SD*x)
def f_price(x): return m.run(price=x)
def f_con(x):   return m.run(contractor=x)
def f_wet(x):   return m.run(wet=x)
def f_util(x):  return m.run(utilities=x)
def f_staff(x): return m.run(staff=x)
def f_maint(x): return m.run(maint=x)
def f_cat(x):   return m.run(catering=x)
def f_mkt(x):   return m.run(marketing=x)
VARS = [  # name, fn, low-case multiplier, high-case multiplier, reliability (from Finance's workbooks), note
 ("School volume (all years)", f_vol, 0.8, 1.2, "Medium", "-20% / +20%"),
 ("Price per student", f_price, 0.9, 1.1, "Medium", "-10% / +10% (cost-plus; market ceiling untested)"),
 ("Contractor day rates", f_con, 1.25, 0.75, "Medium", "+25% / -25%"),
 ("Wet-hire volume", f_wet, 0.0, 1.5, "Low", "0% / 150% of plan"),
 ("Utilities allowance", f_util, 1.5, 0.5, "Low", "+50% / -50%"),
 ("Core staff cost", f_staff, 1.2, 0.8, "Low", "+20% / -20%"),
 ("Heritage and grounds maintenance", f_maint, 3.0, 0.5, "Low", "x3 / x0.5"),
 ("Catering cost per night", f_cat, 1.2, 0.8, "Medium", "+20% / -20%"),
 ("Marketing spend", f_mkt, 1.3, 0.7, "Medium", "+30% / -30%"),
]
sens = []
def solve(fn, target, lo, hi):
    """smallest/largest multiplier where Y5 surplus crosses target; returns None if never crossed in [lo,hi]."""
    flo, fhi = fn(lo)['surplus'][4]-target, fn(hi)['surplus'][4]-target
    if flo*fhi > 0: return None
    for _ in range(80):
        mid = (lo+hi)/2; fm = fn(mid)['surplus'][4]-target
        if fm*flo > 0: lo, flo = mid, fm
        else: hi = mid
    return (lo+hi)/2
for name, fn, a, b, rel, note in VARS:
    ya, yb = float(fn(a)['surplus'][4]), float(fn(b)['surplus'][4])
    na, nb = m.metrics(fn(a))['npv10_3'], m.metrics(fn(b))['npv10_3']
    rng = (0.0, 4.0) if name != "Wet-hire volume" else (0.0, 6.0)
    thr = {}
    for lab, tgt in (("zero", 0.0), ("hurdle_headline", H_HEAD), ("hurdle_conservative", H_CONS)):
        thr[lab] = solve(fn, tgt, *rng)
    sens.append(dict(name=name, low_mult=a, high_mult=b, y5_low=ya, y5_high=yb, swing=abs(ya-yb),
                     npv_low=na, npv_high=nb, reliability=rel, note=note, thresholds=thr))
sens.sort(key=lambda d: -d['swing'])
OUT['sensitivity'] = dict(base_y5=base_y5, hurdle_headline=H_HEAD, hurdle_conservative=H_CONS, target=TARGET, rows=sens)

# ---------------- 4. Monte Carlo ----------------
def mc(n=10000, seed=42, g_lo=g_low):
    rng = np.random.default_rng(seed)
    tri = lambda a, c, b: rng.triangular(a, c, b, n)
    # Ranges follow the reliability ratings in Finance's workbooks (High/Medium/Low) and are centred on his plan,
    # except launch volume (team's own Low/Base/High = 50/72/90), growth (benchmark) and heritage maintenance (one-sided risk).
    D = dict(launch=tri(50/72, 1.0, 90/72), growth=tri(g_lo, g_mean, g_high),
             price=tri(0.90, 1.0, 1.10), contractor=tri(0.80, 1.0, 1.20), wet=tri(0.0, 1.0, 2.0),
             utilities=tri(0.5, 1.0, 1.5), staff=tri(0.75, 1.0, 1.25), catering=tri(0.8, 1.0, 1.2),
             maint=tri(0.5, 1.0, 3.0))
    y5 = np.zeros(n); npv10 = np.zeros(n); avg5 = np.zeros(n); arr = np.zeros(n)
    for i in range(n):
        sd = path(D['launch'][i], D['growth'][i])
        r = m.run(sd, price=D['price'][i], contractor=D['contractor'][i], wet=D['wet'][i], utilities=D['utilities'][i],
                  staff=D['staff'][i], catering=D['catering'][i], maint=D['maint'][i])
        y5[i] = r['surplus'][4]; avg5[i] = r['surplus'].mean()
        npv10[i] = m.npv(0.03, m.cash_flows(r, 10))
    return D, y5, npv10, avg5
def rank(a): return np.argsort(np.argsort(a)).astype(float)
def spearman(a, b): return float(np.corrcoef(rank(a), rank(b))[0, 1])
OUT['mc'] = {}
for label, glo in (("A_benchmark_growth", g_low), ("B_stress_growth", 0.0)):
    D, y5, npv10, avg5 = mc(g_lo=glo)
    pct = lambda a: {p: float(np.percentile(a, p)) for p in (5, 25, 50, 75, 95)}
    OUT['mc'][label] = dict(
        growth_low=glo, n=10000, seed=42,
        y5=pct(y5), npv10=pct(npv10), avg5=pct(avg5),
        p_y5_lt_zero=float((y5 < 0).mean()), p_y5_lt_hurdle_headline=float((y5 < H_HEAD).mean()),
        p_y5_lt_hurdle_cons=float((y5 < H_CONS).mean()), p_y5_lt_target=float((y5 < TARGET).mean()),
        p_npv10_lt_zero=float((npv10 < 0).mean()),
        p_avg5_lt_3pct_endow=float((avg5/m.ENDOWMENT < 0.03).mean()),
        p_avg5_lt_3pct_inv=float((avg5/m.AVG_INVEST < 0.03).mean()),
        finance_plan_percentile_y5=float((y5 < base_y5).mean()*100),
        spearman_y5={k: spearman(v, y5) for k, v in D.items()},
        spearman_npv={k: spearman(v, npv10) for k, v in D.items()},
        mean_y5=float(y5.mean()))
    np.save(f'mc_{label}_y5.npy', y5); np.save(f'mc_{label}_npv10.npy', npv10)

# ---------------- 5. Model-risk stress: cost lines that look light in Finance's workbook ----------------
# Illustrative, NOT Finance's numbers. Each adjustment is anchored to a stated source or to Finance's own rates.
SUP_COST = 70611.818                         # Finance's own General Supervisor cost incl. on-costs (Staffing tab, M6)
base_staff_y5 = float(m.CORE_STAFF[4])
def stress(cat=1.0, fte=0.0, ins_add=0.0, day_price=None):
    r = m.run(catering=cat, staff=(base_staff_y5 + fte*SUP_COST)/base_staff_y5, extra_fixed=ins_add, day_price=day_price)
    return float(r['surplus'][4])
MR = [
 ("Finance plan (no change)", dict(), "Drake (2026a, 2026b)"),
 ("Catering +50% ($32 per night)", dict(cat=1.5), "NSW Sport catered vs uncatered day price implies about $22.50 per student-day; compare with Finance's catering line"),
 ("One extra kitchen/cleaning FTE", dict(fte=1), "Finance's FINAL core staff list is a Site Manager and three General Supervisors, with no cook, cleaner or night staff; costed at Finance's own supervisor rate"),
 ("Two extra kitchen/cleaning FTE", dict(fte=2), "As above"),
 ("Day price at $83", dict(day_price=83.0), "NSW Sport catered day price; Finance's FINAL cost-plus day price is $98.60"),
 ("Insurance $30,000 (from $10,000)", dict(ins_add=20000.0), "Analyst assumption: public liability only is costed; overnight child care and heritage property cover not itemised"),
 ("Mild combination: catering +50%, one FTE, insurance $30k", dict(cat=1.5, fte=1, ins_add=20000.0), "Combination of the above"),
 ("Heavier combination: + second FTE and day price $83", dict(cat=1.5, fte=2, ins_add=20000.0, day_price=83.0), "Combination of the above"),
]
OUT['model_risk'] = dict(
    hurdle_conservative=H_CONS, hurdle_headline=H_HEAD, supervisor_cost=SUP_COST,
    rows=[dict(name=n, y5_surplus=stress(**kw), delta=stress(**kw)-base_y5, basis=b) for n, kw, b in MR])
json.dump(OUT, open('results.json', 'w'), indent=1, default=float)
print('done')
