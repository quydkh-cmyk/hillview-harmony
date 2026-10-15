"""
Hillview Harmony: Business Analytics model.
Author: Quang, N. (2026), MQBS3010 Capstone, Group 1I, Macquarie University.

Sources (Harvard):
  Drake, J. (2026a) Hillview cost structure and break-even (FINAL) [Excel workbook]. MQBS3010 Group 1I. Unpublished.
  Drake, J. (2026b) Hillview performance and returns (FINAL) [Excel workbook]. MQBS3010 Group 1I. Unpublished.
  Drake, J. (2026c) Hillview capital allocation (FINAL) [Excel workbook]. MQBS3010 Group 1I. Unpublished.
Run order: analysis.py -> tradeoff2.py -> charts.py (all from this folder; Python 3, numpy, matplotlib).
"""
import sys, json, itertools, time
sys.path.insert(0, '.')
import numpy as np
import model as m
R = json.load(open('results.json'))
g_low, g_mean, g_high = R['growth']['low'], R['growth']['mean'], R['growth']['high']
def path(launch, g):
    sd = np.zeros(5); sd[1] = m.PLAN_SD[1]*launch
    for t in (2, 3, 4): sd[t] = sd[t-1]*(1+g)
    return sd
DISC = 0.25                      # discount offered to local / under-resourced schools (assumption)
HURDLE_C = 30000.0
def conservation(r, maint_mult):
    """Five-year conservation spend: heritage maintenance + restoration + management plan + surplus-funded conservation capex."""
    return float(m.MAINTENANCE*maint_mult*5 + m.RESTORATION + m.CMP_Y3 + r['pol_capex'].sum())
def package(a, mm, s, sd=None, mult=None):
    mult = mult or {}
    r = m.run(sd, price=(1-a*DISC)*mult.get('price', 1.0), maint=mm*mult.get('maint', 1.0), pol_share=s,
              contractor=mult.get('contractor', 1.0), wet=mult.get('wet', 1.0), utilities=mult.get('utilities', 1.0),
              staff=mult.get('staff', 1.0), catering=mult.get('catering', 1.0))
    return r
def det(a, mm, s):
    out = {}
    for case, sd in (('Finance', m.PLAN_SD), ('Base', path(1.0, g_mean)), ('Low', path(50/72, g_low))):
        r = package(a, mm, s, sd)
        out[case] = dict(y5=float(r['surplus'][4]), min_cash=float(m.min_cash(r)), arr=float(r['surplus'][4]/m.den5(r)),
                         cons=conservation(r, mm), npv5=float(m.npv(0.03, m.cash_flows(r, 10)[:6])))
    return out
# common random numbers
def draws(n, seed=42):
    rng = np.random.default_rng(seed); tri = lambda a, c, b: rng.triangular(a, c, b, n)
    return dict(launch=tri(50/72, 1.0, 90/72), growth=tri(g_low, g_mean, g_high), price=tri(0.90, 1.0, 1.10),
                contractor=tri(0.80, 1.0, 1.20), wet=tri(0.0, 1.0, 2.0), utilities=tri(0.5, 1.0, 1.5),
                staff=tri(0.75, 1.0, 1.25), catering=tri(0.8, 1.0, 1.2), maint=tri(0.5, 1.0, 3.0))
def mc(a, mm, s, D, n):
    y5 = np.zeros(n); cash = np.zeros(n)
    for i in range(n):
        sd = path(D['launch'][i], D['growth'][i])
        r = package(a, mm, s, sd, dict(price=D['price'][i], maint=D['maint'][i], contractor=D['contractor'][i], wet=D['wet'][i],
                                       utilities=D['utilities'][i], staff=D['staff'][i], catering=D['catering'][i]))
        y5[i] = r['surplus'][4]; cash[i] = m.min_cash(r)
    return y5, cash
if __name__ == '__main__':
    t0 = time.time()
    A = [0, .1, .2, .3, .4, .5]; M = [1, 2, 3]; S = [.5, .75, 1.0]
    n = 2000; D = draws(n)
    rows = []
    for a, mm, s in itertools.product(A, M, S):
        d = det(a, mm, s); y5, cash = mc(a, mm, s, D, n)
        rows.append(dict(a=a, m=mm, s=s, det=d, med=float(np.median(y5)), p_lt=float((y5 < HURDLE_C).mean()),
                         p_cash_neg=float((cash < 0).mean()), p5=float(np.percentile(y5, 5))))
    print('time', time.time()-t0)
    json.dump(rows, open('tradeoff_grid.json', 'w'))
    for r in rows:
        print(f"a={r['a']:.1f} m={r['m']} s={r['s']:.2f} | Fin y5={r['det']['Finance']['y5']/1e3:7.1f} Base={r['det']['Base']['y5']/1e3:7.1f} "
              f"Low cash={r['det']['Low']['min_cash']/1e3:7.1f} Fin cash={r['det']['Finance']['min_cash']/1e3:6.1f} cons={r['det']['Finance']['cons']/1e3:6.1f} | med={r['med']/1e3:6.1f} P<30k={r['p_lt']:.3f} Pcash<0={r['p_cash_neg']:.3f}")
