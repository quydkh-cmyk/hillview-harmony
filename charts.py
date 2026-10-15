"""
Hillview Harmony: Business Analytics model.
Author: Quang, N. (2026), MQBS3010 Capstone, Group 1I, Macquarie University.

Sources (Harvard):
  Drake, J. (2026a) Hillview cost structure and break-even (FINAL) [Excel workbook]. MQBS3010 Group 1I. Unpublished.
  Drake, J. (2026b) Hillview performance and returns (FINAL) [Excel workbook]. MQBS3010 Group 1I. Unpublished.
  Drake, J. (2026c) Hillview capital allocation (FINAL) [Excel workbook]. MQBS3010 Group 1I. Unpublished.
Run order: analysis.py -> tradeoff2.py -> charts.py (all from this folder; Python 3, numpy, matplotlib).
"""
import json, numpy as np
import matplotlib; matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib.ticker import FuncFormatter
R = json.load(open('results.json'))
plt.rcParams.update({'text.parse_math': False, 'font.family': 'Liberation Sans', 'font.size': 10, 'axes.edgecolor': '#c9c8c3',
                     'axes.labelcolor': '#52514e', 'xtick.color': '#52514e', 'ytick.color': '#52514e',
                     'figure.facecolor': '#fcfcfb', 'axes.facecolor': '#fcfcfb', 'savefig.facecolor': '#fcfcfb'})
INK, INK2 = '#0b0b0b', '#52514e'
COL = {'Base': '#2a78d6', 'Low': '#eb6834', 'High': '#1baf7a', 'Finance plan': '#52514e'}
usd = FuncFormatter(lambda v, p: f"-${abs(v)/1000:,.0f}k" if v < 0 else f"${v/1000:,.0f}k")
def clean(ax):
    for s in ('top', 'right'): ax.spines[s].set_visible(False)
    ax.grid(axis='y', color='#e6e5e1', lw=0.6); ax.set_axisbelow(True)
def title(fig, t, sub):
    fig.text(0.06, 0.965, t, fontsize=13, fontweight='bold', color=INK, va='top')
    fig.text(0.06, 0.915, sub, fontsize=9, color=INK2, va='top')
def src(fig, s): fig.text(0.06, 0.008, s, fontsize=7.5, color=INK2, va='bottom')

# ---- Fig 1: predicted student-days vs break-even
fig, ax = plt.subplots(figsize=(8, 4.8)); fig.subplots_adjust(top=0.80, bottom=0.17, left=0.10, right=0.76)
yrs = [2, 3, 4, 5]
for k in ('High', 'Finance plan', 'Base', 'Low'):
    v = R['scenarios'][k]['sd'][1:]
    ax.plot(yrs, v, color=COL[k], lw=2, ls='--' if k == 'Finance plan' else '-', marker='o', ms=5,
            mfc='#fcfcfb', mec=COL[k], mew=1.6)
    lab = {'Finance plan': "Finance plan (~10%/yr)", 'Base': f"Base ({R['grow']['Base']*100:.1f}%/yr)",
           'Low': f"Low ({R['grow']['Low']*100:.1f}%/yr)", 'High': f"High ({R['grow']['High']*100:.1f}%/yr)"}[k]
    dy = {'Finance plan': 13, 'Base': -13}.get(k, 0)
    ax.annotate(f"{lab}: {v[-1]:,.0f}", (5, v[-1]), xytext=(8, dy), textcoords='offset points', va='center', fontsize=8.5, color=INK)
be = [9687, 10035, 10318, 10486]
ax.step(yrs, be, where='mid', color='#8a8984', lw=1.4, ls=':')
ax.annotate("Accounting break-even (schools only, Finance): 9,687 to 10,486", (2.05, 8650), fontsize=8.5, color=INK2, va='center')
ax.set_xticks(yrs); ax.set_xticklabels([f"Year {y}" for y in yrs]); ax.set_xlim(1.8, 5.2)
ax.set_ylabel("School student-days per year"); ax.yaxis.set_major_formatter(FuncFormatter(lambda v, p: f"{v:,.0f}"))
ax.set_ylim(0, 20000); clean(ax)
title(fig, "Only the Base and High paths clear break-even by Year 5",
      "Predicted school student-days, Years 2-5. Year 1 is the set-up year (no programs).")
src(fig, "Source: team calculations; growth range from three comparable charities' financial statements (4.8%-8.5%, Crusader Union excluding\ncapital-campaign donations); break-even and plan from Finance's model. Low, Base and High: 50, 72 and 90 students.")
fig.savefig('figures/fig1_predicted_demand.png', dpi=200); plt.close(fig)

# ---- Fig 2: Y5 surplus by scenario
fig, ax = plt.subplots(figsize=(8.6, 4.8)); fig.subplots_adjust(top=0.80, bottom=0.17, left=0.10, right=0.70)
order = ['Low', 'Base', 'Finance plan', 'High']
vals = [R['scenarios'][k]['y5_surplus'] for k in order]
bars = ax.bar(range(4), vals, color=[COL[k] for k in order], width=0.55)
for i, v in enumerate(vals):
    ax.annotate(f"{'-' if v<0 else ''}${abs(v)/1000:,.0f}k", (i, v), xytext=(0, 5 if v >= 0 else -16), textcoords='offset points',
                ha='center', fontsize=9.5, color=INK, fontweight='bold')
ax.axhline(0, color=INK2, lw=0.8)
hh = R['sensitivity']['hurdle_headline']
ax.axhline(hh, color='#8a8984', lw=1.1, ls=':')
ax.annotate(f"3% hurdle\n= ${hh/1000:,.1f}k", (3.45, hh), xytext=(4, 0), textcoords='offset points', fontsize=8.5, color=INK2, va='center')
ax.set_xticks(range(4)); ax.set_xticklabels(['Low', 'Base', 'Finance plan', 'High']); ax.set_xlim(-0.5, 3.45); ax.set_ylim(-190000, 480000)
ax.yaxis.set_major_formatter(usd); ax.set_ylabel("Year 5 operating surplus"); clean(ax)
title(fig, f"Year 5 surplus ranges from a ${abs(min(vals))/1000:,.0f}k loss to ${max(vals)/1000:,.0f}k across scenarios",
      "Operating surplus after depreciation, Year 5, against the 3% hurdle on average investment.")
src(fig, "Source: team calculations using Finance's cost model (matches Finance's Year 5 surplus of $203,657 within $1). Dotted line: hurdle surplus.")
fig.savefig('figures/fig2_scenario_y5_surplus.png', dpi=200); plt.close(fig)

# ---- Fig 3: tornado
S = R['sensitivity']; rows = S['rows']; b = S['base_y5']
fig, ax = plt.subplots(figsize=(8.4, 5.2)); fig.subplots_adjust(top=0.82, bottom=0.14, left=0.33, right=0.95)
rows_r = rows[::-1]
for i, r in enumerate(rows_r):
    lo, hi = r['y5_low'], r['y5_high']
    adv, fav = min(lo, hi), max(lo, hi)
    ax.barh(i, adv - b, left=b, color='#eb6834', height=0.55)
    ax.barh(i, fav - b, left=b, color='#2a78d6', height=0.55)
    ax.annotate(f"${adv/1000:,.0f}k", (adv, i), xytext=(-4, 0), textcoords='offset points', ha='right', va='center', fontsize=8.5, color=INK)
    ax.annotate(f"${fav/1000:,.0f}k", (fav, i), xytext=(4, 0), textcoords='offset points', ha='left', va='center', fontsize=8.5, color=INK)
ax.set_yticks(range(len(rows_r))); ax.set_yticklabels([f"{r['name']}\n({r['note'].split(' (')[0]})" for r in rows_r], fontsize=8.5, color=INK)
ax.axvline(b, color=INK, lw=1); ax.set_xlim(0, 400000)
ax.xaxis.set_major_formatter(usd); ax.set_xlabel("Year 5 operating surplus")
for s in ('top', 'right', 'left'): ax.spines[s].set_visible(False)
ax.tick_params(axis='y', length=0); ax.grid(axis='x', color='#e6e5e1', lw=0.6); ax.set_axisbelow(True)
from matplotlib.patches import Patch
ax.legend(handles=[Patch(color='#eb6834', label='Adverse case'), Patch(color='#2a78d6', label='Favorable case')],
          loc='lower right', frameon=False, fontsize=8.5)
title(fig, "Volume and price move Year 5 profit most; no single cost line breaks the project",
      f"One-at-a-time sensitivity around the Finance plan (Year 5 surplus ${b/1000:,.0f}k).")
src(fig, "Source: team calculations. Ranges reflect the reliability ratings in Finance's model.")
fig.savefig('figures/fig3_tornado.png', dpi=200); plt.close(fig)

# ---- Fig 4 & 5: Monte Carlo histograms
def hist(key, arr, fname, ttl, sub, xlab, marks, xlim, bins):
    y = np.load(arr)
    fig, ax = plt.subplots(figsize=(8, 4.8)); fig.subplots_adjust(top=0.80, bottom=0.17, left=0.10, right=0.95)
    ax.hist(np.clip(y, *xlim), bins=bins, range=xlim, color='#2a78d6', edgecolor='#fcfcfb', lw=0.6)
    ax.set_ylim(0, ax.get_ylim()[1]*1.32); ymax = ax.get_ylim()[1]
    for (x, lab, dy) in marks:
        ax.axvline(x, color=INK if 'Finance' in lab else '#8a8984', lw=1.3, ls='--' if 'Finance' in lab else ':')
        ax.annotate(lab, (x, ymax*dy), xytext=(4, 0), textcoords='offset points', fontsize=8.5, color=INK, va='center', bbox=dict(fc='#fcfcfb', ec='none', alpha=0.92, pad=1.5))
    ax.xaxis.set_major_formatter(usd); ax.set_xlabel(xlab); ax.set_ylabel("Simulation runs (of 10,000)"); clean(ax)
    title(fig, ttl, sub); src(fig, "Source: team Monte Carlo (10,000 runs, seed 42, triangular inputs); tails above/below axis limits are bunched into the end bars.")
    fig.savefig(fname, dpi=200); plt.close(fig); return y
A = R['mc']['A_benchmark_growth']
hist('y5', 'mc_A_benchmark_growth_y5.npy', 'figures/fig4_montecarlo_y5.png',
     f"About {A['p_y5_lt_hurdle_headline']*100:.0f}% of simulated futures fall below the 3% hurdle in Year 5",
     "Distribution of Year 5 operating surplus, benchmark-growth case.", "Year 5 operating surplus",
     [(S['hurdle_headline'], f"3% hurdle $13.1k\n({A['p_y5_lt_hurdle_headline']*100:.0f}% of runs below)", 0.93),
      (b, f"Finance plan ${b/1000:,.0f}k\n(percentile {A['finance_plan_percentile_y5']:.0f})", 0.80),
      (A['y5']['50'], f"Median ${A['y5']['50']/1000:,.0f}k", 0.68)], (-150000, 500000), 52)
print('charts ok')
