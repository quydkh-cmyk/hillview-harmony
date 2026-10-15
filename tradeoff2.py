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
import tradeoff as T, model as m
n = 10000; D = T.draws(n); A = [0, .1, .2, .3, .4, .5]
def mc_cond(a, mm, s, cond):
    y5 = np.zeros(n); cash = np.zeros(n); active = (D['launch'] >= 1.0) if cond else np.ones(n, bool)
    for i in range(n):
        sd = T.path(D['launch'][i], D['growth'][i]); aa = a if active[i] else 0.0
        r = T.package(aa, mm, s, sd, dict(price=D['price'][i], maint=D['maint'][i], contractor=D['contractor'][i], wet=D['wet'][i],
                                          utilities=D['utilities'][i], staff=D['staff'][i], catering=D['catering'][i]))
        y5[i] = r['surplus'][4]; cash[i] = m.min_cash(r)
    return y5, cash
def stats(a, mm, s, cond):
    y5, c = mc_cond(a, mm, s, cond)
    return dict(a=a, m=mm, s=s, cond=cond, med=float(np.median(y5)), p5=float(np.percentile(y5, 5)), p95=float(np.percentile(y5, 95)),
                p_lt=float((y5 < 13097.72).mean()), p_lt30=float((y5 < 30000).mean()), p_cash=float((c < 0).mean()))
out = {'curves': {}}
for key, (mm, cond) in {'uncond': (1, False), 'cond': (1, True), 'uncond_m2': (2, False)}.items():
    out['curves'][key] = [stats(a, mm, .75, cond) for a in A]
out['finance'] = stats(0, 1, .5, False); out['rec'] = stats(.1, 1, .75, True)
out['m2_s5'] = stats(0, 2, .5, False); out['m3_s5'] = stats(0, 3, .5, False); out['rec_m2'] = stats(.1, 2, .75, True)
out['rec_uncond'] = stats(.1, 1, .75, False)
out['det_rec'] = T.det(.1, 1, .75); out['det_a20'] = T.det(.2, 1, .75); out['det_fin'] = T.det(0, 1, .5); out['det_s75'] = T.det(0, 1, .75); out['det_m2'] = T.det(0, 2, .5); out['det_a10'] = T.det(.1, 1, .5)
json.dump(out, open('tradeoff_v2.json', 'w'))
INK, INK2 = '#0b0b0b', '#52514e'
COL = {'uncond': '#eb6834', 'cond': '#2a78d6', 'uncond_m2': '#8a5cc2'}
LAB = {'uncond': 'Discount offered from the start', 'cond': 'Discount only once Year 2 bookings reach 9,700', 'uncond_m2': 'From the start, with maintenance doubled'}
plt.rcParams.update({'text.parse_math': False, 'font.family': 'Liberation Sans', 'font.size': 10, 'axes.edgecolor': '#c9c8c3',
                     'axes.labelcolor': INK2, 'xtick.color': INK2, 'ytick.color': INK2, 'figure.facecolor': '#fcfcfb', 'axes.facecolor': '#fcfcfb'})
fig, axs = plt.subplots(1, 2, figsize=(11, 4.8)); fig.subplots_adjust(top=0.78, bottom=0.17, left=0.08, right=0.97, wspace=0.22)
for ax in axs:
    for s_ in ('top', 'right'): ax.spines[s_].set_visible(False)
    ax.grid(axis='y', color='#e6e5e1', lw=0.6); ax.set_axisbelow(True)
    ax.xaxis.set_major_formatter(FuncFormatter(lambda v, p: f"{v*100:.0f}%")); ax.set_xlabel("Share of school student-days offered at a 25% discount")
for k in ('uncond', 'cond', 'uncond_m2'):
    xs = [r['a'] for r in out['curves'][k]]
    for ax, f in ((axs[0], 'med'), (axs[1], 'p_lt')):
        ax.plot(xs, [r[f] for r in out['curves'][k]], color=COL[k], lw=2, marker='o', ms=5, mfc='#fcfcfb', mew=1.5, label=LAB[k])
axs[0].axhline(13098, color='#8a8984', lw=1.3, ls=':'); axs[0].text(0.0, 19000, "$13,100 hurdle", ha='left', fontsize=8.5, color=INK2)
axs[0].yaxis.set_major_formatter(FuncFormatter(lambda v, p: f"-${abs(v)/1000:.0f}k" if v < 0 else f"${v/1000:.0f}k"))
axs[0].set_ylabel("Median simulated Year 5 surplus"); axs[0].set_title("What each commitment costs in surplus", loc='left', fontsize=10.5, fontweight='bold', color=INK)
axs[1].axhline(0.20, color='#8a8984', lw=1.3, ls=':'); axs[1].text(0.5, 0.205, "20% tolerance", ha='right', fontsize=8.5, color=INK2)
axs[1].yaxis.set_major_formatter(FuncFormatter(lambda v, p: f"{v*100:.0f}%")); axs[1].set_ylim(0, 0.8)
axs[1].set_ylabel("Simulated runs below the $13,100 hurdle"); axs[1].set_title("What each commitment costs in risk", loc='left', fontsize=10.5, fontweight='bold', color=INK)
rc = out['curves']['cond'][1]
axs[0].plot([0.1], [rc['med']], marker='*', ms=15, color=INK, zorder=5, ls=''); axs[1].plot([0.1], [rc['p_lt']], marker='*', ms=15, color=INK, zorder=5, ls='')
axs[1].annotate("Recommended: 10% access,\nonly once the trigger is met", (0.1, rc['p_lt']), xytext=(0.17, 0.06), fontsize=8.5, color=INK, arrowprops=dict(arrowstyle='-', color=INK2, lw=0.8))
axs[0].legend(frameon=False, fontsize=8.3, loc='lower left', bbox_to_anchor=(0.0, 0.0))
fig.text(0.06, 0.965, "Heritage and access commitments compete for the same headroom", fontsize=13, fontweight='bold', color=INK, va='top')
fig.text(0.06, 0.915, "Conservation share of surplus set to 75%. Each 10% of student-days at a 25% discount costs about $34,000 of Year 5 surplus in Finance's plan.", fontsize=9, color=INK2, va='top')
fig.text(0.06, 0.012, "Source: team calculations; scenario and simulation model (10,000 runs per point). Discount applied from Year 2, so costs are slightly overstated.", fontsize=7.5, color=INK2, va='bottom')
fig.savefig('figures/fig5_access_tradeoff.png', dpi=200); plt.close(fig)
for k,v in out.items():
    if k!='curves' and not k.startswith('det'): print(k, {a:(round(b,3) if isinstance(b,float) else b) for a,b in v.items()})
print({k:[(r['a'],round(r['med']/1e3),round(r['p_lt'],3)) for r in v] for k,v in out['curves'].items()})
for k in ('det_a20','det_rec','det_fin','det_s75','det_m2','det_a10'):
    d=out[k]['Finance']; print(k,{a:round(b/1e3,1) if a!='arr' else round(b,3) for a,b in d.items()})
