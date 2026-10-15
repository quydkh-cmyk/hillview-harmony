import json
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as L
R = json.load(open('results.json'))
T2 = json.load(open('tradeoff_v2.json'))
def cite(src):
    if 'Capital Allocation' in src: return '(Drake, 2026c) ' + src.replace('Finance FINAL: ', '').replace('Finance: ', '').replace('Capital Allocation, ', '')
    if '2.5' in src: return '(Drake, 2026b) ' + src.replace('Finance FINAL 2.5: ', '')
    if src.startswith('Finance'): return '(Drake, 2026a) ' + src.replace('Finance FINAL: ', '').replace('Finance: ', '')
    return src
wb = Workbook()
F = 'Arial'
BLUE = Font(name=F, size=10, color='0000FF'); BLK = Font(name=F, size=10); GRN = Font(name=F, size=10, color='008000')
B = Font(name=F, size=10, bold=True); H = Font(name=F, size=10, bold=True, color='FFFFFF'); T = Font(name=F, size=14, bold=True)
IT = Font(name=F, size=9, italic=True, color='595959')
HF = PatternFill('solid', fgColor='1F3864'); SUB = PatternFill('solid', fgColor='D9E1F2'); YEL = PatternFill('solid', fgColor='FFF2CC')
USD = '$#,##0;($#,##0);-'; PCT = '0.0%;(0.0%);-'; NUM = '#,##0;(#,##0);-'; USD2 = '$#,##0.00'
thin = Side(style='thin', color='BFBFBF')
def put(ws, ref, v, font=BLK, fmt=None, fill=None, bold=False, wrap=False, al=None):
    c = ws[ref]; c.value = v; c.font = font
    if bold: c.font = Font(name=F, size=font.size, bold=True, color=font.color)
    if fmt: c.number_format = fmt
    if fill: c.fill = fill
    if wrap or al: c.alignment = Alignment(wrap_text=wrap, vertical='top', horizontal=al)
    return c
def hdr(ws, row, labels, col=1):
    for i, t in enumerate(labels):
        c = ws.cell(row=row, column=col+i, value=t); c.font = H; c.fill = HF; c.alignment = Alignment(horizontal='center', wrap_text=True, vertical='center')

# ======================= README =======================
ws = wb.active; ws.title = 'README'
put(ws, 'A1', 'Hillview Harmony: Business Analytics scenario model', T)
put(ws, 'A2', 'Quang, N. (2026) | MQBS3010 Capstone, Group 1I, Macquarie University | supports the Business Analytics chapter', IT)
rows = [
 ('Author', 'Quang, N. (Business Analytics), MQBS3010 Capstone Team Group 1I, Session 2, 2026.'),
 ('Purpose', 'Scenario, sensitivity and simulation analysis behind the Business Analytics chapter. It re-runs the cost engine in Finance\'s FINAL workbooks (Drake, 2026a; 2026b; 2026c) under Low, Base and High demand, then tests how the Year 5 surplus compares with the 3% hurdle.'),
 ('How it ties to Finance', 'The "Finance plan" scenario reproduces Finance\'s Year 5 operating surplus of $203,657 and headline ARR of 46.6% (Drake, 2026b, Calculations!H84 and H85) to within $1. Every figure quoted in the chapter is listed on the Chapter_Check sheet with a live link to the cell that produces it.'),
 ('Colour code', 'Blue text = input copied from a source or an analyst assumption. Black = formula. Green = link to another sheet. Yellow fill = key analyst assumption to review.'),
 ('Sheets', 'Summary (scenario results) > Inputs > Scenarios (live formulas, four blocks) > Returns (cash flows, lowest cash, payback) > Sensitivity and MonteCarlo (static output of analysis.py: Python, seed 42, 10,000 runs) > ModelRisk (live stress test) > Comparators (the three benchmark charities) > Recommendations (static output of tradeoff2.py) > Chapter_Check > References.'),
 ('Units', 'Real 2026 dollars, ex GST. Year 1 is the set-up year with no school programs. A student-day is one student attending one program day.'),
 ('Key assumptions to review', '1) Low, Base and High use overnight camp sizes of 50, 72 and 90 students (69%, 100% and 125% of Base); 90 is the most students Finance\'s plan allows. 2) Growth rates (4.8%, 6.2%, 8.5%) are the minimum, mean and maximum of four year-on-year revenue observations at three charities (Quang, 2026; Crusader Union, 2025; Stewart House, 2025; YMCA Camping Ltd, 2024, 2025); Crusader Union excludes capital-campaign donations. 3) Wet and dry hire stay at Finance\'s volumes in every scenario. 4) Cash flows after Year 5 are held flat at the Year 5 level and are used only to date payback. 5) The Year 3 restoration, conservation plan and 50% surplus reinvestment follow Finance\'s Section 2.5 layer (Drake, 2026b). 6) Hurdle: 3% of average investment ($13,100, Finance\'s basis); on the full $1 million endowment it is $30,000 (reported as a limit).'),
 ('Limitations', 'Contractor cost scales with program days. Simulated inputs are independent and their ranges are our judgment, anchored to Finance\'s reliability ratings (Drake, 2026a). Growth comes from revenue, not student-days. Capacity (250 days x 270 students) is Finance\'s.'),
 ('How to cite', 'In text: (Quang, 2026). Reference: Quang, N. (2026) Hillview Harmony: Business Analytics scenario model [Excel workbook]. MQBS3010 Group 1I, Macquarie University. Unpublished. Python code: Quang, N. (2026) Hillview Harmony: Business Analytics model [Python code]. MQBS3010 Group 1I, Macquarie University. Unpublished.'),
 ('Sources', 'All sources are cited in Harvard style on the Inputs, Comparators, ModelRisk and Chapter_Check sheets and listed in full on the References sheet. Finance workbooks: (Drake, 2026a) cost structure and break-even; (Drake, 2026b) performance and returns; (Drake, 2026c) capital allocation.'),
 ('AI use (edit before submitting)', 'Claude (Anthropic) was used to help rebuild Finance\'s cost engine in Python, run the scenarios, sensitivity and simulation, and build this workbook. The author set the questions, checked every input against Finance\'s workbooks and reconciled the outputs to the chapter. Delete or edit this line to match the unit\'s AI disclosure rules.'),
]
for i, (k, v) in enumerate(rows, start=4):
    put(ws, f'A{i}', k, B, al='left'); put(ws, f'B{i}', v, BLK, wrap=True)
    ws.row_dimensions[i].height = 18 * (1 + len(v) // 115)
ws.column_dimensions['A'].width = 30; ws.column_dimensions['B'].width = 125

ADDR = {}

# ======================= Inputs =======================
wi = wb.create_sheet('Inputs')
put(wi, 'A1', 'Inputs (copied from Finance workbooks unless marked as analyst assumption)', T)
hdr(wi, 3, ['Scalar input', 'Value', 'Unit', 'Source'])
S = {}
scal = [
 ('DAY_PRICE', 'Day program price', 98.602898, USD2, '$ per student-day', 'Finance FINAL: Pricing!C38 (cost-plus)'),
 ('RES_PRICE', 'Residential program price', 172.142983, USD2, '$ per student-night', 'Finance: Pricing!C47'),
 ('DRY_RATE', 'Dry-hire rate', 700, USD, '$ per room-day', 'Finance: Assumptions!C8'),
 ('WET_PRICE', 'Wet-hire price', 11306.6667, USD, '$ per room-day', 'Finance: Wet Hire Costs!C25'),
 ('WET_COST', 'Wet-hire bought-in services', 9476, USD, '$ per room-day', 'Finance: Wet Hire Costs!I17'),
 ('CONTRACT', 'Contractor cost per program day (all activities)', 3955, USD, '$ per program day', 'Finance: Contractors!C20'),
 ('CATER', 'Catering', 21.33, USD2, '$ per student-night', 'Finance: Cost Register!D6'),
 ('SUPPLIES', 'Workshop + market-garden supplies', 6.65414, '$#,##0.00000', '$ per school student-day', 'Finance: Cost Register!D9 + D10'),
 ('FEE', 'Payment fees', 0.02, PCT, '% of revenue', 'Finance: Cost Register!D15'),
 ('NIGHTS', 'Nights per residential student-day', '=2/3', '0.000', 'ratio', 'Finance: 3-day, 2-night camp'),
 ('AVG_INV', 'Average investment before reinvestment (2.2 ARR Base)', 281956.567, USD, '$', 'Finance FINAL: Capital Allocation, ARR Base!B8'),
 ('ENDOW', 'Endowment', 1000000, USD, '$', 'Capstone brief'),
 ('CAPEX', 'Capex (Year 0)', 452583.19, USD, '$', 'Finance: Capital Allocation, Table 2.3!B18'),
 ('SETUP', 'One-off set-up expenses (Year 0)', 89500, USD, '$', 'Finance: Capital Allocation, Table 2.3!B33'),
 ('WC', 'Day-to-day working capital', 55664.972, USD, '$', 'Finance: Capital Allocation, Other Uses!B24'),
 ('HURDLE', 'Hurdle rate', 0.03, PCT, '% per year', 'Capstone brief / Finance'),
 ('TARGET', 'Working target rate', 0.035, PCT, '% per year', 'Finance (CPI-aligned)'),
 ('CAP', 'Capacity (250 days x 270 students)', 67500, NUM, 'student-days per year', 'Finance: Assumptions!C26 x C27'),
 ('WETEQ', 'Wet-hire equipment, Year 3 (cash; depreciation already in 2.4)', 35727, USD, '$', 'Finance FINAL 2.5: Inputs!C31'),
 ('REST', 'Heritage restoration works, Year 3', 100000, USD, '$', 'Finance FINAL 2.5: Inputs!C32 (25-year life)'),
 ('RESTLIFE', 'Restoration useful life', 25, '0', 'years', 'Finance FINAL 2.5: Inputs!E32'),
 ('CMP', 'Conservation management plan, Year 3 (expensed)', 57500, USD, '$', 'Finance FINAL 2.5: Inputs!C33'),
 ('POLSHARE', 'Share of surplus reinvested in conservation', 0.5, PCT, '% of surplus', 'Finance FINAL 2.5: Inputs!B46'),
 ('POLSTART', 'Policy start year', 3, '0', 'year', 'Finance FINAL 2.5: Inputs!B47'),
 ('POLLIFE', 'Policy capex useful life', 25, '0', 'years', 'Finance FINAL 2.5: Inputs!B49'),
]
r = 4
for key, lab, val, fmt, unit, src in scal:
    put(wi, f'A{r}', lab); put(wi, f'B{r}', val, BLUE if not str(val).startswith('=') else BLK, fmt); put(wi, f'C{r}', unit, IT); put(wi, f'D{r}', cite(src), IT)
    S[key] = f"Inputs!$B${r}"; r += 1
r += 1
put(wi, f'A{r}', 'Benchmark growth observations (Quang, 2026; Crusader Union, 2025; Stewart House, 2025; YMCA Camping Ltd, 2024, 2025)', B); r += 1
hdr(wi, r, ['Observation', 'Prior-year revenue ($)', 'Latest revenue ($)', 'YoY growth']); r += 1
obs = [('Stewart House CY2023 to CY2024', 3162237, 3316064), ('Crusader Union FY23/24 to FY24/25 (excl. capital-campaign donations)', 16813900, 18249398),
       ('YMCA Camping FY2023 to FY2024', 18921663, 19825707), ('YMCA Camping FY2024 to FY2025', 19825707, 21166872)]
o0 = r
for n, a, b in obs:
    put(wi, f'A{r}', n); put(wi, f'B{r}', a, BLUE, USD); put(wi, f'C{r}', b, BLUE, USD); put(wi, f'D{r}', f'=C{r}/B{r}-1', BLK, '0.00%'); r += 1
o1 = r - 1
put(wi, f'A{r}', 'Minimum (Low case)', B); put(wi, f'D{r}', f'=MIN(D{o0}:D{o1})', BLK, '0.00%'); S['G_LOW'] = f'Inputs!$D${r}'; r += 1
put(wi, f'A{r}', 'Mean (Base case)', B); put(wi, f'D{r}', f'=AVERAGE(D{o0}:D{o1})', BLK, '0.00%'); S['G_BASE'] = f'Inputs!$D${r}'; r += 1
put(wi, f'A{r}', 'Maximum (High case)', B); put(wi, f'D{r}', f'=MAX(D{o0}:D{o1})', BLK, '0.00%'); S['G_HIGH'] = f'Inputs!$D${r}'; r += 2
put(wi, f'A{r}', 'Launch-volume multipliers (overnight camp size: Low 50 is our assumption, Base 72 is Finance\'s Year 2 size, High 90 is Finance\'s maximum)', B); r += 1
hdr(wi, r, ['Scenario', 'Overnight students', 'Multiplier vs Base', '']); r += 1
for n, st in (('Low', 50), ('Base', 72), ('High', 90)):
    put(wi, f'A{r}', n); put(wi, f'B{r}', st, BLUE, NUM, fill=YEL); put(wi, f'C{r}', f'=B{r}/72', BLK, '0.000'); S['L_' + n.upper()] = f'Inputs!$C${r}'; r += 1
r += 1
put(wi, f'A{r}', 'Finance plan by year', B); r += 1
hdr(wi, r, ['Line', 'Unit', 'Year 1', 'Year 2', 'Year 3', 'Year 4', 'Year 5', 'Source']); r += 1
Y = {}
yearly = [
 ('DAY_SD', 'Day-program student-days', 'student-days', [0, 3240, 3321, 3402, 3474], NUM, 'Finance: Volumes!C5:G5'),
 ('RES_SD', 'Residential student-days', 'student-days', [0, 6480, 7371, 8343, 9423], NUM, 'Finance: Volumes!C6:G6'),
 ('PROG', 'Program days operated (cycles)', 'days', [0, 72, 81.9, 92.7, 104.7], '#,##0.0', 'Finance: Volumes!C10:G10'),
 ('DRY', 'Dry-hire room-days', 'room-days', [0, 12, 27, 0, 0], NUM, 'Finance: Volumes!C8:G8 (Low reliability)'),
 ('WET', 'Wet-hire room-days', 'room-days', [0, 0, 0, 36, 45], NUM, 'Finance: Volumes!C9:G9 (Low reliability)'),
 ('STAFF', 'Core staff cost', '$', [102492.613, 310902.953, 310902.953, 310902.953, 310902.953], USD, 'Finance FINAL: Staffing!P25:T25 (Site Manager + 3 General Supervisors)'),
 ('OFIX', 'Other fixed costs', '$', [164443, 229443, 234147, 256121, 257171], USD, 'Finance: Calculations!C37:G37'),
 ('DEP', 'Depreciation', '$', [57880, 57880, 65025, 51099, 51099], USD, 'Finance: Assumptions!C32:C36'),
]
for key, lab, unit, vals, fmt, src in yearly:
    put(wi, f'A{r}', lab); put(wi, f'B{r}', unit, IT)
    for j, v in enumerate(vals): put(wi, f'{L(3+j)}{r}', v, BLUE, fmt)
    put(wi, f'H{r}', cite(src), IT); Y[key] = r; r += 1
for key, lab, fmt, f in (('SD', 'School student-days (Finance plan)', NUM, lambda c: f'={c}{Y["DAY_SD"]}+{c}{Y["RES_SD"]}'),
                         ('SHARE', 'Day-program share of student-days', PCT, lambda c: f'=IF({c}{Y["DAY_SD"]}+{c}{Y["RES_SD"]}=0,0,{c}{Y["DAY_SD"]}/({c}{Y["DAY_SD"]}+{c}{Y["RES_SD"]}))'),
                         ('SPC', 'Students per program day (cycle fill)', '#,##0.0', lambda c: f'=IF({c}{Y["PROG"]}=0,0,({c}{Y["DAY_SD"]}+{c}{Y["RES_SD"]})/{c}{Y["PROG"]})')):
    put(wi, f'A{r}', lab); put(wi, f'B{r}', 'calculated', IT)
    for j in range(5): put(wi, f'{L(3+j)}{r}', f(L(3+j)), BLK, fmt)
    Y[key] = r; r += 1
wi.column_dimensions['A'].width = 52; wi.column_dimensions['B'].width = 24
for c in 'CDEFG': wi.column_dimensions[c].width = 15
wi.column_dimensions['H'].width = 40
wi.freeze_panes = 'A4'
def yin(key, j): return f"Inputs!{L(3+j)}${Y[key]}"

# ======================= Scenarios =======================
sc = wb.create_sheet('Scenarios')
put(sc, 'A1', 'Scenario engine (formulas; reproduces Finance cost logic)', T)
put(sc, 'A2', 'Low / Base / High: Year 2 volume = Finance Year 2 x launch multiplier, then grows at the benchmark rate. Finance plan: volumes copied from Finance.', IT)
BL = {}  # scenario -> dict of row numbers
def block(name, start, launch=None, growth=None):
    rr = start
    put(sc, f'A{rr}', name, H, fill=HF)
    for j, y in enumerate(range(1, 6)): put(sc, f'{L(3+j)}{rr}', f'Year {y}', H, fill=HF, al='center')
    put(sc, f'B{rr}', 'Unit', H, fill=HF); rr += 1
    rows = {}
    def line(key, label, unit, fn, fmt=USD, font=BLK, bold=False):
        nonlocal rr
        put(sc, f'A{rr}', label, B if bold else BLK); put(sc, f'B{rr}', unit, IT)
        for j in range(5):
            c = L(3+j); put(sc, f'{c}{rr}', fn(j, c), font, fmt, bold=bold)
        rows[key] = rr; rr += 1
    if launch is None:
        line('sd', 'School student-days', 'student-days', lambda j, c: f'={yin("SD", j)}', NUM, GRN)
    else:
        def f(j, c):
            if j == 0: return 0
            if j == 1: return f'={yin("SD", 1)}*{launch}'
            return f'={L(2+j)}{rows["sd"] if "sd" in rows else rr}*(1+{growth})'
        line('sd', 'School student-days', 'student-days', f, NUM)
    line('occ', 'Occupancy (% of capacity)', '%', lambda j, c: f'={c}{rows["sd"]}/{S["CAP"]}', PCT)
    line('day', 'Day-program student-days', 'student-days', lambda j, c: f'={c}{rows["sd"]}*{yin("SHARE", j)}', NUM)
    line('res', 'Residential student-days', 'student-days', lambda j, c: f'={c}{rows["sd"]}-{c}{rows["day"]}', NUM)
    line('night', 'Residential student-nights', 'student-nights', lambda j, c: f'={c}{rows["res"]}*{S["NIGHTS"]}', NUM)
    line('cyc', 'Program days (cycles)', 'days', lambda j, c: f'=IF({yin("SPC", j)}=0,0,{c}{rows["sd"]}/{yin("SPC", j)})', '#,##0.0')
    line('rsch', 'School revenue', '$', lambda j, c: f'={c}{rows["day"]}*{S["DAY_PRICE"]}+{c}{rows["night"]}*{S["RES_PRICE"]}')
    line('rdry', 'Dry-hire revenue', '$', lambda j, c: f'={yin("DRY", j)}*{S["DRY_RATE"]}')
    line('rwet', 'Wet-hire revenue', '$', lambda j, c: f'={yin("WET", j)}*{S["WET_PRICE"]}')
    line('rev', 'Total revenue', '$', lambda j, c: f'={c}{rows["rsch"]}+{c}{rows["rdry"]}+{c}{rows["rwet"]}', USD, bold=True)
    line('fix', 'Fixed cash costs (core staff + other fixed)', '$', lambda j, c: f'={yin("STAFF", j)}+{yin("OFIX", j)}')
    line('con', 'Program contractors', '$', lambda j, c: f'={c}{rows["cyc"]}*{S["CONTRACT"]}')
    line('cat', 'Catering', '$', lambda j, c: f'={c}{rows["night"]}*{S["CATER"]}')
    line('sup', 'Program supplies', '$', lambda j, c: f'={c}{rows["sd"]}*{S["SUPPLIES"]}')
    line('fee', 'Payment fees', '$', lambda j, c: f'={c}{rows["rev"]}*{S["FEE"]}')
    line('wsv', 'Wet-hire event services', '$', lambda j, c: f'={yin("WET", j)}*{S["WET_COST"]}')
    line('opc', 'Total operating costs (cash)', '$', lambda j, c: f'={c}{rows["fix"]}+SUM({c}{rows["con"]}:{c}{rows["wsv"]})', USD, bold=True)
    line('pre', 'Operating surplus before depreciation', '$', lambda j, c: f'={c}{rows["rev"]}-{c}{rows["opc"]}', USD, bold=True)
    line('dep', 'Depreciation', '$', lambda j, c: f'={yin("DEP", j)}', USD, GRN)
    line('sur', 'Operating surplus after depreciation', '$', lambda j, c: f'={c}{rows["pre"]}-{c}{rows["dep"]}', USD, bold=True)
    for j in range(7):
        sc[f'{L(1+j)}{rows["sur"]}'].fill = SUB
    BL[name] = rows
    return rr + 1
# special handling: growth reference to previous column needs row of sd -> patch by building sd first row index
def block_growth(name, start, launch, growth):
    # sd row is start+1; formula references previous column same row
    nxt = block.__wrapped__(name, start, launch, growth) if hasattr(block, '__wrapped__') else None
    return nxt
# simple approach: pre-compute sd row = start+1 and build custom lambda
def build(name, start, launch=None, growth=None):
    sd_row = start + 1
    global block
    rr = start
    put(sc, f'A{rr}', name, H, fill=HF); put(sc, f'B{rr}', 'Unit', H, fill=HF)
    for j in range(5): put(sc, f'{L(3+j)}{rr}', f'Year {j+1}', H, fill=HF, al='center')
    rr += 1; rows = {}
    def line(key, label, unit, fn, fmt=USD, font=BLK, bold=False):
        nonlocal rr
        put(sc, f'A{rr}', label, B if bold else BLK); put(sc, f'B{rr}', unit, IT)
        for j in range(5):
            c = L(3+j); put(sc, f'{c}{rr}', fn(j, c), font, fmt, bold=bold)
        rows[key] = rr; rr += 1
    if launch is None:
        line('sd', 'School student-days', 'student-days', lambda j, c: f'={yin("SD", j)}', NUM, GRN)
    else:
        def f(j, c):
            if j == 0: return 0
            if j == 1: return f'={yin("SD", 1)}*{launch}'
            return f'={L(2+j)}{sd_row}*(1+{growth})'
        line('sd', 'School student-days', 'student-days', f, NUM)
    line('occ', 'Occupancy (% of capacity)', '%', lambda j, c: f'={c}{rows["sd"]}/{S["CAP"]}', PCT)
    line('day', 'Day-program student-days', 'student-days', lambda j, c: f'={c}{rows["sd"]}*{yin("SHARE", j)}', NUM)
    line('res', 'Residential student-days', 'student-days', lambda j, c: f'={c}{rows["sd"]}-{c}{rows["day"]}', NUM)
    line('night', 'Residential student-nights', 'student-nights', lambda j, c: f'={c}{rows["res"]}*{S["NIGHTS"]}', NUM)
    line('cyc', 'Program days (cycles)', 'days', lambda j, c: f'=IF({yin("SPC", j)}=0,0,{c}{rows["sd"]}/{yin("SPC", j)})', '#,##0.0')
    line('rsch', 'School revenue', '$', lambda j, c: f'={c}{rows["day"]}*{S["DAY_PRICE"]}+{c}{rows["night"]}*{S["RES_PRICE"]}')
    line('rdry', 'Dry-hire revenue', '$', lambda j, c: f'={yin("DRY", j)}*{S["DRY_RATE"]}')
    line('rwet', 'Wet-hire revenue', '$', lambda j, c: f'={yin("WET", j)}*{S["WET_PRICE"]}')
    line('rev', 'Total revenue', '$', lambda j, c: f'={c}{rows["rsch"]}+{c}{rows["rdry"]}+{c}{rows["rwet"]}', USD, bold=True)
    line('fix', 'Fixed cash costs (core staff + other fixed)', '$', lambda j, c: f'={yin("STAFF", j)}+{yin("OFIX", j)}')
    line('con', 'Program contractors', '$', lambda j, c: f'={c}{rows["cyc"]}*{S["CONTRACT"]}')
    line('cat', 'Catering', '$', lambda j, c: f'={c}{rows["night"]}*{S["CATER"]}')
    line('sup', 'Program supplies', '$', lambda j, c: f'={c}{rows["sd"]}*{S["SUPPLIES"]}')
    line('fee', 'Payment fees', '$', lambda j, c: f'={c}{rows["rev"]}*{S["FEE"]}')
    line('wsv', 'Wet-hire event services', '$', lambda j, c: f'={yin("WET", j)}*{S["WET_COST"]}')
    line('opc', 'Total operating costs (cash)', '$', lambda j, c: f'={c}{rows["fix"]}+SUM({c}{rows["con"]}:{c}{rows["wsv"]})', USD, bold=True)
    line('pre', 'Operating surplus before depreciation', '$', lambda j, c: f'={c}{rows["rev"]}-{c}{rows["opc"]}', USD, bold=True)
    line('dep', 'Depreciation', '$', lambda j, c: f'={yin("DEP", j)}', USD, GRN)
    line('sur24', 'Surplus after depreciation (2.4 basis, before reinvestment layer)', '$', lambda j, c: f'={c}{rows["pre"]}-{c}{rows["dep"]}', USD)
    line('drest', 'Restoration depreciation (Year 3 onward)', '$', lambda j, c: f'=IF({j+1}>=3,{S["REST"]}/{S["RESTLIFE"]},0)')
    line('cmp', 'Conservation management plan (Year 3, expensed)', '$', lambda j, c: f'=IF({j+1}=3,{S["CMP"]},0)')
    pc = lambda j, rrow: (f'{L(2+j)}{rrow}' if j > 0 else '0')
    line('base', 'Surplus before reinvestment-policy capex', '$', lambda j, c: f'={c}{rows["sur24"]}-{c}{rows["drest"]}-{c}{rows["cmp"]}-{pc(j, rr+2)}')
    line('pcap', 'Policy capex (share of surplus, if positive)', '$', lambda j, c: f'=IF(AND({j+1}>={S["POLSTART"]},{c}{rows["base"]}>0),{S["POLSHARE"]}*{c}{rows["base"]},0)')
    line('pcum', 'Cumulative policy depreciation', '$', lambda j, c: (f'={pc(j, rr)}+{c}{rows["pcap"]}/{S["POLLIFE"]}'))
    line('sur', 'Operating surplus (Finance 2.5 basis)', '$', lambda j, c: f'={c}{rows["base"]}-{c}{rows["pcap"]}/{S["POLLIFE"]}', USD, bold=True)
    line('reinv', 'Reinvestment capex paid in cash', '$', lambda j, c: f'=IF({j+1}=3,{S["WETEQ"]}+{S["REST"]},0)+{c}{rows["pcap"]}')
    for j in range(7): sc[f'{L(1+j)}{rows["sur"]}'].fill = SUB
    BL[name] = rows
    return rr + 1
nxt = 4
nxt = build('Low', nxt, S['L_LOW'], S['G_LOW'])
nxt = build('Base', nxt, S['L_BASE'], S['G_BASE'])
nxt = build('Finance plan', nxt)
nxt = build('High', nxt, S['L_HIGH'], S['G_HIGH'])
sc.column_dimensions['A'].width = 44; sc.column_dimensions['B'].width = 16
for c in 'CDEFG': sc.column_dimensions[c].width = 15
sc.freeze_panes = 'C4'

# ======================= Returns =======================
rt = wb.create_sheet('Returns')
put(rt, 'A1', 'Cash flows and return measures', T)
put(rt, 'A2', 'Year 0 = capex + set-up. Year 1 = set-up-year operating loss + working capital. Years 2-5 = cash surplus before depreciation, less the Year 3 conservation plan and reinvestment capex. Years 6-10 (used only to date payback) held flat at the Year 5 cash surplus, no further reinvestment. Working capital recovered in Year 10.', IT)
RT = {}
r = 4
for name in ('Low', 'Base', 'Finance plan', 'High'):
    b = BL[name]
    put(rt, f'A{r}', name, H, fill=HF)
    for t in range(11): put(rt, f'{L(3+t)}{r}', f'Year {t}', H, fill=HF, al='center')
    put(rt, f'B{r}', 'Unit', H, fill=HF); r += 1
    put(rt, f'A{r}', 'Cash flow'); put(rt, f'B{r}', '$', IT)
    for t in range(11):
        c = L(3+t)
        if t == 0: f = f'=-({S["CAPEX"]}+{S["SETUP"]})'
        elif t == 1: f = f'=Scenarios!C{b["pre"]}-{S["WC"]}-Scenarios!C{b["cmp"]}-Scenarios!C{b["reinv"]}'
        elif t <= 5: f = f'=Scenarios!{L(2+t)}{b["pre"]}-Scenarios!{L(2+t)}{b["cmp"]}-Scenarios!{L(2+t)}{b["reinv"]}'
        else: f = f'=Scenarios!G{b["pre"]}' + (f'+{S["WC"]}' if t == 10 else '')
        put(rt, f'{c}{r}', f, GRN, USD)
    cf = r; r += 1
    put(rt, f'A{r}', 'Cumulative cash flow'); put(rt, f'B{r}', '$', IT)
    for t in range(11):
        c = L(3+t); put(rt, f'{c}{r}', f'={c}{cf}' if t == 0 else f'={L(2+t)}{r}+{c}{cf}', BLK, USD)
    cum = r; r += 1
    m = {}
    def met(key, lab, f, fmt):
        global r
        put(rt, f'A{r}', lab, B); put(rt, f'C{r}', f, BLK, fmt); m[key] = r; r += 1
    put(rt, f'A{r}', 'Year-end cash balance (endowment less launch spend, plus cash surplus less reinvestment)'); put(rt, f'B{r}', '$', IT)
    for t in range(6):
        c = L(3+t)
        put(rt, f'{c}{r}', f'={S["ENDOW"]}-({S["CAPEX"]}+{S["SETUP"]})' if t == 0 else f'={L(2+t)}{r}+Scenarios!{L(2+t)}{b["pre"]}-Scenarios!{L(2+t)}{b["cmp"]}-Scenarios!{L(2+t)}{b["reinv"]}', BLK, USD)
    cashrow = r; r += 1
    met('mincash', 'Lowest year-end cash balance, Years 1-5', f'=MIN(D{cashrow}:H{cashrow})', USD)
    met('pay', 'Payback year (first year cumulative cash >= 0)', f'=IF(M{cum}<0,"none within 10 yrs",COUNTIF(C{cum}:M{cum},"<0"))', '0')
    RT[name] = dict(cf=cf, cum=cum, cash=cashrow, **m); r += 1
rt.column_dimensions['A'].width = 48; rt.column_dimensions['B'].width = 8
for t in range(11): rt.column_dimensions[L(3+t)].width = 14

# ======================= Summary =======================
sm = wb.create_sheet('Summary', 1)
put(sm, 'A1', 'Scenario summary', T)
put(sm, 'A2', 'All figures are formulas linked to Scenarios and Returns. Real 2026 dollars.', IT)
names = ['Low', 'Base', 'Finance plan', 'High']
hdr(sm, 4, ['Measure'] + names)
def srow(r, lab, fn, fmt, bold=False):
    put(sm, f'A{r}', lab, B if bold else BLK)
    for i, n in enumerate(names): put(sm, f'{L(2+i)}{r}', fn(n), GRN, fmt)
srow(5, 'Launch volume multiplier (vs Finance Year 2)', lambda n: {'Low': f'={S["L_LOW"]}', 'Base': f'={S["L_BASE"]}', 'High': f'={S["L_HIGH"]}', 'Finance plan': 1}[n], '0.00')
srow(6, 'Annual demand growth', lambda n: {'Low': f'={S["G_LOW"]}', 'Base': f'={S["G_BASE"]}', 'High': f'={S["G_HIGH"]}', 'Finance plan': f'=(Inputs!G{Y["SD"]}/Inputs!D{Y["SD"]})^(1/3)-1'}[n], PCT)
srow(7, 'Year 5 school student-days', lambda n: f'=Scenarios!G{BL[n]["sd"]}', NUM)
srow(8, 'Year 5 occupancy (% of capacity)', lambda n: f'=Scenarios!G{BL[n]["occ"]}', PCT)
srow(9, 'Year 5 revenue', lambda n: f'=Scenarios!G{BL[n]["rev"]}', USD)
srow(10, 'Year 5 operating surplus (Finance 2.5 basis)', lambda n: f'=Scenarios!G{BL[n]["sur"]}', USD, True)
srow(11, 'Year 5 total cost / revenue (incl. depreciation)', lambda n: f'=(Scenarios!G{BL[n]["opc"]}+Scenarios!G{BL[n]["dep"]})/Scenarios!G{BL[n]["rev"]}', PCT)
srow(12, 'Average investment at Year 5 (2.2 base + reinvestment / 2)', lambda n: f'={S["AVG_INV"]}+({S["WETEQ"]}+{S["REST"]}+SUM(Scenarios!C{BL[n]["pcap"]}:G{BL[n]["pcap"]}))/2', USD)
srow(13, 'Hurdle surplus: 3% of average investment', lambda n: f'={S["HURDLE"]}*{L(2+names.index(n))}12', USD)
srow(14, 'Hurdle surplus: 3% of $1M endowment', lambda n: f'={S["HURDLE"]}*{S["ENDOW"]}', USD)
srow(15, 'ARR, headline (Year 5 operating surplus / average investment)', lambda n: f'=Scenarios!G{BL[n]["sur"]}/{L(2+names.index(n))}12', PCT, True)
srow(16, 'ARR, conservative (Year 5 operating surplus / $1M)', lambda n: f'=Scenarios!G{BL[n]["sur"]}/{S["ENDOW"]}', PCT)
srow(17, 'ARR, 5-year average operating surplus / $1M (includes Year 1 loss)', lambda n: f'=AVERAGE(Scenarios!C{BL[n]["sur"]}:G{BL[n]["sur"]})/{S["ENDOW"]}', PCT)
srow(18, 'Lowest year-end cash balance, Years 1-5', lambda n: f'=Returns!C{RT[n]["mincash"]}', USD, True)
srow(19, 'Payback year (cash flows after Year 5 held flat at Year 5)', lambda n: f'=Returns!C{RT[n]["pay"]}', '0')
srow(20, 'Meets 3% hurdle on headline ARR?', lambda n: f'=IF(Scenarios!G{BL[n]["sur"]}>={L(2+names.index(n))}13,"Yes","No")', '@')
srow(21, 'Meets 3% hurdle on conservative ARR ($30k)?', lambda n: f'=IF(Scenarios!G{BL[n]["sur"]}>={L(2+names.index(n))}14,"Yes","No")', '@')
put(sm, 'A23', 'Reading the table: headline ARR (Finance\'s basis, Drake, 2026b) and conservative ARR on the $1 million endowment give different hurdles ($13,100 and $30,000). The chapter measures on the headline basis and reports the conservative basis as a limit.', IT)
sm.column_dimensions['A'].width = 62
for c in 'BCDE': sm.column_dimensions[c].width = 16

# ======================= Sensitivity (static) =======================
se = wb.create_sheet('Sensitivity')
put(se, 'A1', 'One-at-a-time sensitivity on the Finance plan (static output from analysis.py)', T)
put(se, 'A2', 'Year 5 surplus after depreciation. Thresholds = multiplier of the plan value at which Year 5 surplus falls to the stated level. "Not reached" = the surplus stays above that level across the tested range.', IT)
hdr(se, 4, ['Variable', 'Adverse multiplier', 'Favourable multiplier', 'Y5 surplus (adverse)', 'Y5 surplus (favourable)', 'Swing ($)', 'Reliability (Finance)', 'Multiplier at Y5 surplus = $0', 'Multiplier at 3% of avg. investment ($13.1k)', 'Multiplier at 3% of $1M ($30k)'])
Sd = R['sensitivity']
for i, x in enumerate(Sd['rows'], start=5):
    ADDR['se:'+x['name']] = i
    t = x['thresholds']
    vals = [x['name'], x['low_mult'], x['high_mult'], x['y5_low'], x['y5_high'], x['swing'], x['reliability'],
            t['zero'] if t['zero'] else 'Not reached', t['hurdle_headline'] if t['hurdle_headline'] else 'Not reached', t['hurdle_conservative'] if t['hurdle_conservative'] else 'Not reached']
    fmts = [None, '0.00', '0.00', USD, USD, USD, None, '0.00', '0.00', '0.00']
    for j, (v, f) in enumerate(zip(vals, fmts)): put(se, f'{L(1+j)}{i}', v, BLUE, f)
put(se, f'A{5+len(Sd["rows"])+1}', f'Finance plan Year 5 surplus: ${Sd["base_y5"]:,.0f}. Volume threshold applies the multiplier to every year, including wet-hire-independent school volume; wet hire stays at plan. Schools-only accounting break-even (Finance FINAL): 10,486 student-days in Year 5, an 18.7% margin of safety.', IT)
se.column_dimensions['A'].width = 34
for j in range(2, 11): se.column_dimensions[L(j)].width = 18
se.row_dimensions[4].height = 44

# ======================= MonteCarlo (static) =======================
mc = wb.create_sheet('MonteCarlo')
put(mc, 'A1', 'Monte Carlo (static output from analysis.py: 10,000 runs, seed 42, triangular distributions, inputs independent)', T)
hdr(mc, 3, ['Input', 'Minimum', 'Most likely', 'Maximum', 'Basis'])
G = R['growth']
dist = [
 ('Launch volume multiplier (Year 2)', 50/72, 1.0, 90/72, 'Low / Base / High overnight camp size: 50 / 72 / 90 students (90 = Finance maximum)', '0.00'),
 ('Annual demand growth, case A (benchmark)', G['low'], G['mean'], G['high'], 'Min / mean / max of 4 comparable YoY observations', '0.0%'),
 ('Annual demand growth, case B (stress)', 0.0, G['mean'], G['high'], 'Floor set to 0% for survivorship bias (analyst judgement)', '0.0%'),
 ('Price multiplier', 0.9, 1.0, 1.1, 'Medium reliability; cost-plus price, market ceiling untested', '0.00'),
 ('Contractor day-rate multiplier', 0.8, 1.0, 1.2, 'Medium reliability', '0.00'),
 ('Wet-hire volume multiplier', 0.0, 1.0, 2.0, 'Low reliability: 0 to double the plan', '0.00'),
 ('Utilities multiplier', 0.5, 1.0, 1.5, 'Low reliability', '0.00'),
 ('Core staff cost multiplier', 0.75, 1.0, 1.25, 'Low reliability', '0.00'),
 ('Catering cost multiplier', 0.8, 1.0, 1.2, 'Medium reliability', '0.00'),
 ('Heritage maintenance multiplier', 0.5, 1.0, 3.0, 'Low reliability; one-sided heritage risk (State-listed, fair condition)', '0.00'),
]
for i, d in enumerate(dist, start=4):
    put(mc, f'A{i}', d[0])
    for j in range(3): put(mc, f'{L(2+j)}{i}', d[1+j], BLUE, d[5])
    put(mc, f'E{i}', d[4], IT)
r = 4 + len(dist) + 1
hdr(mc, r, ['Result', 'Case A (benchmark growth)', 'Case B (stress growth)']); r += 1
A, Bq = R['mc']['A_benchmark_growth'], R['mc']['B_stress_growth']
res = [
 ('Year 5 surplus: 5th percentile', A['y5']['5'], Bq['y5']['5'], USD), ('Year 5 surplus: median', A['y5']['50'], Bq['y5']['50'], USD),
 ('Year 5 surplus: 95th percentile', A['y5']['95'], Bq['y5']['95'], USD),
 ('P(Year 5 surplus < $0)', A['p_y5_lt_zero'], Bq['p_y5_lt_zero'], PCT),
 ('P(Year 5 surplus < 3% of average investment, $13.1k)', A['p_y5_lt_hurdle_headline'], Bq['p_y5_lt_hurdle_headline'], PCT),
 ('P(Year 5 surplus < 3% of $1M, $30k)', A['p_y5_lt_hurdle_cons'], Bq['p_y5_lt_hurdle_cons'], PCT),
 ('P(5-year average surplus < 3% of $1M)', A['p_avg5_lt_3pct_endow'], Bq['p_avg5_lt_3pct_endow'], PCT),
 ('Finance plan percentile within Year 5 distribution', A['finance_plan_percentile_y5']/100, Bq['finance_plan_percentile_y5']/100, PCT),
]
for lab, a, b2, f in res:
    ADDR['mc:'+lab] = r
    put(mc, f'A{r}', lab); put(mc, f'B{r}', a, BLUE, f); put(mc, f'C{r}', b2, BLUE, f); r += 1
r += 1
hdr(mc, r, ['Rank correlation (Spearman) with Year 5 surplus', 'Case A', 'Case B']); r += 1
lab = {'launch': 'Launch volume', 'growth': 'Annual growth', 'price': 'Price', 'contractor': 'Contractor rates', 'wet': 'Wet-hire volume',
       'utilities': 'Utilities', 'staff': 'Core staff', 'catering': 'Catering', 'maint': 'Heritage maintenance'}
for k in sorted(A['spearman_y5'], key=lambda k: -abs(A['spearman_y5'][k])):
    ADDR['sp:'+k] = r
    put(mc, f'A{r}', lab[k]); put(mc, f'B{r}', A['spearman_y5'][k], BLUE, '0.00'); put(mc, f'C{r}', Bq['spearman_y5'][k], BLUE, '0.00'); r += 1
mc.column_dimensions['A'].width = 54
for c in 'BCD': mc.column_dimensions[c].width = 22
mc.column_dimensions['E'].width = 70

# ======================= ModelRisk (live formulas) =======================
mr = wb.create_sheet('ModelRisk')
FP = BL['Finance plan']
put(mr, 'A1', 'Model-risk stress: cost lines that look light in Finance\'s workbook', T)
put(mr, 'A2', 'Illustrative adjustments to Year 5 surplus. Blue cells are analyst assumptions anchored to the source in column D; change them to test other values. Not Finance numbers.', IT)
hdr(mr, 4, ['Assumption', 'Value', 'Unit', 'Basis'])
mra = [
 ('Catering uplift', 0.5, '% of Finance catering cost', '(NSW Sport and Recreation, 2026) catered vs uncatered day price implies about $22.50 per student-day; compare with the catering line in (Drake, 2026a)', '0%'),
 ('Extra kitchen / cleaning staff', 1, 'FTE', '(Drake, 2026a) Staffing tab lists a Site Manager and 3 General Supervisors; no cook, cleaner or night staff', '0'),
 ('Cost per extra FTE', 70611.818, '$ per year', '(Drake, 2026a) Staffing tab: General Supervisor annual cost incl. on-costs', USD),
 ('Extra insurance', 20000, '$ per year', 'Our assumption: (Drake, 2026a) costs $10,000 of public liability cover only', USD),
 ('Alternative day-program price', 83, '$ per student-day', '(NSW Sport and Recreation, 2026) catered day program price; cost-plus day price in (Drake, 2026a) is $98.60', USD2),
]
for i, (a, v, u, b, f) in enumerate(mra, start=5):
    put(mr, f'A{i}', a); put(mr, f'B{i}', v, BLUE, f, fill=YEL); put(mr, f'C{i}', u, IT); put(mr, f'D{i}', b, IT)
put(mr, 'A11', 'Finance plan Year 5 operating surplus (2.5 basis)', B); put(mr, 'B11', f"=Scenarios!G{FP['sur']}", GRN, USD)
put(mr, 'A12', 'Finance Year 5 catering cost'); put(mr, 'B12', f"=Scenarios!G{FP['cat']}", GRN, USD)
put(mr, 'A13', 'Finance Year 5 day-program student-days'); put(mr, 'B13', f"=Scenarios!G{FP['day']}", GRN, NUM)
put(mr, 'A14', 'Finance day-program price'); put(mr, 'B14', f"={S['DAY_PRICE']}", GRN, USD2)
put(mr, 'A15', 'Hurdle surplus: 3% of $1M'); put(mr, 'B15', f"={S['HURDLE']}*{S['ENDOW']}", GRN, USD)
hdr(mr, 17, ['Case', 'Change to Year 5 surplus (simple formula)', 'Year 5 surplus (simple formula)', 'Above $30k hurdle?', 'Python result incl. policy offset (authoritative)'])
cases = [
 ('Finance plan (no change)', '=0'),
 ('Catering uplift', '=-B5*B12'),
 ('Extra kitchen / cleaning staff', '=-B6*B7'),
 ('Extra insurance', '=-B8'),
 ('Day price at alternative', '=-(B14-B9)*B13'),
 ('Mild: catering + one FTE + insurance', '=B19+B20+B21'),
 ('Heavier: catering + two FTE + insurance + day price', '=B19+2*B20+B21+B22'),
]
for i, (lab, f) in enumerate(cases, start=18):
    put(mr, f'A{i}', lab, B if i >= 23 else BLK); put(mr, f'B{i}', f, BLK, USD)
    put(mr, f'C{i}', f'=$B$11+B{i}', BLK, USD, bold=i >= 23); put(mr, f'D{i}', f'=IF(C{i}>=$B$15,"Yes","No")', BLK)
MRR = R['model_risk']['rows']
for k, idx in zip(range(18, 25), (0, 1, 2, 5, 4, 6, 7)):
    put(mr, f'E{k}', MRR[idx]['y5_surplus'], BLUE, USD)
put(mr, 'A27', 'Column E is the exact Python result. It differs from the simple formula by a few thousand dollars because lower surplus also lowers the 50% conservation reinvestment and its depreciation (Section 2.5 layer).', IT)
put(mr, 'A26', 'Note: the heavier case uses 2 x the per-FTE cost in B7 regardless of B6, so B6 only drives the single-FTE row and the mild case.', IT)
mr.column_dimensions['A'].width = 52; mr.column_dimensions['B'].width = 24; mr.column_dimensions['C'].width = 28; mr.column_dimensions['D'].width = 90; mr.column_dimensions['E'].width = 34

# ======================= Comparators (live formulas) =======================
cp = wb.create_sheet('Comparators')
put(cp, 'A1', 'Comparator charities: cost structure and growth', T)
put(cp, 'A2', 'Latest-year figures reclassified onto a common cost structure by (Quang, 2026), from the statements of (Stewart House, 2025), (Crusader Union, 2025) and (YMCA Camping Ltd, 2025). Ratios are formulas.', IT)
hdr(cp, 4, ['Latest year, $', 'Stewart House (CY2024)', 'Crusader Union (FY2024-25)', 'YMCA Camping (FY2025)', 'Source'])
CPR = {}
cpl = [
 ('rev', 'Total revenue', [3316064, 19009565, 21166872], USD, '(Quang, 2026) Comparable Benchmark!B8:D8'),
 ('emp', 'Employee expenses', [3187539, 10169973, 10319033], USD, '(Quang, 2026) Comparable Benchmark!B9:D9'),
 ('dep', 'Depreciation and amortization', [194899, 1848860, 787682], USD, '(Quang, 2026) Comparable Benchmark!B10:D10'),
 ('adm', 'Admin-like overhead', [906189, 3786924, 2472638], USD, '(Quang, 2026) Comparable Benchmark!B11:D11'),
 ('ops', 'Operations and program costs', [633489, 2922135, 5984943], USD, '(Quang, 2026) Comparable Benchmark!B12:D12'),
 ('oth', 'Other (gain or loss on disposals)', [0, 0, -33360], USD, '(Quang, 2026) Comparable Benchmark!B13:D13'),
]
r = 5
for k, lab, vals, f, src in cpl:
    put(cp, f'A{r}', lab)
    for j, v in enumerate(vals): put(cp, f'{L(2+j)}{r}', v, BLUE, f)
    put(cp, f'E{r}', src, IT); CPR[k] = r; r += 1
put(cp, f'A{r}', 'Total expenses', B)
for j in range(3): put(cp, f'{L(2+j)}{r}', f'=SUM({L(2+j)}{CPR["emp"]}:{L(2+j)}{CPR["oth"]})', BLK, USD, bold=True)
CPR['exp'] = r; r += 1
put(cp, f'A{r}', 'Surplus or (deficit)')
for j in range(3): put(cp, f'{L(2+j)}{r}', f'={L(2+j)}{CPR["rev"]}-{L(2+j)}{CPR["exp"]}', BLK, USD)
CPR['sur'] = r; r += 2
hdr(cp, r, ['Ratios (% of revenue)', 'Stewart House', 'Crusader Union', 'YMCA Camping', 'Finance plan, Year 5']); r += 1
for k, lab, num in (('rx', 'Total expense ratio', 'exp'), ('re', 'Employee cost ratio', 'emp'), ('rs', 'Surplus or (deficit) margin', 'sur')):
    put(cp, f'A{r}', lab, B)
    for j in range(3): put(cp, f'{L(2+j)}{r}', f'={L(2+j)}{CPR[num]}/{L(2+j)}{CPR["rev"]}', BLK, PCT)
    CPR[k] = r; r += 1
put(cp, f'E{CPR["rx"]}', f"=Summary!D11", GRN, PCT)
put(cp, f'F{CPR["rx"]}', '(Drake, 2026b), total cost including depreciation / revenue', IT)
r += 1
put(cp, f'A{r}', 'Crusader Union: capital-campaign adjustment', B); r += 1
adj = [
 ('Revenue excluding capital-campaign donations, FY2024-25', 18249398, USD, '(Crusader Union, 2025) Annual Report FY2024-25'),
 ('Revenue excluding capital-campaign donations, FY2023-24', 16813900, USD, '(Crusader Union, 2025) Annual Report FY2024-25, prior-year column'),
]
for lab, v, f, src in adj:
    put(cp, f'A{r}', lab); put(cp, f'C{r}', v, BLUE, f); put(cp, f'E{r}', src, IT); CPR[lab[:30]] = r; r += 1
r_cur, r_prev = r - 2, r - 1
put(cp, f'A{r}', 'Capital-campaign donations, FY2024-25'); put(cp, f'C{r}', f'=C{CPR["rev"]}-C{r_cur}', BLK, USD); CPR['campaign'] = r; r += 1
put(cp, f'A{r}', 'Capital-campaign donations, FY2023-24'); put(cp, f'C{r}', f'=16822335-C{r_prev}', BLK, USD); put(cp, f'E{r}', 'FY2023-24 reported revenue $16,822,335 (Quang, 2026) Growth Trend!B29', IT); CPR['campaign_prev'] = r; r += 1
put(cp, f'A{r}', 'Expense ratio excluding campaign donations'); put(cp, f'C{r}', f'=C{CPR["exp"]}/C{r_cur}', BLK, PCT); CPR['crx_adj'] = r; r += 2
put(cp, f'A{r}', 'Stewart House: donations share of income, CY2024', B); r += 1
put(cp, f'A{r}', 'Donations: schools + salary contributions + corporate and other'); put(cp, f'B{r}', '=214419+1197181+1214314', BLUE, USD); put(cp, f'E{r}', '(Stewart House, 2025) statement of surplus or deficit, 2024 column', IT); CPR['don'] = r; r += 1
put(cp, f'A{r}', 'Share of total revenue'); put(cp, f'B{r}', f'=B{CPR["don"]}/B{CPR["rev"]}', BLK, PCT); CPR['donshare'] = r; r += 2
hdr(cp, r, ['Year-on-year revenue growth', 'Prior year, $', 'Latest year, $', 'Growth', 'Source']); r += 1
gobs = [('Stewart House CY2023 to CY2024', 3162237, 3316064, '(Stewart House, 2025)'),
        ('Crusader Union FY2023-24 to FY2024-25 (excl. capital campaign)', 16813900, 18249398, '(Crusader Union, 2025)'),
        ('YMCA Camping FY2023 to FY2024', 18921663, 19825707, '(YMCA Camping Ltd, 2024)'),
        ('YMCA Camping FY2024 to FY2025', 19825707, 21166872, '(YMCA Camping Ltd, 2025)')]
g0 = r
for n, a, b2, src in gobs:
    put(cp, f'A{r}', n); put(cp, f'B{r}', a, BLUE, USD); put(cp, f'C{r}', b2, BLUE, USD); put(cp, f'D{r}', f'=C{r}/B{r}-1', BLK, '0.00%'); put(cp, f'E{r}', src, IT); r += 1
put(cp, f'A{r}', 'Minimum / mean / maximum', B)
put(cp, f'B{r}', f'=MIN(D{g0}:D{r-1})', BLK, '0.0%'); put(cp, f'C{r}', f'=AVERAGE(D{g0}:D{r-1})', BLK, '0.0%'); put(cp, f'D{r}', f'=MAX(D{g0}:D{r-1})', BLK, '0.0%')
put(cp, f'E{r}', 'Feeds the Low, Base and High growth rates on the Inputs sheet', IT)
cp.column_dimensions['A'].width = 62
for c in 'BCD': cp.column_dimensions[c].width = 26
cp.column_dimensions['E'].width = 70; cp.column_dimensions['F'].width = 50

# ======================= Recommendations (static, from tradeoff2.py + tradeoff.py) =======================
import sys; sys.path.insert(0, '.')
import numpy as np, tradeoff as TR
Dn = TR.draws(10000)
y5_s75, cash_s75 = TR.mc(0, 1, 0.75, Dn, 10000)
trigger_share = float((Dn['launch'] >= 1.0).mean())
HC = 13097.72
rc = wb.create_sheet('Recommendations')
put(rc, 'A1', 'Recommendations: cost of heritage and access commitments (static output of tradeoff2.py)', T)
put(rc, 'A2', 'Deterministic columns: Finance plan volumes, Year 5. Simulation columns: 10,000 runs, seed 42, same random draws in every row. Discount = 25% on the share of school student-days offered to local or under-resourced schools (our assumption); conservation share = share of surplus reinvested in conservation (Finance: 50%, Drake, 2026b).', IT)
hdr(rc, 4, ['Package', 'Access share (25% discount)', 'Maintenance multiple', 'Conservation share', 'Discount timing', 'Y5 surplus (Finance plan)', 'ARR (Finance plan)', 'Lowest cash (Finance plan)', 'Five-year conservation spend', 'Simulated median Y5 surplus', 'Share of runs below $13,100', 'Share of runs with negative cash'])
rc.row_dimensions[4].height = 46
det = lambda k: T2[k]['Finance']
def cur(key, i): return T2['curves'][key][i]
Rr = [
 ('Finance plan', 0, 1, 0.5, 'n/a', det('det_fin') if 'det_fin' in T2 else None, T2['finance']),
 ('75% conservation share', 0, 1, 0.75, 'n/a', det('det_s75'), dict(med=float(np.median(y5_s75)), p_lt=float((y5_s75 < HC).mean()), p_cash=float((cash_s75 < 0).mean()))),
 ('Maintenance doubled', 0, 2, 0.5, 'n/a', det('det_m2'), T2['m2_s5']),
 ('Maintenance tripled', 0, 3, 0.5, 'n/a', None, T2['m3_s5']),
 ('10% access from the start (50% share)', 0.1, 1, 0.5, 'From the start', det('det_a10'), None),
 ('10% access from the start (75% share)', 0.1, 1, 0.75, 'From the start', None, cur('uncond', 1)),
 ('20% access from the start (75% share)', 0.2, 1, 0.75, 'From the start', det('det_a20'), cur('uncond', 2)),
 ('10% access, conditional (75% share): recommended', 0.1, 1, 0.75, 'After Year 2 bookings reach 9,700', det('det_rec'), cur('cond', 1)),
 ('20% access, conditional (75% share)', 0.2, 1, 0.75, 'After Year 2 bookings reach 9,700', None, cur('cond', 2)),
 ('30% access, conditional (75% share)', 0.3, 1, 0.75, 'After Year 2 bookings reach 9,700', None, cur('cond', 3)),
 ('Recommended package with maintenance doubled', 0.1, 2, 0.75, 'After Year 2 bookings reach 9,700', None, T2['rec_m2']),
]
r = 5
for lab, a, mm, sh, tm, d, mcr in Rr:
    put(rc, f'A{r}', lab, B if 'recommended' in lab.lower() else BLK)
    put(rc, f'B{r}', a, BLUE, '0%'); put(rc, f'C{r}', mm, BLUE, '0.0"x"'); put(rc, f'D{r}', sh, BLUE, '0%'); put(rc, f'E{r}', tm, IT)
    if d:
        put(rc, f'F{r}', d['y5'], BLUE, USD); put(rc, f'G{r}', d['arr'], BLUE, PCT); put(rc, f'H{r}', d['min_cash'], BLUE, USD); put(rc, f'I{r}', d['cons'], BLUE, USD)
    if mcr:
        put(rc, f'J{r}', mcr['med'], BLUE, USD); put(rc, f'K{r}', mcr['p_lt'], BLUE, PCT); put(rc, f'L{r}', mcr['p_cash'], BLUE, PCT)
    ADDR['rc:' + lab] = r; r += 1
r += 1
put(rc, f'A{r}', 'Share of simulated futures in which Year 2 bookings reach the 9,700 trigger (launch multiplier of 1.00 or more)', B); put(rc, f'F{r}', trigger_share, BLUE, PCT); ADDR['rc:trigger'] = r; r += 1
put(rc, f'A{r}', 'Finance plan school student-days in Year 5 (for the 10% share)'); put(rc, f'F{r}', f"=Scenarios!G{BL['Finance plan']['sd']}", GRN, NUM); ADDR['rc:sd5'] = r; r += 1
put(rc, f'A{r}', 'Student-days at a 10% share'); put(rc, f'F{r}', f'=0.1*F{r-1}', BLK, NUM); ADDR['rc:sd10'] = r; r += 2
put(rc, f'A{r}', 'Blank cells were not run. The model applies the discount from Year 2, which overstates its cost slightly. We have no data on how demand responds to price, so higher prices are not tested. Contingency of $92,575 is held until Year 2 volume reaches break-even (Drake, 2026b, Inputs!B9; Drake, 2026a, Break-Even!D44).', IT)
rc.column_dimensions['A'].width = 52
for c in 'BCDFGHIJKL': rc.column_dimensions[c].width = 17
rc.column_dimensions['E'].width = 34

# ======================= Chapter_Check =======================
cc = wb.create_sheet('Chapter_Check')
put(cc, 'A1', 'Chapter check: every figure in the chapter against this workbook', T)
put(cc, 'A2', 'Chapter value = the number as written in the chapter (rounded). Workbook value = live link to the cell that produces it (blue = static value from a Finance workbook). Check = OK if the two agree within the tolerance.', IT)
hdr(cc, 4, ['Chapter part', 'Statement in the chapter', 'Chapter value', 'Workbook value', 'Tolerance', 'Check', 'Where it comes from'])
FPB, LB, BB, HB = BL['Finance plan'], BL['Low'], BL['Base'], BL['High']
mcA = lambda lab: f"MonteCarlo!B{ADDR['mc:'+lab]}"
RCr = lambda lab, col: f"Recommendations!{col}{ADDR['rc:'+lab]}"
SEr = lambda name, col: f"Sensitivity!{col}{ADDR['se:'+name]}"
chk = [
 ('Comparators', 'Stewart House expense ratio', 1.48, f"=Comparators!B{CPR['rx']}", 0.005, PCT, '(Quang, 2026); (Stewart House, 2025)'),
 ('Comparators', 'Crusader Union expense ratio', 0.99, f"=Comparators!C{CPR['rx']}", 0.005, PCT, '(Quang, 2026); (Crusader Union, 2025)'),
 ('Comparators', 'YMCA Camping expense ratio', 0.92, f"=Comparators!D{CPR['rx']}", 0.005, PCT, '(Quang, 2026); (YMCA Camping Ltd, 2025)'),
 ('Comparators', 'Employee costs: YMCA Camping', 0.49, f"=Comparators!D{CPR['re']}", 0.005, PCT, 'Comparators sheet'),
 ('Comparators', 'Employee costs: Crusader Union', 0.53, f"=Comparators!C{CPR['re']}", 0.005, PCT, 'Comparators sheet'),
 ('Comparators', 'Employee costs: Stewart House', 0.96, f"=Comparators!B{CPR['re']}", 0.005, PCT, 'Comparators sheet'),
 ('Comparators', 'YMCA Camping surplus margin', 0.077, f"=Comparators!D{CPR['rs']}", 0.0005, PCT, 'Comparators sheet'),
 ('Comparators', 'Growth range: lowest', 0.048, f"=Inputs!D{int(S['G_LOW'].split('$')[-1])}", 0.0005, '0.0%', 'Inputs sheet'),
 ('Comparators', 'Growth range: mean', 0.062, f"=Inputs!D{int(S['G_BASE'].split('$')[-1])}", 0.0005, '0.0%', 'Inputs sheet'),
 ('Comparators', 'Growth range: highest', 0.085, f"=Inputs!D{int(S['G_HIGH'].split('$')[-1])}", 0.0005, '0.0%', 'Inputs sheet'),
 ('Comparators', 'Crusader Union capital-campaign donations, latest year', 760000, f"=Comparators!C{CPR['campaign']}", 500, USD, '(Crusader Union, 2025)'),
 ('Comparators', 'Crusader Union capital-campaign donations, year before', 8000, f"=Comparators!C{CPR['campaign_prev']}", 500, USD, '(Crusader Union, 2025)'),
 ('Comparators', 'Crusader Union expense ratio without donations', 1.03, f"=Comparators!C{CPR['crx_adj']}", 0.005, PCT, 'Comparators sheet'),
 ('Comparators', 'Stewart House deficit as share of revenue', -0.48, f"=Comparators!B{CPR['rs']}", 0.005, PCT, 'Comparators sheet'),
 ('Comparators', 'Stewart House donations share of income', 0.79, f"=Comparators!B{CPR['donshare']}", 0.005, PCT, '(Stewart House, 2025)'),
 ('Demand', 'Finance Year 2 launch volume (student-days)', 9720, f"=Inputs!D{Y['SD']}", 0.5, NUM, '(Drake, 2026a) Volumes'),
 ('Demand', 'Low launch size as share of Base (50 of 72)', 0.69, f"={S['L_LOW']}", 0.005, PCT, 'Inputs sheet'),
 ('Demand', 'High launch size as share of Base (90 of 72)', 1.25, f"={S['L_HIGH']}", 0.005, PCT, 'Inputs sheet'),
 ('Demand', 'Year 5 student-days, Low', 7764, '=Summary!B7', 0.5, NUM, 'Summary'),
 ('Demand', 'Year 5 student-days, Base', 11654, '=Summary!C7', 0.5, NUM, 'Summary'),
 ('Demand', 'Year 5 student-days, High', 15535, '=Summary!E7', 0.5, NUM, 'Summary'),
 ('Demand', 'Finance Year 5 student-days', 12897, '=Summary!D7', 0.5, NUM, 'Summary'),
 ('Demand', 'Finance Year 5 break-even (student-days)', 10486, 10486, 0.5, NUM, '(Drake, 2026a) Break-Even!G44'),
 ('Demand', 'Base Year 2 volume above Finance Year 2 break-even (9,687)', 0.003, f"=Scenarios!D{BB['sd']}/9687-1", 0.0005, '0.0%', '(Drake, 2026a) Break-Even!D44'),
 ('Demand', 'Occupancy, lowest (Low, Year 2)', 0.10, f"=Scenarios!D{LB['occ']}", 0.005, PCT, 'Scenarios'),
 ('Demand', 'Occupancy, highest (High, Year 5)', 0.23, f"=Scenarios!G{HB['occ']}", 0.005, PCT, 'Scenarios'),
 ('Demand', 'Capacity (student-days)', 67500, f"={S['CAP']}", 0.5, NUM, '(Drake, 2026a) Assumptions'),
 ('Return', 'Year 5 surplus, Low', -92000, '=Summary!B10', 500, USD, 'Summary'),
 ('Return', 'Year 5 surplus, Base', 133000, '=Summary!C10', 500, USD, 'Summary'),
 ('Return', 'Year 5 surplus, Finance', 204000, '=Summary!D10', 500, USD, 'Summary; (Drake, 2026b) Calculations!H84'),
 ('Return', 'Year 5 surplus, High', 350000, '=Summary!E10', 500, USD, 'Summary'),
 ('Return', 'Average investment', 437000, '=Summary!D12', 500, USD, '(Drake, 2026b) Calculations!H83'),
 ('Return', 'Hurdle surplus (3% of average investment)', 13100, '=Summary!D13', 50, USD, 'Summary'),
 ('Return', 'Finance ARR, Year 5', 0.47, '=Summary!D15', 0.005, PCT, '(Drake, 2026b) Calculations!H85'),
 ('Sensitivity', 'Volume -20%: Year 5 surplus', 57000, '=' + SEr('School volume (all years)', 'D'), 500, USD, 'Sensitivity'),
 ('Sensitivity', 'Volume +20%: Year 5 surplus', 348000, '=' + SEr('School volume (all years)', 'E'), 500, USD, 'Sensitivity'),
 ('Sensitivity', 'Price -10%: Year 5 surplus', 69000, '=' + SEr('Price per student', 'D'), 500, USD, 'Sensitivity'),
 ('Sensitivity', 'Price +10%: Year 5 surplus', 336000, '=' + SEr('Price per student', 'E'), 500, USD, 'Sensitivity'),
 ('Sensitivity', 'Contractor rates +25%: Year 5 surplus', 104000, '=' + SEr('Contractor day rates', 'D'), 500, USD, 'Sensitivity'),
 ('Sensitivity', 'Contractor rates -25%: Year 5 surplus', 302000, '=' + SEr('Contractor day rates', 'E'), 500, USD, 'Sensitivity'),
 ('Sensitivity', 'No wet hire: Year 5 surplus', 134000, '=' + SEr('Wet-hire volume', 'D'), 500, USD, 'Sensitivity'),
 ('Sensitivity', 'Wet hire share of Year 5 revenue', 0.26, f"=Scenarios!G{FPB['rwet']}/Scenarios!G{FPB['rev']}", 0.005, PCT, 'Scenarios'),
 ('Sensitivity', 'Volume at which surplus reaches the hurdle (share of plan)', 0.74, '=' + SEr('School volume (all years)', 'I'), 0.005, PCT, 'Sensitivity'),
 ('Sensitivity', 'Price at which surplus reaches the hurdle (share of plan)', 0.86, '=' + SEr('Price per student', 'I'), 0.005, PCT, 'Sensitivity'),
 ('Sensitivity', 'Contractor rise at which surplus reaches the hurdle', 0.48, '=' + SEr('Contractor day rates', 'I') + '-1', 0.005, PCT, 'Sensitivity'),
 ('Sensitivity', 'Low case volume as share of Finance Year 5 volume', 0.60, '=Summary!B7/Summary!D7', 0.005, PCT, 'Summary'),
 ('Simulation', 'Median Year 5 surplus', 114000, '=' + mcA('Year 5 surplus: median'), 500, USD, 'MonteCarlo'),
 ('Simulation', '5th percentile', -56000, '=' + mcA('Year 5 surplus: 5th percentile'), 500, USD, 'MonteCarlo'),
 ('Simulation', '95th percentile', 290000, '=' + mcA('Year 5 surplus: 95th percentile'), 500, USD, 'MonteCarlo'),
 ('Simulation', 'Runs below the $13,100 hurdle (benchmark growth)', 0.17, '=' + mcA('P(Year 5 surplus < 3% of average investment, $13.1k)'), 0.005, PCT, 'MonteCarlo'),
 ('Simulation', 'Runs below the hurdle if growth can fall to zero', 0.25, f"=MonteCarlo!C{ADDR['mc:P(Year 5 surplus < 3% of average investment, $13.1k)']}", 0.005, PCT, 'MonteCarlo'),
 ('Simulation', 'Rank correlation: launch volume', 0.70, f"=MonteCarlo!B{ADDR['sp:launch']}", 0.005, '0.00', 'MonteCarlo'),
 ('Simulation', 'Rank correlation: price', 0.44, f"=MonteCarlo!B{ADDR['sp:price']}", 0.005, '0.00', 'MonteCarlo'),
 ('Simulation', 'Finance plan percentile', 0.80, '=' + mcA('Finance plan percentile within Year 5 distribution'), 0.005, PCT, 'MonteCarlo'),
 ('Simulation', 'Runs with negative year-end cash', 0.28, '=' + RCr('Finance plan', 'L'), 0.005, PCT, 'Recommendations'),
 ('Simulation', 'Runs below $30,000 (limit 1)', 0.21, '=' + mcA('P(Year 5 surplus < 3% of $1M, $30k)'), 0.005, PCT, 'MonteCarlo'),
 ('Data', 'Finance total cost as share of revenue, Year 5', 0.89, '=Summary!D11', 0.005, PCT, 'Summary; (Drake, 2026b)'),
 ('Data', 'Stress test (mild): Year 5 surplus', 52000, '=ModelRisk!E23', 500, USD, 'ModelRisk'),
 ('Data', 'Stress test (heavier): Year 5 result', -71000, '=ModelRisk!E24', 500, USD, 'ModelRisk'),
 ('Limitations', 'Hurdle on the $1 million endowment', 30000, '=Summary!D14', 1, USD, 'Summary'),
 ('Limitations', 'Finance ARR on $1 million', 0.204, '=Summary!D16', 0.0005, '0.0%', 'Summary'),
 ('Limitations', 'Finance payback (years, at the Year 5 cash flow)', 8.2, 8.239, 0.05, '0.0', '(Drake, 2026b) Calculations!C116'),
 ('Position', 'Share of runs clearing the hurdle: benchmark growth', 0.83, '=1-' + mcA('P(Year 5 surplus < 3% of average investment, $13.1k)'), 0.005, PCT, 'MonteCarlo'),
 ('Position', 'Share of runs clearing the hurdle: zero-growth floor', 0.75, f"=1-MonteCarlo!C{ADDR['mc:P(Year 5 surplus < 3% of average investment, $13.1k)']}", 0.005, PCT, 'MonteCarlo'),
 ('Recommend 1', 'Contingency released only at break-even', 92600, 92575.436, 50, USD, '(Drake, 2026b) Inputs!B9'),
 ('Recommend 1', 'Year 2 break-even trigger (student-days)', 9700, 9687, 20, NUM, '(Drake, 2026a) Break-Even!D44'),
 ('Recommend 2', '75% share adds conservation spending', 86000, f"={RCr('75% conservation share','I')}-{RCr('Finance plan','I')}", 500, USD, 'Recommendations'),
 ('Recommend 2', '75% share costs Year 5 surplus', 3500, f"={RCr('Finance plan','F')}-{RCr('75% conservation share','F')}", 100, USD, 'Recommendations'),
 ('Recommend 2', 'Doubled maintenance costs Year 5 surplus', 19000, f"={RCr('Finance plan','F')}-{RCr('Maintenance doubled','F')}", 500, USD, 'Recommendations'),
 ('Recommend 2', 'Doubled maintenance adds conservation spending', 80000, f"={RCr('Maintenance doubled','I')}-{RCr('Finance plan','I')}", 500, USD, 'Recommendations'),
 ('Recommend 2', 'Doubled maintenance deepens lowest cash by', 60000, f"={RCr('Finance plan','H')}-{RCr('Maintenance doubled','H')}", 500, USD, 'Recommendations'),
 ('Recommend 2', 'Doubled maintenance: runs below the hurdle', 0.25, '=' + RCr('Maintenance doubled', 'K'), 0.005, PCT, 'Recommendations'),
 ('Recommend 2', 'Tripled maintenance: runs below the hurdle', 0.35, '=' + RCr('Maintenance tripled', 'K'), 0.005, PCT, 'Recommendations'),
 ('Recommend 3', '10% access costs Year 5 surplus', 34000, f"={RCr('Finance plan','F')}-{RCr('10% access from the start (50% share)','F')}", 500, USD, 'Recommendations'),
 ('Recommend 3', '10% access costs lowest cash', 55000, f"={RCr('Finance plan','H')}-{RCr('10% access from the start (50% share)','H')}", 500, USD, 'Recommendations'),
 ('Recommend 3', 'Student-days at a 10% share', 1300, f"=Recommendations!F{ADDR['rc:sd10']}", 20, NUM, 'Recommendations'),
 ('Recommend 3', 'From the start, 10% share: runs below the hurdle', 0.25, '=' + RCr('10% access from the start (75% share)', 'K'), 0.005, PCT, 'Recommendations'),
 ('Recommend 3', 'From the start, 20% share: runs below the hurdle', 0.35, '=' + RCr('20% access from the start (75% share)', 'K'), 0.005, PCT, 'Recommendations'),
 ('Recommend 3', 'Trigger met in share of simulated futures', 0.45, f"=Recommendations!F{ADDR['rc:trigger']}", 0.01, PCT, 'Recommendations'),
 ('Recommend 3', 'Conditional 10% share: runs below the hurdle', 0.18, '=' + RCr('10% access, conditional (75% share): recommended', 'K'), 0.005, PCT, 'Recommendations'),
 ('Recommend 3', 'Conditional 20% share: runs below the hurdle', 0.21, '=' + RCr('20% access, conditional (75% share)', 'K'), 0.005, PCT, 'Recommendations'),
 ('Recommend 3', 'Conditional 30% share: runs below the hurdle', 0.25, '=' + RCr('30% access, conditional (75% share)', 'K'), 0.005, PCT, 'Recommendations'),
 ('Recommend 3', '20% share, from the start: Year 5 surplus', 134000, '=' + RCr('20% access from the start (75% share)', 'F'), 500, USD, 'Recommendations'),
 ('Recommend 3', '20% share, from the start: lowest cash', 80000, '=' + RCr('20% access from the start (75% share)', 'H'), 500, USD, 'Recommendations'),
 ('Recommend 3', '20% share, from the start: five-year conservation', 418000, '=' + RCr('20% access from the start (75% share)', 'I'), 500, USD, 'Recommendations'),
 ('Recommend 3', 'Finance plan five-year conservation', 431000, '=' + RCr('Finance plan', 'I'), 500, USD, 'Recommendations'),
 ('Package', 'Year 5 surplus with the package', 167000, '=' + RCr('10% access, conditional (75% share): recommended', 'F'), 500, USD, 'Recommendations'),
 ('Package', 'ARR with the package', 0.37, '=' + RCr('10% access, conditional (75% share): recommended', 'G'), 0.005, PCT, 'Recommendations'),
 ('Package', 'Five-year conservation with the package', 468000, '=' + RCr('10% access, conditional (75% share): recommended', 'I'), 500, USD, 'Recommendations'),
 ('Package', 'Median simulated surplus with the package', 99000, '=' + RCr('10% access, conditional (75% share): recommended', 'J'), 500, USD, 'Recommendations'),
 ('Package', 'Runs below the hurdle with the package', 0.18, '=' + RCr('10% access, conditional (75% share): recommended', 'K'), 0.005, PCT, 'Recommendations'),
 ('Package', 'Runs below the hurdle with maintenance doubled', 0.28, '=' + RCr('Recommended package with maintenance doubled', 'K'), 0.005, PCT, 'Recommendations'),
]
r = 5
for part, stmt, cval, wf, tol, fmt, where in chk:
    put(cc, f'A{r}', part); put(cc, f'B{r}', stmt); put(cc, f'C{r}', cval, BLUE, fmt)
    put(cc, f'D{r}', wf, GRN if isinstance(wf, str) else BLUE, fmt); put(cc, f'E{r}', tol, BLUE, '0.0000' if tol < 1 else '#,##0.0')
    put(cc, f'F{r}', f'=IF(ABS(D{r}-C{r})<=E{r},"OK","CHECK")', BLK); put(cc, f'G{r}', where, IT); r += 1
put(cc, f'A{r+1}', 'Rows needing a check', B); put(cc, f'D{r+1}', f'=COUNTIF(F5:F{r-1},"CHECK")', BLK, '0')
cc.column_dimensions['A'].width = 16; cc.column_dimensions['B'].width = 62
for c in 'CDE': cc.column_dimensions[c].width = 16
cc.column_dimensions['F'].width = 10; cc.column_dimensions['G'].width = 46
cc.freeze_panes = 'A5'

# ======================= References (Harvard) =======================
rf = wb.create_sheet('References')
put(rf, 'A1', 'References', T)
refs = [
 "Crusader Union (2025) Annual report 2024–25.",
 "Drake, J. (2026a) Hillview cost structure and break-even (FINAL) [Excel workbook]. MQBS3010 Group 1I, Macquarie University. Unpublished.",
 "Drake, J. (2026b) Hillview performance and returns (FINAL) [Excel workbook]. MQBS3010 Group 1I, Macquarie University. Unpublished.",
 "Drake, J. (2026c) Hillview capital allocation (FINAL) [Excel workbook]. MQBS3010 Group 1I, Macquarie University. Unpublished.",
 "Flyvbjerg, B. (2009) 'Survival of the unfittest: why the worst infrastructure gets built and what we can do about it', Oxford Review of Economic Policy, 25(3), pp. 344–367.",
 "NSW Sport and Recreation (2026) Camp price list: school programs. Available at: https://sport.nsw.gov.au/sites/default/files/2025-03/Website_Camp_Pricelist_Centres_2026.pdf",
 "Quang, N. (2026) Hillview comparable charity benchmarking [Excel workbook]. MQBS3010 Group 1I, Macquarie University. Unpublished.",
 "Quang, N. (2026) Hillview Harmony: Business Analytics scenario model [Excel workbook]. MQBS3010 Group 1I, Macquarie University. Unpublished.",
 "Quang, N. (2026) Hillview Harmony: Business Analytics model [Python code]. MQBS3010 Group 1I, Macquarie University. Unpublished.",
 "Saltelli, A., Ratto, M., Andres, T., Campolongo, F., Cariboni, J., Gatelli, D., Saisana, M. and Tarantola, S. (2008) Global sensitivity analysis: the primer. Chichester: Wiley.",
 "Savage, S.L. (2012) The flaw of averages: why we underestimate risk in the face of uncertainty. Hoboken: Wiley.",
 "Stewart House (2025) Financial statements for the year ended 31 December 2024.",
 "United Nations (2015) Transforming our world: the 2030 Agenda for Sustainable Development. A/RES/70/1. New York: United Nations.",
 "YMCA Camping Ltd (2024) Financial statements for the year ended 30 June 2024.",
 "YMCA Camping Ltd (2025) Financial statements for the year ended 30 June 2025.",
]
for i, t in enumerate(refs, start=3): put(rf, f'A{i}', t, BLK, wrap=True)
put(rf, f'A{len(refs)+4}', 'Titles of books, reports and journals are italic in the chapter; plain text is used here. The three Drake entries are lettered by workbook so that (Drake, 2026a) and (Drake, 2026b) match the chapter; (Drake, 2026c) is cited only in this workbook.', IT)
rf.column_dimensions['A'].width = 160

for w in wb.worksheets:
    w.sheet_view.showGridLines = False
wb.save('Hillview_Analytics_Model_Quang.xlsx')
print('saved', BL['Finance plan']['sur'])
