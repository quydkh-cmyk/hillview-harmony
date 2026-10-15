# Hillview Harmony: Business Analytics model (Python)

Quang, N. (2026) *Hillview Harmony: Business Analytics model* [Python code]. MQBS3010 Group 1I, Macquarie University. Unpublished.

Rebuilds the cost engine in Finance's FINAL workbooks (Drake, 2026a; 2026b; 2026c) and runs the scenarios, sensitivity, Monte Carlo and recommendation tests in the Business Analytics chapter. The Finance plan reproduces Finance's Year 5 surplus of $203,657 to within $1.

## Run (Python 3 with numpy, matplotlib, openpyxl)
    python3 analysis.py      # scenarios, sensitivity, Monte Carlo, stress test -> results.json
    python3 tradeoff2.py     # recommendation tests -> tradeoff_v2.json, figures/fig5_access_tradeoff.png
    python3 charts.py        # figures 1-4
    python3 build_xlsx.py    # Excel workbook (Hillview_Analytics_Model_Quang.xlsx)

Seed 42, 10,000 runs, triangular inputs. Results are identical on every run.

## Files
- model.py: Finance's cost engine (inputs cited to Drake, 2026a; 2026b; 2026c).
- analysis.py: Figures 1 to 4 inputs, stress test, simulation.
- tradeoff.py, tradeoff2.py: heritage and access commitments (Figure 5).
- charts.py: Figures 1 to 4.
- build_xlsx.py: writes the workbook; Chapter_Check lists every chapter figure against it.

## References
Drake, J. (2026a) *Hillview cost structure and break-even (FINAL)* [Excel workbook]. MQBS3010 Group 1I, Macquarie University. Unpublished.
Drake, J. (2026b) *Hillview performance and returns (FINAL)* [Excel workbook]. MQBS3010 Group 1I, Macquarie University. Unpublished.
Drake, J. (2026c) *Hillview capital allocation (FINAL)* [Excel workbook]. MQBS3010 Group 1I, Macquarie University. Unpublished.
