# OceanEmbes-SIH
## DrossVault dMRV dashboard

`site/index.html` is the standalone desktop website: a single HTML file with a sidebar layout on wide screens. Open it directly in a browser, or serve it:

```
cd site && python3 -m http.server 8000   # then open http://localhost:8000
```

`dashboard/index.html` is the same page without the document wrapper (used for the published artifact). Both are built from:

- `data/DrossVault_Digital_MRV_Executive_Dashboard.xlsx`: 30-day hourly MRV dataset for a 100 TPD dross processing line
- `data/Dross_Vault_dMRV_Hackathon_Submission.pdf`: dMRV design for CBG plants (credit rules, registry, business model)

Tabs: Overview, Energy & savings, Operations & alerts, Data quality, Carbon credits, Rollout & risks.

To rebuild after editing the workbook or `dashboard/src.html`:

```
pip install openpyxl
python3 dashboard/build_data.py
```
