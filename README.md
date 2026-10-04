# OceanEmbes-SIH
## DrossVault dMRV dashboard

`dashboard/index.html` is a self-contained dashboard (open it in a browser, no server needed) built from:

- `data/DrossVault_Digital_MRV_Executive_Dashboard.xlsx`: 30-day hourly MRV dataset for a 100 TPD dross processing line
- `data/Dross_Vault_dMRV_Hackathon_Submission.pdf`: dMRV design for CBG plants (credit rules, registry, business model)

Tabs: Overview, Energy & savings, Operations & alerts, Data quality, Carbon credits, Rollout & risks.

To rebuild after editing the workbook or `dashboard/src.html`:

```
pip install openpyxl
python3 dashboard/build_data.py
```
