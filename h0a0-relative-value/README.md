# H0A0 Historical Relative Value Engine

Standalone Python project for analyzing ICE/BAML H0A0 historical constituent files and ranking high-yield bonds/issuers as rich or cheap.

## Default source data path

```powershell
P:\jmorris\ICE H0A0 Historical Index Data
```

## V1.1 model updates

V1.1 tightens the relative-value logic:

- Adds issuer-level rating/sector peer residuals.
- Adds issuer residual historical percentiles, so issuer cheapness is not just raw spread percentile.
- Adds distance-from-1Y-median and distance-from-1Y-75th-percentile fields.
- Moves distressed/event issuers out of the normal Issuer RV tab.
- Adds a separate `issuer_distressed_event_latest.csv` output and `Issuer Distressed` Excel tab.
- Adds confidence fields for bond and issuer screens.
- Keeps bond-level peer groups dependent on sector, rating, and maturity.

## Install

```powershell
cd "C:\H0A0 Relative Value"
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install pandas numpy openpyxl --trusted-host pypi.org --trusted-host files.pythonhosted.org
```

## Run 250-file V1 test

```powershell
.\.venv\Scripts\python.exe .\h0a0_relative_value.py --history-dir "P:\jmorris\ICE H0A0 Historical Index Data" --max-files 250 --rebuild-cache
```

## Run full history

```powershell
.\.venv\Scripts\python.exe .\h0a0_relative_value.py --history-dir "P:\jmorris\ICE H0A0 Historical Index Data" --rebuild-cache
```

## Outputs

Outputs are written to `output/`:

- `bond_rv_latest.csv`
- `issuer_rv_latest.csv`
- `issuer_distressed_event_latest.csv`
- `issuer_switch_candidates.csv`
- `distressed_event_latest.csv`
- `data_quality_report.txt`
- `h0a0_relative_value_report.xlsx`

## Notes

The project is intentionally standalone. It does not depend on Flask or the AI Tool Shed dashboard. Raw ICE/BAML files should stay on the P drive and should not be committed to GitHub.
