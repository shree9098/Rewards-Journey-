# YUONE Sales Analysis

Generates a cleaned pricing analysis from the source Excel workbook.

## Setup

```powershell
python -m pip install -r requirements.txt
```

## Run

```powershell
python sales_analysis.py
```

The default input is `Rewards_YUONE_RP_PRICING_FINAL.xlsx` and the default output
is `YUONE_SALES_ANALYSIS_2025_2026.xlsx`. Custom paths are also supported:

```powershell
python sales_analysis.py input.xlsx --output analysis.xlsx
```

If the input contains a `YEAR` column, or a `DATE` column from which a year can be
derived, the output includes a yearly comparison. Otherwise it includes an honest
overall summary. To intentionally assign undated data to a year, pass
`--default-year 2026`.
