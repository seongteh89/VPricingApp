# VPricingApp

Web app for consolidating subcontractor BQ quotation files into a vendor-named quote comparison workbook.

## Run Locally

```powershell
.\.venv\Scripts\python.exe -m uvicorn vpricing_app.main:app --app-dir src --host 127.0.0.1 --port 8000
```

Open:

```text
http://127.0.0.1:8000/
```

## Test

```powershell
.\.venv\Scripts\python.exe -m pytest -v
```

## Current Features

- Upload original BQ workbook.
- Upload vendor Excel or PDF quotes.
- Preview warnings before export.
- Generate vendor-named quote comparison workbook.
- Update an existing generated comparison when a vendor resends a quote.
- Flag vendor-added items, missing items, BQ changes, PDF review rows, and total mismatches.
