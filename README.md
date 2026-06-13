# Leasing Payment Overview

A small pandas tool that reads raw leasing data from an Excel workbook, cleans it,
and produces a per-contract report showing how much has been paid and how much
debt remains.

Built as the Module 1 project of a self-directed Python data & automation course.
The dataset is synthetic (no real client data).

## What it does

1. **Loads** two sheets from `lizing_pregled.xlsx`:
   - `Ugovori` — one row per leasing contract (asset, provider, total amount, number of installments)
   - `Rate` — individual monthly payments
2. **Cleans** the payment data: removes duplicate rows and flags suspicious
   amounts (zero, negative, or implausibly large) for manual review.
3. **Aggregates** total paid per contract with `groupby`.
4. **Joins** the paid totals back to the contracts table with `merge`.
5. **Calculates** remaining debt per contract (total − paid).
6. **Exports** a finished report to `lizing_izvestaj.xlsx`.

One script turns raw input into a ready report — no manual steps. Adding new
payment rows and re-running recalculates everything automatically.

## Tech

- Python 3.13
- pandas, openpyxl

## Run

```bash
python -m venv venv
venv\Scripts\activate          # Windows
pip install -r requirements.txt
python Obrada1.py
```

## Key idea

Clean before you calculate. `groupby` will faithfully sum whatever it is given,
so duplicates and bad entries are removed *before* aggregation — otherwise the
totals look tidy but are wrong ("garbage in, garbage out").
