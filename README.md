# Manpower Validation System

Simple Streamlit app for manpower production validation:

- Upload Excel (`.xlsx`, `.xlsm`, `.xls`)
- Map your columns (advisor, classification, AC, lives, and date)
- View totals per classification (A/B/C/etc.) by month and quarter
- View advisor-level AC/lives details inside each classification

## Run

```bash
cd "C:\Users\qlngn\manpower-validator"
pip install -r requirements.txt
streamlit run app.py
```

## Notes

- Date mapping supports either:
  - one date column, or
  - month + year columns
- You can filter by classification, month, and quarter.
- Export advisor details as CSV.

