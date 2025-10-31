# Attendance Processor (Streamlit)

This repository contains a small Streamlit app that processes attendance Excel files and computes worked hours per day and missing hours.

How to run

1. Create a virtual environment (recommended):

```bash
python -m venv .venv
source .venv/bin/activate
```

2. Install dependencies:

```bash
pip install -r requirements.txt
```

3. Run the app:

```bash
streamlit run main.py
```

Usage

- Upload an Excel file (.xlsx or .xls) containing columns named `Дата`, `приход`, and `уход`.
- The app will pick the first sheet that contains those columns, compute worked hours per day, show a table, and provide a download link for results as an Excel file.

Notes

- The app tolerates repeated rows where the `Дата` column is filled once and subsequent rows are empty — it will forward-fill dates.
- Time strings like `09:15 (1)` or `09:15 (нет)` are handled; `(нет)` is treated as missing.
