# PDF Financial Data to Excel App

A Python Streamlit app that scans uploaded PDF files, extracts key financial figures from annual statements, and exports a polished Excel workbook.

## Features

- Upload multiple PDFs.
- Detects common financial metrics (Revenue, Net Income, Assets, Liabilities, EBITDA, EPS, etc.).
- Supports currency detection for **INR, USD, EUR, and GBP**.
- Attempts to map values to years for yearly statements.
- Generates a nicely formatted Excel file with:
  - `Financial Data` sheet (raw extracted rows)
  - `Yearly Summary` sheet (count/sum/mean by year + metric + currency)
  - `Metric x Year` sheet (pivot-style view)

## Run locally

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
streamlit run app.py
```

Then open the URL shown in your terminal (usually `http://localhost:8501`).

## Notes

- Extraction quality depends on how text is represented in the PDF.
- Scanned PDFs (image-only) may need OCR first.
- Year mapping is best-effort and works best on table-like annual statement text.
