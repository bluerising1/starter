# starter

Python utility to auto-sort bank statement Excel files, categorize your spending, and generate a spending chart.

## What it does
- Reads your bank statement Excel file.
- Normalizes transaction columns (`Date`, `Description`, `Amount` or `Credit`+`Debit`).
- Auto-sorts transactions by date.
- Categorizes expenses (rent, groceries, utilities, transport, shopping, etc.).
- Labels each category as **Necessity** or **Non-necessity**.
- Exports:
  - `bank_statement_report.xlsx` with detailed and summary sheets.
  - `spending_chart.png` bar chart for your spending by category.

## Setup
```bash
python -m venv .venv
source .venv/bin/activate
pip install pandas matplotlib openpyxl
```

## Run
```bash
python bank_statement_analyzer.py --input your_statement.xlsx --output-dir output
```

Optional:
```bash
python bank_statement_analyzer.py --input your_statement.xlsx --sheet "Sheet1"
```

## Expected input columns
Use any of the following equivalent names:
- Date: `Date`, `Txn Date`, `Transaction Date`, `Posted Date`
- Description: `Description`, `Narration`, `Transaction Details`, `Merchant`, `Remarks`
- Amount: `Amount` (or provide both `Credit` and `Debit`)

## Output files
In your output directory:
- `bank_statement_report.xlsx`
  - `Detailed Transactions`
  - `Spending by Category`
  - `Necessity Split`
- `spending_chart.png`
