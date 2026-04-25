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

## Bhagavad Gita Instagram image generator
A chapter-wise image generator that calls OpenAI Images API and keeps progression state between runs.

### Install
```bash
pip install openai
```

### Run
```bash
export OPENAI_API_KEY="your_api_key"
python gita_instagram_image_generator.py
```

Each run generates the **next 3 chapter posts** by default and saves files to `output/gita_images/` plus state in `output/gita_state.json`.

Optional:
```bash
python gita_instagram_image_generator.py --force-chapter 11
python gita_instagram_image_generator.py --model gpt-image-1 --size 1024x1280
python gita_instagram_image_generator.py --posts-per-run 3
```
