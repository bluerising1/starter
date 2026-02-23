#!/usr/bin/env python3
"""Analyze bank statement Excel files, auto-categorize expenses, and generate charts.

Usage:
    python bank_statement_analyzer.py --input statement.xlsx

The script expects an Excel file with at least these columns:
- Date
- Description
- Amount

It will:
1. Clean and normalize the statement.
2. Auto-sort transactions by date.
3. Categorize spending into common buckets.
4. Mark categories as Necessity vs Non-necessity.
5. Export detailed and summary reports to Excel.
6. Save a spending chart as PNG.
"""

from __future__ import annotations

import argparse
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Pattern, Sequence, Tuple

import matplotlib.pyplot as plt
import pandas as pd


DEFAULT_COLUMN_ALIASES: Dict[str, Sequence[str]] = {
    "date": ("date", "txn date", "transaction date", "posted date"),
    "description": (
        "description",
        "narration",
        "transaction details",
        "merchant",
        "remarks",
    ),
    "amount": ("amount", "debit", "withdrawal", "transaction amount", "value"),
    "credit": ("credit", "deposit"),
    "debit": ("debit", "withdrawal"),
}


@dataclass(frozen=True)
class CategoryRule:
    category: str
    necessity: str
    patterns: Sequence[Pattern[str]]


CATEGORY_RULES: List[CategoryRule] = [
    CategoryRule(
        category="Rent / Housing",
        necessity="Necessity",
        patterns=[re.compile(p, re.IGNORECASE) for p in [r"rent", r"landlord", r"mortgage"]],
    ),
    CategoryRule(
        category="Groceries",
        necessity="Necessity",
        patterns=[
            re.compile(p, re.IGNORECASE)
            for p in [r"grocery", r"supermarket", r"mart", r"whole ?foods", r"aldi", r"walmart"]
        ],
    ),
    CategoryRule(
        category="Utilities",
        necessity="Necessity",
        patterns=[
            re.compile(p, re.IGNORECASE)
            for p in [r"electric", r"water", r"gas bill", r"internet", r"broadband", r"utility"]
        ],
    ),
    CategoryRule(
        category="Transport",
        necessity="Necessity",
        patterns=[
            re.compile(p, re.IGNORECASE)
            for p in [r"fuel", r"petrol", r"diesel", r"uber", r"lyft", r"metro", r"bus", r"train"]
        ],
    ),
    CategoryRule(
        category="Healthcare",
        necessity="Necessity",
        patterns=[re.compile(p, re.IGNORECASE) for p in [r"hospital", r"pharmacy", r"doctor", r"clinic"]],
    ),
    CategoryRule(
        category="Education",
        necessity="Necessity",
        patterns=[re.compile(p, re.IGNORECASE) for p in [r"tuition", r"school", r"course", r"udemy", r"coursera"]],
    ),
    CategoryRule(
        category="Dining / Food Delivery",
        necessity="Non-necessity",
        patterns=[
            re.compile(p, re.IGNORECASE)
            for p in [r"restaurant", r"cafe", r"zomato", r"swiggy", r"doordash", r"ubereats"]
        ],
    ),
    CategoryRule(
        category="Shopping",
        necessity="Non-necessity",
        patterns=[re.compile(p, re.IGNORECASE) for p in [r"amazon", r"flipkart", r"shopping", r"store"]],
    ),
    CategoryRule(
        category="Entertainment",
        necessity="Non-necessity",
        patterns=[
            re.compile(p, re.IGNORECASE)
            for p in [r"netflix", r"spotify", r"prime", r"movie", r"cinema", r"game"]
        ],
    ),
    CategoryRule(
        category="Travel",
        necessity="Non-necessity",
        patterns=[re.compile(p, re.IGNORECASE) for p in [r"airlines", r"hotel", r"booking", r"trip", r"vacation"]],
    ),
]


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Auto-sort and analyze a bank statement Excel file.")
    parser.add_argument("--input", required=True, help="Path to bank statement Excel file")
    parser.add_argument("--output-dir", default="output", help="Directory for reports and chart")
    parser.add_argument(
        "--sheet",
        default=0,
        help="Excel sheet name or index to read (default: first sheet)",
    )
    return parser.parse_args()


def _find_column(columns: Sequence[str], aliases: Sequence[str]) -> str | None:
    alias_set = {a.strip().lower() for a in aliases}
    for column in columns:
        if str(column).strip().lower() in alias_set:
            return str(column)
    return None


def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    columns = [str(c) for c in df.columns]

    date_col = _find_column(columns, DEFAULT_COLUMN_ALIASES["date"])
    description_col = _find_column(columns, DEFAULT_COLUMN_ALIASES["description"])
    amount_col = _find_column(columns, DEFAULT_COLUMN_ALIASES["amount"])

    credit_col = _find_column(columns, DEFAULT_COLUMN_ALIASES["credit"])
    debit_col = _find_column(columns, DEFAULT_COLUMN_ALIASES["debit"])

    if not date_col or not description_col:
        raise ValueError(
            "Input file must contain date and description columns. "
            "Accepted names include Date/Transaction Date and Description/Narration."
        )

    normalized = pd.DataFrame()
    normalized["Date"] = pd.to_datetime(df[date_col], errors="coerce")
    normalized["Description"] = df[description_col].astype(str).str.strip()

    if amount_col:
        normalized["Amount"] = pd.to_numeric(df[amount_col], errors="coerce")
    elif credit_col and debit_col:
        credit = pd.to_numeric(df[credit_col], errors="coerce").fillna(0)
        debit = pd.to_numeric(df[debit_col], errors="coerce").fillna(0)
        normalized["Amount"] = credit - debit
    else:
        raise ValueError(
            "Input file must have either an Amount column or both Credit and Debit columns."
        )

    normalized = normalized.dropna(subset=["Date", "Description", "Amount"])
    normalized["Amount"] = normalized["Amount"].astype(float)

    return normalized


def categorize(description: str) -> Tuple[str, str]:
    for rule in CATEGORY_RULES:
        if any(pattern.search(description) for pattern in rule.patterns):
            return rule.category, rule.necessity
    return "Other", "Non-necessity"


def analyze_statement(df: pd.DataFrame) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    sorted_df = df.sort_values("Date").copy()

    categories = sorted_df["Description"].apply(categorize)
    sorted_df["Category"] = [c for c, _ in categories]
    sorted_df["Necessity"] = [n for _, n in categories]

    expenses = sorted_df[sorted_df["Amount"] < 0].copy()
    expenses["Spend"] = expenses["Amount"].abs()

    summary_by_category = (
        expenses.groupby(["Category", "Necessity"], as_index=False)["Spend"].sum().sort_values("Spend", ascending=False)
    )

    summary_by_need = (
        expenses.groupby("Necessity", as_index=False)["Spend"].sum().sort_values("Spend", ascending=False)
    )

    return sorted_df, summary_by_category, summary_by_need


def create_chart(summary_by_category: pd.DataFrame, chart_path: Path) -> None:
    if summary_by_category.empty:
        fig, ax = plt.subplots(figsize=(8, 4))
        ax.text(0.5, 0.5, "No expense data available", ha="center", va="center", fontsize=12)
        ax.axis("off")
        fig.tight_layout()
        fig.savefig(chart_path, dpi=150)
        plt.close(fig)
        return

    fig, ax = plt.subplots(figsize=(10, 6))
    ax.bar(summary_by_category["Category"], summary_by_category["Spend"], color="#4C78A8")
    ax.set_title("Spending by Category")
    ax.set_xlabel("Category")
    ax.set_ylabel("Amount Spent")
    ax.tick_params(axis="x", rotation=35)
    fig.tight_layout()
    fig.savefig(chart_path, dpi=150)
    plt.close(fig)


def save_outputs(
    detailed_df: pd.DataFrame,
    summary_by_category: pd.DataFrame,
    summary_by_need: pd.DataFrame,
    output_dir: Path,
) -> Tuple[Path, Path]:
    output_dir.mkdir(parents=True, exist_ok=True)
    report_path = output_dir / "bank_statement_report.xlsx"
    chart_path = output_dir / "spending_chart.png"

    with pd.ExcelWriter(report_path, engine="openpyxl") as writer:
        detailed_df.to_excel(writer, sheet_name="Detailed Transactions", index=False)
        summary_by_category.to_excel(writer, sheet_name="Spending by Category", index=False)
        summary_by_need.to_excel(writer, sheet_name="Necessity Split", index=False)

    create_chart(summary_by_category, chart_path)

    return report_path, chart_path


def main() -> None:
    args = parse_args()

    input_path = Path(args.input)
    if not input_path.exists():
        raise FileNotFoundError(f"Input file not found: {input_path}")

    df = pd.read_excel(input_path, sheet_name=args.sheet)
    normalized = normalize_columns(df)
    detailed, summary_by_category, summary_by_need = analyze_statement(normalized)

    report_path, chart_path = save_outputs(
        detailed_df=detailed,
        summary_by_category=summary_by_category,
        summary_by_need=summary_by_need,
        output_dir=Path(args.output_dir),
    )

    print("Analysis complete")
    print(f"Report: {report_path}")
    print(f"Chart:  {chart_path}")


if __name__ == "__main__":
    main()
