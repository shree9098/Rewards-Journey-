"""Generate a sales and pricing analysis workbook from an Excel source file."""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

import pandas as pd


REQUIRED_COLUMNS = {"B2B", "RP PRICING", "MRP"}
OUTPUT_FILENAME = "YUONE_SALES_ANALYSIS_2025_2026.xlsx"


def normalize_columns(data: pd.DataFrame) -> pd.DataFrame:
    """Return a copy with consistent column names and known typos corrected."""
    normalized = data.copy()
    normalized.columns = [
        str(column).strip().strip("'\"").upper() for column in normalized.columns
    ]
    return normalized.rename(columns={"RP RPICING": "RP PRICING"})


def load_transactions(source_file: Path) -> pd.DataFrame:
    """Load and combine every worksheet containing the required pricing columns."""
    frames: list[pd.DataFrame] = []

    with pd.ExcelFile(source_file) as workbook:
        for sheet_name in workbook.sheet_names:
            try:
                frame = normalize_columns(
                    pd.read_excel(workbook, sheet_name=sheet_name)
                )
            except (ValueError, TypeError) as error:
                print(f"Warning: skipped sheet '{sheet_name}': {error}", file=sys.stderr)
                continue

            if REQUIRED_COLUMNS.issubset(frame.columns):
                frame["SOURCE_SHEET"] = sheet_name
                frames.append(frame)

    if not frames:
        required = ", ".join(sorted(REQUIRED_COLUMNS))
        raise ValueError(f"No worksheet contains all required columns: {required}")

    return pd.concat(frames, ignore_index=True)


def prepare_analysis(data: pd.DataFrame, default_year: int | None = None) -> pd.DataFrame:
    """Clean numeric fields and calculate pricing metrics."""
    result = data.copy()
    for column in REQUIRED_COLUMNS:
        result[column] = pd.to_numeric(result[column], errors="coerce")

    result = result.dropna(subset=list(REQUIRED_COLUMNS)).copy()
    if result.empty:
        raise ValueError("No rows contain valid B2B, RP PRICING, and MRP values")

    if "YEAR" in result.columns:
        result["YEAR"] = pd.to_numeric(result["YEAR"], errors="coerce").astype("Int64")
    elif "DATE" in result.columns:
        result["YEAR"] = pd.to_datetime(result["DATE"], errors="coerce").dt.year.astype("Int64")
    elif default_year is not None:
        result["YEAR"] = default_year

    result["TARGET_RP"] = (result["B2B"] * 1.20).clip(upper=result["MRP"]).round(2)
    result["CURRENT_PROFIT"] = (result["RP PRICING"] - result["B2B"]).round(2)
    result["TARGET_PROFIT"] = (result["TARGET_RP"] - result["B2B"]).round(2)
    return result


def build_summary(data: pd.DataFrame) -> tuple[pd.DataFrame, str]:
    """Build a yearly summary when real years exist, otherwise an overall summary."""
    aggregations = {
        "Total_Products": ("B2B", "size"),
        "Avg_Cost_B2B": ("B2B", "mean"),
        "Avg_Current_Retail_RP": ("RP PRICING", "mean"),
        "Avg_Max_Retail_MRP": ("MRP", "mean"),
        "Projected_Target_RP_Avg": ("TARGET_RP", "mean"),
        "Total_Current_Profit_Pool": ("CURRENT_PROFIT", "sum"),
        "Total_Target_Profit_Pool": ("TARGET_PROFIT", "sum"),
    }

    if "YEAR" in data.columns and data["YEAR"].notna().any():
        summary = data.dropna(subset=["YEAR"]).groupby("YEAR").agg(**aggregations).reset_index()
        sheet_name = "Yearly_Sales_Comparison"
    else:
        summary = data.assign(SCOPE="All data").groupby("SCOPE").agg(**aggregations).reset_index()
        sheet_name = "Sales_Summary"

    numeric_columns = summary.select_dtypes(include="number").columns
    summary[numeric_columns] = summary[numeric_columns].round(2)
    return summary, sheet_name


def generate_sales_analysis(
    source_file: Path, output_file: Path, default_year: int | None = None
) -> Path:
    """Generate the analysis workbook and return its path."""
    if not source_file.is_file():
        raise FileNotFoundError(f"Source workbook does not exist: {source_file}")
    if source_file.resolve() == output_file.resolve():
        raise ValueError("Source and output files must be different")

    output_file.parent.mkdir(parents=True, exist_ok=True)
    analysis = prepare_analysis(load_transactions(source_file), default_year)
    summary, summary_sheet = build_summary(analysis)

    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        summary.to_excel(writer, sheet_name=summary_sheet, index=False)
        analysis.to_excel(writer, sheet_name="Analyzed_Raw_Data", index=False)

    return output_file


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "source",
        nargs="?",
        type=Path,
        default=Path("Rewards_YUONE_RP_PRICING_FINAL.xlsx"),
        help="input Excel workbook",
    )
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        default=Path(OUTPUT_FILENAME),
        help="output Excel workbook",
    )
    parser.add_argument(
        "--default-year",
        type=int,
        help="year to use only when the source has no YEAR or DATE column",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    try:
        output = generate_sales_analysis(args.source, args.output, args.default_year)
    except (FileNotFoundError, ValueError, OSError, ImportError) as error:
        print(f"Error: {error}", file=sys.stderr)
        return 1

    print(f"Sales analysis created: {output.resolve()}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
