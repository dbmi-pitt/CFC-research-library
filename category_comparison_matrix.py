"""Create 2x2 category comparison matrices from reference folder assignments.

Input is usually:
    reports/CFC_Reference_Folder_Assignment.xlsx

Output:
    reports/CFC_Category_Comparison_Matrix.xlsx
"""

from __future__ import annotations

import argparse
import time
from pathlib import Path

import pandas as pd

from cfc_research_library import SECTIONS


DEFAULT_INPUT = "reports/CFC_Reference_Folder_Assignment.xlsx"
DEFAULT_OUTPUT = "reports/CFC_Category_Comparison_Matrix.xlsx"


def split_categories(value: object) -> set[str]:
    if value is None or pd.isna(value):
        return set()
    return {part.strip() for part in str(value).split(",") if part.strip()}


def build_matrices(df: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    rows = []
    article_rows = []
    for category in SECTIONS:
        tp = fp = fn = tn = 0
        for _, row in df.iterrows():
            existing_categories = split_categories(row.get("Matched_Categories")) or split_categories(row.get("Primary_Category"))
            existing_yes = category in existing_categories
            api_yes = category == str(row.get("API_Suggested_Category", "")).strip()

            if existing_yes and api_yes:
                cell = "TP"
                tp += 1
            elif not existing_yes and api_yes:
                cell = "FP"
                fp += 1
            elif existing_yes and not api_yes:
                cell = "FN"
                fn += 1
            else:
                cell = "TN"
                tn += 1

            if cell != "TN":
                article_rows.append(
                    {
                        "Category": category,
                        "Matrix_Cell": cell,
                        "PMID": row.get("PMID", ""),
                        "Title": row.get("Title", ""),
                        "Existing_Category": row.get("Primary_Category", ""),
                        "Existing_Matched_Categories": row.get("Matched_Categories", ""),
                        "API_Suggested_Category": row.get("API_Suggested_Category", ""),
                        "API_Secondary_Category": row.get("API_Secondary_Category", ""),
                        "API_Relevance_Decision": row.get("API_Relevance_Decision", ""),
                        "API_Confidence": row.get("API_Confidence", ""),
                        "API_Analysis": row.get("API_Analysis", ""),
                        "PubMed_URL": row.get("PubMed_URL", ""),
                    }
                )

        total = tp + fp + fn + tn
        agreement = (tp + tn) / total if total else 0
        rows.append(
            {
                "Category": category,
                "True_Positive": tp,
                "False_Positive": fp,
                "False_Negative": fn,
                "True_Negative": tn,
                "Total": total,
                "Agreement_Rate": agreement,
                "Needs_Review": fp + fn,
            }
        )
    return pd.DataFrame(rows), pd.DataFrame(article_rows)


def write_workbook(summary: pd.DataFrame, articles: pd.DataFrame, input_path: Path, output_path: Path) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    run_log = pd.DataFrame(
        [
            {"Field": "Input", "Value": str(input_path)},
            {"Field": "Output", "Value": str(output_path)},
            {"Field": "Generated", "Value": time.ctime()},
            {"Field": "Interpretation", "Value": "Existing category membership compared with API_Suggested_Category."},
        ]
    )
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        summary.to_excel(writer, sheet_name="2x2_By_Category", index=False)
        articles.to_excel(writer, sheet_name="Article_Details", index=False)
        articles[articles["Matrix_Cell"].isin(["FP", "FN"])].to_excel(writer, sheet_name="Disagreements", index=False)
        run_log.to_excel(writer, sheet_name="Run_Log", index=False)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Create 2x2 category comparison matrices.")
    parser.add_argument("--input", default=DEFAULT_INPUT)
    parser.add_argument("--output", default=DEFAULT_OUTPUT)
    parser.add_argument("--sheet", default="Reference_Assignments")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    input_path = Path(args.input)
    output_path = Path(args.output)
    df = pd.read_excel(input_path, sheet_name=args.sheet, dtype={"PMID": str})
    summary, articles = build_matrices(df)
    write_workbook(summary, articles, input_path, output_path)
    print(f"Comparison matrix workbook written: {output_path.resolve()}")


if __name__ == "__main__":
    main()
