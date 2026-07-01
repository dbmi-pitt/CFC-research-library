"""Create a separate OpenAI-assisted folder assignment workbook.

This script does not modify cfc_research_library.py or the source workbook.

Default input:
    reports/CFC_All_Categories_Master_Review_Report.xlsx

Default output:
    reports/CFC_Reference_Folder_Assignment.xlsx
"""

from __future__ import annotations

import argparse
import json
import os
import re
import textwrap
import time
from pathlib import Path
from typing import Any

import pandas as pd

from cfc_research_library import SECTIONS


DEFAULT_INPUT = "reports/CFC_All_Categories_Master_Review_Report.xlsx"
DEFAULT_OUTPUT = "reports/CFC_Reference_Folder_Assignment.xlsx"


def load_env_file(path: Path = Path(".env")) -> None:
    if not path.exists():
        return
    for raw_line in path.read_text(encoding="utf-8").splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        os.environ.setdefault(key.strip(), value.strip().strip('"').strip("'"))


def categories_payload() -> list[dict[str, str]]:
    return [
        {
            "category": section.name,
            "description": section.description,
            "inclusion": section.inclusion,
            "exclusion": section.exclusion,
        }
        for section in SECTIONS.values()
    ]


def normalize(value: Any) -> str:
    if value is None:
        return ""
    return str(value).strip()


def relevant_articles(df: pd.DataFrame, include_screen_out: bool) -> pd.DataFrame:
    data = df.copy()
    for column in ["Reviewer_Decision", "Eligibility_Decision", "Title", "Abstract"]:
        if column not in data:
            data[column] = ""

    reviewer = data["Reviewer_Decision"].fillna("").astype(str).str.lower()
    eligibility = data["Eligibility_Decision"].fillna("").astype(str).str.lower()
    has_title = data["Title"].fillna("").astype(str).str.strip().ne("")

    keep = has_title & ~reviewer.eq("not approved")
    if not include_screen_out:
        keep = keep & ~eligibility.eq("screen out")

    filtered = data[keep].copy()
    filtered["Reference_Relevance_Reason"] = "Included for OpenAI folder assignment"
    return filtered.reset_index(drop=True)


def parse_json(text: str) -> dict:
    cleaned = text.strip()
    if cleaned.startswith("```"):
        cleaned = re.sub(r"^```(?:json)?", "", cleaned, flags=re.IGNORECASE).strip()
        cleaned = re.sub(r"```$", "", cleaned).strip()
    try:
        return json.loads(cleaned)
    except json.JSONDecodeError:
        start = cleaned.find("{")
        end = cleaned.rfind("}")
        if start >= 0 and end > start:
            return json.loads(cleaned[start : end + 1])
        raise


def assign_batch(client: Any, model: str, batch: pd.DataFrame) -> dict[int, dict[str, str]]:
    articles = []
    for idx, row in batch.iterrows():
        articles.append(
            {
                "row_index": int(idx),
                "pmid": normalize(row.get("PMID")),
                "title": normalize(row.get("Title")),
                "abstract": normalize(row.get("Abstract"))[:3000],
                "primary_category": normalize(row.get("Primary_Category")),
                "matched_categories": normalize(row.get("Matched_Categories")),
                "suggested_labels": normalize(row.get("Suggested_Labels")),
                "eligibility_decision": normalize(row.get("Eligibility_Decision")),
            }
        )

    prompt = textwrap.dedent(
        f"""
        Assign each article to the single best CFC research library folder/category.

        Use these category definitions and inclusion/exclusion criteria:
        {json.dumps(categories_payload(), indent=2)}

        Articles to classify:
        {json.dumps(articles, indent=2)}

        Return only valid JSON in this shape:
        {{
          "assignments": [
            {{
              "row_index": 0,
              "api_suggested_category": "Dermatology",
              "api_secondary_category": "Allergy and Immunology",
              "api_relevance_decision": "Relevant",
              "api_confidence": "High",
              "api_analysis": "Brief reason for the category and relevance decision."
            }}
          ]
        }}

        api_relevance_decision must be one of: Relevant, Possibly relevant, Not relevant.
        api_confidence must be one of: High, Medium, Low.
        Use the article title and abstract. Do not invent facts.
        """
    ).strip()

    response = client.responses.create(model=model, input=prompt)
    payload = parse_json(getattr(response, "output_text", ""))
    results: dict[int, dict[str, str]] = {}
    for item in payload.get("assignments", []):
        try:
            row_index = int(item["row_index"])
        except Exception:
            continue
        results[row_index] = {
            "API_Suggested_Category": normalize(item.get("api_suggested_category")),
            "API_Secondary_Category": normalize(item.get("api_secondary_category")),
            "API_Relevance_Decision": normalize(item.get("api_relevance_decision")),
            "API_Confidence": normalize(item.get("api_confidence")),
            "API_Analysis": normalize(item.get("api_analysis")),
        }
    return results


def add_api_assignments(df: pd.DataFrame, model: str, batch_size: int) -> pd.DataFrame:
    api_key = os.getenv("OPENAI_API_KEY")
    if not api_key:
        raise RuntimeError("OPENAI_API_KEY is required. Add it to .env or your environment.")

    from openai import OpenAI

    client = OpenAI(api_key=api_key)
    output = df.copy()
    for column in [
        "API_Suggested_Category",
        "API_Secondary_Category",
        "API_Relevance_Decision",
        "API_Confidence",
        "API_Analysis",
    ]:
        output[column] = ""

    for start in range(0, len(output), batch_size):
        batch = output.iloc[start : start + batch_size]
        try:
            assignments = assign_batch(client, model, batch)
        except Exception as exc:
            assignments = {}
            for idx, row in batch.iterrows():
                output.at[idx, "API_Suggested_Category"] = normalize(row.get("Primary_Category"))
                output.at[idx, "API_Relevance_Decision"] = "Possibly relevant"
                output.at[idx, "API_Confidence"] = "Low"
                output.at[idx, "API_Analysis"] = f"OpenAI batch failed; fallback used Primary_Category. Error: {type(exc).__name__}"
        for idx, assignment in assignments.items():
            for column, value in assignment.items():
                output.at[idx, column] = value

    return output


def write_reference_workbook(df: pd.DataFrame, output_path: Path, input_path: Path, model: str) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    preferred = [
        "API_Suggested_Category",
        "API_Secondary_Category",
        "API_Relevance_Decision",
        "API_Confidence",
        "Primary_Category",
        "Matched_Categories",
        "Reviewer_Decision",
        "PMID",
        "Title",
        "Authors",
        "Journal",
        "Publication_Year",
        "Abstract",
        "Suggested_Labels",
        "Eligibility_Decision",
        "API_Analysis",
        "Reference_Relevance_Reason",
        "PubMed_URL",
    ]
    columns = [column for column in preferred if column in df.columns] + [
        column for column in df.columns if column not in preferred
    ]
    summary = (
        df.groupby(["API_Suggested_Category", "API_Relevance_Decision"], dropna=False)
        .size()
        .reset_index(name="Article_Count")
        .sort_values(["API_Suggested_Category", "API_Relevance_Decision"])
    )
    run_log = pd.DataFrame(
        [
            {"Field": "Input", "Value": str(input_path)},
            {"Field": "Output", "Value": str(output_path)},
            {"Field": "OpenAI model", "Value": model},
            {"Field": "Rows", "Value": len(df)},
            {"Field": "Generated", "Value": time.ctime()},
        ]
    )
    criteria = pd.DataFrame(categories_payload())
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        df[columns].to_excel(writer, sheet_name="Reference_Assignments", index=False)
        summary.to_excel(writer, sheet_name="Category_Summary", index=False)
        criteria.to_excel(writer, sheet_name="Criteria", index=False)
        run_log.to_excel(writer, sheet_name="Run_Log", index=False)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Create OpenAI-assisted reference folder assignments.")
    parser.add_argument("--input", default=DEFAULT_INPUT)
    parser.add_argument("--output", default=DEFAULT_OUTPUT)
    parser.add_argument("--sheet", default="Review_Report")
    parser.add_argument("--model", default=os.getenv("OPENAI_CATEGORY_MODEL", "gpt-4.1-mini"))
    parser.add_argument("--batch-size", type=int, default=5)
    parser.add_argument("--limit", type=int, help="Optional limit for testing.")
    parser.add_argument("--include-screen-out", action="store_true")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    load_env_file()
    input_path = Path(args.input)
    output_path = Path(args.output)
    df = pd.read_excel(input_path, sheet_name=args.sheet, dtype={"PMID": str})
    filtered = relevant_articles(df, include_screen_out=args.include_screen_out)
    if args.limit:
        filtered = filtered.head(args.limit).copy()
    assigned = add_api_assignments(filtered, model=args.model, batch_size=args.batch_size)
    write_reference_workbook(assigned, output_path, input_path, args.model)
    print(f"Reference assignment workbook written: {output_path.resolve()}")


if __name__ == "__main__":
    main()
