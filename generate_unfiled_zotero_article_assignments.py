"""Generate a clean Excel workbook for unfiled Zotero full-text assignments.

This script pulls unfiled, article-like records from Zotero, tries to use
Zotero indexed full text plus the title/abstract, asks OpenAI to apply the
new CFC inclusion/exclusion criteria, and creates a workbook with:

    - All_Unfiled
    - Excluded
    - Criteria
    - Run_Log

The output name is intentionally different from prior workbooks so it will not
overwrite edited files.

Run from the project folder:
    uv run python generate_unfiled_zotero_article_assignments.py
"""

from __future__ import annotations

import argparse
import json
import os
import re
import textwrap
import time
from pathlib import Path
from urllib.parse import quote
from urllib.request import Request, urlopen

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

from cfc_research_library import CATEGORY_COLORS, SECTIONS


DEFAULT_OUTPUT = "reports/Unfiled_Zotero_FullText_AI_Assignment.xlsx"
DEFAULT_OPENAI_MODEL = "gpt-4.1-mini"
ARTICLE_ITEM_TYPES = {"journalArticle", "newspaperArticle", "bookSection"}
EXCLUDED_CATEGORIES = {"Historical Articles", "Conferences"}
BAD_TITLES = {"pdf", "full text pdf", "full text", "pubmed entry"}
DEFAULT_CRITERIA_FILE = Path(
    r"C:\Users\lexim\.codex\attachments\772c3eed-9d2d-47af-8885-b8f6a5fa19aa\pasted-text.txt"
)


def load_env_file(path: Path = Path(".env")) -> None:
    if not path.exists():
        return
    for line in path.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        os.environ.setdefault(key.strip(), value.strip().strip('"').strip("'"))


def require_env(name: str) -> str:
    value = os.getenv(name)
    if not value:
        raise RuntimeError(f"Missing required environment variable: {name}")
    return value


def extract_pmid(extra: object) -> str:
    if not extra:
        return ""
    match = re.search(r"\bPMID:\s*(\d+)\b", str(extra), flags=re.IGNORECASE)
    return match.group(1) if match else ""


def author_text(creators: list[dict] | None) -> str:
    names = []
    for creator in creators or []:
        if creator.get("creatorType") not in {"author", "editor"}:
            continue
        name = " ".join(
            part for part in [creator.get("firstName", ""), creator.get("lastName", "")] if part
        ).strip()
        if not name:
            name = creator.get("name", "")
        if name:
            names.append(name)
    return ", ".join(names)


def year_from_date(value: object) -> str:
    if not value:
        return ""
    match = re.search(r"\b(19\d{2}|20\d{2})\b", str(value))
    return match.group(1) if match else ""


def clean_note(value: object) -> str:
    text = re.sub(r"<[^>]+>", " ", str(value or ""))
    return re.sub(r"\s+", " ", text).strip()


def display_title(data: dict, key: str) -> str:
    for field in ("title", "filename", "publicationTitle", "bookTitle", "proceedingsTitle"):
        value = str(data.get(field, "") or "").strip()
        if value:
            return value
    note = clean_note(data.get("note", ""))
    if note:
        return note[:120]
    return f"Untitled {data.get('itemType', 'item')} ({key})"


def item_text(data: dict, title: str) -> str:
    tags = " ".join(tag.get("tag", "") for tag in data.get("tags", []) if isinstance(tag, dict))
    parts = [
        title,
        data.get("abstractNote", ""),
        clean_note(data.get("note", "")),
        data.get("publicationTitle", ""),
        data.get("extra", ""),
        tags,
    ]
    return ". ".join(str(part) for part in parts if part)


def parse_json_response(text: str) -> dict:
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


def allowed_categories() -> list[str]:
    return [category for category in SECTIONS if category not in EXCLUDED_CATEGORIES]


def load_criteria_text(criteria_file: Path | None = None) -> str:
    if criteria_file and criteria_file.exists():
        return criteria_file.read_text(encoding="utf-8", errors="replace")
    if DEFAULT_CRITERIA_FILE.exists():
        return DEFAULT_CRITERIA_FILE.read_text(encoding="utf-8", errors="replace")
    criteria_rows = []
    for name, section in SECTIONS.items():
        if name in EXCLUDED_CATEGORIES:
            continue
        criteria_rows.append(
            f"{section.name}\n"
            f"Description: {section.description}\n"
            f"Inclusion: {section.inclusion}\n"
            f"Exclusion: {section.exclusion}"
        )
    return "\n\n".join(criteria_rows)


def zotero_api_get_json(url: str, api_key: str) -> dict | list:
    request = Request(
        url,
        headers={"Zotero-API-Key": api_key, "Accept": "application/json"},
    )
    try:
        with urlopen(request, timeout=30) as response:
            return json.loads(response.read().decode("utf-8"))
    except Exception:
        return {}


def zotero_fulltext_for_item(group_id: str, api_key: str, item_key: str) -> tuple[str, str]:
    encoded_key = quote(item_key)
    base = f"https://api.zotero.org/groups/{group_id}/items/{encoded_key}"
    texts = []
    sources = []

    for key, label in [(item_key, "item")]:
        url = f"https://api.zotero.org/groups/{group_id}/items/{quote(key)}/fulltext"
        data = zotero_api_get_json(url, api_key)
        if isinstance(data, dict) and data.get("content"):
            texts.append(str(data.get("content", "")))
            sources.append(label)

    children = zotero_api_get_json(f"{base}/children?limit=100", api_key)
    if isinstance(children, list):
        for child in children:
            child_data = child.get("data", {})
            if child_data.get("itemType") != "attachment":
                continue
            child_key = child.get("key") or child_data.get("key")
            if not child_key:
                continue
            fulltext = zotero_api_get_json(
                f"https://api.zotero.org/groups/{group_id}/items/{quote(child_key)}/fulltext",
                api_key,
            )
            if isinstance(fulltext, dict) and fulltext.get("content"):
                texts.append(str(fulltext.get("content", "")))
                sources.append(f"attachment:{child_key}")

    combined = "\n\n".join(text for text in texts if text).strip()
    return combined, ", ".join(sources)


def fetch_unfiled_articles() -> pd.DataFrame:
    from pyzotero import zotero

    group_id = require_env("ZOTERO_GROUP_ID")
    api_key = require_env("ZOTERO_API_KEY")
    zot = zotero.Zotero(group_id, "group", api_key)
    rows = []
    for item in zot.everything(zot.items()):
        data = item.get("data", {})
        item_type = data.get("itemType", "")
        if data.get("collections"):
            continue
        if item_type not in ARTICLE_ITEM_TYPES:
            continue
        key = item.get("key", "")
        title = display_title(data, key)
        if title.strip().lower() in BAD_TITLES:
            continue
        full_text, full_text_source = zotero_fulltext_for_item(group_id, api_key, key)
        rows.append(
            {
                "Zotero Item Key": key,
                "Title": title,
                "Authors": author_text(data.get("creators")),
                "Year": year_from_date(data.get("date", "")),
                "Journal / Publication": data.get("publicationTitle", "")
                or data.get("proceedingsTitle", "")
                or data.get("bookTitle", ""),
                "DOI": data.get("DOI", ""),
                "PMID": extract_pmid(data.get("extra")),
                "Item Type": item_type,
                "Abstract": data.get("abstractNote", "") or clean_note(data.get("note", "")),
                "Full Text Found?": "Yes" if full_text else "No",
                "Full Text Source": full_text_source,
                "URL": data.get("url", ""),
                "_Text_For_AI": "\n\n".join(
                    part
                    for part in [
                        item_text(data, title),
                        f"ZOTERO INDEXED FULL TEXT:\n{full_text[:18000]}" if full_text else "",
                    ]
                    if part
                ),
            }
        )
    return pd.DataFrame(rows).fillna("")


def classify_openai_batch(client, batch: pd.DataFrame, model_name: str, criteria_text: str) -> list[dict]:
    articles = []
    for idx, row in batch.iterrows():
        articles.append(
            {
                "row_index": int(idx),
                "zotero_item_key": str(row.get("Zotero Item Key", "")),
                "title": str(row.get("Title", "")),
                "abstract": str(row.get("Abstract", ""))[:5000],
                "journal": str(row.get("Journal / Publication", "")),
                "year": str(row.get("Year", "")),
                "doi": str(row.get("DOI", "")),
                "pmid": str(row.get("PMID", "")),
                "full_text_found": str(row.get("Full Text Found?", "")),
                "article_text": str(row.get("_Text_For_AI", ""))[:20000],
            }
        )

    prompt = textwrap.dedent(
        f"""
        Classify each unfiled Zotero article for a CFC syndrome research library.

        Use the available Zotero indexed full text when present. If full text is not present,
        use title, abstract, journal, and metadata.

        Apply the overall inclusion/exclusion criteria and category-specific criteria strictly.

        Output rules:
        - If the article is not CFC-specific, does not include CFC data, or is focused only on Noonan syndrome without CFC relevance, set decision to "Excluded".
        - If excluded, leave ai_category blank and explain why.
        - If included, choose one primary category in ai_category.
        - If the article clearly belongs in more than one category, mention additional categories in additional_categories.
        - Do not add multiple categories just because the article mentions multiple symptoms; only add them when the article substantially belongs in those sections.
        - Do not use Historical Articles or Conferences.

        New inclusion/exclusion criteria and category definitions:
        {criteria_text}

        Articles:
        {json.dumps(articles, indent=2)}

        Return only valid JSON in this exact shape:
        {{
          "articles": [
            {{
              "row_index": 0,
              "decision": "Included",
              "ai_category": "Cardiology",
              "additional_categories": ["Genetics"],
              "confidence": "High",
              "reasoning": "Brief rationale based on the full text/abstract."
            }}
          ]
        }}
        """
    ).strip()

    response = client.responses.create(model=model_name, input=prompt)
    return parse_json_response(response.output_text).get("articles", [])


def add_openai_categories(df: pd.DataFrame, model_name: str, batch_size: int, criteria_text: str) -> pd.DataFrame:
    from openai import OpenAI

    if df.empty:
        df["AI Decision"] = ""
        df["AI Categorization"] = ""
        df["AI Confidence"] = ""
        df["AI Reasoning"] = ""
        return df

    client = OpenAI(api_key=require_env("OPENAI_API_KEY"))
    output = df.copy()
    output["AI Decision"] = ""
    output["AI Categorization"] = ""
    output["AI Confidence"] = ""
    output["AI Reasoning"] = ""

    for start in range(0, len(output), batch_size):
        batch = output.iloc[start : start + batch_size]
        try:
            assignments = classify_openai_batch(client, batch, model_name, criteria_text)
        except Exception as exc:
            assignments = []
            for idx, _ in batch.iterrows():
                output.at[idx, "AI Decision"] = "Needs review"
                output.at[idx, "AI Categorization"] = ""
                output.at[idx, "AI Confidence"] = "Low"
                output.at[idx, "AI Reasoning"] = f"OpenAI batch failed. Error: {type(exc).__name__}"

        for item in assignments:
            idx = item.get("row_index")
            if idx not in output.index:
                continue
            decision = str(item.get("decision", "") or "").strip() or "Needs review"
            category = str(item.get("ai_category", "") or "").strip()
            if category in EXCLUDED_CATEGORIES or category not in SECTIONS:
                category = ""
            additional = item.get("additional_categories", [])
            if isinstance(additional, list):
                additional = [str(value).strip() for value in additional if str(value).strip() in SECTIONS and str(value).strip() not in EXCLUDED_CATEGORIES]
            else:
                additional = []
            category_text = category
            if category and additional:
                category_text = f"{category}; Additional: {', '.join(additional)}"
            output.at[idx, "AI Decision"] = decision
            output.at[idx, "AI Categorization"] = category_text
            output.at[idx, "AI Confidence"] = str(item.get("confidence", "") or "").strip().replace("Moderate", "Medium")
            output.at[idx, "AI Reasoning"] = str(item.get("reasoning", "") or item.get("rationale", "") or "").strip()
    return output


def build_output_frame(df: pd.DataFrame) -> pd.DataFrame:
    output = pd.DataFrame(
        {
            "Zotero Item Key": df["Zotero Item Key"],
            "Title": df["Title"],
            "AI Decision": df["AI Decision"],
            "AI Categorization": df["AI Categorization"],
            "AI Confidence": df["AI Confidence"],
            "AI Reasoning": df["AI Reasoning"],
            "Full Text Found?": df["Full Text Found?"],
            "Lexi Category Assignment": "",
            "Lexi Notes": "",
        }
    )
    return output


def criteria_frame() -> pd.DataFrame:
    return pd.DataFrame(
        [
            {
                "Category": section.name,
                "Description": section.description,
                "Inclusion": section.inclusion,
                "Exclusion": section.exclusion,
            }
            for name, section in SECTIONS.items()
            if name not in EXCLUDED_CATEGORIES
        ]
    )


def write_workbook(path: Path, report: pd.DataFrame, openai_model: str, criteria_text: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    excluded = report[report["AI Decision"].str.lower().eq("excluded")].copy()
    summary = report.groupby(["AI Decision", "AI Categorization"], dropna=False).size().reset_index(name="Article Count")
    run_log = pd.DataFrame(
        [
            {"Metric": "Clean unfiled article rows", "Value": len(report)},
            {"Metric": "Excluded rows", "Value": len(excluded)},
            {"Metric": "OpenAI model", "Value": openai_model},
            {"Metric": "Dropdown categories excluded", "Value": ", ".join(sorted(EXCLUDED_CATEGORIES))},
            {"Metric": "Generated", "Value": time.ctime()},
        ]
    )
    criteria_text_frame = pd.DataFrame({"Criteria Used": [criteria_text]})

    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        report.to_excel(writer, sheet_name="All_Unfiled", index=False)
        excluded.to_excel(writer, sheet_name="Excluded", index=False)
        summary.to_excel(writer, sheet_name="Summary", index=False)
        criteria_frame().to_excel(writer, sheet_name="Criteria", index=False)
        criteria_text_frame.to_excel(writer, sheet_name="Full_Criteria_Text", index=False)
        run_log.to_excel(writer, sheet_name="Run_Log", index=False)

    format_workbook(path)


def format_workbook(path: Path) -> None:
    wb = load_workbook(path)
    widths = {
        "Zotero Item Key": 16,
        "Title": 72,
        "AI Decision": 16,
        "AI Categorization": 38,
        "AI Confidence": 16,
        "AI Reasoning": 70,
        "Full Text Found?": 18,
        "Lexi Category Assignment": 28,
        "Lexi Notes": 42,
    }

    for sheet_name in wb.sheetnames:
        sheet = wb[sheet_name]
        sheet.freeze_panes = "A2"
        sheet.auto_filter.ref = sheet.dimensions
        for cell in sheet[1]:
            cell.font = Font(bold=True)
            cell.fill = PatternFill(start_color="D9EAD3", end_color="D9EAD3", fill_type="solid")

        headers = [cell.value for cell in sheet[1]]
        max_row = max(sheet.max_row, 2)
        if sheet_name in {"All_Unfiled", "Excluded"} and "Lexi Category Assignment" in headers:
            col = get_column_letter(headers.index("Lexi Category Assignment") + 1)
            validation = DataValidation(
                type="list",
                formula1=f"='Criteria'!$A$2:$A${len(allowed_categories()) + 1}",
                allow_blank=True,
            )
            sheet.add_data_validation(validation)
            validation.add(f"{col}2:{col}{max_row}")

        if sheet_name in {"All_Unfiled", "Excluded"}:
            for category_header in ("AI Categorization", "Lexi Category Assignment"):
                if category_header in headers:
                    col_index = headers.index(category_header) + 1
                    for row_num in range(2, max_row + 1):
                        value = str(sheet.cell(row=row_num, column=col_index).value or "")
                        category = value.split(";")[0].strip()
                        color = CATEGORY_COLORS.get(category, "FFFFFF")
                        sheet.cell(row=row_num, column=col_index).fill = PatternFill(
                            start_color=color,
                            end_color=color,
                            fill_type="solid",
                        )

        for idx, header in enumerate(headers, start=1):
            sheet.column_dimensions[get_column_letter(idx)].width = widths.get(
                str(header),
                42 if sheet_name in {"Criteria", "Full_Criteria_Text"} else 24,
            )
            if str(header) in {"Title", "AI Reasoning", "Lexi Notes", "Criteria Used"}:
                for row_num in range(2, min(max_row, 250) + 1):
                    sheet.cell(row=row_num, column=idx).alignment = Alignment(wrap_text=True, vertical="top")

    wb.save(path)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Generate clean unfiled Zotero article category workbook.")
    parser.add_argument("--output", default=DEFAULT_OUTPUT, help="Workbook output path.")
    parser.add_argument("--openai-model", default=os.getenv("OPENAI_CATEGORY_MODEL", DEFAULT_OPENAI_MODEL))
    parser.add_argument("--batch-size", type=int, default=8)
    parser.add_argument(
        "--criteria-file",
        type=Path,
        default=DEFAULT_CRITERIA_FILE if DEFAULT_CRITERIA_FILE.exists() else None,
        help="Text file containing updated inclusion/exclusion criteria.",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    load_env_file()
    criteria_text = load_criteria_text(args.criteria_file)
    articles = fetch_unfiled_articles()
    articles = add_openai_categories(articles, args.openai_model, args.batch_size, criteria_text)
    report = build_output_frame(articles)
    output_path = Path(args.output)
    write_workbook(output_path, report, args.openai_model, criteria_text)
    print(f"Workbook written: {output_path.resolve()}")
    print(f"Rows: {len(report)}")


if __name__ == "__main__":
    main()
