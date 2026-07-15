"""Create a 2x2 human-vs-OpenAI comparison workbook.

Default input:
    reports/CFC_2017_2022_Review_Comparison.xlsx

Default output:
    reports/CFC_2017_2022_2x2_Matrix.xlsx
"""

from __future__ import annotations

import argparse
import re
import tempfile
import time
from pathlib import Path
from urllib.parse import urlparse
from urllib.request import urlretrieve

import pandas as pd
from openpyxl import load_workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter


DEFAULT_INPUT = "reports/CFC_2017_2022_Review_Comparison.xlsx"
DEFAULT_OUTPUT = "reports/CFC_2017_2022_2x2_Matrix.xlsx"
DEFAULT_SHEET = "Review_Comparison"
DEFAULT_SCREENING_HISTORY_URL = (
    "https://docs.google.com/spreadsheets/d/1BUvWcV6XgYiOL3cCrYAHjkccb24C38OK/"
    "edit?usp=sharing&ouid=106518116377917721454&rtpof=true&sd=true"
)

CATEGORIES = [
    "Allergy and Immunology",
    "Cardiology",
    "Dermatology",
    "Development and Behavior",
    "Endocrinology",
    "Gastroenterology",
    "General and Reviews",
    "Genetics",
    "Growth",
    "Gynecology",
    "Neurology",
    "Oncology",
    "Ophthalmology",
    "Orthopedic",
    "Otolaryngology",
    "Pulmonology",
    "Research Studies",
    "Seizures",
    "Treatments",
]

CATEGORY_COLORS = {
    "Allergy and Immunology": "D9EAD3",
    "Cardiology": "F4CCCC",
    "Dermatology": "FCE5CD",
    "Development and Behavior": "D9D2E9",
    "Endocrinology": "D0E0E3",
    "Gastroenterology": "FFF2CC",
    "General and Reviews": "EADCF8",
    "Genetics": "CFE2F3",
    "Growth": "E2F0CB",
    "Gynecology": "FCE4EC",
    "Neurology": "C9DAF8",
    "Oncology": "E6B8AF",
    "Ophthalmology": "D9EAD3",
    "Orthopedic": "EAD7C2",
    "Otolaryngology": "D0E0E3",
    "Pulmonology": "D9EAF7",
    "Research Studies": "E6E6E6",
    "Seizures": "D9D2E9",
    "Treatments": "D5A6BD",
    "Historical Articles": "EFEFEF",
    "Conferences": "D9D9D9",
}


def normalize_text(value: object) -> str:
    text = "" if value is None else str(value)
    text = re.sub(r"<[^>]+>", " ", text)
    text = re.sub(r"[^a-z0-9]+", " ", text.lower())
    return re.sub(r"\s+", " ", text).strip()


def normalize_pmid(value: object) -> str:
    if value is None:
        return ""
    match = re.search(r"\b(\d{4,})\b", str(value))
    return match.group(1) if match else ""


def normalize_doi(value: object) -> str:
    if value is None:
        return ""
    text = str(value).strip().lower().replace("_", " ")
    text = re.sub(r"^https?://(?:dx\.)?doi\.org/", "", text)
    text = re.sub(r"^doi:\s*", "", text)
    return text.strip().rstrip(".")


def find_column(df: pd.DataFrame, candidates: tuple[str, ...]) -> str | None:
    normalized = {normalize_text(column): column for column in df.columns}
    for candidate in candidates:
        key = normalize_text(candidate)
        if key in normalized:
            return normalized[key]
    return None


def materialize_source(source: str) -> Path:
    parsed = urlparse(source)
    if parsed.scheme in {"http", "https"}:
        match = re.search(r"/spreadsheets/d/([^/]+)", source)
        download_url = source
        suffix = ".xlsx"
        if match:
            download_url = f"https://docs.google.com/spreadsheets/d/{match.group(1)}/export?format=xlsx"
        target = Path(tempfile.gettempdir()) / f"cfc_human_screening_{int(time.time())}{suffix}"
        urlretrieve(download_url, target)
        return target
    return Path(source)


def screening_direction(value: object) -> str:
    text = normalize_text(value)
    if not text:
        return ""
    out_terms = (
        "screen out",
        "screened out",
        "exclude",
        "excluded",
        "not approved",
        "not relevant",
        "no",
    )
    in_terms = (
        "screen in",
        "screened in",
        "include",
        "included",
        "approved",
        "relevant",
        "yes",
    )
    if any(term in text for term in out_terms):
        return "out"
    if any(term in text for term in in_terms):
        return "in"
    return ""


def readable_screen(value: str) -> str:
    if value == "in":
        return "Screen in"
    if value == "out":
        return "Screen out"
    return "Needs review"


def load_human_history(source: str | None) -> dict[str, dict[str, dict[str, str]]]:
    empty = {"pmid": {}, "doi": {}, "title": {}}
    if not source:
        return empty

    path = materialize_source(source)
    if not path.exists():
        raise RuntimeError(f"Human screening file was not found: {path}")

    if path.suffix.lower() in {".xlsx", ".xls"}:
        sheets = pd.read_excel(path, sheet_name=None, dtype=str)
        frames = [frame.assign(_history_sheet=name) for name, frame in sheets.items()]
        history_df = pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()
    elif path.suffix.lower() == ".csv":
        history_df = pd.read_csv(path, dtype=str)
        history_df["_history_sheet"] = path.name
    elif path.suffix.lower() == ".tsv":
        history_df = pd.read_csv(path, sep="\t", dtype=str)
        history_df["_history_sheet"] = path.name
    else:
        raise RuntimeError("Human screening history must be .xlsx, .xls, .csv, .tsv, or an accessible Google Sheets URL.")

    if history_df.empty:
        return empty

    pmid_col = find_column(history_df, ("PMID", "PubMed ID", "PubMed_ID", "PMID Number"))
    doi_col = find_column(history_df, ("DOI", "Digital Object Identifier"))
    title_col = find_column(history_df, ("Title", "Article Title", "Article_Title"))
    decision_col = find_column(history_df, ("Eligibility_Decision", "Eligibility Decision", "Decision", "Review_Status", "Review Status"))
    category_col = find_column(history_df, ("Category", "Primary_Category", "Primary Category", "Folder", "Zotero Folder", "Suggested_Labels", "Suggested Labels"))
    notes_col = find_column(history_df, ("Notes", "Reviewer_Notes", "Reviewer Notes", "Rationale", "Comment", "Comments"))

    history = {"pmid": {}, "doi": {}, "title": {}}
    for idx, row in history_df.iterrows():
        pmid = normalize_pmid(row.get(pmid_col)) if pmid_col else ""
        doi = normalize_doi(row.get(doi_col)) if doi_col else ""
        title = normalize_text(row.get(title_col)) if title_col else ""
        decision = str(row.get(decision_col, "") or "").strip() if decision_col else ""
        category = str(row.get(category_col, "") or "").strip() if category_col else ""
        notes = str(row.get(notes_col, "") or "").strip() if notes_col else ""
        record = {
            "decision": decision,
            "screen": screening_direction(decision),
            "category": category,
            "notes": notes,
            "source": f"{path.name}:{row.get('_history_sheet', 'Sheet')}:{idx + 2}",
        }
        if pmid:
            history["pmid"][pmid] = record
        if doi:
            history["doi"][doi] = record
        if title:
            history["title"][title] = record
    return history


def lookup_history(row: pd.Series, history: dict[str, dict[str, dict[str, str]]]) -> dict[str, str] | None:
    pmid = normalize_pmid(row.get("PubMed ID") or row.get("PMID"))
    doi = normalize_doi(row.get("DOI"))
    title = normalize_text(row.get("Article Title") or row.get("Title"))
    if pmid and pmid in history.get("pmid", {}):
        return history["pmid"][pmid]
    if doi and doi in history.get("doi", {}):
        return history["doi"][doi]
    if title and title in history.get("title", {}):
        return history["title"][title]
    return None


def choose_human_category(row: pd.Series, history_match: dict[str, str] | None) -> str:
    if history_match and history_match.get("category"):
        return history_match["category"]
    return str(row.get("If in Zotero - Category", "") or "").strip()


CATEGORY_ALIASES = {
    "allergy": "Allergy and Immunology",
    "allergy and immunology": "Allergy and Immunology",
    "cardio": "Cardiology",
    "cardiology": "Cardiology",
    "derm": "Dermatology",
    "dermatology": "Dermatology",
    "development": "Development and Behavior",
    "development and behavior": "Development and Behavior",
    "behavior": "Development and Behavior",
    "endocrinology": "Endocrinology",
    "gastro": "Gastroenterology",
    "gastroenterology": "Gastroenterology",
    "general": "General and Reviews",
    "general_and_reviews": "General and Reviews",
    "general and reviews": "General and Reviews",
    "reviews": "General and Reviews",
    "genetic": "Genetics",
    "genetics": "Genetics",
    "growth": "Growth",
    "gynecology": "Gynecology",
    "neuro": "Neurology",
    "neuology": "Neurology",
    "neurology": "Neurology",
    "oncology": "Oncology",
    "ophthalmology": "Ophthalmology",
    "orthopedic": "Orthopedic",
    "orthopedics": "Orthopedic",
    "ent": "Otolaryngology",
    "otolaryngology ent": "Otolaryngology",
    "otolaryngology_ent": "Otolaryngology",
    "otolaryngology": "Otolaryngology",
    "pulmonology": "Pulmonology",
    "research": "Research Studies",
    "research studies": "Research Studies",
    "seizure": "Seizures",
    "seizures": "Seizures",
    "treatment": "Treatments",
    "treatments": "Treatments",
    "exclude": "Excluded",
    "excluded": "Excluded",
    "exclusion": "Excluded",
}


def split_categories(value: object) -> set[str]:
    if value is None or pd.isna(value):
        return set()
    text = str(value).strip().lower()
    if not text:
        return set()
    protected = {
        "allergy and immunology": "allergy_immunology",
        "development and behavior": "development_behavior",
        "general and reviews": "general_reviews",
    }
    for phrase, token in protected.items():
        text = text.replace(phrase, token)
    text = re.sub(r"\badditional\s*:\s*", ";", text)
    text = text.replace("&", ";")
    text = re.sub(r"\band\b", ";", text)
    text = text.replace(",", ";").replace("/", ";").replace("|", ";")
    for phrase, token in protected.items():
        text = text.replace(token, phrase)
    parts = [re.sub(r"\s+", " ", part).strip(" .;:?") for part in text.split(";")]
    categories: set[str] = set()
    for part in parts:
        if not part:
            continue
        if part in CATEGORY_ALIASES:
            categories.add(CATEGORY_ALIASES[part])
            continue
        for alias, category in CATEGORY_ALIASES.items():
            if re.search(rf"\b{re.escape(alias)}\b", part):
                categories.add(category)
    return {category for category in categories if category in CATEGORIES or category == "Excluded"}


def is_exclusion_category(value: object) -> bool:
    return "Excluded" in split_categories(value)


def zotero_screen_from_category(value: object) -> str:
    text = str(value or "").strip()
    if not text or normalize_text(text) in {"not in zotero", "unfiled"}:
        return ""
    if is_exclusion_category(text):
        return "out"
    if split_categories(text):
        return "in"
    return ""


def enrich_review_rows(review_df: pd.DataFrame, history: dict[str, dict[str, dict[str, str]]]) -> pd.DataFrame:
    rows = []
    for _, row in review_df.iterrows():
        match = lookup_history(row, history)
        openai_screen = screening_direction(row.get("OpenAI Screening In or Out") or row.get("Relevance"))
        zotero_category = str(row.get("If in Zotero - Category", "") or "").strip()
        human_screen = match.get("screen", "") if match else ""
        if not human_screen:
            human_screen = zotero_screen_from_category(zotero_category)
        human_category = choose_human_category(row, match)
        if is_exclusion_category(human_category):
            human_category = ""
        openai_category = str(row.get("OpenAI Assigned Category", "") or "").strip()
        if openai_screen == "out":
            openai_category = ""
        human_categories = split_categories(human_category)
        openai_categories = split_categories(openai_category)
        matrix_cell = ""
        if human_screen == "in" and openai_screen == "in":
            matrix_cell = "TP"
        elif human_screen == "out" and openai_screen == "in":
            matrix_cell = "FP"
        elif human_screen == "in" and openai_screen == "out":
            matrix_cell = "FN"
        elif human_screen == "out" and openai_screen == "out":
            matrix_cell = "TN"

        rows.append(
            {
                "PubMed ID": row.get("PubMed ID", ""),
                "Article Title": row.get("Article Title", ""),
                "Date": row.get("Date", ""),
                "Human Found In Sheets": "Yes" if match else "No",
                "Human Screening": readable_screen(human_screen) if human_screen else "",
                "OpenAI Screening": readable_screen(openai_screen) if openai_screen else "",
                "Screening Matrix Cell": matrix_cell or "Not counted",
                "Screening Match?": "Yes" if human_screen and openai_screen and human_screen == openai_screen else ("No" if human_screen and openai_screen else ""),
                "Human Category": human_category,
                "OpenAI Category": openai_category,
                "Human Categories Normalized": "; ".join(category for category in CATEGORIES if category in human_categories),
                "OpenAI Categories Normalized": "; ".join(category for category in CATEGORIES if category in openai_categories),
                "Category Exact Match?": "Yes" if human_categories and openai_categories and human_categories == openai_categories else ("No" if human_categories and openai_categories else ""),
                "Category Match?": "Yes" if human_categories and openai_categories and bool(human_categories & openai_categories) else ("No" if human_categories and openai_categories else ""),
                "AI Subset of Human?": "Yes" if openai_categories and human_categories and openai_categories.issubset(human_categories) else ("No" if openai_categories and human_categories else ""),
                "Human Subset of AI?": "Yes" if openai_categories and human_categories and human_categories.issubset(openai_categories) else ("No" if openai_categories and human_categories else ""),
                "Found in Zotero?": row.get("Found in Zotero?", ""),
                "Zotero Category": zotero_category,
                "Human Source": match.get("source", "") if match else "",
                "Human Notes": match.get("notes", "") if match else "",
            }
        )
    return pd.DataFrame(rows)


def build_overall_2x2(details: pd.DataFrame) -> pd.DataFrame:
    counts = details["Screening Matrix Cell"].value_counts()
    tp = int(counts.get("TP", 0))
    fp = int(counts.get("FP", 0))
    fn = int(counts.get("FN", 0))
    tn = int(counts.get("TN", 0))
    total = tp + fp + fn + tn
    precision = tp / (tp + fp) if tp + fp else ""
    recall = tp / (tp + fn) if tp + fn else ""
    f1 = 2 * precision * recall / (precision + recall) if precision != "" and recall != "" and precision + recall else ""
    return pd.DataFrame(
        [
            {
                "Comparison": "Human screening vs OpenAI screening",
                "True_Positive": tp,
                "False_Positive": fp,
                "False_Negative": fn,
                "True_Negative": tn,
                "Total_Counted": total,
                "Agreement_Rate": (tp + tn) / total if total else 0,
                "Precision": precision,
                "Recall": recall,
                "F1_Score": f1,
                "Sensitivity": recall,
                "Specificity": tn / (tn + fp) if tn + fp else "",
                "Needs_Review": fp + fn,
                "Not_Counted": int((details["Screening Matrix Cell"] == "Not counted").sum()),
            }
        ]
    )


def build_screening_matrix_table(details: pd.DataFrame) -> pd.DataFrame:
    counts = details["Screening Matrix Cell"].value_counts()
    tp = int(counts.get("TP", 0))
    fp = int(counts.get("FP", 0))
    fn = int(counts.get("FN", 0))
    tn = int(counts.get("TN", 0))
    return pd.DataFrame(
        [
            {
                "Human \\ OpenAI": "Human Screen In",
                "OpenAI Screen In": tp,
                "OpenAI Screen Out": fn,
                "Row Total": tp + fn,
            },
            {
                "Human \\ OpenAI": "Human Screen Out",
                "OpenAI Screen In": fp,
                "OpenAI Screen Out": tn,
                "Row Total": fp + tn,
            },
            {
                "Human \\ OpenAI": "Column Total",
                "OpenAI Screen In": tp + fp,
                "OpenAI Screen Out": fn + tn,
                "Row Total": tp + fp + fn + tn,
            },
        ]
    )


def build_category_2x2(details: pd.DataFrame) -> pd.DataFrame:
    rows = []
    for category in CATEGORIES:
        tp = fp = fn = tn = 0
        for _, row in details.iterrows():
            human_yes = category in split_categories(row.get("Human Category"))
            openai_yes = category in split_categories(row.get("OpenAI Category"))
            if human_yes and openai_yes:
                tp += 1
            elif not human_yes and openai_yes:
                fp += 1
            elif human_yes and not openai_yes:
                fn += 1
            else:
                tn += 1
        total = tp + fp + fn + tn
        precision = tp / (tp + fp) if tp + fp else ""
        recall = tp / (tp + fn) if tp + fn else ""
        f1 = 2 * precision * recall / (precision + recall) if precision != "" and recall != "" and precision + recall else ""
        specificity = tn / (tn + fp) if tn + fp else ""
        rows.append(
            {
                "Category": category,
                "True_Positive": tp,
                "False_Positive": fp,
                "False_Negative": fn,
                "True_Negative": tn,
                "Total": total,
                "Agreement_Rate": (tp + tn) / total if total else 0,
                "Precision": precision,
                "Recall": recall,
                "F1_Score": f1,
                "Specificity": specificity,
                "Needs_Review": fp + fn,
            }
        )
    return pd.DataFrame(rows)


def build_category_overlap_summary(details: pd.DataFrame) -> pd.DataFrame:
    comparable = details[
        details["Human Categories Normalized"].astype(str).str.strip().ne("")
        & details["OpenAI Categories Normalized"].astype(str).str.strip().ne("")
    ].copy()
    denominator = len(comparable)

    def percent_yes(column: str) -> float | str:
        if denominator == 0:
            return ""
        return float((comparable[column] == "Yes").sum() / denominator)

    return pd.DataFrame(
        [
            {
                "Metric": "Exact category-list match",
                "Percent": percent_yes("Category Exact Match?"),
                "Numerator": int((comparable["Category Exact Match?"] == "Yes").sum()) if denominator else 0,
                "Denominator": denominator,
                "Meaning": "AI and human category lists are identical after normalization.",
            },
            {
                "Metric": "Any category overlap",
                "Percent": percent_yes("Category Match?"),
                "Numerator": int((comparable["Category Match?"] == "Yes").sum()) if denominator else 0,
                "Denominator": denominator,
                "Meaning": "AI and human share at least one category; this is fairer for multi-category articles.",
            },
            {
                "Metric": "AI subset of human",
                "Percent": percent_yes("AI Subset of Human?"),
                "Numerator": int((comparable["AI Subset of Human?"] == "Yes").sum()) if denominator else 0,
                "Denominator": denominator,
                "Meaning": "AI chose only categories that humans also chose, even if it missed an additional human category.",
            },
            {
                "Metric": "Human subset of AI",
                "Percent": percent_yes("Human Subset of AI?"),
                "Numerator": int((comparable["Human Subset of AI?"] == "Yes").sum()) if denominator else 0,
                "Denominator": denominator,
                "Meaning": "AI covered all human categories, even if it added an extra category.",
            },
        ]
    )


def write_workbook(
    output_path: Path,
    overall: pd.DataFrame,
    category_summary: pd.DataFrame,
    details: pd.DataFrame,
    input_path: Path,
    history_source: str | None,
) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    disagreements = details[
        (details["Screening Match?"] == "No")
        | (details["Category Match?"] == "No")
        | (details["Screening Matrix Cell"].isin(["FP", "FN"]))
    ].copy()
    missing_human = details[details["Human Found In Sheets"] == "No"].copy()
    matrix_table = build_screening_matrix_table(details)
    overlap_summary = build_category_overlap_summary(details)
    run_log = pd.DataFrame(
        [
            {"Field": "Input workbook", "Value": str(input_path)},
            {"Field": "Input sheet", "Value": DEFAULT_SHEET},
            {"Field": "Human screening source", "Value": history_source or ""},
            {"Field": "Output workbook", "Value": str(output_path)},
            {"Field": "Generated", "Value": time.ctime()},
            {"Field": "TP", "Value": "Human screened in and OpenAI screened in."},
            {"Field": "FP", "Value": "Human screened out and OpenAI screened in."},
            {"Field": "FN", "Value": "Human screened in and OpenAI screened out."},
            {"Field": "TN", "Value": "Human screened out and OpenAI screened out."},
            {"Field": "Category normalization", "Value": "Uses human-aligned category aliases; splits semicolon/comma/slash/and lists; excludes Historical Articles and Conferences."},
            {"Field": "Criteria alignment", "Value": "Aligned with stricter CFC rules: meaningful CFC discussion required; passing mentions/figure-only/reference-only CFC allusions should be excluded upstream."},
            {"Field": "Zotero Exclusion rule", "Value": "Items found in Zotero's Exclusion/Excluded folder are treated as human screen-out, not as a clinical category."},
            {"Field": "Zotero category rule", "Value": "When Google Sheets history is missing, a non-exclusion Zotero folder is treated as the human category/screen-in evidence."},
        ]
    )

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        matrix_table.to_excel(writer, sheet_name="Screening_Matrix", index=False)
        overlap_summary.to_excel(writer, sheet_name="Category_Overlap", index=False)
        overall.to_excel(writer, sheet_name="Overall_2x2", index=False)
        category_summary.to_excel(writer, sheet_name="Category_2x2", index=False)
        details.to_excel(writer, sheet_name="Article_Details", index=False)
        disagreements.to_excel(writer, sheet_name="Disagreements", index=False)
        missing_human.to_excel(writer, sheet_name="Missing_Human_Sheet", index=False)
        run_log.to_excel(writer, sheet_name="Run_Log", index=False)
    format_workbook(output_path)


def format_workbook(path: Path) -> None:
    wb = load_workbook(path)
    match_fill = PatternFill(start_color="D9EAD3", end_color="D9EAD3", fill_type="solid")
    mismatch_fill = PatternFill(start_color="F4CCCC", end_color="F4CCCC", fill_type="solid")
    review_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")

    for ws in wb.worksheets:
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        for cell in ws[1]:
            cell.font = Font(bold=True)
        headers = [cell.value for cell in ws[1]]
        max_row = max(ws.max_row, 2)
        for header in ("Screening Match?", "Category Match?"):
            if header in headers:
                letter = get_column_letter(headers.index(header) + 1)
                cell_range = f"{letter}2:{letter}{max_row}"
                ws.conditional_formatting.add(cell_range, FormulaRule(formula=[f'${letter}2="Yes"'], fill=match_fill))
                ws.conditional_formatting.add(cell_range, FormulaRule(formula=[f'${letter}2="No"'], fill=mismatch_fill))
        if "Screening Matrix Cell" in headers:
            letter = get_column_letter(headers.index("Screening Matrix Cell") + 1)
            cell_range = f"{letter}2:{letter}{max_row}"
            ws.conditional_formatting.add(cell_range, FormulaRule(formula=[f'OR(${letter}2="FP",${letter}2="FN")'], fill=mismatch_fill))
            ws.conditional_formatting.add(cell_range, FormulaRule(formula=[f'${letter}2="Not counted"'], fill=review_fill))
        if "Category" in headers:
            col = headers.index("Category") + 1
            for row in range(2, max_row + 1):
                category = str(ws.cell(row=row, column=col).value or "")
                color = CATEGORY_COLORS.get(category, "FFFFFF")
                ws.cell(row=row, column=col).fill = PatternFill(start_color=color, end_color=color, fill_type="solid")
        widths = {
            "PubMed ID": 12,
            "Article Title": 62,
            "Date": 14,
            "Human Screening": 18,
            "OpenAI Screening": 18,
            "Screening Matrix Cell": 20,
            "Screening Match?": 18,
            "Human Category": 30,
            "OpenAI Category": 26,
            "Category Match?": 18,
            "Zotero Category": 30,
            "Human Source": 36,
            "Human Notes": 40,
        }
        for index, header in enumerate(headers, start=1):
            ws.column_dimensions[get_column_letter(index)].width = widths.get(str(header), 18)
    wb.save(path)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Create a 2x2 human-vs-OpenAI comparison workbook.")
    parser.add_argument("--input", default=DEFAULT_INPUT, help="Review comparison workbook to analyze.")
    parser.add_argument("--output", default=DEFAULT_OUTPUT, help="2x2 matrix workbook to create.")
    parser.add_argument("--sheet", default=DEFAULT_SHEET, help="Sheet in the review workbook to read.")
    parser.add_argument(
        "--screening-history",
        default=DEFAULT_SCREENING_HISTORY_URL,
        help="Human screening .xlsx/.csv/.tsv file or accessible Google Sheets URL.",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    input_path = Path(args.input)
    output_path = Path(args.output)
    review_df = pd.read_excel(input_path, sheet_name=args.sheet, dtype=str)
    history = load_human_history(args.screening_history)
    details = enrich_review_rows(review_df, history)
    overall = build_overall_2x2(details)
    category_summary = build_category_2x2(details)
    write_workbook(output_path, overall, category_summary, details, input_path, args.screening_history)
    print(f"2x2 comparison workbook written: {output_path.resolve()}")


if __name__ == "__main__":
    main()
