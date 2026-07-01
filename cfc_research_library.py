"""Update a CFC syndrome literature review workbook from PubMed and Zotero.

This replaces the original Colab-only notebook with a reusable local program:

    python cfc_research_library.py --category Dermatology --output reports/Dermatology_Master_Review_Report.xlsx

Credentials are read from environment variables so API keys are not stored in code.
"""

from __future__ import annotations

import argparse
import json
import os
import re
import tempfile
import textwrap
import time
from dataclasses import asdict, dataclass
from datetime import date
from pathlib import Path
from typing import Any, Iterable, TypeAlias
from urllib.parse import urlparse
from urllib.request import urlretrieve

pd = None
Entrez = None
zotero = None
SentenceTransformer = None
util = None

DataFrame: TypeAlias = Any


CFC_TERMS = (
    "cardiofaciocutaneous",
    "cardio-facio-cutaneous",
    "cardio-facio-cutaneous syndrome",
    "cardiofaciocutaneous syndrome",
    "CFC syndrome",
)

CFC_GENES = ("BRAF", "MAP2K1", "MAP2K2", "KRAS")

REVIEWER_DECISIONS = ("Needs reviewed", "Approved", "Not approved")

DEFAULT_SCREENING_HISTORY_URL = (
    "https://docs.google.com/spreadsheets/d/1BUvWcV6XgYiOL3cCrYAHjkccb24C38OK/"
    "edit?usp=sharing&ouid=106518116377917721454&rtpof=true&sd=true"
)

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

OTHER_RASOPATHY_SIGNALS = (
    "costello syndrome",
    "hras-positive",
    "hras positive",
    "noonan syndrome",
    "legius syndrome",
    "neurofibromatosis type 1",
)

STRONG_CFC_SIGNALS = (
    "cardiofaciocutaneous syndrome",
    "cardio-facio-cutaneous syndrome",
    "cfc syndrome",
    "cardiofaciocutaneous",
    "cardio-facio-cutaneous",
)

GENERAL_INCLUSION_CRITERIA = [
    "Directly discusses cardiofaciocutaneous syndrome, CFC syndrome, or a CFC-specific subgroup within RASopathy research.",
    "Includes confirmed or strongly suspected CFC cases, or uses CFC-associated RAS/MAPK variants to model relevant clinical biology.",
    "Provides clinical, genetic, mechanistic, management, or review evidence relevant to the selected Zotero library section.",
]

GENERAL_EXCLUSION_CRITERIA = [
    "Focuses only on Noonan, Costello, or other RASopathies without CFC-specific data or interpretation.",
    "Mentions RAS/MAPK biology without a clear CFC clinical, genetic, or phenotypic connection.",
    "Is unrelated to humans, CFC-relevant model systems, clinical care, diagnosis, treatment, or literature synthesis.",
]


@dataclass(frozen=True)
class LibrarySection:
    name: str
    description: str
    inclusion: str
    exclusion: str
    query: str


def cfc_clause() -> str:
    return '(cardiofaciocutaneous[All Fields] OR "cardio-facio-cutaneous"[All Fields] OR "CFC syndrome"[All Fields] OR "RASopathies"[MeSH Terms])'


SECTIONS: dict[str, LibrarySection] = {
    "Allergy and Immunology": LibrarySection(
        "Allergy and Immunology",
        "Immune dysregulation, hypersensitivity, eczema, allergic rhinitis, and recurrent infections in CFC and related RASopathies.",
        "Papers that directly investigate immune dysregulation, allergic disease, or inflammatory responses in individuals with confirmed or strongly suspected CFC syndrome, including RAS/MAPK mutations specifically linked to CFC.",
        "Studies focused solely on other RASopathies without CFC-specific data, generalized immunology papers without clinical relevance to CFC, or pathway studies lacking immune or allergy outcomes.",
        f'(("Hypersensitivity"[MeSH Terms] OR "Immunologic Deficiency Syndromes"[MeSH Terms] OR atopy[All Fields] OR allergic[All Fields] OR immune[All Fields]) AND {cfc_clause()})',
    ),
    "Cardiology": LibrarySection(
        "Cardiology",
        "Cardiac structure, hypertrophic cardiomyopathy, rhythm abnormalities, outflow obstruction, and contractile function in CFC.",
        "Research examining cardiac structure, function, or electrophysiology in CFC syndrome, including mechanistic studies using BRAF or MAPK mutations known to cause CFC.",
        "Cardiac studies on Noonan, Costello, or other RASopathies without a CFC subgroup, or broad cardiomyopathy papers not tied to CFC-relevant mutations.",
        f'(("Heart Defects, Congenital"[MeSH Terms] OR "Cardiomyopathies"[MeSH Terms] OR arrhythmia[All Fields] OR hypertrophic[All Fields]) AND {cfc_clause()})',
    ),
    "Dermatology": LibrarySection(
        "Dermatology",
        "Skin, hair, and nail findings in CFC, including eczema-like rashes, hyperkeratosis, keratosis pilaris, sparse or curly hair, and fragile nails.",
        "Studies describing skin, hair, or nail abnormalities in CFC syndrome, or mechanistic dermatologic work using CFC-associated RAS/MAPK mutations.",
        "Dermatology papers on other RASopathies without CFC-specific findings, or general dermatologic pathway studies not linked to CFC phenotypes.",
        f'(("Skin Manifestations"[MeSH Terms] OR "Skin Diseases, Genetic"[MeSH Terms] OR dermatology[All Fields] OR cutaneous[All Fields] OR hair[All Fields] OR nail[All Fields]) AND {cfc_clause()})',
    ),
    "Development and Behavior": LibrarySection(
        "Development and Behavior",
        "Cognitive, motor, speech, behavioral, and growth-related developmental outcomes in CFC.",
        "Papers addressing neurodevelopment, cognition, motor milestones, or behavioral phenotypes in individuals with CFC, including mechanistic work connecting RAS/MAPK dysregulation to developmental outcomes.",
        "Behavioral or developmental studies on other RASopathies without CFC representation, or broad neurodevelopmental research not tied to CFC-relevant mutations.",
        f'(("Developmental Disabilities"[MeSH Terms] OR "Behavioral Symptoms"[MeSH Terms] OR "Intellectual Disability"[MeSH Terms] OR cognition[All Fields] OR behavior[All Fields]) AND {cfc_clause()})',
    ),
    "Endocrinology": LibrarySection(
        "Endocrinology",
        "Hormonal, metabolic, and organ-related abnormalities associated with CFC syndrome and other RASopathies, including endocrine function, growth regulation, pubertal timing, gastrointestinal and renal organ development, feeding difficulties, failure to thrive, growth hormone deficiency, delayed or atypical puberty, renal anomalies, and electrolyte imbalances.",
        "Studies examining hormonal regulation, growth hormone function, pubertal development, metabolic abnormalities, or endocrine organ involvement in individuals with CFC; mechanistic work using CFC-associated mutations to model endocrine dysfunction.",
        "Endocrine studies focused solely on other RASopathies without CFC data; general metabolic or hormonal pathway papers not linked to CFC-specific clinical features.",
        f'(("Endocrine System Diseases"[MeSH Terms] OR hormone[All Fields] OR renal[All Fields] OR metabolic[All Fields] OR puberty[All Fields] OR pubertal[All Fields] OR "growth hormone"[All Fields] OR electrolyte[All Fields]) AND {cfc_clause()})',
    ),
    "Gastroenterology": LibrarySection(
        "Gastroenterology",
        "Feeding difficulties, reflux, constipation, motility disorders, nutrition, and failure to thrive in CFC.",
        "Studies investigating feeding difficulties, reflux, motility disorders, or gastrointestinal dysfunction in CFC syndrome, including mechanistic work using CFC-relevant RAS/MAPK mutations.",
        "GI studies on Noonan, Costello, or other RASopathies without CFC-specific findings, or general GI physiology papers not tied to CFC.",
        f'(("Gastrointestinal Diseases"[MeSH Terms] OR "Feeding and Eating Disorders"[MeSH Terms] OR reflux[All Fields] OR constipation[All Fields] OR motility[All Fields]) AND {cfc_clause()})',
    ),
    "General and Reviews": LibrarySection(
        "General and Reviews",
        "Foundational, overview, diagnostic, clinical-spectrum, and review papers about CFC within the RASopathy family.",
        "Foundational papers, clinical overviews, and reviews that explicitly discuss CFC syndrome, its diagnostic criteria, phenotype spectrum, or its place within the RASopathy family.",
        "Broad RASopathy reviews that do not meaningfully address CFC, or historical papers lacking clinical or genetic relevance to CFC.",
        f'({cfc_clause()} AND (Review[Publication Type] OR guideline[Publication Type] OR diagnosis[All Fields] OR phenotype[All Fields]))',
    ),
    "Genetics": LibrarySection(
        "Genetics",
        "Molecular etiology, mutation spectrum, genotype-phenotype correlation, and diagnostic genetics of CFC.",
        "Studies that present genetic analyses, mutation identification, or molecular characterization of CFC syndrome, including work that uses CFC-associated RAS/MAPK mutations to explore developmental mechanisms.",
        "Genetic studies focused solely on other RASopathies without CFC data, or general molecular pathway papers not directly linked to CFC-specific mutations or phenotypes.",
        f'(("Genetic Phenomena"[MeSH Terms] OR genetics[Subheading] OR mutation[All Fields] OR BRAF[All Fields] OR MAP2K1[All Fields] OR MAP2K2[All Fields] OR KRAS[All Fields]) AND {cfc_clause()})',
    ),
    "Growth": LibrarySection(
        "Growth",
        "Physical growth patterns, stature, skeletal maturation, and pubertal development across CFC syndrome and the broader Noonan-spectrum RASopathies, including growth hormone signaling, bone age progression, linear growth velocity, timing of puberty, short stature, delayed or disproportionate growth, growth hormone deficiency, and atypical pubertal onset.",
        "Studies addressing growth, stature, bone age, or pubertal development in CFC or Noonan-spectrum RASopathies; research using CFC-associated or Noonan-associated RAS/MAPK mutations to model growth hormone or pubertal regulation.",
        "Growth studies unrelated to RAS/MAPK signaling or lacking CFC/Noonan-spectrum relevance; general endocrinology papers without developmental or pubertal data tied to RASopathies.",
        '((growth[All Fields] OR stature[All Fields] OR "short stature"[All Fields] OR "bone age"[All Fields] OR "growth hormone"[All Fields] OR puberty[All Fields] OR pubertal[All Fields] OR "linear growth"[All Fields]) AND (cardiofaciocutaneous[All Fields] OR "cardio-facio-cutaneous"[All Fields] OR "CFC syndrome"[All Fields] OR "Noonan Syndrome"[MeSH Terms] OR "noonan spectrum"[All Fields] OR RASopathies[MeSH Terms]))',
    ),
    "Gynecology": LibrarySection(
        "Gynecology",
        "Reproductive, pubertal, menstrual, genital tract, and reproductive endocrine findings in CFC.",
        "Studies describing gynecologic, reproductive, or pubertal findings in individuals with CFC or using CFC-associated mutations to model reproductive endocrine function.",
        "Papers focused solely on other RASopathies without CFC data, or general reproductive endocrinology studies not tied to CFC-specific mutations or phenotypes.",
        f'(("Gynecology"[MeSH Terms] OR puberty[All Fields] OR menstrual[All Fields] OR reproductive[All Fields] OR ovarian[All Fields] OR uterine[All Fields]) AND {cfc_clause()})',
    ),
    "Neurology": LibrarySection(
        "Neurology",
        "Brain structure, hypotonia, motor delay, neuroimaging, coordination, and neurological physiology in CFC.",
        "Studies that document neurological findings, neuroimaging results, tone or motor abnormalities, or neurodevelopmental physiology in individuals with CFC, including mechanistic work using CFC-associated RAS/MAPK mutations.",
        "Neurologic studies focused solely on other RASopathies without CFC data, or general neuroscience papers not tied to CFC-specific phenotypes.",
        f'(("Nervous System Malformations"[MeSH Terms] OR "Neurodevelopmental Disorders"[MeSH Terms] OR hypotonia[All Fields] OR neuroimaging[All Fields] OR brain[All Fields]) AND {cfc_clause()})',
    ),
    "Oncology": LibrarySection(
        "Oncology",
        "Cancer risk, tumor formation, oncogenic RAS/MAPK mechanisms, and malignancy surveillance in CFC.",
        "Studies examining tumor development, cancer risk, or oncogenic mechanisms in individuals with CFC or using CFC-associated RAS/MAPK mutations to model malignancy.",
        "Oncology papers focused solely on other RASopathies without CFC data, or general cancer pathway studies not directly linked to CFC-specific mutations or phenotypes.",
        f'(("Neoplasms"[MeSH Terms] OR "Neoplastic Syndromes, Hereditary"[MeSH Terms] OR tumor[All Fields] OR cancer[All Fields] OR leukemia[All Fields]) AND {cfc_clause()})',
    ),
    "Ophthalmology": LibrarySection(
        "Ophthalmology",
        "Ocular structure, visual function, ptosis, strabismus, refractive errors, optic nerve findings, and corneal abnormalities in CFC.",
        "Research documenting ocular structure, function, or visual outcomes in CFC, or mechanistic studies using CFC-associated mutations to model eye development.",
        "Ophthalmology papers on other RASopathies without CFC representation, or general eye development studies not linked to CFC.",
        f'(("Eye Abnormalities"[MeSH Terms] OR "Eye Manifestations"[MeSH Terms] OR ophthalmology[All Fields] OR ocular[All Fields] OR strabismus[All Fields] OR nystagmus[All Fields]) AND {cfc_clause()})',
    ),
    "Orthopedic": LibrarySection(
        "Orthopedic",
        "Skeletal development, bone density, joint laxity, mineralization, posture, and mobility findings in CFC.",
        "Studies examining skeletal development, bone density, joint structure, or orthopedic manifestations in individuals with CFC, including mechanistic work using CFC-associated RAS/MAPK mutations.",
        "Orthopedic or bone metabolism studies focused solely on other RASopathies without CFC data, or general musculoskeletal research not tied to CFC-specific phenotypes.",
        f'(("Musculoskeletal Abnormalities"[MeSH Terms] OR skeletal[All Fields] OR bone[All Fields] OR orthopedic[All Fields] OR joint[All Fields]) AND {cfc_clause()})',
    ),
    "Otolaryngology": LibrarySection(
        "Otolaryngology",
        "Ear, nose, throat, hearing, airway, swallowing, sinus, laryngeal, and craniofacial ENT findings in CFC.",
        "Studies examining hearing, airway, swallowing, or ENT-related structural findings in CFC, including those using CFC-associated mutations to model craniofacial development.",
        "ENT studies on other RASopathies without CFC-specific data, or general airway/ENT research not tied to CFC phenotypes.",
        f'(("Otorhinolaryngologic Diseases"[MeSH Terms] OR hearing[All Fields] OR airway[All Fields] OR otitis[All Fields] OR laryngeal[All Fields] OR swallowing[All Fields]) AND {cfc_clause()})',
    ),
    "Pulmonology": LibrarySection(
        "Pulmonology",
        "Airway malformations, chronic lung disease, recurrent respiratory infections, aspiration, wheezing, cough, and sleep-disordered breathing in CFC.",
        "Research addressing lung function, airway structure, respiratory infections, or breathing disorders in individuals with CFC.",
        "Pulmonary studies on other RASopathies without CFC data, or general respiratory physiology papers not linked to CFC.",
        f'(("Respiratory Tract Diseases"[MeSH Terms] OR pulmonology[All Fields] OR respiratory[All Fields] OR lung[All Fields] OR apnea[All Fields] OR aspiration[All Fields]) AND {cfc_clause()})',
    ),
    "Research Studies": LibrarySection(
        "Research Studies",
        "Original experimental, clinical, molecular, developmental, and translational studies forming the scientific backbone for CFC pathophysiology.",
        "Studies presenting original data, experimental models, or clinical analyses directly involving individuals with CFC or using CFC-associated RAS/MAPK mutations to investigate developmental or molecular mechanisms.",
        "Research focused solely on other RASopathies without CFC representation, or general pathway studies not linked to CFC-specific clinical or genetic findings.",
        f'({cfc_clause()} NOT Review[Publication Type])',
    ),
    "Seizures": LibrarySection(
        "Seizures",
        "Epilepsy, seizure types, EEG findings, epileptic encephalopathy, medication response, and seizure mechanisms in CFC.",
        "Research focused on seizure types, EEG findings, epilepsy mechanisms, or seizure management in individuals with CFC, including studies using CFC-associated mutations to model epileptogenesis.",
        "Epilepsy studies on other RASopathies without CFC representation, or general seizure research not linked to CFC-specific genetic or clinical features.",
        f'(("Epilepsy"[MeSH Terms] OR "Seizures"[MeSH Terms] OR seizure[All Fields] OR epilepsy[All Fields] OR EEG[All Fields]) AND {cfc_clause()})',
    ),
    "Treatments": LibrarySection(
        "Treatments",
        "Therapeutic approaches, management strategies, targeted RAS/MAPK interventions, supportive therapies, and multidisciplinary care in CFC.",
        "Studies assessing treatments, interventions, or management strategies specifically for CFC, or therapeutic research using CFC-associated mutations to test targeted approaches.",
        "Treatment papers focused solely on other RASopathies without CFC data, or general therapeutic studies not applicable to CFC-specific clinical needs.",
        f'(("Therapeutics"[MeSH Terms] OR treatment[All Fields] OR therapy[All Fields] OR management[All Fields] OR inhibitor[All Fields] OR intervention[All Fields]) AND {cfc_clause()})',
    ),
    "Historical Articles": LibrarySection(
        "Historical Articles",
        "Historical, biography, and early descriptive papers relevant to the discovery and naming of CFC.",
        "Historical papers that clarify the origin, early recognition, diagnostic framing, or naming of CFC syndrome.",
        "Historical material that does not address CFC or CFC-relevant RASopathy classification.",
        f'({cfc_clause()} AND (Historical Article[Publication Type] OR Biography[Publication Type]))',
    ),
    "Conferences": LibrarySection(
        "Conferences",
        "Conference abstracts, proceedings, and congress papers that mention CFC.",
        "Conference material with explicit CFC clinical, genetic, mechanistic, or treatment relevance.",
        "Conference material about other RASopathies or general biology without CFC-specific information.",
        f'({cfc_clause()} AND Congresses[Publication Type])',
    ),
}


SECTION_ALIASES = {
    "ENT": "Otolaryngology",
    "Otolaryngology ENT": "Otolaryngology",
    "Ophtalmology": "Ophthalmology",
}


FINAL_COLUMNS = [
    "Primary_Category",
    "Matched_Categories",
    "Reviewer_Decision",
    "PMID",
    "Title",
    "Authors",
    "Journal",
    "Publication_Year",
    "Publication_Date",
    "DOI",
    "Abstract",
    "Suggested_Labels",
    "Eligibility_Decision",
    "Eligibility_Rationale",
    "History_Match",
    "History_Decision",
    "History_Source",
    "Section_Inclusion_Criteria",
    "Section_Exclusion_Criteria",
    "Found_in_Zotero",
    "Review_Status",
    "Reviewer_Notes",
    "PubMed_URL",
    "Last_Seen",
]


def normalize_category(name: str) -> str:
    cleaned = name.strip()
    return SECTION_ALIASES.get(cleaned, cleaned)


def load_env_file(path: Path = Path(".env")) -> None:
    """Load simple KEY=VALUE entries from .env without overriding existing variables."""
    if not path.exists():
        return
    for raw_line in path.read_text(encoding="utf-8").splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        key = key.strip()
        value = value.strip().strip('"').strip("'")
        if key and key not in os.environ:
            os.environ[key] = value


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


def require_env(name: str) -> str:
    value = os.getenv(name)
    if not value:
        raise RuntimeError(f"Missing required environment variable: {name}")
    return value


def load_runtime_dependencies() -> None:
    global pd, Entrez, zotero, SentenceTransformer, util

    try:
        import pandas as pandas_module
        from Bio import Entrez as entrez_module
        from pyzotero import zotero as zotero_module
        from sentence_transformers import SentenceTransformer as sentence_transformer_class
        from sentence_transformers import util as sentence_transformer_util
    except ModuleNotFoundError as exc:
        missing = exc.name or "a required package"
        raise RuntimeError(
            f"Missing Python package '{missing}'. Install dependencies with: python -m pip install -r requirements.txt"
        ) from exc

    pd = pandas_module
    Entrez = entrez_module
    zotero = zotero_module
    SentenceTransformer = sentence_transformer_class
    util = sentence_transformer_util


def materialize_history_source(source: str) -> Path:
    parsed = urlparse(source)
    if parsed.scheme in {"http", "https"}:
        match = re.search(r"/spreadsheets/d/([^/]+)", source)
        download_url = source
        suffix = ".xlsx"
        if match:
            download_url = f"https://docs.google.com/spreadsheets/d/{match.group(1)}/export?format=xlsx"
        target = Path(tempfile.gettempdir()) / f"cfc_screening_history_{int(time.time())}{suffix}"
        urlretrieve(download_url, target)
        return target
    return Path(source)


def find_column(df: DataFrame, candidates: tuple[str, ...]) -> str | None:
    normalized = {normalize_text(column): column for column in df.columns}
    for candidate in candidates:
        key = normalize_text(candidate)
        if key in normalized:
            return normalized[key]
    return None


def load_screening_history(source: str | None) -> dict[str, dict[str, dict[str, str]]]:
    empty = {"pmid": {}, "title": {}}
    if not source:
        return empty

    path = materialize_history_source(source)
    if not path.exists():
        raise RuntimeError(f"Screening history file was not found: {path}")

    suffix = path.suffix.lower()
    if suffix in {".xlsx", ".xls"}:
        sheets = pd.read_excel(path, sheet_name=None, dtype=str)
        frames = [frame.assign(_history_sheet=name) for name, frame in sheets.items()]
        history_df = pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()
    elif suffix == ".csv":
        history_df = pd.read_csv(path, dtype=str)
        history_df["_history_sheet"] = path.name
    elif suffix == ".tsv":
        history_df = pd.read_csv(path, sep="\t", dtype=str)
        history_df["_history_sheet"] = path.name
    else:
        raise RuntimeError("Screening history must be an .xlsx, .xls, .csv, .tsv, or accessible Google Sheets URL.")

    if history_df.empty:
        return empty

    pmid_col = find_column(history_df, ("PMID", "PubMed ID", "PubMed_ID", "PMID Number"))
    title_col = find_column(history_df, ("Title", "Article Title", "Article_Title"))
    decision_col = find_column(history_df, ("Eligibility_Decision", "Eligibility Decision", "Decision", "Review_Status", "Review Status"))
    status_col = find_column(history_df, ("Review_Status", "Review Status", "Status"))

    history = {"pmid": {}, "title": {}}
    for idx, row in history_df.iterrows():
        pmid = normalize_pmid(row.get(pmid_col)) if pmid_col else ""
        title = normalize_text(row.get(title_col)) if title_col else ""
        decision = str(row.get(decision_col, "") or row.get(status_col, "") or "").strip()
        source_label = f"{path.name}:{row.get('_history_sheet', 'Sheet')}:{idx + 2}"
        record = {"decision": decision, "source": source_label}
        if pmid:
            history["pmid"][pmid] = record
        if title:
            history["title"][title] = record
    return history


def lookup_history_match(pmid: str, title: str, history: dict[str, dict[str, dict[str, str]]]) -> dict[str, str] | None:
    if pmid and pmid in history.get("pmid", {}):
        return history["pmid"][pmid]
    normalized_title = normalize_text(title)
    if normalized_title and normalized_title in history.get("title", {}):
        return history["title"][normalized_title]
    return None


def extract_pmid_from_zotero_extra(extra: str | None) -> str | None:
    if not extra:
        return None
    match = re.search(r"\bPMID:\s*(\d+)\b", extra, flags=re.IGNORECASE)
    return match.group(1) if match else None


def fetch_zotero_pmids(group_id: str, api_key: str) -> set[str]:
    zot = zotero.Zotero(group_id, "group", api_key)
    items = zot.everything(zot.items())
    pmids = set()
    for item in items:
        data = item.get("data", {})
        pmid = extract_pmid_from_zotero_extra(data.get("extra"))
        if pmid:
            pmids.add(pmid)
    return pmids


def publication_date_filter(from_year: int | None, to_year: int | None) -> tuple[str | None, str | None]:
    if not from_year and not to_year:
        return None, None
    current_year = date.today().year
    if from_year and (from_year < 1800 or from_year > current_year):
        raise RuntimeError(f"--from-year must be between 1800 and {current_year}.")
    if to_year and (to_year < 1800 or to_year > current_year):
        raise RuntimeError(f"--to-year must be between 1800 and {current_year}.")
    from_year = from_year or 1800
    to_year = to_year or current_year
    if from_year > to_year:
        raise RuntimeError("--from-year must be earlier than or equal to --to-year.")
    return f"{from_year}/01/01", f"{to_year}/12/31"


def format_year_filter(from_year: int | None, to_year: int | None) -> str:
    if from_year and to_year:
        return f"{from_year}-{to_year}"
    if from_year:
        return f"{from_year} onward"
    if to_year:
        return f"through {to_year}"
    return "all years"


def search_pubmed(query: str, retmax: int, from_year: int | None, to_year: int | None) -> list[str]:
    mindate, maxdate = publication_date_filter(from_year, to_year)
    search_kwargs = {
        "db": "pubmed",
        "term": query,
        "retmax": str(retmax),
        "sort": "pub+date",
    }
    if mindate and maxdate:
        search_kwargs.update({"datetype": "pdat", "mindate": mindate, "maxdate": maxdate})
    handle = Entrez.esearch(**search_kwargs)
    record = Entrez.read(handle)
    return [str(pmid) for pmid in record.get("IdList", [])]


def chunks(values: list[str], size: int) -> Iterable[list[str]]:
    for index in range(0, len(values), size):
        yield values[index : index + size]


def fetch_pubmed_records(pmids: list[str], batch_size: int = 200) -> list[dict]:
    records: list[dict] = []
    for batch in chunks(pmids, batch_size):
        handle = Entrez.efetch(db="pubmed", id=",".join(batch), retmode="xml")
        data = Entrez.read(handle)
        records.extend(data.get("PubmedArticle", []))
    return records


def stringify_abstract(article: dict) -> str:
    abstract = article.get("MedlineCitation", {}).get("Article", {}).get("Abstract", {})
    parts = abstract.get("AbstractText", [])
    values = []
    for part in parts:
        label = getattr(part, "attributes", {}).get("Label")
        text = str(part)
        values.append(f"{label}: {text}" if label else text)
    return " ".join(values).strip()


def article_date(article: dict) -> tuple[str, str]:
    pubmed_data = article.get("PubmedData", {})
    history = pubmed_data.get("History", [])
    for event in history:
        if event.attributes.get("PubStatus") == "pubmed":
            year = str(event.get("Year", ""))
            month = str(event.get("Month", "01")).zfill(2)
            day = str(event.get("Day", "01")).zfill(2)
            return year, f"{year}-{month}-{day}"
    journal_issue = article.get("MedlineCitation", {}).get("Article", {}).get("Journal", {}).get("JournalIssue", {})
    pub_date = journal_issue.get("PubDate", {})
    year = str(pub_date.get("Year", ""))
    return year, year


def doi_from_article(article: dict) -> str:
    ids = article.get("PubmedData", {}).get("ArticleIdList", [])
    for article_id in ids:
        if article_id.attributes.get("IdType") == "doi":
            return str(article_id)
    return ""


def parse_pubmed_article(
    article: dict,
    section: LibrarySection,
    zotero_pmids: set[str],
    screening_history: dict[str, dict[str, dict[str, str]]] | None = None,
) -> dict:
    medline = article.get("MedlineCitation", {})
    article_data = medline.get("Article", {})
    pmid = str(medline.get("PMID", ""))
    title = str(article_data.get("ArticleTitle", "")).strip()
    abstract = stringify_abstract(article)
    authors = []
    for author in article_data.get("AuthorList", []):
        name = f"{author.get('LastName', '')} {author.get('Initials', '')}".strip()
        if name:
            authors.append(name)
    year, publication_date = article_date(article)
    journal = str(article_data.get("Journal", {}).get("Title", ""))
    history_match = lookup_history_match(pmid, title, screening_history or {"pmid": {}, "title": {}})
    eligibility, rationale = screen_article(title, abstract, section)
    found_in_zotero = pmid in zotero_pmids
    return {
        "PMID": pmid,
        "Title": title,
        "Authors": ", ".join(authors) if authors else "N/A",
        "Journal": journal,
        "Publication_Year": year,
        "Publication_Date": publication_date,
        "DOI": doi_from_article(article),
        "Abstract": abstract,
        "Primary_Category": section.name,
        "Matched_Categories": section.name,
        "Suggested_Labels": "",
        "Eligibility_Decision": eligibility,
        "Eligibility_Rationale": rationale,
        "History_Match": bool(history_match),
        "History_Decision": history_match.get("decision", "") if history_match else "",
        "History_Source": history_match.get("source", "") if history_match else "",
        "Section_Inclusion_Criteria": section.inclusion,
        "Section_Exclusion_Criteria": section.exclusion,
        "Found_in_Zotero": found_in_zotero,
        "Reviewer_Decision": "Needs reviewed",
        "Review_Status": "Already in Zotero" if found_in_zotero else "Unreviewed",
        "Reviewer_Notes": "",
        "PubMed_URL": f"https://pubmed.ncbi.nlm.nih.gov/{pmid}/" if pmid else "",
        "Last_Seen": date.today().isoformat(),
    }


def screen_article(title: str, abstract: str, section: LibrarySection) -> tuple[str, str]:
    text = f"{title} {abstract}".lower()
    title_text = title.lower()
    has_cfc_term = any(term.lower() in text for term in CFC_TERMS)
    has_strong_cfc_signal = any(term in text for term in STRONG_CFC_SIGNALS)
    has_cfc_in_title = any(term in title_text for term in STRONG_CFC_SIGNALS)
    has_cfc_gene = any(gene.lower() in text for gene in CFC_GENES)
    has_rasopathy = "rasopath" in text
    has_other_rasopathy_in_title = any(signal in title_text for signal in OTHER_RASOPATHY_SIGNALS)
    section_terms = [token.lower() for token in re.findall(r"[A-Za-z]{5,}", section.description)]
    has_section_signal = any(token in text for token in section_terms[:12])

    if has_other_rasopathy_in_title and not has_cfc_in_title:
        return "Screen out", "Title is centered on another RASopathy and does not identify CFC as the study population."
    if has_strong_cfc_signal and has_section_signal:
        return "Needs human review", "Mentions CFC and matches the section topic; confirm the article includes CFC-specific data before screening in."
    if has_strong_cfc_signal:
        return "Needs human review", "Mentions CFC, but the section-specific topic signal is weak."
    if has_rasopathy and has_cfc_gene:
        return "Needs human review", "Mentions RASopathy plus a CFC-associated gene; check whether CFC data are present."
    return "Screen out", "No clear CFC-specific signal in title or abstract."


def add_suggested_labels(df: DataFrame, model_choice: str, label_count: int) -> DataFrame:
    if df.empty:
        return df
    model = SentenceTransformer(model_choice)
    labels = list(SECTIONS)
    label_embeddings = model.encode(labels)
    texts = (df["Title"].fillna("") + ". " + df["Abstract"].fillna("")).tolist()
    article_embeddings = model.encode(texts)
    scores = util.cos_sim(article_embeddings, label_embeddings)
    suggestions = []
    for row_index, row_scores in enumerate(scores):
        primary = df.iloc[row_index]["Primary_Category"]
        pairs = sorted(
            [(float(score), label) for score, label in zip(row_scores, labels)],
            key=lambda item: item[0],
            reverse=True,
        )
        suggestions.append(", ".join([label for _, label in pairs if label != primary][:label_count]))
    df = df.copy()
    df["Suggested_Labels"] = suggestions
    return df


def merge_with_existing(existing: DataFrame, new_rows: DataFrame) -> DataFrame:
    if existing.empty:
        return new_rows
    for column in FINAL_COLUMNS:
        if column not in existing:
            existing[column] = ""
    existing = existing[FINAL_COLUMNS].copy()
    existing_pmids = set(existing["PMID"].astype(str))
    truly_new = new_rows[~new_rows["PMID"].astype(str).isin(existing_pmids)]
    refreshed = pd.concat([existing, truly_new], ignore_index=True)
    return refreshed.drop_duplicates(subset=["PMID"], keep="first")


def filter_report_year_range(report: DataFrame, from_year: int | None, to_year: int | None) -> DataFrame:
    if (not from_year and not to_year) or report.empty or "Publication_Year" not in report:
        return report
    filtered = report.copy()
    years = pd.to_numeric(filtered["Publication_Year"], errors="coerce")
    mask = pd.Series(True, index=filtered.index)
    if from_year:
        mask = mask & (years >= from_year)
    if to_year:
        mask = mask & (years <= to_year)
    return filtered[mask].reset_index(drop=True)


def load_existing(path: Path) -> tuple[DataFrame, int]:
    if not path.exists():
        return pd.DataFrame(columns=FINAL_COLUMNS), 1
    existing = pd.read_excel(path, sheet_name="Review_Report", dtype={"PMID": str})
    run_count = 1
    try:
        instructions = pd.read_excel(path, sheet_name="Instructions", header=None)
        line = instructions[instructions[0].astype(str).str.contains("Run Count:", na=False)].iloc[0, 0]
        run_count = int(re.search(r"\d+", str(line)).group()) + 1
    except Exception:
        run_count = 1
    return existing, run_count


def criteria_frame() -> DataFrame:
    return pd.DataFrame(
        [
            {
                "Section": section.name,
                "Description": section.description,
                "Inclusion": section.inclusion,
                "Exclusion": section.exclusion,
                "PubMed_Query": section.query,
            }
            for section in SECTIONS.values()
        ]
    )


def build_deep_research_prompt(section: LibrarySection | None, recent_df: DataFrame) -> str:
    article_lines = []
    for _, row in recent_df.head(50).iterrows():
        article_lines.append(
            f"- PMID {row['PMID']}: {row['Title']} ({row.get('Publication_Year', '')}). "
            f"Decision: {row.get('Eligibility_Decision', '')}. URL: {row.get('PubMed_URL', '')}"
        )
    articles = "\n".join(article_lines) if article_lines else "No new candidate articles were found in this run."
    if section:
        scope = f"the Zotero section: {section.name}"
        section_context = textwrap.dedent(
            f"""
            Section description:
            {section.description}

            Inclusion criteria:
            {section.inclusion}

            Exclusion criteria:
            {section.exclusion}
            """
        ).strip()
    else:
        scope = "all Zotero library sections in the CFC research library"
        section_context = (
            "Use each article's Primary_Category and Matched_Categories fields to keep the original "
            "Zotero folder structure intact while reviewing the combined update."
        )

    return textwrap.dedent(
        f"""
        Conduct a deep research review update for {scope}.

        {section_context}

        General inclusion criteria:
        {json.dumps(GENERAL_INCLUSION_CRITERIA, indent=2)}

        General exclusion criteria:
        {json.dumps(GENERAL_EXCLUSION_CRITERIA, indent=2)}

        Candidate articles from the latest PubMed/Zotero update:
        {articles}

        Produce:
        1. A concise summary of the most important new findings.
        2. A table of articles that should be added to Zotero, with rationale.
        3. Articles that should be excluded or require human review, with rationale.
        4. Any missing search terms or adjacent concepts that should be considered for the next update.
        Use PubMed IDs and source links for every article-specific claim.
        """
    ).strip()


def launch_openai_deep_research(prompt: str) -> dict[str, str]:
    """Start an OpenAI deep research run and return lightweight metadata."""
    api_key = os.getenv("OPENAI_API_KEY")
    if not api_key:
        raise RuntimeError("OPENAI_API_KEY is required because deep research runs by default.")

    from openai import OpenAI

    client = OpenAI(api_key=api_key)
    response = client.responses.create(
        model=os.getenv("OPENAI_DEEP_RESEARCH_MODEL", "o3-deep-research"),
        input=prompt,
        background=True,
        tools=[
            {"type": "web_search_preview"},
            {"type": "code_interpreter", "container": {"type": "auto"}},
        ],
    )
    return {
        "response_id": getattr(response, "id", ""),
        "model": getattr(response, "model", os.getenv("OPENAI_DEEP_RESEARCH_MODEL", "o3-deep-research")),
        "status": getattr(response, "status", "submitted"),
        "submitted_at": time.ctime(),
    }


def write_workbook(
    path: Path,
    report: DataFrame,
    run_count: int,
    section: LibrarySection | None,
    prompt: str,
    from_year: int | None,
    to_year: int | None,
    openai_run: dict[str, str] | None = None,
) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    report = report.copy()
    if "Reviewer_Decision" in report:
        report["Reviewer_Decision"] = report["Reviewer_Decision"].replace("", pd.NA).fillna("Needs reviewed")
    sort_columns = [column for column in ("Primary_Category", "Title", "Publication_Year") if column in report]
    if sort_columns:
        report = report.sort_values(sort_columns, kind="stable").reset_index(drop=True)
    instructions = pd.DataFrame(
        [
            [f"Report generated on: {time.ctime()}"],
            [f"Run Count: {run_count}"],
            [f"Primary Category: {section.name if section else 'All categories'}"],
            [f"Publication Date Filter: {format_year_filter(from_year, to_year)}"],
            [""],
            ["How to Use This File"],
            ["1. Review_Report contains candidate articles and preserves Reviewer_Notes across runs."],
            ["2. Criteria contains the library descriptions plus inclusion/exclusion criteria for every Zotero section."],
            ["3. Deep_Research_Brief contains a prompt you can run with OpenAI deep research or an agent workflow."],
            ["4. Use Reviewer_Decision to mark Needs reviewed, Approved, or Not approved."],
            ["5. Re-run this program periodically; only unseen PubMed records are appended."],
        ]
    )
    run_log = pd.DataFrame(
        [
            {"Field": "Category", "Value": section.name if section else "All categories"},
            {"Field": "PubMed Query", "Value": section.query if section else "Multiple category queries"},
            {"Field": "Publication date filter", "Value": format_year_filter(from_year, to_year)},
            {"Field": "Rows in report", "Value": len(report)},
            {"Field": "Generated", "Value": time.ctime()},
        ]
    )
    category_summary = (
        report.groupby("Primary_Category", dropna=False)
        .size()
        .reset_index(name="Article_Count")
        .sort_values(["Primary_Category"])
    )
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        report[FINAL_COLUMNS].to_excel(writer, sheet_name="Review_Report", index=False)
        category_summary.to_excel(writer, sheet_name="Category_Summary", index=False)
        criteria_frame().to_excel(writer, sheet_name="Criteria", index=False)
        pd.DataFrame({"Deep_Research_Brief": [prompt]}).to_excel(writer, sheet_name="Deep_Research_Brief", index=False)
        if openai_run:
            pd.DataFrame([openai_run]).to_excel(writer, sheet_name="OpenAI_Run", index=False)
        run_log.to_excel(writer, sheet_name="Run_Log", index=False)
        instructions.to_excel(writer, sheet_name="Instructions", index=False, header=False)
    apply_review_dropdown_formatting(path)


def apply_review_dropdown_formatting(path: Path) -> None:
    from openpyxl import load_workbook
    from openpyxl.formatting.rule import FormulaRule
    from openpyxl.styles import Font, PatternFill
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.datavalidation import DataValidation

    wb = load_workbook(path)
    ws = wb["Review_Report"]
    headers = [cell.value for cell in ws[1]]
    if "Reviewer_Decision" not in headers:
        wb.save(path)
        return

    decision_col = headers.index("Reviewer_Decision") + 1
    decision_letter = get_column_letter(decision_col)
    primary_category_col = headers.index("Primary_Category") + 1 if "Primary_Category" in headers else None
    matched_category_col = headers.index("Matched_Categories") + 1 if "Matched_Categories" in headers else None
    max_row = max(ws.max_row, 2)
    decision_range = f"{decision_letter}2:{decision_letter}{max_row}"

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    validation = DataValidation(
        type="list",
        formula1=f'"{",".join(REVIEWER_DECISIONS)}"',
        allow_blank=True,
    )
    ws.add_data_validation(validation)
    validation.add(decision_range)

    fills = {
        "Needs reviewed": PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid"),
        "Approved": PatternFill(start_color="D9EAD3", end_color="D9EAD3", fill_type="solid"),
        "Not approved": PatternFill(start_color="F4CCCC", end_color="F4CCCC", fill_type="solid"),
    }
    for decision, fill in fills.items():
        ws.conditional_formatting.add(
            decision_range,
            FormulaRule(formula=[f'${decision_letter}2="{decision}"'], fill=fill),
        )

    if primary_category_col:
        for row in range(2, max_row + 1):
            category_cell = ws.cell(row=row, column=primary_category_col)
            color = CATEGORY_COLORS.get(str(category_cell.value), "FFFFFF")
            category_cell.fill = PatternFill(start_color=color, end_color=color, fill_type="solid")
            category_cell.font = Font(bold=True)
            if matched_category_col:
                ws.cell(row=row, column=matched_category_col).fill = PatternFill(
                    start_color=color,
                    end_color=color,
                    fill_type="solid",
                )

    summary_ws = wb["Category_Summary"] if "Category_Summary" in wb.sheetnames else None
    if summary_ws:
        summary_headers = [cell.value for cell in summary_ws[1]]
        if "Primary_Category" in summary_headers:
            summary_category_col = summary_headers.index("Primary_Category") + 1
            for row in range(2, summary_ws.max_row + 1):
                category_cell = summary_ws.cell(row=row, column=summary_category_col)
                color = CATEGORY_COLORS.get(str(category_cell.value), "FFFFFF")
                category_cell.fill = PatternFill(start_color=color, end_color=color, fill_type="solid")
                category_cell.font = Font(bold=True)
        summary_ws.freeze_panes = "A2"
        summary_ws.auto_filter.ref = summary_ws.dimensions

    for cell in ws[1]:
        cell.font = Font(bold=True)

    widths = {
        "PMID": 12,
        "Title": 55,
        "Primary_Category": 24,
        "Matched_Categories": 38,
        "Eligibility_Decision": 20,
        "Reviewer_Decision": 18,
        "Reviewer_Notes": 36,
        "PubMed_URL": 34,
    }
    for index, header in enumerate(headers, start=1):
        width = widths.get(str(header), 18)
        ws.column_dimensions[get_column_letter(index)].width = width

    wb.save(path)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Update a CFC syndrome literature review report.")
    parser.add_argument("--category", default="Dermatology", help="Zotero/library section to update.")
    parser.add_argument(
        "--all-categories",
        action="store_true",
        help="Search every configured Zotero/library section and export one combined workbook.",
    )
    parser.add_argument("--output", default="reports/CFC_Master_Review_Report.xlsx", help="Workbook output path.")
    parser.add_argument("--retmax", type=int, default=10000, help="Maximum PubMed IDs to return.")
    parser.add_argument(
        "--since-year",
        type=int,
        default=2025,
        help="Only search PubMed records published from this year onward. Use 0 to search all years.",
    )
    parser.add_argument(
        "--from-year",
        type=int,
        help="First publication year to include. Overrides --since-year when provided.",
    )
    parser.add_argument(
        "--to-year",
        type=int,
        help="Last publication year to include. Use with --from-year for a bounded date range.",
    )
    parser.add_argument(
        "--embedding-model",
        default="microsoft/BiomedNLP-PubMedBERT-base-uncased-abstract-fulltext",
        help="Sentence-transformers model used for suggested labels.",
    )
    parser.add_argument("--suggested-labels", type=int, default=2, help="Number of secondary labels to suggest.")
    parser.add_argument("--skip-zotero", action="store_true", help="Skip Zotero API lookup and mark all as not found in Zotero.")
    parser.add_argument(
        "--screening-history",
        default=DEFAULT_SCREENING_HISTORY_URL,
        help="Previously screened .xlsx/.csv/.tsv file or accessible Google Sheets URL. Matches by PMID first, then title.",
    )
    parser.add_argument(
        "--include-previously-screened",
        action="store_true",
        help="Keep screening-history matches in the new report instead of filtering them out.",
    )
    parser.add_argument(
        "--skip-openai-deep-research",
        action="store_true",
        help="Do not submit the generated brief to OpenAI deep research. Use only for testing.",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    load_env_file()
    load_runtime_dependencies()
    category = normalize_category(args.category)
    run_all_categories = args.all_categories or category.lower() in {"all", "all categories"}
    from_year = args.from_year if args.from_year is not None else (args.since_year or None)
    to_year = args.to_year
    if not run_all_categories and category not in SECTIONS:
        choices = ", ".join(sorted(SECTIONS))
        raise SystemExit(f"Unknown category '{args.category}'. Choose one of: {choices}")

    section = None if run_all_categories else SECTIONS[category]
    sections_to_run = list(SECTIONS.values()) if run_all_categories else [section]
    Entrez.email = require_env("ENTREZ_EMAIL")
    if not args.skip_openai_deep_research:
        require_env("OPENAI_API_KEY")
    output_path = Path(args.output)

    existing, run_count = load_existing(output_path)
    existing_pmids = set(existing.get("PMID", pd.Series(dtype=str)).astype(str))
    screening_history = load_screening_history(args.screening_history)
    history_pmids = set(screening_history["pmid"])

    zotero_pmids = set()
    if not args.skip_zotero:
        zotero_pmids = fetch_zotero_pmids(require_env("ZOTERO_GROUP_ID"), require_env("ZOTERO_API_KEY"))

    pmid_categories: dict[str, list[str]] = {}
    section_counts: list[dict[str, object]] = []
    for active_section in sections_to_run:
        section_pmids = search_pubmed(active_section.query, args.retmax, from_year, to_year)
        section_counts.append({"Section": active_section.name, "PubMed_Records_Found": len(section_pmids)})
        for pmid in section_pmids:
            pmid_categories.setdefault(pmid, [])
            if active_section.name not in pmid_categories[pmid]:
                pmid_categories[pmid].append(active_section.name)

    pubmed_pmids = list(pmid_categories)
    new_pmids = [
        pmid
        for pmid in pubmed_pmids
        if pmid not in existing_pmids and (args.include_previously_screened or pmid not in history_pmids)
    ]
    records = fetch_pubmed_records(new_pmids) if new_pmids else []
    rows = []
    for record in records:
        pmid = str(record.get("MedlineCitation", {}).get("PMID", ""))
        matched_categories = pmid_categories.get(pmid, [])
        primary_section_name = matched_categories[0] if matched_categories else sections_to_run[0].name
        primary_section = SECTIONS[primary_section_name]
        row = parse_pubmed_article(record, primary_section, zotero_pmids, screening_history)
        row["Matched_Categories"] = ", ".join(matched_categories) if matched_categories else primary_section.name
        rows.append(row)
    if not args.include_previously_screened:
        rows = [row for row in rows if not row["History_Match"]]
    new_df = pd.DataFrame(rows, columns=FINAL_COLUMNS)
    new_df = filter_report_year_range(new_df, from_year, to_year)
    new_df = add_suggested_labels(new_df, args.embedding_model, args.suggested_labels)
    report = merge_with_existing(existing, new_df)
    report = filter_report_year_range(report, from_year, to_year)
    prompt = build_deep_research_prompt(section, new_df)
    openai_run = None if args.skip_openai_deep_research else launch_openai_deep_research(prompt)
    write_workbook(output_path, report, run_count, section, prompt, from_year, to_year, openai_run)

    print(f"Category: {section.name if section else 'All categories'}")
    print(f"Publication date filter: {format_year_filter(from_year, to_year)}")
    print(f"Unique PubMed records found: {len(pubmed_pmids)}")
    if run_all_categories:
        print("Section counts:")
        for item in section_counts:
            print(f"  - {item['Section']}: {item['PubMed_Records_Found']}")
    if args.screening_history:
        print(f"Previously screened PMIDs loaded: {len(history_pmids)}")
    if openai_run:
        print(f"OpenAI deep research submitted: {openai_run.get('response_id', '')} ({openai_run.get('status', '')})")
    print(f"New records appended: {len(new_df)}")
    print(f"Workbook written: {output_path.resolve()}")


if __name__ == "__main__":
    main()
