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
import textwrap
import time
from dataclasses import asdict, dataclass
from datetime import date
from pathlib import Path
from typing import Iterable

pd = None
Entrez = None
zotero = None
SentenceTransformer = None
util = None


CFC_TERMS = (
    "cardiofaciocutaneous",
    "cardio-facio-cutaneous",
    "cardio-facio-cutaneous syndrome",
    "cardiofaciocutaneous syndrome",
    "CFC syndrome",
)

CFC_GENES = ("BRAF", "MAP2K1", "MAP2K2", "KRAS")

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
        "Hormonal, metabolic, renal, and growth-related abnormalities in CFC.",
        "Research examining hormonal, metabolic, renal, or growth-related abnormalities in CFC syndrome, including studies using CFC-associated mutations to model endocrine dysfunction.",
        "Endocrine studies on other RASopathies without CFC data, or general metabolic pathway papers not linked to CFC clinical features.",
        f'(("Endocrine System Diseases"[MeSH Terms] OR "Growth Disorders"[MeSH Terms] OR hormone[All Fields] OR renal[All Fields] OR metabolic[All Fields]) AND {cfc_clause()})',
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
    "PMID",
    "Title",
    "Authors",
    "Journal",
    "Publication_Year",
    "Publication_Date",
    "DOI",
    "Abstract",
    "Primary_Category",
    "Suggested_Labels",
    "Eligibility_Decision",
    "Eligibility_Rationale",
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


def search_pubmed(query: str, retmax: int) -> list[str]:
    handle = Entrez.esearch(db="pubmed", term=query, retmax=str(retmax), sort="pub+date")
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


def parse_pubmed_article(article: dict, section: LibrarySection, zotero_pmids: set[str]) -> dict:
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
        "Suggested_Labels": "",
        "Eligibility_Decision": eligibility,
        "Eligibility_Rationale": rationale,
        "Section_Inclusion_Criteria": section.inclusion,
        "Section_Exclusion_Criteria": section.exclusion,
        "Found_in_Zotero": found_in_zotero,
        "Review_Status": "Already in Zotero" if found_in_zotero else "Unreviewed",
        "Reviewer_Notes": "",
        "PubMed_URL": f"https://pubmed.ncbi.nlm.nih.gov/{pmid}/" if pmid else "",
        "Last_Seen": date.today().isoformat(),
    }


def screen_article(title: str, abstract: str, section: LibrarySection) -> tuple[str, str]:
    text = f"{title} {abstract}".lower()
    has_cfc_term = any(term.lower() in text for term in CFC_TERMS)
    has_cfc_gene = any(gene.lower() in text for gene in CFC_GENES)
    has_rasopathy = "rasopath" in text
    section_terms = [token.lower() for token in re.findall(r"[A-Za-z]{5,}", section.description)]
    has_section_signal = any(token in text for token in section_terms[:12])

    if has_cfc_term and has_section_signal:
        return "Screen in", "Mentions CFC and matches the selected section topic."
    if has_cfc_term:
        return "Needs human review", "Mentions CFC, but the section-specific topic signal is weak."
    if has_rasopathy and has_cfc_gene:
        return "Needs human review", "Mentions RASopathy plus a CFC-associated gene; check whether CFC data are present."
    return "Screen out", "No clear CFC-specific signal in title or abstract."


def add_suggested_labels(df: pd.DataFrame, model_choice: str, label_count: int) -> pd.DataFrame:
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


def merge_with_existing(existing: pd.DataFrame, new_rows: pd.DataFrame) -> pd.DataFrame:
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


def load_existing(path: Path) -> tuple[pd.DataFrame, int]:
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


def criteria_frame() -> pd.DataFrame:
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


def build_deep_research_prompt(section: LibrarySection, recent_df: pd.DataFrame) -> str:
    article_lines = []
    for _, row in recent_df.head(50).iterrows():
        article_lines.append(
            f"- PMID {row['PMID']}: {row['Title']} ({row.get('Publication_Year', '')}). "
            f"Decision: {row.get('Eligibility_Decision', '')}. URL: {row.get('PubMed_URL', '')}"
        )
    articles = "\n".join(article_lines) if article_lines else "No new candidate articles were found in this run."
    return textwrap.dedent(
        f"""
        Conduct a deep research review update for the Zotero section: {section.name}.

        Section description:
        {section.description}

        Inclusion criteria:
        {section.inclusion}

        Exclusion criteria:
        {section.exclusion}

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
    """Start an optional OpenAI deep research run and return lightweight metadata."""
    api_key = os.getenv("OPENAI_API_KEY")
    if not api_key:
        raise RuntimeError("OPENAI_API_KEY is required for --run-openai-deep-research")

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
    report: pd.DataFrame,
    run_count: int,
    section: LibrarySection,
    prompt: str,
    openai_run: dict[str, str] | None = None,
) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    instructions = pd.DataFrame(
        [
            [f"Report generated on: {time.ctime()}"],
            [f"Run Count: {run_count}"],
            [f"Primary Category: {section.name}"],
            [""],
            ["How to Use This File"],
            ["1. Review_Report contains candidate articles and preserves Reviewer_Notes across runs."],
            ["2. Criteria contains the library descriptions plus inclusion/exclusion criteria for every Zotero section."],
            ["3. Deep_Research_Brief contains a prompt you can run with OpenAI deep research or an agent workflow."],
            ["4. Re-run this program periodically; only unseen PubMed records are appended."],
        ]
    )
    run_log = pd.DataFrame(
        [
            {"Field": "Category", "Value": section.name},
            {"Field": "PubMed Query", "Value": section.query},
            {"Field": "Rows in report", "Value": len(report)},
            {"Field": "Generated", "Value": time.ctime()},
        ]
    )
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        report[FINAL_COLUMNS].to_excel(writer, sheet_name="Review_Report", index=False)
        criteria_frame().to_excel(writer, sheet_name="Criteria", index=False)
        pd.DataFrame({"Deep_Research_Brief": [prompt]}).to_excel(writer, sheet_name="Deep_Research_Brief", index=False)
        if openai_run:
            pd.DataFrame([openai_run]).to_excel(writer, sheet_name="OpenAI_Run", index=False)
        run_log.to_excel(writer, sheet_name="Run_Log", index=False)
        instructions.to_excel(writer, sheet_name="Instructions", index=False, header=False)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Update a CFC syndrome literature review report.")
    parser.add_argument("--category", default="Dermatology", help="Zotero/library section to update.")
    parser.add_argument("--output", default="reports/CFC_Master_Review_Report.xlsx", help="Workbook output path.")
    parser.add_argument("--retmax", type=int, default=10000, help="Maximum PubMed IDs to return.")
    parser.add_argument(
        "--embedding-model",
        default="microsoft/BiomedNLP-PubMedBERT-base-uncased-abstract-fulltext",
        help="Sentence-transformers model used for suggested labels.",
    )
    parser.add_argument("--suggested-labels", type=int, default=2, help="Number of secondary labels to suggest.")
    parser.add_argument("--skip-zotero", action="store_true", help="Skip Zotero API lookup and mark all as not found in Zotero.")
    parser.add_argument(
        "--run-openai-deep-research",
        action="store_true",
        help="Submit the generated brief to OpenAI deep research in background mode. Requires OPENAI_API_KEY.",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    load_runtime_dependencies()
    category = normalize_category(args.category)
    if category not in SECTIONS:
        choices = ", ".join(sorted(SECTIONS))
        raise SystemExit(f"Unknown category '{args.category}'. Choose one of: {choices}")

    section = SECTIONS[category]
    Entrez.email = require_env("ENTREZ_EMAIL")
    output_path = Path(args.output)

    existing, run_count = load_existing(output_path)
    existing_pmids = set(existing.get("PMID", pd.Series(dtype=str)).astype(str))

    zotero_pmids = set()
    if not args.skip_zotero:
        zotero_pmids = fetch_zotero_pmids(require_env("ZOTERO_GROUP_ID"), require_env("ZOTERO_API_KEY"))

    pubmed_pmids = search_pubmed(section.query, args.retmax)
    new_pmids = [pmid for pmid in pubmed_pmids if pmid not in existing_pmids]
    records = fetch_pubmed_records(new_pmids) if new_pmids else []
    rows = [parse_pubmed_article(record, section, zotero_pmids) for record in records]
    new_df = pd.DataFrame(rows, columns=FINAL_COLUMNS)
    new_df = add_suggested_labels(new_df, args.embedding_model, args.suggested_labels)
    report = merge_with_existing(existing, new_df)
    prompt = build_deep_research_prompt(section, new_df)
    openai_run = launch_openai_deep_research(prompt) if args.run_openai_deep_research else None
    write_workbook(output_path, report, run_count, section, prompt, openai_run)

    print(f"Category: {section.name}")
    print(f"PubMed records found: {len(pubmed_pmids)}")
    print(f"New records appended: {len(new_df)}")
    print(f"Workbook written: {output_path.resolve()}")


if __name__ == "__main__":
    main()
