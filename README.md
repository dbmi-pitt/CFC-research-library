# CFC Research Library

This project updates cardiofaciocutaneous syndrome literature review workbooks from PubMed and a Zotero group library.

It started as a Colab notebook and is now a reusable local program with:

- all Zotero library sections from the attached descriptions
- PubMed searches for each section
- section-specific inclusion and exclusion criteria
- Zotero cross-checking by PMID
- suggested secondary labels using biomedical embeddings
- an OpenAI deep research brief prompt saved into the workbook

## Setup

Install the Python libraries:

```powershell
python -m pip install -r requirements.txt
```

Set credentials as environment variables. You can copy `.env.example` as a reference, but keep real keys out of git.

Required:

- `ENTREZ_EMAIL`
- `ZOTERO_GROUP_ID`
- `ZOTERO_API_KEY`

Optional:

- `OPENAI_API_KEY`
- `OPENAI_DEEP_RESEARCH_MODEL`

## Run

```powershell
python cfc_research_library.py --category Dermatology --output reports/Dermatology_Master_Review_Report.xlsx
```

Useful categories include:

`Allergy and Immunology`, `Cardiology`, `Dermatology`, `Development and Behavior`, `Endocrinology`, `Gastroenterology`, `General and Reviews`, `Genetics`, `Gynecology`, `Neurology`, `Oncology`, `Ophthalmology`, `Orthopedic`, `Otolaryngology`, `Pulmonology`, `Research Studies`, `Seizures`, and `Treatments`.

The workbook contains:

- `Review_Report`: article-level screening list
- `Criteria`: library descriptions, inclusion criteria, exclusion criteria, and PubMed queries
- `Deep_Research_Brief`: a ready-to-use prompt for OpenAI deep research or an agent workflow
- `Run_Log`: run metadata
- `Instructions`: user-facing notes

## Notes

The program preserves existing `Review_Status` and `Reviewer_Notes` when appending newly discovered PubMed records. API keys from the original notebook were intentionally moved to environment variables.
