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
- `OPENAI_API_KEY`

Optional:

- `OPENAI_DEEP_RESEARCH_MODEL`

## Run

Easiest option on Windows:

1. Create a local `.env` file using `.env.example` as the template.
2. Add your real `ENTREZ_EMAIL`, `ZOTERO_GROUP_ID`, `ZOTERO_API_KEY`, and `OPENAI_API_KEY`.
3. Double-click `Run CFC Research Update.bat`.

The launcher creates/uses the local Python environment, installs the required libraries, filters out previously screened articles from the default Google Sheets history, searches for articles published from 2025 onward, runs all categories, submits OpenAI deep research, and writes:

`reports/CFC_All_Categories_Master_Review_Report.xlsx`

Command-line option:

```powershell
python cfc_research_library.py --category Dermatology --output reports/Dermatology_Master_Review_Report.xlsx
```

To export one combined workbook for every category while preserving category labels:

```powershell
uv run python cfc_research_library.py --all-categories --since-year 2025 --output reports/CFC_All_Categories_Master_Review_Report.xlsx
```

The combined workbook keeps `Primary_Category` and `Matched_Categories` columns so reviewers can filter or sort by the original library sections. By default, this command searches publication dates from 2025 onward, reads the prior Google Sheets screening history, filters out already-screened papers, and submits an OpenAI deep research run.

The prior screening history defaults to:

`https://docs.google.com/spreadsheets/d/1BUvWcV6XgYiOL3cCrYAHjkccb24C38OK/edit?usp=sharing&ouid=106518116377917721454&rtpof=true&sd=true`

To use a different prior screening file:

```powershell
uv run python cfc_research_library.py --all-categories --since-year 2025 --output reports/CFC_All_Categories_Master_Review_Report.xlsx --screening-history "C:\path\to\screening_history.xlsx"
```

If the Google Sheet is not publicly downloadable, export it as `.xlsx` or `.csv` and pass the downloaded file path to `--screening-history`.

For testing only, you can skip OpenAI deep research:

```powershell
uv run python cfc_research_library.py --all-categories --since-year 2025 --output reports/CFC_All_Categories_Master_Review_Report.xlsx --skip-openai-deep-research
```

Useful categories include:

`Allergy and Immunology`, `Cardiology`, `Dermatology`, `Development and Behavior`, `Endocrinology`, `Gastroenterology`, `General and Reviews`, `Genetics`, `Gynecology`, `Neurology`, `Oncology`, `Ophthalmology`, `Orthopedic`, `Otolaryngology`, `Pulmonology`, `Research Studies`, `Seizures`, and `Treatments`.

The workbook contains:

- `Review_Report`: article-level screening list sorted by category, with category columns and reviewer decision first
- `Category_Summary`: article counts by primary category
- `Criteria`: library descriptions, inclusion criteria, exclusion criteria, and PubMed queries
- `Deep_Research_Brief`: a ready-to-use prompt for OpenAI deep research or an agent workflow
- `OpenAI_Run`: OpenAI response ID and status when deep research is submitted
- `Run_Log`: run metadata
- `Instructions`: user-facing notes

## Notes

The program preserves existing `Review_Status` and `Reviewer_Notes` when appending newly discovered PubMed records. API keys from the original notebook were intentionally moved to environment variables.

Eligibility labels are intentionally conservative. The script no longer auto-labels CFC mentions as `Screen in`; likely matches are marked `Needs human review`, while titles centered on another RASopathy, such as Costello syndrome without CFC in the title, are screened out unless there is clear CFC-specific evidence.

Reviewer decisions are separate from automated eligibility labels. In `Review_Report`, use the `Reviewer_Decision` dropdown:

- `Needs reviewed`: yellow
- `Approved`: green
- `Not approved`: red

The `Primary_Category` and `Matched_Categories` columns appear at the front of the sheet and are color coded by category.
