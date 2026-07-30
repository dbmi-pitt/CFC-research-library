# CFC Research Library

This project updates a cardiofaciocutaneous syndrome research library by combining PubMed search results, prior human screening history, Zotero library state, vector-based category suggestions, and OpenAI-assisted review outputs into a review workbook for curators.

The main workflow is designed for a rare disease organization that needs to:

- rerun topic-specific PubMed searches across the library's Zotero sections
- avoid re-screening articles that were already reviewed in prior update rounds
- compare new PubMed hits against the current Zotero library
- generate structured update suggestions for human review
- submit a deep research brief and save the returned report as markdown next to the workbook

It started as a Colab notebook and is now a reusable local Python program.

## What The Workflow Produces

A standard library update run produces:

- an Excel workbook with candidate articles, category suggestions, screening fields, and run metadata
- an OpenAI deep research prompt stored inside the workbook
- a markdown deep research report written beside the workbook after the Responses API result is retrieved

Typical outputs:

- `reports/CFC_All_Categories_Master_Review_Report.xlsx`
- `reports/CFC_All_Categories_Master_Review_Report_deep_research.md`

## Workflow Overview

### Systems Swimlane

```mermaid
flowchart LR
    subgraph User[User / Curator]
        U1[Run batch file or CLI command]
        U2[Review workbook]
        U3[Decide approve / not approve / needs review]
    end

    subgraph Local[Local Workflow Script]
        L1[cfc_research_library.py]
        L2[Section-specific PubMed queries]
        L3[Merge and classify candidate records]
        L4[Write workbook]
        L5[Poll OpenAI Responses API]
        L6[Write deep research markdown]
    end

    subgraph History[Prior Review History]
        H1[Google Sheet or local xlsx/csv/tsv]
        H2[Previously screened PMIDs and titles]
    end

    subgraph PubMed[NCBI / PubMed]
        P1[PubMed search]
        P2[PubMed article metadata]
    end

    subgraph Zotero[Zotero Group Library]
        Z1[Current library holdings]
        Z2[Folder/category presence by PMID]
    end

    subgraph Models[Local Models + OpenAI]
        M1[Sentence-transformer similarity scoring]
        M2[OpenAI category/review calls]
        M3[OpenAI deep research background run]
    end

    U1 --> L1
    L1 --> L2
    L2 --> P1
    P1 --> P2
    P2 --> L3
    H1 --> H2
    H2 --> L3
    Z1 --> Z2
    Z2 --> L3
    L3 --> M1
    M1 --> L3
    L3 --> M2
    L3 --> M3
    L3 --> L4
    L4 --> U2
    L4 --> L5
    M3 --> L5
    L5 --> L6
    L6 --> U2
    U2 --> U3
```

### Information Flow

```mermaid
flowchart TD
    A[Library section definitions and PubMed queries] --> B[Run PubMed searches]
    B --> C[Retrieve current PubMed record set]
    D[Prior screening history] --> E[Match by PMID and title]
    C --> E
    F[Current Zotero library] --> G[Match by PMID to existing library entries]
    C --> G
    C --> H[Parse article metadata]
    H --> I[Assign Primary_Category and Matched_Categories candidates]
    I --> J[Vector similarity scoring for secondary label suggestions]
    E --> K[Filter or flag previously screened records]
    G --> L[Flag already-in-Zotero records]
    J --> M[Candidate update table]
    K --> M
    L --> M
    M --> N[Workbook sheets: Review_Report, Category_Summary, Criteria, Run_Log]
    M --> O[Deep research brief prompt]
    O --> P[OpenAI background deep research run]
    P --> Q[Poll Responses API until completed]
    Q --> R[Deep research markdown report]
    N --> S[Human review and library update decisions]
    R --> S
```

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
- `OPENAI_DEEP_RESEARCH_TIMEOUT_SECONDS`
- `OPENAI_DEEP_RESEARCH_POLL_INTERVAL_SECONDS`

## How To Use It For A Library Update

A normal update cycle is:

1. Start from the current Zotero library and prior screening history.
2. Run the script for one category or all categories.
3. Let the script retrieve new PubMed results, compare them against Zotero and prior history, and write the workbook.
4. Wait for the deep research background run to complete; the script now polls the Responses API after workbook creation.
5. Review both the workbook and the generated markdown report.
6. Use the workbook to decide which articles should be added to Zotero, excluded, or escalated for human review.

The script is intended to support human curation, not replace it. Reviewers should make the final inclusion and categorization decisions.

## Run

Easiest option on Windows:

1. Create a local `.env` file using `.env.example` as the template.
2. Add your real `ENTREZ_EMAIL`, `ZOTERO_GROUP_ID`, `ZOTERO_API_KEY`, and `OPENAI_API_KEY`.
3. Double-click `Run CFC Research Update.bat`.

The launcher creates or reuses the local Python environment, installs the required libraries, filters out previously screened articles from the default Google Sheets history, searches for articles published from 2025 onward, runs all categories, submits OpenAI deep research, writes the workbook, polls for the deep research result, and writes the markdown report.

Command-line example for one category:

```powershell
python cfc_research_library.py --category Dermatology --output reports/Dermatology_Master_Review_Report.xlsx
```

To export one combined workbook for every category while preserving category labels:

```powershell
uv run python cfc_research_library.py --all-categories --since-year 2025 --output reports/CFC_All_Categories_Master_Review_Report.xlsx
```

The combined workbook keeps `Primary_Category` and `Matched_Categories` columns so reviewers can filter or sort by the original library sections. By default, this command searches publication dates from 2025 onward, reads the prior Google Sheets screening history, filters out already-screened papers, submits an OpenAI deep research run, and waits for the markdown report to be returned.

The prior screening history defaults to:

`https://docs.google.com/spreadsheets/d/1BUvWcV6XgYiOL3cCrYAHjkccb24C38OK/edit?usp=sharing&ouid=106518116377917721454&rtpof=true&sd=true`

To use a different prior screening file:

```powershell
uv run python cfc_research_library.py --all-categories --since-year 2025 --output reports/CFC_All_Categories_Master_Review_Report.xlsx --screening-history "C:\path\to\screening_history.xlsx"
```

If the Google Sheet is not publicly downloadable, export it as `.xlsx`, `.csv`, or `.tsv` and pass the downloaded file path to `--screening-history`.

For testing only, you can skip OpenAI deep research:

```powershell
uv run python cfc_research_library.py --all-categories --since-year 2025 --output reports/CFC_All_Categories_Master_Review_Report.xlsx --skip-openai-deep-research
```

Useful categories include:

`Allergy and Immunology`, `Cardiology`, `Dermatology`, `Development and Behavior`, `Endocrinology`, `Gastroenterology`, `General and Reviews`, `Genetics`, `Growth`, `Gynecology`, `Neurology`, `Oncology`, `Ophthalmology`, `Orthopedic`, `Otolaryngology`, `Pulmonology`, `Research Studies`, `Seizures`, and `Treatments`.

## Output Files

The workbook contains:

- `Review_Report`: article-level screening list sorted by category, with category columns and reviewer decision first
- `Category_Summary`: article counts by primary category
- `Criteria`: library descriptions, inclusion criteria, exclusion criteria, and PubMed queries
- `Deep_Research_Brief`: a ready-to-use prompt for OpenAI deep research or an agent workflow
- `OpenAI_Run`: OpenAI response ID and submission status for the deep research job
- `Run_Log`: run metadata
- `Instructions`: user-facing notes

The markdown report contains the returned deep research narrative from the OpenAI Responses API. It is written in the same directory as the workbook, using the workbook stem plus `_deep_research.md`.

## Notes

The program preserves existing `Review_Status` and `Reviewer_Notes` when appending newly discovered PubMed records. API keys from the original notebook were intentionally moved to environment variables.

Eligibility labels are intentionally conservative. The script no longer auto-labels CFC mentions as `Screen in`; likely matches are marked `Needs human review`, while titles centered on another RASopathy, such as Costello syndrome without CFC in the title, are screened out unless there is clear CFC-specific evidence.

Reviewer decisions are separate from automated eligibility labels. In `Review_Report`, use the `Reviewer_Decision` dropdown:

- `Needs reviewed`: yellow
- `Approved`: green
- `Not approved`: red

The `Primary_Category` and `Matched_Categories` columns appear at the front of the sheet and are color coded by category.
