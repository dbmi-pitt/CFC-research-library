# Prompt Changelog

Use this file to track prompt edits and F1 effects.

## 2026-07-15

- File changed: prompts/review_classification_prompt.txt; prompts/priority_rules.txt; cfc_research_library.py
- Category targeted: all categories
- What changed: made the supplied inclusion/exclusion criteria authoritative and tightened additional category rules so secondary categories are only used when absolutely necessary.
- Expected effect: fewer unnecessary secondary categories, better precision, and cleaner adherence to Lexi's category criteria.
- F1 result after rerun:

## 2026-07-15

- File changed: prompts/review_classification_prompt.txt; cfc_research_library.py
- Category targeted: additional category behavior
- What changed: made additional categories opt-in by default, removed the example that showed a filled secondary category, and stopped fallback mode from copying deep-learning suggestions into the secondary category column.
- Expected effect: return to mostly one-category assignments and reduce random additional categories.
- F1 result after rerun:

## 2026-07-15

- File changed: prompts/review_classification_prompt.txt; prompts/priority_rules.txt
- Category targeted: relevance screening before category assignment
- What changed: added human-screened calibration examples showing that Noonan-only, broad RASopathy, broad pathway/model, gene-curation, and non-CFC syndrome papers should be excluded unless they contain substantive CFC evidence or direct CFC comparison.
- Expected effect: fewer weakly related articles kept in the review set and fewer random category assignments.
- F1 result after rerun:

## 2026-07-15

- File changed: prompts/review_classification_prompt.txt; prompts/priority_rules.txt; prompts/article_types.txt; prompts/category_guidance.json
- Category targeted: PromptTest_08 calibration
- What changed: restored the stricter PromptTest_08 classification style, including stricter relevance screening, conservative Treatments logic, detailed case-report handling, and Genetics/Neurology/Development priority rules.
- Expected effect: classify future review-comparison runs more like CFC_2017_2022_Review_Comparison_PromptTest_08.xlsx.
- F1 result after rerun:

## Template

- Date:
- File changed:
- Category targeted:
- What changed:
- Expected effect:
- F1 result after rerun:
