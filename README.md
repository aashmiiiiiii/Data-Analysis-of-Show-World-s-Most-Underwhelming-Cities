# Experience Analysis — Outscraper Review Pipeline

Reads a Google Maps review export from Outscraper and runs keyword-based sentiment analysis across 200+ attractions. Outputs a styled Excel report grouped by attraction name.

---

## What it does

-Ingested 100K+ Google Maps reviews from an Outscraper Excel export across 200 tourist attractions.
-Built a custom NLP keyword detection engine using compiled regex with wildcard support (e.g. recommend* matches recommended, recommending, etc.) across 8 experience categories and 209 patterns.
-Implemented negation-aware false-positive filtering, phrases like "not worth it" or "don't recommend" are correctly scored as negative rather than triggering positive matches.
-Translated all 209 experience keywords into 8 languages (Spanish, French, German, Japanese, Chinese, Arabic, Portuguese) via Google Translate API with deduplication to minimise API calls.
-Generated styled Excel reports with colour-coded scoring, frozen panes, and alternating row formatting across 3 analysis-ready sheets: Attraction Summary, Category Breakdown, and Keyword Translations.

---

## Requirements

```
pip install pandas openpyxl deep-translator
```

Python 3.10+ required (uses `float | None` type hints).

---

## Usage

```bash
# quick run, no translation
py -3 experience_analysis.py --no-translate

# full run with translations (~10-15 min)
py -3 experience_analysis.py

# custom input/output paths
py -3 experience_analysis.py --input my_file.xlsx --output my_report.xlsx
```

The input file defaults to `Outscraper-20241216130344s54.xlsx` in the same folder. The output file is timestamped automatically (e.g. `experience_analysis_20240101_120000.xlsx`).

---

## Output

The Excel report has 3 sheets:

| Sheet | Contents |
|---|---|
| Attraction Summary | One row per attraction — overall score, avg rating, per-category keyword counts |
| Category Breakdown | Per attraction × category with score and sentiment label |
| Keyword Translations | All 209 patterns translated into Spanish, French, German, Japanese, Chinese, Arabic, Portuguese |

Scores are colour-coded: green (70+), yellow (40–70), red (under 40).

---

## Categories

The analysis covers 8 experience categories:

- Recommendation
- Overall Impression
- Value for Money
- Service Quality
- Atmosphere
- Cleanliness
- Crowd & Wait Times
- Uniqueness & Memorability

---

## Notes

- The Outscraper Excel export has two sheets — the script auto-detects which one contains the review data
- Translation step makes ~600 API calls to Google Translate, hence the time estimate. Use `--no-translate` during testing
- Non-English reviews won't match many patterns since the keyword list is English only
