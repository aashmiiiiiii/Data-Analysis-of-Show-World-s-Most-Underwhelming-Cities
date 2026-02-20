# Experience Analysis — Outscraper Review Pipeline

Reads a Google Maps review export from Outscraper and runs keyword-based sentiment analysis across 200+ attractions. Outputs a styled Excel report grouped by attraction name.

---

## What it does

- Scans 100K+ reviews for experience-related keywords using regex with wildcard support (`recommend*`, `disappoint*`, etc.)
- Handles negation so phrases like "not worth it" or "don't recommend" score correctly as negative
- Scores each attraction on a 0–100 experience scale per category
- Translates all keywords into 8 languages via Google Translate
- Exports results to a formatted Excel file with colour-coded scores

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
