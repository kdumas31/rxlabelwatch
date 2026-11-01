# RxLabelWatch — Formulary Safety Monitor

RxLabelWatch is a Streamlit app that checks your formulary against FDA's
**Drug Safety-related Labeling Changes (SrLC)** database and flags changes that
may affect hazardous-drug handling (NIOSH criteria and USP 800 language).

## What it does

1. **Load data.** Upload a formulary export (Epic `.xlsx`/`.csv`), then either:
   - **Fetch from FDA (date range)**, the default. Pick a date range (default:
     last 6 months) and the app runs SrLC's Date Search across all label
     sections itself, or
   - **Upload SrLC export**, the backup. Upload a CSV/Excel file exported from
     the SrLC site.
2. **Crosswalk.** Drug names are normalized (salts, dosage forms, strengths,
   and diluents are stripped) and fuzzy-matched (`rapidfuzz` token sort ratio,
   default threshold 78). Each fuzzy hit must also pass an ingredient-level
   check: every word of the shorter name has to appear in the other name,
   either exactly or as a near-identical spelling (90% similar or more). This
   rejects look-alike generics like amoxicillin/oxacillin and
   prednisone/prednisolone. For fetched data, the app matches on *Active
   Ingredient* and also on *Drug Name* (the brand), and keeps every matching
   label in range. The per-drug limit is adjustable, and 0 means no limit.
3. **Change details.** For each match, the app downloads SrLC's per-drug
   **Download Data** CSV and keeps only the supplements approved in your date
   range (dates come from the drug's SrLC detail page). If the CSV isn't
   available, it falls back to scraping the detail page.
4. **Handling text.** The app looks up the current label on **openFDA**, first
   by application number and then by generic name. It pulls the How Supplied,
   Storage & Handling, References, Carcinogenesis/Mutagenesis/Fertility,
   Pregnancy, Lactation, and Females & Males of Reproductive Potential
   sections. SrLC doesn't track handling sections, so this fills that gap.
5. **Hazard scan.** The NIOSH/USP 800 keyword scan runs on the SrLC change
   text and the openFDA text. The **risk level comes from the SrLC change text
   by default**. openFDA flags are still shown, labeled as context, and kept in
   the QA file. The *Count openFDA text in risk level* toggle adds them to the
   risk level. Keywords must start at a word boundary. Stems such as
   `carcinogen` still match *carcinogenicity*. Short acronyms such as `ppe`,
   `bsc`, `sds`, `cstd`, `n95`, `niosh`, and `usp 800` (plus `respirator`)
   must match as whole words. Highlighting follows the same rule.
6. **Review and export.** Record pharmacist decisions, then download:
   - the **report** (`.xlsx`: All Matches, High Risk, Action Required, P&T
     Summary), and
   - the **QA source data** (`.xlsx`: Run Info, SrLC Results, Matches, SrLC
     Change Text, openFDA Handling), with source URLs and fetch timestamps.
     Use it to check the raw data before you share results.

## Run locally

```bash
python -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt
streamlit run app.py
```

To try it out, use `samples/sample_formulary.csv` as the formulary.

## Optional openFDA API key

Without a key, openFDA allows about 1,000 requests per day per IP address.
A free key ([get one here](https://open.fda.gov/apis/authentication/)) raises
that to about 120,000 per day. The app makes one or two openFDA calls per
matched drug, and results are cached for 24 hours.

- **Local:** copy `.streamlit/secrets.toml.example` to
  `.streamlit/secrets.toml` and set `OPENFDA_API_KEY`. This file is
  git-ignored.
- **Streamlit Community Cloud:** open the app, then go to **Settings →
  Secrets** and paste:
  `OPENFDA_API_KEY = "your-key"`

The app runs fine without a key, and the key is never written to the exports.

## Tests

`python tests/test_core.py` runs offline checks of the keyword rule and the
near-miss drug pairs.

## Data sources

| Source | Used for | Access |
|---|---|---|
| [FDA SrLC](https://www.accessdata.fda.gov/scripts/cder/safetylabelingchanges/index.cfm) Date Search | Labeling changes in a date range | Scripted form POST with all section checkboxes |
| SrLC per-drug Download Data CSV (`?event=downloadDetails.page&DrugNameID=N`) | Changed label text by section and supplement | Plain GET |
| SrLC per-drug detail page | Supplement approval dates (and the fallback text) | Plain GET |
| [openFDA drug label API](https://open.fda.gov/apis/drug/label/) | Current handling, carcinogenicity, and reproductive text | JSON API, optional key |

Network calls are cached: the SrLC search for 6 hours, and per-drug data for 24
hours. The app waits about 0.4 seconds between FDA calls.

## Limitations

- The hazard scan is keyword-based and doesn't read context, so phrases like
  "no evidence of carcinogenicity" still trigger a flag. openFDA's
  carcinogenesis and pregnancy sections exist on almost every label, which is
  why openFDA text is left out of the risk level by default.
- Some products, such as Thalomid, have no openFDA label. Their flags come
  from SrLC only.
- FDA's Download Data CSV isn't strictly valid CSV, because it contains
  unescaped quotes. The app parses it with a pattern anchored on its fixed
  trailing columns. If FDA changes the format, the app falls back to the
  detail page.
