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
   default threshold 78). For fetched data, the app matches on
   *Active Ingredient* and also on *Drug Name* (the brand).
3. **Change details.** For each match, the app downloads SrLC's per-drug
   **Download Data** CSV and keeps only the supplements approved in your date
   range (dates come from the drug's SrLC detail page). If the CSV isn't
   available, it falls back to scraping the detail page.
4. **Handling text.** The app looks up the current label on **openFDA**, first
   by application number and then by generic name. It pulls the How Supplied,
   Storage & Handling, References, Carcinogenesis/Mutagenesis/Fertility,
   Pregnancy, Lactation, and Females & Males of Reproductive Potential
   sections. SrLC doesn't track handling sections, so this fills that gap.
5. **Hazard scan.** The same NIOSH/USP 800 keyword scan runs on the SrLC
   change text and the openFDA text. Each flag shows which source triggered
   it, and you can choose whether openFDA text counts toward the overall risk
   level.
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

- The hazard scan is keyword-based. openFDA's carcinogenesis and pregnancy
  sections exist on almost every label, so with openFDA counted, most matches
  score HIGH. Use the per-source risk columns, or turn off *Count openFDA text
  in risk level*.
- Some keywords also match inside other words (for example, `ppe` matches
  "appear"). The scan is unchanged from the original app.
- Matching keeps the top 5 SrLC rows per formulary drug, so a common generic
  such as methotrexate can have more than 5 labels in range.
- Some products, such as Thalomid, have no openFDA label. Their flags come
  from SrLC only.
- FDA's Download Data CSV isn't strictly valid CSV, because it contains
  unescaped quotes. The app parses it with a pattern anchored on its fixed
  trailing columns. If FDA changes the format, the app falls back to the
  detail page.
