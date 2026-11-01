"""Network access to FDA data sources used by RxLabelWatch.

Plain functions with no Streamlit dependency, so they can be tested headlessly.
app.py wraps them with st.cache_data.

Sources
-------
* SrLC (Drug Safety-related Labeling Changes database)
  - date-range search (results table)
  - per-drug "Download Data" CSV (changed label text by section/supplement)
  - per-drug detail page (used only for its supplement -> date headers, and as
    a last-resort text fallback via fetch_label_page)
* openFDA drug label API (current label handling / carcinogenicity /
  reproductive sections)
"""
from __future__ import annotations

import io
import re
import time
from datetime import date, datetime, timezone
from urllib.parse import parse_qs, urlencode, urlparse

import pandas as pd
import requests
from bs4 import BeautifulSoup

SRLC_BASE = "https://www.accessdata.fda.gov/scripts/cder/safetylabelingchanges/index.cfm"
SRLC_SEARCH_URL = SRLC_BASE + "?event=searchResult.page"
SRLC_DETAIL_URL = SRLC_BASE + "?event=searchdetail.page&DrugNameID={id}"
SRLC_CSV_URL = SRLC_BASE + "?event=downloadDetails.page&DrugNameID={id}"
OPENFDA_LABEL_URL = "https://api.fda.gov/drug/label.json"

# SrLC label-section checkboxes. The search silently redirects (302, no results)
# unless at least one is sent, so send all of them.
SRLC_SECTIONS = ["BW", "C", "WP", "AR", "DI", "USP", "MG"]

REQUEST_DELAY_S = 0.4  # politeness delay after every FDA call
TIMEOUT_S = 30

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0 Safari/537.36"
    ),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "en-US,en;q=0.5",
    "Referer": SRLC_BASE,
}

# openFDA label fields that carry handling / carcinogenicity / reproductive text.
# These are scanned for hazards.
OPENFDA_SCAN_FIELDS = [
    "how_supplied",
    "storage_and_handling",
    "references",
    "carcinogenesis_and_mutagenesis_and_impairment_of_fertility",
    "pregnancy",
    "females_and_males_of_reproductive_potential",
    "lactation",
]
# Extra context kept in the QA export only (not scanned, because they describe
# the whole current label rather than handling).
OPENFDA_CONTEXT_FIELDS = ["recent_major_changes", "boxed_warning"]

_ILLEGAL_XML = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")


class FDAError(Exception):
    """Raised when an FDA source can't be reached or parsed."""


def now_utc_iso() -> str:
    return datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


def clean_text(s) -> str:
    if not isinstance(s, str):
        return ""
    return _ILLEGAL_XML.sub("", s)


def html_to_text(fragment: str) -> str:
    if not fragment:
        return ""
    text = BeautifulSoup(fragment, "html.parser").get_text(separator=" ")
    text = re.sub(r"[ \t\r\f\v]+", " ", text)
    text = re.sub(r"\s*\n\s*", "\n", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return clean_text(text.strip())


def _session() -> requests.Session:
    s = requests.Session()
    s.headers.update(HEADERS)
    return s


def _parse_date(s) -> date | None:
    if isinstance(s, date):
        return s
    try:
        return datetime.strptime(str(s).strip(), "%m/%d/%Y").date()
    except (ValueError, TypeError):
        return None


def drug_name_id_from_url(url: str) -> str:
    """Extract DrugNameID from an SrLC detail/download URL ('' if absent)."""
    if not isinstance(url, str) or "DrugNameID" not in url:
        return ""
    qs = parse_qs(urlparse(url).query)
    val = qs.get("DrugNameID", [""])[0]
    return val if val.isdigit() else ""


# ─── 1. SrLC date-range search ───────────────────────────────────────────────
SRLC_RESULT_COLUMNS = [
    "Drug Name", "Active Ingredient", "Application Number", "Application Type",
    "Supplement Date", "Database Updated", "Link", "DrugNameID",
]


def search_srlc_by_date(date_from: date, date_to: date) -> pd.DataFrame:
    """Run SrLC's Date Search (all label sections) and return the results table.

    Columns: SRLC_RESULT_COLUMNS plus 'fetched_at' and 'source_url'.
    Raises FDAError on network/parse failure.
    """
    if date_from > date_to:
        raise FDAError("Start date is after end date.")
    s = _session()
    try:
        s.get(SRLC_BASE, timeout=TIMEOUT_S).raise_for_status()  # session cookies
        time.sleep(REQUEST_DELAY_S)
        form = [
            ("date_range_from", date_from.strftime("%m/%d/%Y")),
            ("date_range_to", date_to.strftime("%m/%d/%Y")),
        ] + [("section", code) for code in SRLC_SECTIONS] + [("Datesearch", "")]
        resp = s.post(SRLC_SEARCH_URL, data=form, timeout=90, allow_redirects=False)
    except requests.RequestException as e:
        raise FDAError(f"SrLC search request failed: {e}") from e
    finally:
        time.sleep(REQUEST_DELAY_S)
    if resp.status_code in (301, 302, 303):
        raise FDAError("SrLC search redirected instead of returning results (form rejected).")
    if resp.status_code != 200:
        raise FDAError(f"SrLC search returned HTTP {resp.status_code}.")

    soup = BeautifulSoup(resp.text, "lxml")
    table = soup.find("table", id="example")
    fetched_at = now_utc_iso()
    if table is None:
        if "no records" in resp.text.lower() or "no results" in resp.text.lower():
            return pd.DataFrame(columns=SRLC_RESULT_COLUMNS + ["fetched_at", "source_url"])
        raise FDAError("SrLC results table not found — the page layout may have changed.")

    header_cells = [c.get_text(" ", strip=True) for c in table.find("tr").find_all(["th", "td"])]
    rows = []
    for tr in table.find_all("tr")[1:]:
        cells = tr.find_all("td")
        if len(cells) != len(header_cells):
            continue
        rec = {h: c.get_text(" ", strip=True) for h, c in zip(header_cells, cells)}
        link = rec.get("Link", "")
        if not link.startswith("http"):
            a = cells[0].find("a")
            link = ("https://www.accessdata.fda.gov" + a["href"]) if a and a.get("href", "").startswith("/") else link
            rec["Link"] = link
        rec["DrugNameID"] = drug_name_id_from_url(link)
        rows.append(rec)
    df = pd.DataFrame(rows)
    for col in SRLC_RESULT_COLUMNS:
        if col not in df.columns:
            df[col] = ""
    df = df[SRLC_RESULT_COLUMNS].copy()
    df["fetched_at"] = fetched_at
    df["source_url"] = (
        f"{SRLC_SEARCH_URL} (Date Search {date_from:%m/%d/%Y}–{date_to:%m/%d/%Y}, "
        f"sections {','.join(SRLC_SECTIONS)})"
    )
    return df


# ─── 2. SrLC per-drug Download Data CSV ──────────────────────────────────────
# FDA's CSV is not valid CSV: the Description field contains unescaped quotes
# from embedded HTML. The trailing 8 fields are well-formed, so anchor on them.
_SRLC_CSV_RECORD = re.compile(
    r'^"(?P<section>[^"\n]*)","(?P<title>[^"\n]*)","(?P<description>.*?)",'
    r'"(?P<appl_no>\d+)","(?P<appl_type>[A-Z]+)","(?P<drug_name>[^"\n]*)",'
    r'"(?P<active_ingredient>[^"\n]*)","(?P<sub_no>[^"\n]*)","(?P<sub_type>[^"\n]*)",'
    r'"(?P<label_url>[^"\n]*)","(?P<label_url_text>[^"\n]*)"[ \t]*\r?$',
    re.S | re.M,
)


def parse_srlc_csv(raw: str) -> list[dict]:
    body = raw.lstrip("\ufeff\r\n ")
    if body.startswith('"Section Title"'):
        body = body.split("\n", 1)[1] if "\n" in body else ""
    out = []
    for m in _SRLC_CSV_RECORD.finditer(body):
        rec = {k: (v or "").strip() for k, v in m.groupdict().items()}
        rec["description_text"] = html_to_text(m.group("description"))
        out.append(rec)
    return out


def fetch_srlc_csv(drug_name_id: str) -> tuple[list[dict], str, str]:
    """Return (records, source_url, fetched_at). Raises FDAError."""
    url = SRLC_CSV_URL.format(id=drug_name_id)
    try:
        resp = requests.get(url, headers=HEADERS, timeout=TIMEOUT_S)
        resp.raise_for_status()
    except requests.RequestException as e:
        raise FDAError(f"SrLC CSV download failed: {e}") from e
    finally:
        time.sleep(REQUEST_DELAY_S)
    if "csv" not in resp.headers.get("content-type", "").lower():
        raise FDAError("SrLC Download Data did not return a CSV.")
    records = parse_srlc_csv(resp.text)
    if not records and len(resp.text) > 300:
        raise FDAError("SrLC CSV could not be parsed.")
    return records, url, now_utc_iso()


_SUPPL_HEADER = re.compile(r"(\d{2}/\d{2}/\d{4})\s*\(\s*([A-Z]+)-(\d+)\s*\)")


def fetch_srlc_supplement_dates(drug_name_id: str) -> tuple[dict[str, str], str]:
    """Map supplement number -> approval date (MM/DD/YYYY) from the detail page.

    The Download Data CSV has no dates, so the detail page's accordion headers
    ("08/11/2026 (SUPPL-31)") are used to keep only changes in the date range.
    Returns (mapping, source_url). Raises FDAError.
    """
    url = SRLC_DETAIL_URL.format(id=drug_name_id)
    try:
        resp = requests.get(url, headers=HEADERS, timeout=TIMEOUT_S)
        resp.raise_for_status()
    except requests.RequestException as e:
        raise FDAError(f"SrLC detail page failed: {e}") from e
    finally:
        time.sleep(REQUEST_DELAY_S)
    text = BeautifulSoup(resp.text, "lxml").get_text(" ")
    mapping: dict[str, str] = {}
    for d, _kind, num in _SUPPL_HEADER.findall(text):
        mapping.setdefault(num, d)
    return mapping, url


def get_srlc_changes(
    drug_name_id: str,
    date_from: date | None = None,
    date_to: date | None = None,
    on_date: str | None = None,
) -> dict:
    """Changed-label text for one SrLC drug, restricted to the wanted supplements.

    Filtering: if date_from/date_to given, keep supplements approved in range;
    elif on_date given (an uploaded row's supplement date), keep that date's
    supplements; otherwise keep everything. If dates can't be resolved for a
    record it is kept and flagged date_unknown.

    Returns {"records": [...], "csv_url", "detail_url", "fetched_at", "error"}.
    """
    result = {"records": [], "csv_url": SRLC_CSV_URL.format(id=drug_name_id),
              "detail_url": SRLC_DETAIL_URL.format(id=drug_name_id),
              "fetched_at": now_utc_iso(), "error": "", "total_records": 0}
    try:
        records, csv_url, fetched_at = fetch_srlc_csv(drug_name_id)
    except FDAError as e:
        result["error"] = str(e)
        return result
    result.update(csv_url=csv_url, fetched_at=fetched_at, total_records=len(records))
    try:
        sub_dates, _ = fetch_srlc_supplement_dates(drug_name_id)
    except FDAError:
        sub_dates = {}

    d_on = _parse_date(on_date) if on_date else None
    kept = []
    for rec in records:
        sd = sub_dates.get(rec["sub_no"], "")
        rec["sub_date"] = sd
        pd_ = _parse_date(sd)
        rec["date_unknown"] = pd_ is None
        if pd_ is None:
            keep = True
        elif date_from or date_to:
            keep = (date_from is None or pd_ >= date_from) and (date_to is None or pd_ <= date_to)
        elif d_on:
            keep = pd_ == d_on
        else:
            keep = True
        if keep:
            kept.append(rec)
    result["records"] = kept
    return result


def srlc_changes_to_text(records: list[dict]) -> str:
    blocks = []
    for r in records:
        head = r["section"] + (f" — {r['title']}" if r.get("title") else "")
        meta = f"SUPPL-{r['sub_no']}" + (f" ({r['sub_date']})" if r.get("sub_date") else "")
        blocks.append(f"{head} [{meta}]\n{r['description_text']}")
    return "\n\n---\n\n".join(blocks)


# ─── 3. Legacy HTML detail-page scraper (fallback only) ──────────────────────
PRIORITY_SECTION_KEYWORDS = [
    "warning", "precaution", "hazard", "handling", "disposal",
    "pregnancy", "lactation", "reproductive", "carcinogen",
    "mutageni", "genotoxic", "teratogen", "special handling",
    "storage and handling", "how supplied",
]


def fetch_label_page(url: str) -> tuple[str, str]:
    """Returns (extracted_text, error_message). Unchanged legacy scraper; used
    only when the Download Data CSV is unavailable. Note: picks up SrLC's
    glossary text, which can create false keyword hits."""
    if not url or url in ("None", "nan", ""):
        return "", "No URL provided"
    try:
        resp = requests.get(url, headers=HEADERS, timeout=20)
        resp.raise_for_status()
        soup = BeautifulSoup(resp.content, "html.parser")
        for tag in soup(["script", "style", "nav", "footer", "header", "noscript"]):
            tag.decompose()

        priority_blocks = []
        for heading in soup.find_all(["h1", "h2", "h3", "h4", "strong", "b", "p"]):
            heading_text = heading.get_text(separator=" ").lower()
            if any(kw in heading_text for kw in PRIORITY_SECTION_KEYWORDS):
                block = heading.get_text(separator=" ") + "\n"
                for sib in heading.next_siblings:
                    if hasattr(sib, "name") and sib.name in ["h1", "h2", "h3", "h4"]:
                        break
                    if hasattr(sib, "get_text"):
                        block += sib.get_text(separator=" ") + "\n"
                    elif isinstance(sib, str):
                        block += sib
                    if len(block) > 2500:
                        break
                priority_blocks.append(block[:2500])

        if priority_blocks:
            text = "\n\n---\n\n".join(priority_blocks[:8])
        else:
            text = soup.get_text(separator="\n")

        text = re.sub(r"\n{3,}", "\n\n", text)
        text = re.sub(r" {2,}", " ", text)
        return text[:8000], ""
    except requests.exceptions.Timeout:
        return "", "Request timed out — try Manual Fetch in review."
    except requests.exceptions.HTTPError as e:
        return "", f"HTTP {e.response.status_code} — page may require JavaScript or authentication."
    except Exception as e:
        return "", f"Fetch error: {str(e)}"
    finally:
        time.sleep(REQUEST_DELAY_S)


# ─── 4. openFDA current-label lookup ─────────────────────────────────────────
def _openfda_query(search: str, api_key: str | None) -> tuple[dict | None, str, str]:
    """Return (first_result_or_None, public_url_without_key, error)."""
    params = {"search": search, "limit": 1, "sort": "effective_time:desc"}
    public_url = OPENFDA_LABEL_URL + "?" + urlencode(params)
    if api_key:
        params["api_key"] = api_key
    try:
        resp = requests.get(OPENFDA_LABEL_URL, params=params, timeout=TIMEOUT_S)
    except requests.RequestException as e:
        return None, public_url, f"openFDA request failed: {e}"
    finally:
        time.sleep(REQUEST_DELAY_S)
    if resp.status_code == 404:
        return None, public_url, ""  # openFDA's "no matches"
    if resp.status_code == 429:
        return None, public_url, "openFDA rate limit hit (add a free API key)."
    if resp.status_code != 200:
        return None, public_url, f"openFDA HTTP {resp.status_code}"
    try:
        results = resp.json().get("results") or []
    except ValueError:
        return None, public_url, "openFDA returned invalid JSON"
    return (results[0] if results else None), public_url, ""


def _first_ingredient(active_ingredient: str, normalizer=None) -> str:
    first = re.split(r"[;,/]", active_ingredient or "")[0].strip()
    if normalizer:
        first = normalizer(first) or first
    return re.sub(r"[^A-Za-z0-9 \-]", " ", first).strip()


def fetch_openfda_handling(
    appl_no: str,
    appl_type: str,
    active_ingredient: str,
    api_key: str | None = None,
    normalizer=None,
) -> dict:
    """Current-label handling / carcinogenicity / reproductive text from openFDA.

    Tries application number first (e.g. NDA208400), then generic name.
    Returns {"matched_by", "query_url", "fetched_at", "set_id", "effective_time",
             "brand_name", "generic_name", "application_number", "fields": {f: text},
             "context": {f: text}, "error"}.
    """
    out = {"matched_by": "", "query_url": "", "fetched_at": now_utc_iso(), "set_id": "",
           "effective_time": "", "brand_name": "", "generic_name": "", "manufacturer": "",
           "application_number": "", "fields": {}, "context": {}, "error": ""}
    attempts = []
    num = re.sub(r"\D", "", str(appl_no or ""))
    typ = re.sub(r"[^A-Z]", "", str(appl_type or "").upper())
    if num and typ:
        attempts.append(("application_number", f'openfda.application_number:"{typ}{num}"'))
    ing = _first_ingredient(active_ingredient, normalizer)
    if ing:
        attempts.append(("generic_name", f'openfda.generic_name:"{ing}"'))
    if not attempts:
        out["error"] = "No application number or ingredient to search openFDA"
        return out

    errors = []
    for how, search in attempts:
        res, url, err = _openfda_query(search, api_key)
        out["query_url"] = url
        if err:
            errors.append(err)
            if "rate limit" in err:
                break
            continue
        if res:
            ofda = res.get("openfda", {})
            out.update(
                matched_by=how,
                fetched_at=now_utc_iso(),
                set_id=res.get("set_id", ""),
                effective_time=res.get("effective_time", ""),
                brand_name="; ".join(ofda.get("brand_name", [])),
                generic_name="; ".join(ofda.get("generic_name", [])),
                manufacturer="; ".join(ofda.get("manufacturer_name", [])),
                application_number="; ".join(ofda.get("application_number", [])),
                fields={f: clean_text("\n".join(res[f])) for f in OPENFDA_SCAN_FIELDS if res.get(f)},
                context={f: clean_text("\n".join(res[f])) for f in OPENFDA_CONTEXT_FIELDS if res.get(f)},
            )
            return out
    out["error"] = "; ".join(errors) if errors else "No openFDA label found"
    return out


def openfda_fields_to_text(fields: dict) -> str:
    return "\n\n---\n\n".join(f"{k.replace('_', ' ').upper()}\n{v}" for k, v in fields.items())
