"""RxLabelWatch pipeline: crosswalk, enrichment (SrLC CSV + openFDA), QA export.

No Streamlit imports; app.py passes in cached fetchers and a progress callback.
"""
from __future__ import annotations

import io
import re
from dataclasses import dataclass
from datetime import date
from typing import Callable

import pandas as pd

import fda_sources as fda
from rxlw_core import (
    autodetect_col, find_best_matches, ingredient_agreement, normalize_drug_name, scan_for_hazards,
)

SRC_SRLC = "SrLC change text"
SRC_SRLC_PAGE = "SrLC detail page (fallback)"
SRC_OPENFDA = "openFDA current label"
XLSX_CELL_MAX = 32000


@dataclass
class Fetchers:
    """Injectable network functions (app.py passes st.cache_data-wrapped ones)."""
    srlc_changes: Callable = fda.get_srlc_changes
    openfda: Callable = fda.fetch_openfda_handling
    label_page: Callable = fda.fetch_label_page


# ─── Crosswalk (same loop as the original app, plus extra metadata) ──────────
def run_matching(
    fdf: pd.DataFrame,
    f_col: str,
    sdf: pd.DataFrame,
    s_col: str,
    s_url_col: str | None,
    s_date_col: str | None,
    s_type_col: str | None,
    threshold: int = 78,
    extra_name_col: str | None = None,
    progress: Callable[[float, str], None] | None = None,
    max_per_drug: int | None = None,
    require_ingredient_agreement: bool = True,
) -> list[dict]:
    """Fuzzy-match formulary names to SrLC rows with find_best_matches.

    extra_name_col optionally matches a second SrLC column too (e.g. brand
    'Drug Name' in addition to 'Active Ingredient'); results are merged per
    SrLC row (best score kept). max_per_drug=None keeps every SrLC row that
    matches (the original app capped at 5). With require_ingredient_agreement,
    a fuzzy hit must also pass ingredient_agreement (rejects look-alike
    generics such as amoxicillin/oxacillin).
    """
    formulary_drugs = fdf[f_col].dropna().astype(str).unique().tolist()
    name_cols = [s_col] + ([extra_name_col] if extra_name_col and extra_name_col != s_col else [])
    norms_by_col = {c: [normalize_drug_name(d) for d in sdf[c].fillna("").astype(str).tolist()] for c in name_cols}

    brand_col = "Drug Name" if "Drug Name" in sdf.columns else s_col
    ing_col = autodetect_col(sdf, ["active ingredient", "ingredient"])
    appno_col = autodetect_col(sdf, ["application number", "appl no", "application no"])
    apptype_col = autodetect_col(sdf, ["application type", "appl type"])
    id_col = "DrugNameID" if "DrugNameID" in sdf.columns else None

    def cell(row, col):
        if not col:
            return ""
        v = str(row[col])
        return "" if v in ("", "None", "nan") else v

    matches = []
    n = max(len(formulary_drugs), 1)
    for i, drug in enumerate(formulary_drugs):
        if progress:
            progress((i + 1) / n, f"Matching: {drug[:60]}")
        drug_norm = normalize_drug_name(drug)
        if not drug_norm:
            continue
        best: dict[int, tuple[int, str]] = {}
        for c in name_cols:
            for idx, score in find_best_matches(drug_norm, norms_by_col[c], threshold=threshold, limit=None):
                if require_ingredient_agreement and not ingredient_agreement(drug_norm, norms_by_col[c][idx]):
                    continue
                if idx not in best or score > best[idx][0]:
                    best[idx] = (score, c)
        ranked = sorted(best.items(), key=lambda kv: -kv[1][0])
        if max_per_drug:
            ranked = ranked[:max_per_drug]
        for idx, (score, matched_col) in ranked:
            row = sdf.iloc[idx]
            url = cell(row, s_url_col)
            appl_no = cell(row, appno_col)
            if appl_no.replace(".0", "").isdigit():
                appl_no = appl_no.replace(".0", "").zfill(6)
            matches.append({
                "formulary_drug": drug,
                "slrc_drug": cell(row, brand_col) or cell(row, matched_col),
                "matched_on": f"{matched_col}: {cell(row, matched_col)}",
                "match_score": score,
                "date": cell(row, s_date_col),
                "change_type": cell(row, s_type_col),
                "url": url,
                "slrc_row": int(idx),
                "drug_name_id": cell(row, id_col) or fda.drug_name_id_from_url(url),
                "active_ingredient": cell(row, ing_col),
                "appl_no": appl_no,
                "appl_type": cell(row, apptype_col),
                "label_text": "",
                "openfda_text": "",
                "text_source": "",
                "fetch_error": "",
                "srlc_records": [],
                "srlc_csv_url": "",
                "srlc_fetched_at": "",
                "openfda": {},
                "risk_level": "none",
                "risk_by_source": {},
                "flag_sources": {},
                "niosh": {},
                "usp800": {},
                "boxed_warning": False,
            })
    if progress:
        progress(1.0, f"Matched {len(matches)} drug–label pairs.")
    return matches


# ─── Enrichment + hazard scan ────────────────────────────────────────────────
def _flag_labels(findings: dict) -> list[str]:
    labels = ["Boxed warning"] if findings["boxed_warning"] else []
    labels += [f"NIOSH: {c}" for c in findings["niosh"]]
    labels += [f"USP 800: {c}" for c in findings["usp800"]]
    return labels


def apply_scan(m: dict, include_openfda_in_risk: bool = False) -> None:
    """Run scan_for_hazards per source and on the combined text; record which
    source triggered each flag."""
    per_source = {}
    if m.get("label_text"):
        per_source[m.get("text_source") or SRC_SRLC] = scan_for_hazards(m["label_text"])
    if m.get("openfda_text"):
        per_source[SRC_OPENFDA] = scan_for_hazards(m["openfda_text"])

    parts = [m.get("label_text", "")]
    if include_openfda_in_risk:
        parts.append(m.get("openfda_text", ""))
    combined = scan_for_hazards("\n\n".join(p for p in parts if p))

    m["risk_level"] = combined["risk_level"]
    m["niosh"] = combined["niosh"]
    m["usp800"] = combined["usp800"]
    m["boxed_warning"] = combined["boxed_warning"]
    m["risk_by_source"] = {src: f["risk_level"] for src, f in per_source.items()}
    m["openfda_in_risk"] = include_openfda_in_risk
    flag_sources: dict[str, list[str]] = {}
    for src, f in per_source.items():
        for lab in _flag_labels(f):
            flag_sources.setdefault(lab, []).append(src)
    m["flag_sources"] = flag_sources


def enrich_match(
    m: dict,
    fetchers: Fetchers | None = None,
    date_from: date | None = None,
    date_to: date | None = None,
    use_openfda: bool = True,
    api_key: str | None = None,
    include_openfda_in_risk: bool = False,
) -> dict:
    """Fill a match with SrLC change text (CSV), openFDA handling text, and flags."""
    f = fetchers or Fetchers()
    errors = []
    m["label_text"], m["text_source"], m["srlc_records"] = "", "", []

    dnid = m.get("drug_name_id") or fda.drug_name_id_from_url(m.get("url", ""))
    if dnid:
        on_date = None if (date_from or date_to) else (m.get("date") or None)
        ch = f.srlc_changes(dnid, date_from, date_to, on_date)
        m["srlc_csv_url"], m["srlc_fetched_at"] = ch["csv_url"], ch["fetched_at"]
        recs = ch["records"]
        if ch["error"]:
            errors.append(ch["error"])
        if recs:
            m["srlc_records"] = recs
            m["label_text"] = fda.srlc_changes_to_text(recs)
            m["text_source"] = SRC_SRLC
    if not m["label_text"] and m.get("url"):
        text, err = f.label_page(m["url"])
        if text:
            m["label_text"], m["text_source"] = text, SRC_SRLC_PAGE
            m["srlc_fetched_at"] = fda.now_utc_iso()
        elif err:
            errors.append(err)

    if use_openfda:
        ing_raw = m.get("active_ingredient") or m.get("slrc_drug", "")
        first_ing = re.split(r"[;,/]", ing_raw)[0].strip()
        ing = normalize_drug_name(first_ing) or first_ing
        o = f.openfda(m.get("appl_no", ""), m.get("appl_type", ""), ing, api_key)
        m["openfda"] = o
        m["openfda_text"] = fda.openfda_fields_to_text(o.get("fields", {}))
        if o.get("error"):
            errors.append(o["error"])
    m["fetch_error"] = " | ".join(errors)
    apply_scan(m, include_openfda_in_risk)
    return m


def enrich_all(matches, progress=None, **kw) -> list[dict]:
    n = max(len(matches), 1)
    for j, m in enumerate(matches):
        if progress:
            progress((j + 1) / n, f"Fetching FDA data: {m['slrc_drug'][:60]}")
        enrich_match(m, **kw)
    return matches


# ─── QA source-data export ───────────────────────────────────────────────────
def _fmt_hits(d: dict) -> str:
    return "; ".join(f"{c}: {', '.join(k)}" for c, k in d.items())


def _trim(v):
    if isinstance(v, str):
        v = fda.clean_text(v)
        return v if len(v) <= XLSX_CELL_MAX else v[:XLSX_CELL_MAX] + " …[truncated]"
    return v


def build_qa_workbook(slrc_df: pd.DataFrame | None, matches: list[dict], run_info: dict) -> bytes:
    """xlsx with every source record behind the flags, for QA before sharing."""
    match_rows, change_rows, ofda_rows = [], [], []
    for m in matches:
        o = m.get("openfda") or {}
        match_rows.append({
            "Formulary Drug": m["formulary_drug"], "SrLC Drug": m["slrc_drug"],
            "Matched On": m.get("matched_on", ""), "Match Score": m["match_score"],
            "Supplement Date": m.get("date", ""), "Application": f"{m.get('appl_type','')} {m.get('appl_no','')}".strip(),
            "Active Ingredient": m.get("active_ingredient", ""), "DrugNameID": m.get("drug_name_id", ""),
            "Overall Risk": m["risk_level"].upper(),
            "Risk by Source": "; ".join(f"{s}: {r.upper()}" for s, r in m.get("risk_by_source", {}).items()),
            "Flags (source)": "; ".join(
                f"{lab} [{', '.join(s + ('' if m.get('openfda_in_risk') or s != SRC_OPENFDA else ' (context)') for s in srcs)}]"
                for lab, srcs in m.get("flag_sources", {}).items()),
            "openFDA Counted in Risk": "yes" if m.get("openfda_in_risk") else "no (context only)",
            "NIOSH Signals": _fmt_hits(m.get("niosh", {})), "USP 800 Signals": _fmt_hits(m.get("usp800", {})),
            "Boxed Warning": "YES" if m.get("boxed_warning") else "",
            "SrLC Text Source": m.get("text_source", ""), "SrLC Change Rows": len(m.get("srlc_records", [])),
            "openFDA Label": o.get("brand_name", "") + (f" ({o.get('effective_time')})" if o.get("effective_time") else ""),
            "Errors": m.get("fetch_error", ""),
            "SrLC Detail URL": m.get("url", ""), "SrLC CSV URL": m.get("srlc_csv_url", ""),
            "openFDA Query URL": o.get("query_url", ""),
        })
        for r in m.get("srlc_records", []):
            change_rows.append({
                "Formulary Drug": m["formulary_drug"], "SrLC Drug": r.get("drug_name", m["slrc_drug"]),
                "DrugNameID": m.get("drug_name_id", ""), "Application": f"{r.get('appl_type','')} {r.get('appl_no','')}",
                "Supplement": f"{r.get('sub_type','').strip()}-{r.get('sub_no','')}", "Supplement Date": r.get("sub_date", ""),
                "Section": r.get("section", ""), "Title": r.get("title", ""), "Changed Text": r.get("description_text", ""),
                "Label PDF": r.get("label_url", ""), "Source URL": m.get("srlc_csv_url", ""),
                "Fetched At (UTC)": m.get("srlc_fetched_at", ""),
            })
        if m.get("text_source") == SRC_SRLC_PAGE:
            change_rows.append({
                "Formulary Drug": m["formulary_drug"], "SrLC Drug": m["slrc_drug"],
                "DrugNameID": m.get("drug_name_id", ""), "Section": "(HTML page scrape fallback)",
                "Changed Text": m.get("label_text", ""), "Source URL": m.get("url", ""),
                "Fetched At (UTC)": m.get("srlc_fetched_at", ""),
            })
        base = {
            "Formulary Drug": m["formulary_drug"], "SrLC Drug": m["slrc_drug"],
            "Matched By": o.get("matched_by", ""), "openFDA Brand": o.get("brand_name", ""),
            "openFDA Generic": o.get("generic_name", ""), "Manufacturer": o.get("manufacturer", ""),
            "Application Number": o.get("application_number", ""), "set_id": o.get("set_id", ""),
            "effective_time": o.get("effective_time", ""), "Query URL": o.get("query_url", ""),
            "Fetched At (UTC)": o.get("fetched_at", ""),
        }
        fields = list((o.get("fields") or {}).items())
        ctx = list((o.get("context") or {}).items())
        if not fields and not ctx:
            ofda_rows.append({**base, "Field": "", "Scanned": "", "Text": o.get("error", "") or "(openFDA not queried)"})
        for k, v in fields:
            ofda_rows.append({**base, "Field": k, "Scanned": "yes", "Text": v})
        for k, v in ctx:
            ofda_rows.append({**base, "Field": k, "Scanned": "no (context)", "Text": v})

    sheets = {
        "Run Info": pd.DataFrame(list(run_info.items()), columns=["Item", "Value"]),
        "SrLC Results": slrc_df if slrc_df is not None else pd.DataFrame(),
        "Matches": pd.DataFrame(match_rows),
        "SrLC Change Text": pd.DataFrame(change_rows),
        "openFDA Handling": pd.DataFrame(ofda_rows),
    }
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter", engine_kwargs={"options": {"strings_to_urls": False}}) as xw:
        for name, df in sheets.items():
            df = df.copy()
            for c in df.columns:
                if df[c].dtype == object or str(df[c].dtype).startswith("str"):
                    df[c] = df[c].map(_trim)
            df.to_excel(xw, sheet_name=name, index=False)
            ws = xw.sheets[name]
            ws.freeze_panes(1, 0)
            for ci, c in enumerate(df.columns):
                width = 80 if c in ("Changed Text", "Text", "Value") else min(max(len(str(c)) + 2, 14), 45)
                ws.set_column(ci, ci, width)
    return buf.getvalue()
