import streamlit as st
import pandas as pd
import io
import re
from datetime import date, datetime, timedelta
import xlsxwriter

import fda_sources as fda
import pipeline as pl
from rxlw_core import autodetect_col, highlight_keywords

# ─── Page Config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="RxLabelWatch | Formulary Safety Monitor",
    page_icon="⚠️",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─── CSS ───────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;600&family=IBM+Plex+Sans:ital,wght@0,300;0,400;0,600;1,400&display=swap');
html, body, [class*="css"] { font-family: 'IBM Plex Sans', sans-serif; }
h1, h2, h3, h4 { font-family: 'IBM Plex Mono', monospace !important; }
.stApp { background-color: #0d1117; color: #e6edf3; }
.block-container { padding-top: 1.5rem; }
div[data-testid="stSidebar"] { background-color: #161b22; border-right: 1px solid #30363d; }
.rx-card { background: #161b22; border: 1px solid #30363d; border-radius: 8px; padding: 14px 18px; margin-bottom: 10px; }
.tag { display: inline-block; background: #21262d; border: 1px solid #30363d; border-radius: 4px; padding: 2px 8px; font-size: 0.72rem; font-family: 'IBM Plex Mono', monospace; margin-right: 4px; color: #8b949e; }
.tag.niosh { background: #3d1f1f; border-color: #f85149; color: #f85149; }
.tag.usp800 { background: #2d2000; border-color: #d29922; color: #d29922; }
.tag.boxed { background: #1f0a0a; border-color: #ff7b72; color: #ff7b72; font-weight: 600; }
.hl-niosh { background: rgba(248,81,73,0.25); border-radius: 3px; padding: 1px 3px; font-weight: 600; }
.hl-usp800 { background: rgba(210,153,34,0.25); border-radius: 3px; padding: 1px 3px; font-weight: 600; }
.section-header { font-family: 'IBM Plex Mono', monospace; font-size: 0.75rem; color: #8b949e; letter-spacing: 0.08em; text-transform: uppercase; margin-bottom: 0.25rem; }
.label-block { background: #0d1117; border: 1px solid #21262d; border-radius: 6px; padding: 12px 16px; font-size: 0.83rem; line-height: 1.7; white-space: pre-wrap; color: #c9d1d9; max-height: 400px; overflow-y: auto; }
</style>
""", unsafe_allow_html=True)


# ─── Cached FDA access ─────────────────────────────────────────────────────────
# Failures raise inside the cached function so they are NOT cached; the
# wrappers turn them back into error results.
class _NoCache(Exception):
    def __init__(self, payload):
        self.payload = payload


@st.cache_data(ttl=6 * 3600, show_spinner=False)
def cached_srlc_search(date_from: date, date_to: date) -> pd.DataFrame:
    return fda.search_srlc_by_date(date_from, date_to)


@st.cache_data(ttl=24 * 3600, show_spinner=False)
def _cached_srlc_changes(drug_name_id, date_from, date_to, on_date):
    res = fda.get_srlc_changes(drug_name_id, date_from, date_to, on_date)
    if res["error"] and not res["records"]:
        raise _NoCache(res)
    return res


def cached_srlc_changes(drug_name_id, date_from=None, date_to=None, on_date=None):
    try:
        return _cached_srlc_changes(drug_name_id, date_from, date_to, on_date)
    except _NoCache as e:
        return e.payload


@st.cache_data(ttl=24 * 3600, show_spinner=False)
def _cached_openfda(appl_no, appl_type, ingredient, _api_key):
    res = fda.fetch_openfda_handling(appl_no, appl_type, ingredient, _api_key)
    if res["error"] and res["error"] != "No openFDA label found":
        raise _NoCache(res)
    return res


def cached_openfda(appl_no, appl_type, ingredient, api_key=None):
    try:
        return _cached_openfda(appl_no, appl_type, ingredient, api_key)
    except _NoCache as e:
        return e.payload


@st.cache_data(ttl=24 * 3600, show_spinner=False)
def _cached_label_page(url):
    text, err = fda.fetch_label_page(url)
    if err:
        raise _NoCache((text, err))
    return text, err


def cached_label_page(url):
    try:
        return _cached_label_page(url)
    except _NoCache as e:
        return e.payload


FETCHERS = pl.Fetchers(srlc_changes=cached_srlc_changes, openfda=cached_openfda, label_page=cached_label_page)


def get_openfda_key():
    """Optional free openFDA key from .streamlit/secrets.toml or Streamlit Cloud secrets."""
    try:
        key = st.secrets.get("OPENFDA_API_KEY", "")
    except Exception:  # no secrets file at all
        key = ""
    return str(key).strip() or None


def enrich_kwargs():
    rng = st.session_state.srlc_range
    return dict(
        fetchers=FETCHERS,
        date_from=rng[0] if rng else None,
        date_to=rng[1] if rng else None,
        use_openfda=st.session_state.use_openfda,
        api_key=get_openfda_key(),
        include_openfda_in_risk=st.session_state.openfda_in_risk,
    )


def qa_run_info() -> dict:
    rng = st.session_state.srlc_range
    sdf = st.session_state.slrc_df
    matches = st.session_state.matches or []
    info = {
        "Generated (local)": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "SrLC source": "FDA SrLC Date Search (fetched by app)" if st.session_state.slrc_source == "fetch" else "Uploaded SrLC export",
        "Date range": f"{rng[0]:%m/%d/%Y} – {rng[1]:%m/%d/%Y}" if rng else "(uploaded file)",
        "SrLC search URL": fda.SRLC_SEARCH_URL if rng else "",
        "SrLC sections searched": ", ".join(fda.SRLC_SECTIONS) if rng else "",
        "SrLC rows": len(sdf) if sdf is not None else 0,
        "SrLC fetched at (UTC)": (sdf["fetched_at"].iloc[0] if sdf is not None and "fetched_at" in sdf.columns and len(sdf) else ""),
        "Formulary drugs": int(st.session_state.formulary_df[st.session_state.f_name_col].nunique()) if st.session_state.formulary_df is not None and st.session_state.f_name_col else "",
        "Matched pairs": len(matches),
        "Formulary drugs matched": len({m["formulary_drug"] for m in matches}),
        "Match threshold": st.session_state.get("last_threshold", ""),
        "openFDA used": "yes" if st.session_state.use_openfda else "no",
        "openFDA API key configured": "yes" if get_openfda_key() else "no",
        "openFDA text counted in overall risk": "yes" if st.session_state.openfda_in_risk else "no",
        "openFDA scanned fields": ", ".join(fda.OPENFDA_SCAN_FIELDS),
        "openFDA endpoint": fda.OPENFDA_LABEL_URL,
    }
    return info


def qa_download_button(key: str):
    data = pl.build_qa_workbook(st.session_state.slrc_df, st.session_state.matches or [], qa_run_info())
    st.download_button(
        "🧾 Download QA Source Data (.xlsx)",
        data=data,
        file_name=f"rxlabelwatch_source_data_{datetime.now():%Y%m%d}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
        key=key,
        help="SrLC results, the per-drug SrLC change text, and openFDA handling text for every match, with source URLs and fetch times.",
    )


# ─── Session State Init ────────────────────────────────────────────────────────
def init_state():
    defaults = {
        "formulary_df": None,
        "slrc_df": None,
        "slrc_source": "fetch",
        "srlc_range": None,
        "f_name_col": None,
        "s_name_col": None,
        "s_extra_name_col": None,
        "s_url_col": None,
        "s_date_col": None,
        "s_type_col": None,
        "matches": None,
        "use_openfda": True,
        "openfda_in_risk": True,
        "ph_decisions": {},
        "ph_notes": {},
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v

init_state()

# ─── Sidebar ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown('<div class="section-header">RxLabelWatch</div>', unsafe_allow_html=True)
    st.markdown("## ⚠️ Safety Monitor")
    st.divider()
    page = st.radio(
        "Navigate",
        ["1. Load Data", "2. Run Matching", "3. Hazard Review", "4. Export"],
        label_visibility="collapsed",
        key="nav",
    )
    st.divider()
    if st.session_state.matches:
        matches = st.session_state.matches
        total = len(matches)
        high = sum(1 for m in matches if m["risk_level"] == "high")
        medium = sum(1 for m in matches if m["risk_level"] == "medium")
        reviewed = len(st.session_state.ph_decisions)
        st.metric("Matched Drugs", total)
        st.metric("🔴 High Risk", high)
        st.metric("🟡 Medium Risk", medium)
        st.metric("Reviewed", f"{reviewed}/{total}")
        st.divider()
    st.markdown('<div class="section-header">Risk Legend</div>', unsafe_allow_html=True)
    st.markdown("""
🔴 **High** — NIOSH criteria or USP 800 language found
🟡 **Medium** — Pregnancy/lactation, organ toxicity, PPE
🟢 **Low** — Minor keyword signals; review advised
⚠ **None** — No hazardous signals detected
""")
    st.markdown('<div class="section-header">About</div>', unsafe_allow_html=True)
    st.caption("Matches your Epic formulary to FDA Safety-related Labeling Changes (SrLC), adds current openFDA handling text, and scans for NIOSH + USP 800 hazardous drug signals.")
    st.caption("openFDA API key: " + ("configured ✓" if get_openfda_key() else "not set (fine for small runs)"))

# ═══════════════════════════════════════════════════════════════════════════════
# Page 1 — Load Data
# ═══════════════════════════════════════════════════════════════════════════════
if page == "1. Load Data":
    st.markdown("## 📁 Load Data")
    st.caption("Upload your Epic formulary export, then pull FDA Safety-related Labeling Changes (SrLC) for a date range — or upload an SrLC export.")

    col1, col2 = st.columns(2)

    with col1:
        st.markdown("### 💊 Formulary (Epic Export)")
        formulary_file = st.file_uploader(
            "Drop your Epic formulary export here",
            type=["xlsx", "xls", "csv"],
            key="formulary_upload",
        )
        if formulary_file:
            try:
                df = pd.read_csv(formulary_file) if formulary_file.name.endswith(".csv") else pd.read_excel(formulary_file)
                st.session_state.formulary_df = df
            except Exception as e:
                st.error(f"Error loading file: {e}")
        if st.session_state.formulary_df is not None:
            df = st.session_state.formulary_df
            st.success(f"Loaded **{len(df):,} rows**, {len(df.columns)} columns")
            st.dataframe(df.head(8), use_container_width=True)

    with col2:
        st.markdown("### 📋 FDA Safety Labeling Changes (SrLC)")
        mode = st.radio(
            "SrLC source",
            ["Fetch from FDA (date range)", "Upload SrLC export"],
            horizontal=True,
            key="srlc_mode",
        )
        if mode == "Fetch from FDA (date range)":
            today = date.today()
            rng = st.date_input(
                "Supplement approval date range",
                value=(today - timedelta(days=183), today),
                max_value=today,
                format="MM/DD/YYYY",
                key="srlc_dates",
                help="Searches SrLC's Date Search across all label sections (boxed warning, contraindications, warnings & precautions, adverse reactions, drug interactions, specific populations, patient info).",
            )
            if st.button("🔎 Fetch SrLC changes", type="primary", use_container_width=True, key="fetch_srlc"):
                if not isinstance(rng, (tuple, list)) or len(rng) != 2:
                    st.error("Pick both a start and an end date.")
                else:
                    d_from, d_to = rng
                    with st.spinner(f"Searching FDA SrLC {d_from:%m/%d/%Y} – {d_to:%m/%d/%Y}…"):
                        try:
                            sdf = cached_srlc_search(d_from, d_to)
                            st.session_state.slrc_df = sdf
                            st.session_state.slrc_source = "fetch"
                            st.session_state.srlc_range = (d_from, d_to)
                            st.session_state.matches = None
                        except fda.FDAError as e:
                            st.error(f"FDA SrLC search failed: {e}  You can still upload an SrLC export instead.")
            st.caption("Source: [FDA SrLC database](https://www.accessdata.fda.gov/scripts/cder/safetylabelingchanges/index.cfm)")
        else:
            st.caption("Backup path. Download from: [FDA Safety Labeling Changes](https://www.accessdata.fda.gov/scripts/cder/safetylabelingchanges/index.cfm)")
            slrc_file = st.file_uploader(
                "Drop your FDA SrLC export here",
                type=["xlsx", "xls", "csv"],
                key="slrc_upload",
            )
            if slrc_file:
                try:
                    df = pd.read_csv(slrc_file) if slrc_file.name.endswith(".csv") else pd.read_excel(slrc_file)
                    st.session_state.slrc_df = df
                    st.session_state.slrc_source = "upload"
                    st.session_state.srlc_range = None
                except Exception as e:
                    st.error(f"Error loading file: {e}")

        if st.session_state.slrc_df is not None:
            sdf = st.session_state.slrc_df
            if st.session_state.slrc_source == "fetch" and st.session_state.srlc_range:
                r0, r1 = st.session_state.srlc_range
                st.success(f"Fetched **{len(sdf):,} SrLC rows** for {r0:%m/%d/%Y} – {r1:%m/%d/%Y}")
            else:
                st.success(f"Loaded **{len(sdf):,} rows**, {len(sdf.columns)} columns")
            st.dataframe(sdf.drop(columns=[c for c in ("fetched_at", "source_url") if c in sdf.columns]).head(8), use_container_width=True)

    if st.session_state.formulary_df is not None and st.session_state.slrc_df is not None:
        st.divider()
        st.markdown("### ⚙️ Column Mapping")
        st.caption("We auto-detected the most likely columns. Adjust if needed.")

        fdf = st.session_state.formulary_df
        sdf = st.session_state.slrc_df
        fcols = list(fdf.columns)
        scols = [c for c in sdf.columns if c not in ("fetched_at", "source_url")]
        fetched = st.session_state.slrc_source == "fetch"
        sdf_map = sdf[scols]  # don't auto-detect our provenance columns

        # Auto-detect defaults
        f_name_default = autodetect_col(fdf, ["drug", "name", "generic", "description", "medication", "item", "title"]) or fcols[0]
        if fetched and "Active Ingredient" in scols:
            # Formularies list generics; SrLC 'Drug Name' is usually the brand.
            s_name_default = "Active Ingredient"
            s_extra_default = "Drug Name"
        else:
            s_name_default = autodetect_col(sdf_map, ["drug", "name", "proprietary", "nonproprietary", "generic", "product"]) or scols[0]
            s_extra_default = None
        s_url_default = autodetect_col(sdf_map, ["url", "link", "href", "web", "labeling"])
        s_date_default = autodetect_col(sdf_map, ["supplement date", "date", "submission", "approval", "effective"])
        s_type_default = None if fetched else autodetect_col(sdf_map, ["type", "supplement", "change", "category", "description"])

        c1, c2 = st.columns(2)
        with c1:
            st.markdown("**Formulary**")
            st.session_state.f_name_col = st.selectbox(
                "Drug name column", fcols,
                index=fcols.index(f_name_default),
                help="Column with generic drug names from Epic",
            )

        with c2:
            st.markdown("**SrLC Data**")
            ca, cb = st.columns(2)
            st.session_state.s_name_col = ca.selectbox(
                "Drug name column", scols,
                index=scols.index(s_name_default) if s_name_default in scols else 0,
            )
            extra_opts = ["(none)"] + scols
            s_extra_sel = cb.selectbox(
                "Also match on", extra_opts,
                index=extra_opts.index(s_extra_default) if s_extra_default in extra_opts else 0,
                help="Optional second SrLC name column (e.g. brand 'Drug Name').",
            )
            st.session_state.s_extra_name_col = None if s_extra_sel == "(none)" else s_extra_sel
            url_opts = ["(none)"] + scols
            s_url_idx = url_opts.index(s_url_default) if s_url_default in url_opts else 0
            s_url_sel = ca.selectbox("URL / link column", url_opts, index=s_url_idx)
            st.session_state.s_url_col = None if s_url_sel == "(none)" else s_url_sel

            date_opts = ["(none)"] + scols
            s_date_idx = date_opts.index(s_date_default) if s_date_default in date_opts else 0
            s_date_sel = cb.selectbox("Date column", date_opts, index=s_date_idx)
            st.session_state.s_date_col = None if s_date_sel == "(none)" else s_date_sel

            type_opts = ["(none)"] + scols
            s_type_idx = type_opts.index(s_type_default) if s_type_default in type_opts else 0
            s_type_sel = ca.selectbox("Change type column", type_opts, index=s_type_idx)
            st.session_state.s_type_col = None if s_type_sel == "(none)" else s_type_sel

        st.info("✅ Column mapping set. Go to **2. Run Matching** to find matches.")

# ═══════════════════════════════════════════════════════════════════════════════
# Page 2 — Run Matching
# ═══════════════════════════════════════════════════════════════════════════════
elif page == "2. Run Matching":
    st.markdown("## 🔍 Run Matching")

    if st.session_state.formulary_df is None or st.session_state.slrc_df is None:
        st.warning("Load the formulary and SrLC data in Step 1 first.")
        st.stop()

    c1, c2 = st.columns([2, 1])
    threshold = c1.slider(
        "Fuzzy match sensitivity",
        min_value=60, max_value=98, value=78, step=2,
        help="78 is a good default. Lower to catch more matches (more false positives); raise for stricter matching.",
    )
    fetch_pages = c2.toggle(
        "Auto-fetch FDA change details",
        value=True,
        help="Downloads each match's SrLC change data (CSV) and scans it. ~1 sec per match; cached for 24 h.",
    )
    st.session_state.use_openfda = c2.toggle(
        "Add openFDA handling text",
        value=st.session_state.use_openfda,
        help="Pulls the current label's How Supplied, Storage & Handling, References, Carcinogenesis, Pregnancy, Lactation and Reproductive Potential sections (SrLC doesn't track these).",
    )
    st.session_state.openfda_in_risk = c2.toggle(
        "Count openFDA text in risk level",
        value=st.session_state.openfda_in_risk,
        help="When off, the risk level reflects only the SrLC change text; openFDA hits are still listed by source.",
    )

    with st.expander("ℹ️ How matching works"):
        st.markdown("""
The app normalizes drug names on both sides before matching:
- Removes salt forms (HCl, sodium, sulfate, etc.)
- Removes dosage forms (tablet, injection, cream, etc.)
- Removes strengths (e.g., "2.5 mg", "1 gram/200 mL")
- Removes EPIC IV bag diluent phrases ("in dextrose 5% in water")

Then uses **token sort ratio** fuzzy matching — so "VANCOMYCIN HYDROCHLORIDE" and "vancomycin HCl 1g IVPB" will match correctly.

For each match, the app downloads SrLC's per-drug **Download Data** CSV and keeps only the supplements approved in your date range,
then looks up the current label on **openFDA** (by application number, falling back to generic name). Each flag shows which source triggered it.
""")

    if st.button("▶ Run Matching", use_container_width=True, type="primary", key="run_matching"):
        fdf = st.session_state.formulary_df
        sdf = st.session_state.slrc_df
        f_col = st.session_state.f_name_col or fdf.columns[0]
        s_col = st.session_state.s_name_col or sdf.columns[0]
        st.session_state.last_threshold = threshold

        progress = st.progress(0, text="Normalizing and matching…")
        matches = pl.run_matching(
            fdf, f_col, sdf, s_col,
            st.session_state.s_url_col, st.session_state.s_date_col, st.session_state.s_type_col,
            threshold=threshold,
            extra_name_col=st.session_state.s_extra_name_col,
            progress=lambda p, t: progress.progress(p, text=t),
        )
        st.session_state.matches = matches

        if fetch_pages and matches:
            fp = st.progress(0, text="Fetching FDA change details…")
            pl.enrich_all(matches, progress=lambda p, t: fp.progress(p, text=t), **enrich_kwargs())
            fp.empty()
            n_err = sum(1 for m in matches if m["fetch_error"])
            if n_err:
                st.warning(f"{n_err} match(es) had a fetch issue — see Hazard Review for details.")

        st.success(f"**{len(matches)} matches** found across {len(set(m['formulary_drug'] for m in matches))} formulary drugs.")

    if st.session_state.matches:
        matches = st.session_state.matches
        preview = pd.DataFrame([{
            "Formulary Drug": m["formulary_drug"],
            "SrLC Match": m["slrc_drug"],
            "Score": m["match_score"],
            "Date": m["date"],
            "Change Type": m["change_type"],
            "Risk": m["risk_level"].upper(),
            "SrLC Risk": m.get("risk_by_source", {}).get(pl.SRC_SRLC, m.get("risk_by_source", {}).get(pl.SRC_SRLC_PAGE, "")).upper(),
            "openFDA Risk": m.get("risk_by_source", {}).get(pl.SRC_OPENFDA, "").upper(),
            "Flags (source)": "; ".join(f"{lab} [{', '.join(s.split(' ')[0] for s in srcs)}]" for lab, srcs in m.get("flag_sources", {}).items()),
        } for m in matches])
        st.dataframe(preview, use_container_width=True)
        qa_download_button("qa_dl_page2")

# ═══════════════════════════════════════════════════════════════════════════════
# Page 3 — Hazard Review
# ═══════════════════════════════════════════════════════════════════════════════
elif page == "3. Hazard Review":
    st.markdown("## 🔬 Hazard Review")
    st.caption("Review each matched drug's labeling changes. Record your clinical decision for the export.")

    if not st.session_state.matches:
        st.warning("Run matching in Step 2 first.")
        st.stop()

    matches = st.session_state.matches

    # Filters
    fc1, fc2, fc3 = st.columns(3)
    risk_filter = fc1.multiselect(
        "Show Risk Levels",
        ["high", "medium", "low", "none"],
        default=["high", "medium", "low"],
    )
    show_reviewed = fc2.toggle("Show reviewed items", value=True)
    sort_by = fc3.selectbox("Sort by", ["Risk (High first)", "Match Score (High first)", "Drug Name (A–Z)"])

    sort_fn = {
        "Risk (High first)": lambda m: ({"high": 0, "medium": 1, "low": 2, "none": 3}.get(m["risk_level"], 4), m["formulary_drug"]),
        "Match Score (High first)": lambda m: (-m["match_score"], m["formulary_drug"]),
        "Drug Name (A–Z)": lambda m: m["formulary_drug"].lower(),
    }[sort_by]

    filtered = [m for m in matches if m["risk_level"] in risk_filter]
    if not show_reviewed:
        filtered = [m for m in filtered if m["formulary_drug"] not in st.session_state.ph_decisions]
    filtered.sort(key=sort_fn)

    if not filtered:
        st.info("No items match your current filters.")
        st.stop()

    RISK_ICONS = {"high": "🔴", "medium": "🟡", "low": "🟢", "none": "⚪"}
    DECISION_OPTIONS = [
        "",
        "Action Required — HD handling update needed",
        "Monitor — Reassess at next formulary review",
        "No Action — Not hazardous handling related",
        "Escalate — Send to P&T committee",
    ]

    for mi, match in enumerate(filtered):
        drug = match["formulary_drug"]
        wkey = f"{drug}_{match.get('slrc_row', mi)}"
        risk = match["risk_level"]
        icon = RISK_ICONS.get(risk, "⚪")
        reviewed_mark = "✅ " if drug in st.session_state.ph_decisions else ""
        label = f"{icon} {reviewed_mark}{drug}  ←  {match['slrc_drug']}  [{match['match_score']:.0f}% match]"
        if match.get("date"):
            label += f"  ·  {match['date']}"

        with st.expander(label, expanded=(risk == "high" and drug not in st.session_state.ph_decisions)):
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Risk Level", risk.upper())
            m2.metric("Match Score", f"{match['match_score']:.0f}%")
            m3.metric("Change Date", match.get("date") or "—")
            m4.metric("Change Type", (match.get("change_type") or match.get("appl_type") or "—")[:30])

            # Tags
            tags = ""
            if match.get("boxed_warning"):
                tags += '<span class="tag boxed">⚠️ BOXED WARNING</span>'
            for cat in match.get("niosh", {}):
                tags += f'<span class="tag niosh">NIOSH: {cat}</span>'
            for cat in match.get("usp800", {}):
                tags += f'<span class="tag usp800">USP 800: {cat}</span>'
            if tags:
                st.markdown(tags + "<br>", unsafe_allow_html=True)
            if match.get("flag_sources"):
                st.markdown("**Flag sources:** " + " · ".join(
                    f"{lab} — _{', '.join(srcs)}_" for lab, srcs in match["flag_sources"].items()))
            if match.get("risk_by_source"):
                st.caption("Risk by source: " + " · ".join(f"{s}: {r.upper()}" for s, r in match["risk_by_source"].items()))

            col_label, col_review = st.columns([3, 2])

            with col_label:
                if match.get("url"):
                    st.markdown(f"[🔗 Open FDA Labeling Change Page]({match['url']})")

                if match.get("label_text"):
                    with st.expander(f"📄 {match.get('text_source') or 'SrLC change text'} (keyword-highlighted)", expanded=(risk in ["high", "medium"])):
                        highlighted = highlight_keywords(match["label_text"][:4000])
                        st.markdown(
                            f'<div class="label-block">{highlighted}</div>',
                            unsafe_allow_html=True,
                        )
                if match.get("openfda_text"):
                    o = match.get("openfda", {})
                    with st.expander(f"🏷️ openFDA current label — {o.get('brand_name', '')} ({o.get('effective_time', '')})"):
                        highlighted = highlight_keywords(match["openfda_text"][:4000])
                        st.markdown(f'<div class="label-block">{highlighted}</div>', unsafe_allow_html=True)
                if match.get("fetch_error"):
                    st.warning(f"Fetch issue: {match['fetch_error']}")
                if match.get("drug_name_id") or match.get("url"):
                    btn = "🔄 Retry Fetch" if (match.get("fetch_error") or match.get("label_text")) else "Fetch FDA Data Now"
                    if st.button(btn, key=f"fetch_{wkey}"):
                        with st.spinner("Fetching…"):
                            pl.enrich_match(match, **enrich_kwargs())
                        st.rerun()
                else:
                    st.caption("No URL available — review the FDA SrLC page manually.")

            with col_review:
                st.markdown("**Pharmacist Decision**")
                current_dec = st.session_state.ph_decisions.get(drug, "")
                dec_idx = DECISION_OPTIONS.index(current_dec) if current_dec in DECISION_OPTIONS else 0
                decision = st.selectbox(
                    "Decision",
                    DECISION_OPTIONS,
                    index=dec_idx,
                    key=f"dec_{wkey}",
                    label_visibility="collapsed",
                )
                notes = st.text_area(
                    "Notes",
                    value=st.session_state.ph_notes.get(drug, ""),
                    placeholder="Clinical rationale, follow-up actions, P&T notes…",
                    height=110,
                    key=f"notes_{wkey}",
                )
                if st.button("Save Review", key=f"save_{wkey}", use_container_width=True):
                    if decision:
                        st.session_state.ph_decisions[drug] = decision
                    if notes:
                        st.session_state.ph_notes[drug] = notes
                    st.success("Saved ✓")

# ═══════════════════════════════════════════════════════════════════════════════
# Page 4 — Export
# ═══════════════════════════════════════════════════════════════════════════════
elif page == "4. Export":
    st.markdown("## 📊 Export Report")

    if not st.session_state.matches:
        st.warning("No data yet. Complete Steps 1–3 first.")
        st.stop()

    matches = st.session_state.matches
    total = len(matches)
    high = sum(1 for m in matches if m["risk_level"] == "high")
    medium = sum(1 for m in matches if m["risk_level"] == "medium")
    reviewed = len(st.session_state.ph_decisions)
    action_count = sum(1 for v in st.session_state.ph_decisions.values() if "Action" in v)

    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Total Matches", total)
    c2.metric("🔴 High Risk", high)
    c3.metric("🟡 Medium Risk", medium)
    c4.metric("Reviewed", reviewed)
    c5.metric("Action Required", action_count)

    st.divider()

    output = io.BytesIO()
    wb = xlsxwriter.Workbook(output, {"in_memory": True})

    # Formats
    hdr = wb.add_format({"bold": True, "bg_color": "#21262d", "font_color": "#e6edf3",
                          "border": 1, "text_wrap": True, "valign": "vcenter"})
    high_fmt = wb.add_format({"bg_color": "#3d1f1f", "font_color": "#f85149", "text_wrap": True, "valign": "top"})
    med_fmt = wb.add_format({"bg_color": "#2d2000", "font_color": "#d29922", "text_wrap": True, "valign": "top"})
    low_fmt = wb.add_format({"bg_color": "#1a2f1a", "font_color": "#3fb950", "text_wrap": True, "valign": "top"})
    norm_fmt = wb.add_format({"text_wrap": True, "valign": "top"})
    link_fmt = wb.add_format({"font_color": "#58a6ff", "underline": True, "text_wrap": True, "valign": "top"})

    ROW_FMTS = {"high": high_fmt, "medium": med_fmt, "low": low_fmt, "none": norm_fmt}

    HEADERS = [
        "Formulary Drug", "SrLC Match", "Match Score (%)",
        "Change Date", "Change Type", "Risk Level",
        "Boxed Warning", "NIOSH Signals", "USP 800 Signals",
        "FDA URL", "Pharmacist Decision", "Clinical Notes",
        "Flag Sources",
    ]

    def write_sheet(ws, data, show_url=True):
        ws.set_row(0, 28)
        for c, h in enumerate(HEADERS):
            ws.write(0, c, h, hdr)
        for r, m in enumerate(data, start=1):
            drug = m["formulary_drug"]
            rf = ROW_FMTS.get(m["risk_level"], norm_fmt)
            niosh_str = "; ".join(
                f"{cat}: {', '.join(kws)}"
                for cat, kws in m.get("niosh", {}).items()
            )
            usp_str = "; ".join(
                f"{cat}: {', '.join(kws)}"
                for cat, kws in m.get("usp800", {}).items()
            )
            row_vals = [
                m["formulary_drug"],
                m["slrc_drug"],
                m["match_score"],
                m.get("date") or "",
                m.get("change_type") or "",
                m["risk_level"].upper(),
                "YES" if m.get("boxed_warning") else "",
                niosh_str,
                usp_str,
                m.get("url") or "",
                st.session_state.ph_decisions.get(drug, "Pending Review"),
                st.session_state.ph_notes.get(drug, ""),
                "; ".join(f"{lab} [{', '.join(srcs)}]" for lab, srcs in m.get("flag_sources", {}).items()),
            ]
            for c, val in enumerate(row_vals):
                if c == 9 and val and val.startswith("http"):
                    ws.write_url(r, c, val, link_fmt, val[:80])
                else:
                    ws.write(r, c, val, rf)
            ws.set_row(r, 45)
        ws.set_column(0, 1, 32)
        ws.set_column(2, 5, 14)
        ws.set_column(6, 6, 14)
        ws.set_column(7, 8, 55)
        ws.set_column(9, 9, 45)
        ws.set_column(10, 11, 40)
        ws.set_column(12, 12, 60)
        ws.freeze_panes(1, 0)

    # Sheet 1: All matches
    ws_all = wb.add_worksheet("All Matches")
    write_sheet(ws_all, matches)

    # Sheet 2: High Risk
    ws_high = wb.add_worksheet("High Risk")
    write_sheet(ws_high, [m for m in matches if m["risk_level"] == "high"])

    # Sheet 3: Action Required
    ws_action = wb.add_worksheet("Action Required")
    action_items = [m for m in matches if "Action" in st.session_state.ph_decisions.get(m["formulary_drug"], "")]
    write_sheet(ws_action, action_items)

    # Sheet 4: P&T Summary
    ws_pt = wb.add_worksheet("P&T Summary")
    pt_hdr = wb.add_format({"bold": True, "bg_color": "#161b22", "font_color": "#e6edf3", "border": 1})
    ws_pt.write(0, 0, "RxLabelWatch — Formulary Safety Labeling Review", wb.add_format({"bold": True, "font_size": 14, "font_color": "#e6edf3"}))
    ws_pt.write(1, 0, f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}")
    ws_pt.write(3, 0, "Summary", wb.add_format({"bold": True, "font_size": 12}))
    summary_rows = [
        ("Total formulary drugs with SrLC labeling changes", total),
        ("High risk (NIOSH/USP 800 signals)", high),
        ("Medium risk", medium),
        ("Low risk", sum(1 for m in matches if m["risk_level"] == "low")),
        ("No signals detected", sum(1 for m in matches if m["risk_level"] == "none")),
        ("Reviewed by pharmacist", reviewed),
        ("Action Required", action_count),
    ]
    for i, (label_text, val) in enumerate(summary_rows, start=4):
        ws_pt.write(i, 0, label_text)
        ws_pt.write(i, 1, val)
    ws_pt.set_column(0, 0, 50)
    ws_pt.set_column(1, 1, 12)

    wb.close()
    output.seek(0)

    date_str = datetime.now().strftime("%Y%m%d")
    st.download_button(
        "📥 Download Full Report (.xlsx)",
        data=output,
        file_name=f"rxlabelwatch_{date_str}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
        type="primary",
    )
    qa_download_button("qa_dl_page4")
    st.caption("The QA file holds the raw source data behind every flag (SrLC rows, SrLC change text, openFDA label text) with source URLs and fetch times — check it before sharing results with clients.")

    st.markdown("### Preview — All Matches")
    prev_df = pd.DataFrame([{
        "Formulary Drug": m["formulary_drug"],
        "SrLC Match": m["slrc_drug"],
        "Score": m["match_score"],
        "Date": m["date"],
        "Risk": m["risk_level"].upper(),
        "Boxed Warning": "YES" if m.get("boxed_warning") else "",
        "Decision": st.session_state.ph_decisions.get(m["formulary_drug"], "—"),
    } for m in matches])
    st.dataframe(prev_df, use_container_width=True)
