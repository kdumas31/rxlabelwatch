import streamlit as st
import pandas as pd
import re
import io
import time
import html
import requests
from bs4 import BeautifulSoup
from rapidfuzz import fuzz
from datetime import datetime
import xlsxwriter

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

# ─── NIOSH + USP 800 Keyword Definitions ───────────────────────────────────────
NIOSH_CRITERIA = {
    "Carcinogenicity": [
        "carcinogen", "carcinogenic", "oncogenic", "tumor", "malignant",
        "cancer", "neoplasm", "carcinoma", "sarcoma", "leukemia", "lymphoma",
    ],
    "Genotoxicity": [
        "genotoxic", "mutagenic", "mutagenicity", "clastogenic",
        "chromosomal aberration", "dna damage", "ames test", "micronuclei",
    ],
    "Reproductive Toxicity": [
        "teratogenic", "teratogenicity", "embryotoxic", "fetotoxic",
        "fetal toxicity", "reproductive toxicity", "developmental toxicity",
        "birth defect", "spermatogenesis", "anovulation", "infertility",
        "embryo-fetal", "embryofetal",
    ],
    "Organ Toxicity": [
        "organ toxicity", "hepatotoxic", "hepatotoxicity", "nephrotoxic",
        "nephrotoxicity", "neurotoxic", "neurotoxicity", "cardiotoxic",
        "cardiotoxicity", "myelosuppression", "myelotoxic", "pulmonary toxicity",
    ],
    "Antineoplastic / Cytotoxic": [
        "antineoplastic", "cytotoxic", "alkylating agent", "antimetabolite",
        "topoisomerase inhibitor", "mitotic inhibitor", "kinase inhibitor",
        "immunosuppressant", "monoclonal antibody",
    ],
    "Pregnancy / Lactation": [
        "pregnancy", "lactation", "nursing", "breastfeeding",
        "fetus", "neonatal", "placental", "embryo", "maternal",
        "pregnancy category", "avoid pregnancy",
    ],
}

USP800_CRITERIA = {
    "HD Handling Language": [
        "hazardous drug", "hazardous medication", "hazardous agent",
        "handling precaution", "safe handling", "special handling",
        "precautions for handling",
    ],
    "PPE Requirements": [
        "personal protective equipment", "double glov", "chemotherapy gown",
        "nitrile glove", "protective gown", "respiratory protection",
        "n95", "respirator", "face shield", "ppe",
    ],
    "Engineering Controls": [
        "closed system", "cstd", "closed-system transfer device",
        "negative pressure", "ventilated cabinet", "biological safety cabinet",
        "bsc", "engineering control", "containment", "isolator",
    ],
    "Disposal / Spill": [
        "spill kit", "spill clean", "disposal precaution", "waste disposal",
        "chemotherapy waste", "hazardous waste", "incinerat",
    ],
    "Regulatory References": [
        "usp 800", "usp800", "niosh", "osha hazard communication",
        "safety data sheet", "sds", "hazard communication standard",
    ],
}

BOXED_WARNING_PHRASES = [
    "boxed warning", "black box warning", "WARNING\n", "WARNINGS\n",
    "box warning",
]

SALT_FORMS = [
    r"\bhydrochloride\b", r"\bhcl\b", r"\bhydrobromide\b", r"\bhbr\b",
    r"\bsulfate\b", r"\bsulphate\b", r"\bsodium\b", r"\bpotassium\b",
    r"\bcalcium\b", r"\bmagnesium\b", r"\bacetate\b", r"\bcitrate\b",
    r"\bphosphate\b", r"\bnitrate\b", r"\bmaleate\b", r"\btartrate\b",
    r"\bfumarate\b", r"\bsuccinate\b", r"\bmesylate\b", r"\btosylate\b",
    r"\bbesylate\b", r"\bmalate\b", r"\blactate\b", r"\bgluconate\b",
    r"\bchloride\b", r"\bbromide\b", r"\biodide\b",
    r"\bmonohydrate\b", r"\bdihydrate\b", r"\btrihydrate\b",
]

DOSAGE_FORM_WORDS = [
    "injection", "solution", "tablet", "capsule", "cream", "ointment",
    "patch", "gel", "suspension", "syrup", "elixir", "suppository",
    "powder", "lyophilized", "concentrate", "infusion", "emulsion",
    "oral", "intravenous", "iv", "subcutaneous", "intramuscular",
    "topical", "ophthalmic", "otic", "nasal", "inhaler", "inhalation",
    "transdermal", "rectal", "sublingual", "buccal",
    "extended release", "er", "xr", "cr", "sr", "la", "dr", "ir",
    "delayed release", "modified release", "immediate release",
    "ivpb", "piggyback",
]

DOSAGE_NUM_PATTERN = re.compile(
    r"\b\d+(\.\d+)?\s*(mg|mcg|g|gram|ml|unit|units|iu|meq|mmol|%)\b",
    re.IGNORECASE,
)
IN_DILUENT_PATTERN = re.compile(
    r"\bin\s+(dextrose|saline|water|sodium chloride|d5w|ns|lr|lactated|normal saline).*",
    re.IGNORECASE,
)

# ─── Drug Name Normalization ────────────────────────────────────────────────────
def normalize_drug_name(name: str) -> str:
    if not isinstance(name, str):
        return ""
    n = name.lower().strip()
    n = DOSAGE_NUM_PATTERN.sub(" ", n)
    n = IN_DILUENT_PATTERN.sub(" ", n)
    n = re.sub(r"\(.*?\)", " ", n)
    for pattern in SALT_FORMS:
        n = re.sub(pattern, " ", n, flags=re.IGNORECASE)
    for word in DOSAGE_FORM_WORDS:
        n = re.sub(r"\b" + re.escape(word) + r"\b", " ", n)
    n = re.sub(r"[^a-z\s]", " ", n)
    n = re.sub(r"\s+", " ", n).strip()
    return n

# ─── Fuzzy Matching ────────────────────────────────────────────────────────────
def find_best_matches(formulary_norm: str, slrc_norms: list, threshold: int = 78):
    results = []
    for i, norm in enumerate(slrc_norms):
        if not norm:
            continue
        score = fuzz.token_sort_ratio(formulary_norm, norm)
        if score >= threshold:
            results.append((i, score))
    results.sort(key=lambda x: -x[1])
    return results[:5]

# ─── Hazard Analysis ───────────────────────────────────────────────────────────
def scan_for_hazards(text: str) -> dict:
    if not text:
        return {"niosh": {}, "usp800": {}, "boxed_warning": False, "risk_level": "none"}
    t = text.lower()
    niosh_hits = {}
    for category, keywords in NIOSH_CRITERIA.items():
        hits = [kw for kw in keywords if kw.lower() in t]
        if hits:
            niosh_hits[category] = hits
    usp800_hits = {}
    for category, keywords in USP800_CRITERIA.items():
        hits = [kw for kw in keywords if kw.lower() in t]
        if hits:
            usp800_hits[category] = hits
    boxed = any(p.lower() in t for p in BOXED_WARNING_PHRASES)
    high_signals = [
        "carcinogen", "genotoxic", "teratogenic", "antineoplastic", "cytotoxic",
        "hazardous drug", "niosh", "usp 800", "closed system", "cstd",
        "embryo-fetal", "myelosuppression",
    ]
    medium_signals = [
        "reproductive toxicity", "organ toxicity", "hepatotoxic", "nephrotoxic",
        "ppe", "personal protective equipment", "disposal precaution",
        "handling precaution", "pregnancy", "lactation", "mutagenic",
    ]
    if boxed or any(s in t for s in high_signals):
        risk = "high"
    elif any(s in t for s in medium_signals):
        risk = "medium"
    elif niosh_hits or usp800_hits:
        risk = "low"
    else:
        risk = "none"
    return {"niosh": niosh_hits, "usp800": usp800_hits, "boxed_warning": boxed, "risk_level": risk}

def highlight_keywords(raw_text: str) -> str:
    escaped = html.escape(raw_text)
    all_kw = []
    for kws in NIOSH_CRITERIA.values():
        for kw in kws:
            all_kw.append((kw, "hl-niosh"))
    for kws in USP800_CRITERIA.values():
        for kw in kws:
            all_kw.append((kw, "hl-usp800"))
    all_kw.sort(key=lambda x: -len(x[0]))
    seen = set()
    for kw, css in all_kw:
        if kw.lower() in seen:
            continue
        pattern = re.compile(re.escape(kw), re.IGNORECASE)
        if pattern.search(escaped):
            escaped = pattern.sub(
                lambda m: f'<span class="{css}">{m.group()}</span>',
                escaped,
            )
            seen.add(kw.lower())
    return escaped

# ─── FDA Page Fetcher ──────────────────────────────────────────────────────────
HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0 Safari/537.36"
    ),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "en-US,en;q=0.5",
}

PRIORITY_SECTION_KEYWORDS = [
    "warning", "precaution", "hazard", "handling", "disposal",
    "pregnancy", "lactation", "reproductive", "carcinogen",
    "mutageni", "genotoxic", "teratogen", "special handling",
    "storage and handling", "how supplied",
]

def fetch_label_page(url: str) -> tuple[str, str]:
    """Returns (extracted_text, error_message)."""
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

# ─── Column Auto-Detect ────────────────────────────────────────────────────────
def autodetect_col(df: pd.DataFrame, keywords: list[str]) -> str | None:
    for kw in keywords:
        for col in df.columns:
            if kw.lower() in col.lower():
                return col
    return None

# ─── Session State Init ────────────────────────────────────────────────────────
def init_state():
    defaults = {
        "formulary_df": None,
        "slrc_df": None,
        "f_name_col": None,
        "s_name_col": None,
        "s_url_col": None,
        "s_date_col": None,
        "s_type_col": None,
        "matches": None,
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
        ["1. Upload Files", "2. Run Matching", "3. Hazard Review", "4. Export"],
        label_visibility="collapsed",
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
    st.caption("Matches your Epic formulary to FDA Safety Labeling Changes (SrLC) and scans for NIOSH + USP 800 hazardous drug signals.")

# ═══════════════════════════════════════════════════════════════════════════════
# Page 1 — Upload Files
# ═══════════════════════════════════════════════════════════════════════════════
if page == "1. Upload Files":
    st.markdown("## 📁 Upload Files")
    st.caption("Upload your Epic formulary export and the FDA Safety Labeling Changes (SrLC) Excel/CSV file.")

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
                st.success(f"Loaded **{len(df):,} rows**, {len(df.columns)} columns")
                st.dataframe(df.head(8), use_container_width=True)
            except Exception as e:
                st.error(f"Error loading file: {e}")

    with col2:
        st.markdown("### 📋 SrLC Report (from FDA)")
        st.caption("Download from: [FDA Safety Labeling Changes](https://www.accessdata.fda.gov/scripts/cder/safetylabelingchanges/index.cfm)")
        slrc_file = st.file_uploader(
            "Drop your FDA SrLC export here",
            type=["xlsx", "xls", "csv"],
            key="slrc_upload",
        )
        if slrc_file:
            try:
                df = pd.read_csv(slrc_file) if slrc_file.name.endswith(".csv") else pd.read_excel(slrc_file)
                st.session_state.slrc_df = df
                st.success(f"Loaded **{len(df):,} rows**, {len(df.columns)} columns")
                st.dataframe(df.head(8), use_container_width=True)
            except Exception as e:
                st.error(f"Error loading file: {e}")

    if st.session_state.formulary_df is not None and st.session_state.slrc_df is not None:
        st.divider()
        st.markdown("### ⚙️ Column Mapping")
        st.caption("We auto-detected the most likely columns. Adjust if needed.")

        fdf = st.session_state.formulary_df
        sdf = st.session_state.slrc_df
        fcols = list(fdf.columns)
        scols = list(sdf.columns)

        # Auto-detect defaults
        f_name_default = autodetect_col(fdf, ["drug", "name", "generic", "description", "medication", "item", "title"]) or fcols[0]
        s_name_default = autodetect_col(sdf, ["drug", "name", "proprietary", "nonproprietary", "generic", "product"]) or scols[0]
        s_url_default = autodetect_col(sdf, ["url", "link", "href", "web", "labeling"])
        s_date_default = autodetect_col(sdf, ["date", "submission", "approval", "effective"])
        s_type_default = autodetect_col(sdf, ["type", "supplement", "change", "category", "description"])

        c1, c2 = st.columns(2)
        with c1:
            st.markdown("**Formulary**")
            st.session_state.f_name_col = st.selectbox(
                "Drug name column", fcols,
                index=fcols.index(f_name_default),
                help="Column with generic drug names from Epic",
            )

        with c2:
            st.markdown("**SrLC File**")
            ca, cb = st.columns(2)
            st.session_state.s_name_col = ca.selectbox(
                "Drug name column", scols,
                index=scols.index(s_name_default),
            )
            url_opts = ["(none)"] + scols
            s_url_idx = url_opts.index(s_url_default) if s_url_default in url_opts else 0
            s_url_sel = cb.selectbox("URL / link column", url_opts, index=s_url_idx)
            st.session_state.s_url_col = None if s_url_sel == "(none)" else s_url_sel

            date_opts = ["(none)"] + scols
            s_date_idx = date_opts.index(s_date_default) if s_date_default in date_opts else 0
            s_date_sel = ca.selectbox("Date column", date_opts, index=s_date_idx)
            st.session_state.s_date_col = None if s_date_sel == "(none)" else s_date_sel

            type_opts = ["(none)"] + scols
            s_type_idx = type_opts.index(s_type_default) if s_type_default in type_opts else 0
            s_type_sel = cb.selectbox("Change type column", type_opts, index=s_type_idx)
            st.session_state.s_type_col = None if s_type_sel == "(none)" else s_type_sel

        st.info("✅ Column mapping set. Go to **2. Run Matching** to find matches.")

# ═══════════════════════════════════════════════════════════════════════════════
# Page 2 — Run Matching
# ═══════════════════════════════════════════════════════════════════════════════
elif page == "2. Run Matching":
    st.markdown("## 🔍 Run Matching")

    if st.session_state.formulary_df is None or st.session_state.slrc_df is None:
        st.warning("Please upload both files in Step 1 first.")
        st.stop()

    c1, c2 = st.columns([2, 1])
    threshold = c1.slider(
        "Fuzzy match sensitivity",
        min_value=60, max_value=98, value=78, step=2,
        help="78 is a good default. Lower to catch more matches (more false positives); raise for stricter matching.",
    )
    fetch_pages = c2.toggle(
        "Auto-fetch FDA label pages",
        value=True,
        help="Fetches and analyzes the FDA labeling change page for each match. Adds ~1–2 sec per match.",
    )

    with st.expander("ℹ️ How matching works"):
        st.markdown("""
The app normalizes drug names on both sides before matching:
- Removes salt forms (HCl, sodium, sulfate, etc.)
- Removes dosage forms (tablet, injection, cream, etc.)
- Removes strengths (e.g., "2.5 mg", "1 gram/200 mL")
- Removes EPIC IV bag diluent phrases ("in dextrose 5% in water")

Then uses **token sort ratio** fuzzy matching — so "VANCOMYCIN HYDROCHLORIDE" and "vancomycin HCl 1g IVPB" will match correctly.
""")

    if st.button("▶ Run Matching", use_container_width=True, type="primary"):
        fdf = st.session_state.formulary_df
        sdf = st.session_state.slrc_df
        f_col = st.session_state.f_name_col or fdf.columns[0]
        s_col = st.session_state.s_name_col or sdf.columns[0]

        formulary_drugs = fdf[f_col].dropna().astype(str).unique().tolist()
        slrc_names = sdf[s_col].fillna("").astype(str).tolist()
        slrc_norms = [normalize_drug_name(d) for d in slrc_names]

        progress = st.progress(0, text="Normalizing and matching…")
        matches = []

        for i, drug in enumerate(formulary_drugs):
            progress.progress((i + 1) / len(formulary_drugs), text=f"Matching: {drug[:60]}")
            drug_norm = normalize_drug_name(drug)
            if not drug_norm:
                continue
            best = find_best_matches(drug_norm, slrc_norms, threshold=threshold)
            for idx, score in best:
                row = sdf.iloc[idx]
                url = str(row[st.session_state.s_url_col]) if st.session_state.s_url_col else ""
                date_val = str(row[st.session_state.s_date_col]) if st.session_state.s_date_col else ""
                change_type = str(row[st.session_state.s_type_col]) if st.session_state.s_type_col else ""
                matches.append({
                    "formulary_drug": drug,
                    "slrc_drug": slrc_names[idx],
                    "match_score": score,
                    "date": date_val,
                    "change_type": change_type,
                    "url": url if url not in ("", "None", "nan") else "",
                    "label_text": "",
                    "fetch_error": "",
                    "risk_level": "none",
                    "niosh": {},
                    "usp800": {},
                    "boxed_warning": False,
                })

        progress.progress(1.0, text=f"Matched {len(matches)} drug–label pairs.")
        st.session_state.matches = matches

        # Auto-fetch FDA pages
        if fetch_pages and matches:
            urls = [(i, m) for i, m in enumerate(matches) if m["url"]]
            if urls:
                fp = st.progress(0, text="Fetching FDA label pages…")
                for j, (i, m) in enumerate(urls):
                    fp.progress((j + 1) / len(urls), text=f"Fetching: {m['slrc_drug'][:60]}")
                    text, err = fetch_label_page(m["url"])
                    matches[i]["label_text"] = text
                    matches[i]["fetch_error"] = err
                    if text:
                        findings = scan_for_hazards(text)
                        matches[i]["risk_level"] = findings["risk_level"]
                        matches[i]["niosh"] = findings["niosh"]
                        matches[i]["usp800"] = findings["usp800"]
                        matches[i]["boxed_warning"] = findings["boxed_warning"]
                    time.sleep(0.4)
                fp.empty()

        st.success(f"**{len(matches)} matches** found across {len(set(m['formulary_drug'] for m in matches))} formulary drugs.")

        # Preview table
        preview = pd.DataFrame([{
            "Formulary Drug": m["formulary_drug"],
            "SrLC Match": m["slrc_drug"],
            "Score": m["match_score"],
            "Date": m["date"],
            "Change Type": m["change_type"],
            "Risk": m["risk_level"].upper(),
        } for m in matches])
        st.dataframe(preview, use_container_width=True)

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

    for match in filtered:
        drug = match["formulary_drug"]
        risk = match["risk_level"]
        icon = RISK_ICONS.get(risk, "⚪")
        reviewed_mark = "✅ " if drug in st.session_state.ph_decisions else ""
        label = f"{icon} {reviewed_mark}{drug}  ←  {match['slrc_drug']}  [{match['match_score']}% match]"
        if match.get("date"):
            label += f"  ·  {match['date']}"

        with st.expander(label, expanded=(risk == "high" and drug not in st.session_state.ph_decisions)):
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Risk Level", risk.upper())
            m2.metric("Match Score", f"{match['match_score']}%")
            m3.metric("Change Date", match.get("date") or "—")
            m4.metric("Change Type", (match.get("change_type") or "—")[:30])

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

            col_label, col_review = st.columns([3, 2])

            with col_label:
                if match.get("url"):
                    st.markdown(f"[🔗 Open FDA Labeling Change Page]({match['url']})")

                if match.get("label_text"):
                    with st.expander("📄 Extracted Label Content (keyword-highlighted)", expanded=(risk in ["high", "medium"])):
                        highlighted = highlight_keywords(match["label_text"][:4000])
                        st.markdown(
                            f'<div class="label-block">{highlighted}</div>',
                            unsafe_allow_html=True,
                        )
                elif match.get("fetch_error"):
                    st.warning(f"Auto-fetch failed: {match['fetch_error']}")
                    if match.get("url") and st.button("🔄 Retry Fetch", key=f"retry_{drug}"):
                        with st.spinner("Fetching…"):
                            text, err = fetch_label_page(match["url"])
                        match["label_text"] = text
                        match["fetch_error"] = err
                        if text:
                            findings = scan_for_hazards(text)
                            match.update({
                                "risk_level": findings["risk_level"],
                                "niosh": findings["niosh"],
                                "usp800": findings["usp800"],
                                "boxed_warning": findings["boxed_warning"],
                            })
                        st.rerun()
                elif match.get("url"):
                    if st.button("Fetch Label Page Now", key=f"fetch_{drug}"):
                        with st.spinner("Fetching…"):
                            text, err = fetch_label_page(match["url"])
                        match["label_text"] = text
                        match["fetch_error"] = err
                        if text:
                            findings = scan_for_hazards(text)
                            match.update({
                                "risk_level": findings["risk_level"],
                                "niosh": findings["niosh"],
                                "usp800": findings["usp800"],
                                "boxed_warning": findings["boxed_warning"],
                            })
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
                    key=f"dec_{drug}",
                    label_visibility="collapsed",
                )
                notes = st.text_area(
                    "Notes",
                    value=st.session_state.ph_notes.get(drug, ""),
                    placeholder="Clinical rationale, follow-up actions, P&T notes…",
                    height=110,
                    key=f"notes_{drug}",
                )
                if st.button("Save Review", key=f"save_{drug}", use_container_width=True):
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
