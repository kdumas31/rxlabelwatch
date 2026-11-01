"""Matching and hazard-scan logic for RxLabelWatch.

Moved verbatim from app.py so it can be imported and tested without Streamlit.
Behavior is unchanged.
"""
import html
import re

import pandas as pd
from rapidfuzz import fuzz

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

# ─── Column Auto-Detect ────────────────────────────────────────────────────────
def autodetect_col(df: pd.DataFrame, keywords: list[str]) -> str | None:
    for kw in keywords:
        for col in df.columns:
            if kw.lower() in col.lower():
                return col
    return None
