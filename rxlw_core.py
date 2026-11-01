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
def find_best_matches(formulary_norm: str, slrc_norms: list, threshold: int = 78, limit: int | None = 5):
    results = []
    for i, norm in enumerate(slrc_norms):
        if not norm:
            continue
        score = fuzz.token_sort_ratio(formulary_norm, norm)
        if score >= threshold:
            results.append((i, score))
    results.sort(key=lambda x: -x[1])
    return results if limit is None else results[:limit]


def ingredient_agreement(a_norm: str, b_norm: str, token_threshold: int = 95) -> bool:
    """Ingredient-level check on top of the fuzzy score.

    Every word of the shorter normalized name must appear in the other name,
    exactly or as a near-identical spelling (fuzz.ratio >= token_threshold,
    e.g. cyclosporin/cyclosporine = 96). This rejects look-alike generics that
    share a USAN stem, which token_sort_ratio scores highly:
    amoxicillin/oxacillin = 80, prednisone/prednisolone = 91.
    """
    a, b = a_norm.split(), b_norm.split()
    if not a or not b:
        return False
    short, long_ = (a, b) if len(a) <= len(b) else (b, a)
    return all(
        any(tok == other or fuzz.ratio(tok, other) >= token_threshold for other in long_)
        for tok in short
    )

# ─── Keyword matching rule ─────────────────────────────────────────────────────
# Every keyword must start at a word boundary, so "tumor" no longer hits
# "pseudotumor" and "ames test" no longer hits "names test".
# Most keywords are stems and may continue ("carcinogen" -> carcinogenic,
# carcinogenicity; "double glov" -> gloves/gloving; "incinerat" -> incineration).
# Short acronyms and words whose stem would hit an unrelated word are WHOLE
# words (optional plural s), so "ppe" no longer hits "happen"/"appear",
# "bsc"/"sds" can't fire inside other words, and "respirator" no longer hits
# "respiratory".
WHOLE_WORD_KEYWORDS = {
    "ppe", "bsc", "sds", "cstd", "n95", "niosh", "usp 800", "usp800",
    "respirator",
}

HIGH_SIGNALS = [
    "carcinogen", "genotoxic", "teratogenic", "antineoplastic", "cytotoxic",
    "hazardous drug", "niosh", "usp 800", "closed system", "cstd",
    "embryo-fetal", "myelosuppression",
]
MEDIUM_SIGNALS = [
    "reproductive toxicity", "organ toxicity", "hepatotoxic", "nephrotoxic",
    "ppe", "personal protective equipment", "disposal precaution",
    "handling precaution", "pregnancy", "lactation", "mutagenic",
]

_PATTERN_CACHE: dict[str, re.Pattern] = {}


def keyword_regex(kw: str) -> str:
    """Regex source for one keyword under the start-of-word / whole-word rule."""
    k = kw.lower()
    body = re.escape(k.rstrip("\n").strip()) if k.endswith("\n") else re.escape(k)
    if k.endswith("\n"):           # boxed-warning heading phrases like "WARNING\n"
        return r"\b" + body + r"\n"
    if k in WHOLE_WORD_KEYWORDS:
        return r"\b" + body + r"(?:s|es)?\b"
    return r"\b" + body


def keyword_pattern(kw: str) -> re.Pattern:
    p = _PATTERN_CACHE.get(kw)
    if p is None:
        p = _PATTERN_CACHE[kw] = re.compile(keyword_regex(kw), re.IGNORECASE)
    return p


def keyword_in(kw: str, text: str) -> bool:
    return keyword_pattern(kw).search(text) is not None


# ─── Hazard Analysis ───────────────────────────────────────────────────────────
def scan_for_hazards(text: str) -> dict:
    if not text:
        return {"niosh": {}, "usp800": {}, "boxed_warning": False, "risk_level": "none"}
    t = text.lower()
    niosh_hits = {}
    for category, keywords in NIOSH_CRITERIA.items():
        hits = [kw for kw in keywords if keyword_in(kw, t)]
        if hits:
            niosh_hits[category] = hits
    usp800_hits = {}
    for category, keywords in USP800_CRITERIA.items():
        hits = [kw for kw in keywords if keyword_in(kw, t)]
        if hits:
            usp800_hits[category] = hits
    boxed = any(keyword_in(p, t) for p in BOXED_WARNING_PHRASES)
    if boxed or any(keyword_in(s, t) for s in HIGH_SIGNALS):
        risk = "high"
    elif any(keyword_in(s, t) for s in MEDIUM_SIGNALS):
        risk = "medium"
    elif niosh_hits or usp800_hits:
        risk = "low"
    else:
        risk = "none"
    return {"niosh": niosh_hits, "usp800": usp800_hits, "boxed_warning": boxed, "risk_level": risk}


def _highlight_regex():
    css_for = {}
    for kws in NIOSH_CRITERIA.values():
        for kw in kws:
            css_for.setdefault(kw.lower(), "hl-niosh")
    for kws in USP800_CRITERIA.values():
        for kw in kws:
            css_for.setdefault(kw.lower(), "hl-usp800")
    ordered = sorted(css_for, key=len, reverse=True)  # longest first
    groups = "|".join(f"(?P<k{i}>{keyword_regex(kw)})" for i, kw in enumerate(ordered))
    return re.compile(groups, re.IGNORECASE), {f"k{i}": css_for[kw] for i, kw in enumerate(ordered)}


_HL_RE, _HL_CSS = _highlight_regex()


def highlight_keywords(raw_text: str) -> str:
    """Same matching rule as scan_for_hazards; one pass over the escaped text so
    a keyword can never match inside an inserted <span> tag."""
    escaped = html.escape(raw_text)
    return _HL_RE.sub(
        lambda m: f'<span class="{_HL_CSS[m.lastgroup]}">{m.group()}</span>',
        escaped,
    )

# ─── Column Auto-Detect ────────────────────────────────────────────────────────
def autodetect_col(df: pd.DataFrame, keywords: list[str]) -> str | None:
    for kw in keywords:
        for col in df.columns:
            if kw.lower() in col.lower():
                return col
    return None
