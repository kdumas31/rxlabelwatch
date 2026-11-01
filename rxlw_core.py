"""Matching and hazard-scan logic for RxLabelWatch.

Moved verbatim from app.py so it can be imported and tested without Streamlit.
The scan adds NegEx-style negation handling (see "Negation" below).
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


# ─── Negation (NegEx-style) ────────────────────────────────────────────────────
# A keyword occurrence is NEGATED when, inside its own clause:
#   * a negation cue comes before it with at most NEG_WINDOW words between
#     ("no evidence of carcinogenicity", "was not carcinogenic", "non-mutagenic",
#     "negative in the Ames bacterial mutagenicity assay"), or
#   * a post-negation cue follows it within NEG_WINDOW words
#     ("Ames test results were negative", "hepatotoxicity was not observed"),
# and the clause has NO affirmative cue ("may cause", "risk of", "associated
# with", "was positive" ...) that is not itself negated. An affirmative cue
# keeps every hit in the clause.
# Clauses end at . ; : ! ? newline and at contrast words (but, however,
# although, though, whereas, except), so "not teratogenic in rats but may
# cause embryo-fetal toxicity" negates only the first part.
# Pseudo-negations ("do not use", "not recommended", "not been established",
# "no adequate studies", "cannot be ruled out" ...) are masked first, so
# instructions and missing-data statements still count as hits.
# Negated hits are never dropped: scan_for_hazards reports them under
# "negated" with the clause text in "negated_evidence".
NEG_WINDOW = 6

_WORD_RE = re.compile(r"[a-z0-9]+(?:['-][a-z0-9]+)*", re.IGNORECASE)
_CLAUSE_BREAK_RE = re.compile(
    r"(?<!\d)\.(?!\d)|[;:!?\n]|\b(?:but|however|although|though|whereas|except)\b",
    re.IGNORECASE,
)
_PSEUDO_NEG_RE = re.compile(
    r"\b(?:do|should|must|need)\s+not\b"
    r"|\bnot\s+(?:be\s+)?recommended\b|\bnot\s+to\b|\bnot\s+be\s+used\b|\bnot\s+use\b"
    r"|\bnot\s+(?:been\s+)?(?:known|established|studied|evaluated|conducted|performed|determined|assessed)\b"
    r"|\bno\s+(?:adequate|data|available|information|well-controlled|controlled|clinical|longer|known\s+antidote|specific\s+antidote|antidote)\b"
    r"|\b(?:cannot|can\s+not|could\s+not)\s+be\s+(?:ruled\s+out|excluded)\b|\bnot\s+(?:been\s+)?ruled\s+out\b"
    r"|\bnot\s+only\b|\bnot\s+necessarily\b"
    r"|\b(?:not|non-?)\s*pregnant\b",   # "females who are not pregnant must ... pregnancy testing"
    re.IGNORECASE,
)
_PRE_NEG_RE = re.compile(
    r"\b(?:no|not|without|absent|neither|nor|lacks?|lacked|absence\s+of|lack\s+of|free\s+of"
    r"|negative\s+(?:in|for|results?)|non-|failed\s+to)(?=\W|$)|\bnon-",
    re.IGNORECASE,
)
_POST_NEG_RE = re.compile(
    r"\b(?:was|were|is|are|been|be|proved|tested|found|remained)\s+negative\b"
    r"|\bnot\s+(?:been\s+)?(?:observed|seen|detected|found|increased|demonstrated|reported|identified|noted)\b"
    r"|\b(?:was|were|been)\s+ruled\s+out\b",
    re.IGNORECASE,
)
_AFFIRM_RE = re.compile(
    r"\b(?:may|can|could|might|will|to)\s+(?:cause|result\s+in|lead\s+to|increase|produce|induce|impair)\b"
    r"|\bcaus(?:es|ed|ing)\b|\b(?:increased|elevated)\s+(?:risk|incidence)\b|\brisk\s+of\b"
    r"|\bassociated\s+with\b|\b(?:has|have)\s+been\s+reported\b|\bpositive\s+(?:in|for)\b"
    r"|\b(?:was|were|is|are)\s+(?:positive|carcinogenic|teratogenic|mutagenic|genotoxic|clastogenic|embryotoxic|fetotoxic)\b",
    re.IGNORECASE,
)


def _mask(text: str, pattern: re.Pattern) -> str:
    return pattern.sub(lambda m: " " * len(m.group()), text)


def _nwords(text: str) -> int:
    return len(_WORD_RE.findall(text))


def clause_bounds(text: str, start: int, end: int) -> tuple[int, int]:
    """Start/end of the clause holding text[start:end]."""
    lo = 0
    for b in _CLAUSE_BREAK_RE.finditer(text, 0, start):
        lo = b.end()
    nb = _CLAUSE_BREAK_RE.search(text, end)
    return lo, (nb.start() if nb else len(text))


def _affirmed_clause(clause: str) -> bool:
    masked = _mask(clause, _PSEUDO_NEG_RE)
    for a in _AFFIRM_RE.finditer(masked):
        before = masked[: a.start()]
        cue = list(_PRE_NEG_RE.finditer(before))
        if not cue or _nwords(before[cue[-1].end():]) > 2:
            return True
    return False


def is_negated(text: str, start: int, end: int) -> bool:
    """NegEx-style check for the keyword occurrence text[start:end]."""
    lo, hi = clause_bounds(text, start, end)
    clause = text[lo:hi]
    s, e = start - lo, end - lo
    masked = _mask(clause, _PSEUDO_NEG_RE)
    neg = False
    pre = list(_PRE_NEG_RE.finditer(masked, 0, s))
    if pre and _nwords(masked[pre[-1].end():s]) <= NEG_WINDOW:
        neg = True
    if not neg:
        post = _POST_NEG_RE.search(masked, e)
        if post and _nwords(masked[e:post.start()]) <= NEG_WINDOW:
            neg = True
    return neg and not _affirmed_clause(clause)


def keyword_occurrences(kw: str, text: str) -> list[tuple[int, int, bool]]:
    """(start, end, negated) for every occurrence of kw in text."""
    return [(m.start(), m.end(), is_negated(text, m.start(), m.end()))
            for m in keyword_pattern(kw).finditer(text)]


# ─── Hazard Analysis ───────────────────────────────────────────────────────────
def _risk(boxed, has, cat_hits) -> str:
    if boxed or any(has(s) for s in HIGH_SIGNALS):
        return "high"
    if any(has(s) for s in MEDIUM_SIGNALS):
        return "medium"
    return "low" if cat_hits else "none"


def scan_for_hazards(text: str) -> dict:
    """Keyword scan. A keyword counts if at least one occurrence is not negated.
    Keywords whose every occurrence is negated go to "negated" (same
    category layout) and do not raise the risk level; "risk_level_ignoring_negation"
    is the level the old scan would have given."""
    empty = {"niosh": {}, "usp800": {}, "boxed_warning": False, "risk_level": "none",
             "negated": {"niosh": {}, "usp800": {}}, "negated_evidence": [],
             "risk_level_ignoring_negation": "none"}
    if not text:
        return empty
    t = text.lower()
    occ_cache: dict[str, list] = {}

    def occ(kw):
        if kw not in occ_cache:
            occ_cache[kw] = keyword_occurrences(kw, t)
        return occ_cache[kw]

    def affirmed(kw):
        return any(not n for _, _, n in occ(kw))

    def present(kw):
        return bool(occ(kw))

    out = {"niosh": {}, "usp800": {}}
    neg = {"niosh": {}, "usp800": {}}
    evidence = []
    for key, criteria in (("niosh", NIOSH_CRITERIA), ("usp800", USP800_CRITERIA)):
        for category, keywords in criteria.items():
            hits = [kw for kw in keywords if affirmed(kw)]
            nhits = [kw for kw in keywords if present(kw) and not affirmed(kw)]
            if hits:
                out[key][category] = hits
            if nhits:
                neg[key][category] = nhits
                for kw in nhits:
                    s0, e0, _ = occ(kw)[0]
                    lo, hi = clause_bounds(text, s0, e0)
                    evidence.append({"category": category, "keyword": kw,
                                     "clause": " ".join(text[lo:hi].split())[:300]})
    boxed = any(keyword_in(p, t) for p in BOXED_WARNING_PHRASES)
    any_old = {c: 1 for k in ("niosh", "usp800") for c in list(out[k]) + list(neg[k])}
    return {
        "niosh": out["niosh"], "usp800": out["usp800"], "boxed_warning": boxed,
        "risk_level": _risk(boxed, affirmed, out["niosh"] or out["usp800"]),
        "negated": neg, "negated_evidence": evidence,
        "risk_level_ignoring_negation": _risk(boxed, present, any_old),
    }


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
    """Same matching and negation rules as scan_for_hazards. Negated
    occurrences get class "hl-negated" (plus their category class) and a
    tooltip. Matching runs on the raw text and each piece is escaped
    separately, so a keyword can never match inside an inserted tag."""
    out, pos = [], 0
    for m in _HL_RE.finditer(raw_text):
        out.append(html.escape(raw_text[pos:m.start()]))
        css = _HL_CSS[m.lastgroup]
        word = html.escape(m.group())
        if is_negated(raw_text, m.start(), m.end()):
            out.append(f'<span class="{css} hl-negated" title="negated: not counted in risk">{word}</span>')
        else:
            out.append(f'<span class="{css}">{word}</span>')
        pos = m.end()
    out.append(html.escape(raw_text[pos:]))
    return "".join(out)

# ─── Column Auto-Detect ────────────────────────────────────────────────────────
def autodetect_col(df: pd.DataFrame, keywords: list[str]) -> str | None:
    for kw in keywords:
        for col in df.columns:
            if kw.lower() in col.lower():
                return col
    return None
