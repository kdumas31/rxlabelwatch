"""Quick offline checks for keyword matching and the crosswalk.

Run: python tests/test_core.py   (or: pytest tests/)
"""
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from rxlw_core import (  # noqa: E402
    BOXED_WARNING_PHRASES, NIOSH_CRITERIA, USP800_CRITERIA, find_best_matches,
    highlight_keywords, ingredient_agreement, is_negated, keyword_in, normalize_drug_name,
    scan_for_hazards,
)


def _matches(formulary, srlc, threshold=78):
    a, b = normalize_drug_name(formulary), normalize_drug_name(srlc)
    return bool(find_best_matches(a, [b], threshold=threshold)) and ingredient_agreement(a, b)


def test_every_keyword_matches_itself():
    for group in (NIOSH_CRITERIA, USP800_CRITERIA):
        for kws in group.values():
            for kw in kws:
                assert keyword_in(kw, f"x {kw} x"), kw
    for p in BOXED_WARNING_PHRASES:
        assert keyword_in(p, f"x {p.lower()}"), p


def test_stems_still_match_longer_words():
    for kw, word in [
        ("carcinogen", "carcinogenicity"), ("carcinogenic", "carcinogenicity"),
        ("genotoxic", "genotoxicity"), ("mutagenic", "mutagenicity"),
        ("teratogenic", "teratogenicity"), ("embryotoxic", "embryotoxicity"),
        ("hepatotoxic", "hepatotoxicity"), ("cytotoxic", "cytotoxicity"),
        ("double glov", "double gloving"), ("nitrile glove", "nitrile gloves"),
        ("incinerat", "incineration"), ("spill clean", "spill cleanup"),
        ("hazardous drug", "hazardous drugs"), ("tumor", "tumors"),
        ("embryo", "embryonic"), ("respirator", "respirators"), ("ppe", "PPEs"),
    ]:
        assert keyword_in(kw, word.lower()), (kw, word)


def test_short_terms_dont_fire_inside_words():
    for kw, word in [
        ("ppe", "happen"), ("ppe", "appear"), ("ppe", "stopped"), ("ppe", "upper"),
        ("bsc", "absciss"), ("bsc", "obscure"), ("sds", "bsds"), ("n95", "an950"),
        ("cstd", "xcstdx"), ("niosh", "fniosh"), ("respirator", "respiratory"),
        ("tumor", "pseudotumor"), ("ames test", "names test"), ("mutagenic", "nonmutagenic"),
    ]:
        assert not keyword_in(kw, word.lower()), (kw, word)
    r = scan_for_hazards("Such reactions may happen; rash may appear in upper limbs.")
    assert "PPE Requirements" not in r["usp800"], r


def test_highlight_uses_same_rule():
    out = highlight_keywords("It may happen. Wear PPE. Carcinogenicity studies. NIOSH list.")
    assert 'happen' in out and '<span class="hl-usp800">PPE</span>' in out
    assert '<span class="hl-niosh">Carcinogenic</span>ity' in out
    assert out.count("<span") == 3  # PPE, Carcinogenic, NIOSH; nothing inside a tag
    assert '<span class="hl-usp800">NIOSH</span>' in out


def test_near_miss_generics_rejected():
    for f, s in [
        ("Amoxicillin 500 mg capsule", "OXACILLIN SODIUM"),
        ("Prednisone 10 mg tablet", "PREDNISOLONE"),
        ("Hydroxyzine 25 mg tablet", "HYDRALAZINE HYDROCHLORIDE"),
        ("Cefuroxime 500 mg tablet", "CEFOTAXIME SODIUM"),
        ("Clonidine 0.1 mg tablet", "CLOZAPINE"),
    ]:
        assert not _matches(f, s), (f, s)


def test_true_matches_kept():
    for f, s in [
        ("Methotrexate 2.5 mg tablet", "METHOTREXATE SODIUM"),
        ("Methotrexate sodium 25 mg/mL injection", "METHOTREXATE"),
        ("Carbamazepine 200 mg tablet", "CARBAMAZEPINE"),
        ("Leuprolide acetate 22.5 mg injection", "LEUPROLIDE MESYLATE"),
        ("Metformin HCl 500 mg tablet", "METFORMIN HYDROCHLORIDE"),
        ("Cyclosporin 100 mg capsule", "CYCLOSPORINE"),
    ]:
        assert _matches(f, s), (f, s)


# ─── Negation (real-style label sentences) ─────────────────────────────────────
def _neg(r):
    return {c for d in r["negated"].values() for c in d}


def _aff(r):
    return set(r["niosh"]) | set(r["usp800"])


NEGATED_SENTENCES = [
    ("In 2-year studies in rats and mice, there was no evidence of carcinogenicity at doses up to 30 mg/kg/day.",
     "Carcinogenicity"),
    ("Drug X was not carcinogenic in a 104-week oral study in rats.", "Carcinogenicity"),
    ("Drug X was not teratogenic in rats or rabbits at doses up to 8 times the MRHD.", "Reproductive Toxicity"),
    ("Drug X did not show mutagenic or clastogenic potential in a battery of in vitro and in vivo tests.",
     "Genotoxicity"),
    ("No reproductive toxicity was observed in male or female rats.", "Reproductive Toxicity"),
    ("Oral administration to pregnant rabbits during organogenesis was without evidence of teratogenic effects.",
     "Reproductive Toxicity"),
    ("Drug X was negative in the Ames bacterial reverse mutation mutagenicity assay.", "Genotoxicity"),
    ("The Ames test was negative.", "Genotoxicity"),
    ("Drug X was non-mutagenic in the mouse lymphoma assay.", "Genotoxicity"),
    ("Hepatotoxicity was not observed in controlled clinical trials.", "Organ Toxicity"),
    ("Epidemiologic data showed no increased risk of birth defects.", "Reproductive Toxicity"),
]

AFFIRMED_SENTENCES = [
    ("Use during pregnancy is not recommended; Drug X may cause fetal harm and embryo-fetal toxicity.",
     "Reproductive Toxicity"),
    ("Not recommended in pregnancy because Drug X may cause embryo-fetal toxicity.", "Reproductive Toxicity"),
    ("Drug X was not teratogenic in rats but caused embryo-fetal toxicity in rabbits.", "Reproductive Toxicity"),
    ("Drug X can cause embryo-fetal toxicity, including fetal death.", "Reproductive Toxicity"),
    ("Drug X was carcinogenic in mice, producing hepatocellular tumors.", "Carcinogenicity"),
    ("Drug X was clastogenic in the in vitro chromosomal aberration assay.", "Genotoxicity"),
    ("Do not use Drug X during pregnancy.", "Pregnancy / Lactation"),
    ("There are no adequate and well-controlled studies of Drug X in pregnancy.", "Pregnancy / Lactation"),
    ("Safety and efficacy in patients with hepatotoxicity have not been established.", "Organ Toxicity"),
    ("Long-term carcinogenicity studies have not been conducted.", "Carcinogenicity"),
    ("Use negative pressure rooms when compounding.", "Engineering Controls"),
    ("Females of reproductive potential who are not pregnant must comply with the pregnancy testing requirements.",
     "Pregnancy / Lactation"),
]


def test_negated_sentences_dont_raise_risk():
    for text, cat in NEGATED_SENTENCES:
        r = scan_for_hazards(text)
        assert cat in _neg(r) and cat not in _aff(r), (text, r)
        assert r["risk_level"] == "none", (text, r["risk_level"])
        assert r["risk_level_ignoring_negation"] != "none", text  # old scan flagged it
        assert r["negated_evidence"] and all(e["clause"] for e in r["negated_evidence"])


def test_affirmed_sentences_still_flag():
    for text, cat in AFFIRMED_SENTENCES:
        r = scan_for_hazards(text)
        assert cat in _aff(r), (text, r)
        assert r["risk_level"] == r["risk_level_ignoring_negation"], text


def test_negation_is_per_occurrence():
    # A negated mention does not hide an affirmed one elsewhere in the text.
    r = scan_for_hazards("Drug X was not carcinogenic in rats. Drug X was carcinogenic in mice.")
    assert "Carcinogenicity" in _aff(r) and r["risk_level"] == "high"
    # Negation does not cross a sentence boundary.
    t = "No dose adjustment is needed. Hepatotoxicity occurred in 2% of patients."
    i = t.lower().index("hepatotoxicity")
    assert not is_negated(t.lower(), i, i + 14)
    # Window: a cue far before the term does not negate it.
    r = scan_for_hazards("No patients in the study who received the highest dose for twelve weeks developed "
                         "anything other than hepatotoxicity.")
    assert "Organ Toxicity" in _aff(r)


def test_boxed_warning_not_negated():
    r = scan_for_hazards("WARNING\nEMBRYO-FETAL TOXICITY. Do not use in pregnancy.")
    assert r["boxed_warning"] and r["risk_level"] == "high"


def test_highlight_marks_negated():
    out = highlight_keywords("Drug X was not teratogenic in rats but caused embryo-fetal toxicity in rabbits.")
    assert '<span class="hl-niosh hl-negated" title="negated: not counted in risk">teratogenic</span>' in out
    assert '<span class="hl-niosh">embryo-fetal</span>' in out
    assert "<span" not in highlight_keywords("No <b>tags</b> here").replace("&lt;", "")


if __name__ == "__main__":
    for name, fn in list(globals().items()):
        if name.startswith("test_"):
            fn()
            print("ok", name)
