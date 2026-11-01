"""Quick offline checks for keyword matching and the crosswalk.

Run: python tests/test_core.py   (or: pytest tests/)
"""
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from rxlw_core import (  # noqa: E402
    BOXED_WARNING_PHRASES, NIOSH_CRITERIA, USP800_CRITERIA, find_best_matches,
    highlight_keywords, ingredient_agreement, keyword_in, normalize_drug_name, scan_for_hazards,
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


if __name__ == "__main__":
    for name, fn in list(globals().items()):
        if name.startswith("test_"):
            fn()
            print("ok", name)
