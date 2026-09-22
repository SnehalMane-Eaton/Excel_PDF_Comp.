from app.matching.normalizer import normalize


def test_normalize_collapses_punctuation_and_case():
    assert normalize("  Part-N° 42 ") == "part n 42"
