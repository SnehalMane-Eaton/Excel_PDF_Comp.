from pathlib import Path
from app.core.models import SourceField
from app.matching.verification_engine import verify


def test_verify_exact_and_missing_values():
    result = verify(Path("a.xlsx"), Path("a.pdf"), [SourceField("part", "ABC-42"), SourceField("size", "10")], ["Part ABC-42", "other"], 0.95)
    assert len(result.matched) == 1
    assert len(result.missing) == 1
    assert not result.passed
