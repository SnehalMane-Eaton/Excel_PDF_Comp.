import json
from pathlib import Path
from app.core.models import BatchSummary, VerificationResult
from app.reporting.report_generator import write_json


def test_write_json_report(tmp_path):
    summary = BatchSummary([VerificationResult(Path("a.xlsx"), Path("a.pdf"))])
    output = tmp_path / "report.json"
    write_json(summary, output)
    payload = json.loads(output.read_text(encoding="utf-8"))
    assert payload["passed"] == 1
    assert payload["results"][0]["passed"] is True
