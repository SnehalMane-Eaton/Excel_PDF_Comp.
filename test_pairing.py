from pathlib import Path
from app.pairing.file_pairer import pair_files


def test_pair_files_by_normalized_stem(tmp_path):
    excel = tmp_path / "excel"
    pdf = tmp_path / "pdf"
    excel.mkdir(); pdf.mkdir()
    (excel / "Job-001.xlsx").touch()
    (pdf / "job_001.pdf").touch()
    pairs = pair_files(excel, pdf)
    assert len(pairs) == 1
    assert pairs[0].key == "job 001"
