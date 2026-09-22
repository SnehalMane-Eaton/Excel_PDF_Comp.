from app.core.models import VerificationResult


def result_to_dict(result: VerificationResult) -> dict:
    return {
        "excel_file": str(result.excel_file),
        "pdf_file": str(result.pdf_file),
        "passed": result.passed,
        "matched": [field.__dict__ for field in result.matched],
        "missing": [field.__dict__ for field in result.missing],
        "uncertain": [{"field": field.__dict__, "candidate": candidate, "score": score} for field, candidate, score in result.uncertain],
        "errors": result.errors,
    }
