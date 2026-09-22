from pathlib import Path
from typing import Any

from app.extractors.excel_extractor import ExcelExtractor
from app.extractors.pdf_extractor import PDFExtractor
from app.matching.verification_engine import verify_boms
from app.pairing.file_pairer import FilePair, FilePairer


class BatchProcessor:
    """
    End-to-end processor for Excel/PDF BOM verification.

    Workflow:

        folders
          ↓
        file pairing
          ↓
        Excel extraction
          ↓
        PDF extraction
          ↓
        BOM verification
          ↓
        consolidated results
    """

    def __init__(self):
        self.pairer = FilePairer()

    def process_folders(
        self,
        excel_folder: str | Path,
        pdf_folder: str | Path,
    ) -> dict[str, Any]:

        pairs = self.pairer.pair_folders(
            excel_folder,
            pdf_folder,
        )

        pair_results: list[dict[str, Any]] = []

        for pair in pairs:

            pair_result = self._process_pair(pair)

            pair_results.append(pair_result)

        summary = self._build_batch_summary(
            pair_results
        )

        return {
            "excel_folder": str(excel_folder),
            "pdf_folder": str(pdf_folder),
            "pair_count": len(pairs),
            "pairs": pair_results,
            "summary": summary,
        }

    def _process_pair(
        self,
        pair: FilePair,
    ) -> dict[str, Any]:

        result: dict[str, Any] = {
            "key": pair.key,
            "status": pair.status,
            "message": pair.message,
            "match_method": pair.match_method,
            "match_score": pair.match_score,
            "excel_file": (
                str(pair.excel_file)
                if pair.excel_file
                else None
            ),
            "pdf_file": (
                str(pair.pdf_file)
                if pair.pdf_file
                else None
            ),
            "verification": None,
            "error": None,
        }

        # Only a successfully paired Excel/PDF pair
        # can proceed to BOM verification.
        if pair.status != "PAIRED":
            return result

        if pair.excel_file is None or pair.pdf_file is None:
            result["status"] = "REVIEW"
            result["message"] = (
                "Pair is marked as PAIRED but one of the "
                "required files is missing."
            )
            return result

        try:

            excel_data = ExcelExtractor(
                pair.excel_file
            ).extract()

            pdf_data = PDFExtractor(
                pair.pdf_file
            ).extract()

            verification = verify_boms(
                excel_data,
                pdf_data,
            )

            result["verification"] = verification

            result["status"] = "VERIFIED"

            result["message"] = (
                "Excel and PDF BOMs extracted and verified."
            )

            return result

        except Exception as error:

            result["status"] = "PROCESSING_ERROR"

            result["message"] = (
                "BOM verification could not be completed."
            )

            result["error"] = str(error)

            return result

    @staticmethod
    def _build_batch_summary(
        pair_results: list[dict[str, Any]],
    ) -> dict[str, Any]:

        summary = {
            "total_pairs": len(pair_results),
            "verified": 0,
            "paired_not_verified": 0,
            "review": 0,
            "processing_error": 0,
            "excel_only": 0,
            "pdf_only": 0,
            "duplicate_files": 0,

            "total_bom_rows": 0,
            "match": 0,
            "qty_mismatch": 0,
            "bom_excel_only": 0,
            "bom_pdf_only": 0,
            "bom_review": 0,
        }

        for pair_result in pair_results:

            status = pair_result["status"]

            if status == "VERIFIED":
                summary["verified"] += 1

            elif status == "PROCESSING_ERROR":
                summary["processing_error"] += 1

            elif status == "REVIEW":
                summary["review"] += 1

            elif status == "EXCEL_ONLY_FILE":
                summary["excel_only"] += 1

            elif status == "PDF_ONLY_FILE":
                summary["pdf_only"] += 1

            elif status == "DUPLICATE_FILES":
                summary["duplicate_files"] += 1

            else:
                summary["paired_not_verified"] += 1

            verification = pair_result.get(
                "verification"
            )

            if not verification:
                continue

            verification_summary = verification.get(
                "summary",
                {},
            )

            summary["total_bom_rows"] += (
                verification_summary.get("total", 0)
            )

            summary["match"] += (
                verification_summary.get("match", 0)
            )

            summary["qty_mismatch"] += (
                verification_summary.get(
                    "qty_mismatch",
                    0,
                )
            )

            summary["bom_excel_only"] += (
                verification_summary.get(
                    "excel_only",
                    0,
                )
            )

            summary["bom_pdf_only"] += (
                verification_summary.get(
                    "pdf_only",
                    0,
                )
            )

            summary["bom_review"] += (
                verification_summary.get(
                    "review",
                    0,
                )
            )

        return summary