from __future__ import annotations

from copy import deepcopy
from pathlib import Path
from typing import Any

from app.extractors.excel_extractor import ExcelExtractor
from app.extractors.pdf_extractor import PDFExtractor
from app.matching.verification_engine import verify_boms
from app.pairing.file_pairer import FilePair, FilePairer


class BatchProcessor:
    """End-to-end processor for Excel/PDF BOM verification."""

    def __init__(self) -> None:
        self.pairer = FilePairer()

    def process_folders(
        self,
        excel_folder: str | Path,
        pdf_folder: str | Path,
    ) -> dict[str, Any]:
        pairs = self.pairer.pair_folders(excel_folder, pdf_folder)
        pair_results = [self._process_pair(pair) for pair in pairs]
        summary = self._build_batch_summary(pair_results)

        return {
            "excel_folder": str(excel_folder),
            "pdf_folder": str(pdf_folder),
            "pair_count": len(pairs),
            "pairs": pair_results,
            "summary": summary,
        }

    def _process_pair(self, pair: FilePair) -> dict[str, Any]:
        result: dict[str, Any] = {
            "key": pair.key,
            "status": pair.status,
            "message": pair.message,
            "match_method": pair.match_method,
            "match_score": pair.match_score,
            "excel_file": str(pair.excel_file) if pair.excel_file else None,
            "pdf_file": str(pair.pdf_file) if pair.pdf_file else None,
            "excel_data": None,
            "pdf_data": None,
            "excel_records": [],
            "pdf_records": [],
            "verification": None,
            "warnings": [],
            "error": None,
        }

        if pair.status != "PAIRED":
            return result

        if pair.excel_file is None or pair.pdf_file is None:
            result["status"] = "REVIEW"
            result["message"] = (
                "Pair is marked as PAIRED but one of the required files is missing."
            )
            return result

        excel_data: dict[str, Any] | None = None
        pdf_data: dict[str, Any] | None = None

        try:
            excel_data = ExcelExtractor(pair.excel_file).extract()
            result["excel_data"] = deepcopy(excel_data)
            result["excel_records"] = self._flatten_excel_records(excel_data)
            if not result["excel_records"]:
                result["warnings"].append("Excel extraction returned no BOM rows.")

            pdf_data = PDFExtractor(pair.pdf_file).extract()
            result["pdf_data"] = deepcopy(pdf_data)
            result["pdf_records"] = self._flatten_pdf_records(pdf_data)
            if not result["pdf_records"]:
                warning = "PDF extraction returned no current BOM rows."
                if pdf_data.get("removed_rows"):
                    warning += " Removed rows were preserved separately and excluded from the current BOM."
                result["warnings"].append(warning)

            verification = verify_boms(excel_data, pdf_data)
            result["verification"] = deepcopy(verification)
            result["status"] = "VERIFIED"
            result["message"] = "Excel and PDF BOMs extracted and verified."
            return result

        except Exception as error:
            result["status"] = "PROCESSING_ERROR"
            result["message"] = "BOM verification could not be completed."
            result["error"] = str(error)

            if excel_data is not None and result["excel_data"] is None:
                result["excel_data"] = deepcopy(excel_data)
                result["excel_records"] = self._flatten_excel_records(excel_data)

            if pdf_data is not None and result["pdf_data"] is None:
                result["pdf_data"] = deepcopy(pdf_data)
                result["pdf_records"] = self._flatten_pdf_records(pdf_data)

            return result

    @staticmethod
    def _flatten_excel_records(excel_data: dict[str, Any]) -> list[dict[str, Any]]:
        records: list[dict[str, Any]] = []

        for sheet_name, sheet_data in excel_data.get("sheets", {}).items():
            for record in sheet_data.get("records", []):
                records.append(
                    {
                        "source": "excel",
                        "part_number": record.get("Part Number"),
                        "quantity": record.get("QTY"),
                        "description": record.get("Description"),
                        "revision": record.get("REV"),
                        "sheet_name": sheet_name,
                        "excel_row": record.get("_excel_row"),
                    }
                )

        return records

    @staticmethod
    def _flatten_pdf_records(pdf_data: dict[str, Any]) -> list[dict[str, Any]]:
        records: list[dict[str, Any]] = []

        for record in pdf_data.get("bom_rows", []):
            records.append(
                {
                    "source": "pdf",
                    "part_number": record.get("part_number"),
                    "quantity": record.get("quantity"),
                    "description": record.get("description"),
                    "pdf_page": record.get("pdf_page"),
                    "type": record.get("type"),
                    "class": record.get("class"),
                    "location": record.get("location"),
                    "section": record.get("section"),
                    "raw_line": record.get("raw_line"),
                    "model_filename": record.get("model_filename"),
                    "revision": record.get("revision"),
                }
            )

        return records

    @staticmethod
    def _build_batch_summary(pair_results: list[dict[str, Any]]) -> dict[str, Any]:
        summary: dict[str, Any] = {
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
            "warnings": [],
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

            for warning in pair_result.get("warnings", []):
                summary["warnings"].append(
                    {
                        "key": pair_result.get("key"),
                        "status": status,
                        "warning": warning,
                    }
                )

            verification = pair_result.get("verification")
            if not verification:
                continue

            verification_summary = verification.get("summary", {})
            summary["total_bom_rows"] += verification_summary.get("total", 0)
            summary["match"] += verification_summary.get("match", 0)
            summary["qty_mismatch"] += verification_summary.get("qty_mismatch", 0)
            summary["bom_excel_only"] += verification_summary.get("excel_only", 0)
            summary["bom_pdf_only"] += verification_summary.get("pdf_only", 0)
            summary["bom_review"] += verification_summary.get("review", 0)

        return summary
