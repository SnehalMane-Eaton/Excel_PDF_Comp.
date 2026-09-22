from pathlib import Path

from app.ui.main_window import VerificationController, normalize_bom_selection, summary_display_values


def test_normalize_bom_selection_and_summary_mapping():
    assert normalize_bom_selection("PDF") == ("PDF", "Excel")
    assert normalize_bom_selection("Excel") == ("Excel", "PDF")

    display = summary_display_values(
        {
            "total_pairs": 5,
            "verified": 2,
            "match": 8,
            "qty_mismatch": 1,
            "bom_pdf_only": 3,
            "bom_excel_only": 4,
            "bom_review": 6,
            "processing_error": 7,
            "review": 1,
            "excel_only": 1,
            "pdf_only": 1,
            "duplicate_files": 1,
            "paired_not_verified": 1,
        }
    )

    assert display == {
        "total_pairs": "5",
        "verified_pairs": "2",
        "matched": "8",
        "qty_mismatch": "1",
        "pdf_only": "3",
        "excel_only": "4",
        "review": "6",
        "errors": "7",
        "file_issues": "12",
    }


def test_verification_controller_passes_bom_order(tmp_path):
    calls = {}

    controller = VerificationController()
    controller.batch_processor = type(
        "FakeBatchProcessor",
        (),
        {"process_folders": lambda self, excel_folder, pdf_folder: {"summary": {}, "pairs": []}},
    )()
    controller.report_generator = type(
        "FakeReportGenerator",
        (),
        {
            "generate": lambda self, result, output_path, bom1_type, bom2_type: calls.update(
                {
                    "result": result,
                    "output_path": output_path,
                    "bom1_type": bom1_type,
                    "bom2_type": bom2_type,
                }
            )
        },
    )()

    report_path = tmp_path / "out.xlsx"
    result = controller.verify("excel-folder", "pdf-folder", report_path, bom1_type="PDF", bom2_type="Excel")

    assert result == {"summary": {}, "pairs": []}
    assert calls["output_path"] == str(report_path)
    assert calls["bom1_type"] == "PDF"
    assert calls["bom2_type"] == "Excel"
