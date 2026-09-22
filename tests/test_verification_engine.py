from app.matching.verification_engine import verify_boms


def test_verify_boms_preserves_duplicate_quantity_policy_and_blank_review():
    excel_data = {
        "file": "excel.xlsx",
        "sheet_names": ["Main"],
        "sheets": {
            "Main": {
                "records": [
                    {"Part Number": "AA-01", "QTY": 1, "Description": "Excel A", "REV": "A", "_excel_row": 2},
                    {"Part Number": "AA-01", "QTY": None, "Description": "Excel A duplicate", "REV": "A", "_excel_row": 3},
                    {"Part Number": "BB-02", "QTY": None, "Description": "Missing qty", "REV": "B", "_excel_row": 4},
                    {"Part Number": "", "QTY": 0, "Description": "Blank part review", "REV": "", "_excel_row": 5},
                ]
            }
        },
    }
    pdf_data = {
        "file": "bom.pdf",
        "page_count": 1,
        "pages": [],
        "bom_rows": [
            {"part_number": "aa-01", "quantity": 1, "description": "PDF A", "pdf_page": 1, "type": "bom", "class": "c", "location": "L1", "section": "current", "raw_line": "AA-01"},
            {"part_number": "BB-02", "quantity": None, "description": "PDF B", "pdf_page": 2, "type": "bom", "class": "c", "location": "L2", "section": "current", "raw_line": "BB-02"},
        ],
        "removed_rows": [],
    }

    verification = verify_boms(excel_data, pdf_data)
    results = {result["part_number"]: result for result in verification["results"] if result["part_number"]}

    assert results["AA-01"]["status"] == "MATCH"
    assert results["AA-01"]["excel_quantity"] == 1
    assert results["AA-01"]["pdf_quantity"] == 1

    assert results["BB-02"]["status"] == "REVIEW"
    assert results["BB-02"]["remark"] == "Excel quantity is missing or invalid"

    blank_rows = [result for result in verification["results"] if result["status"] == "REVIEW" and result["part_number"] == ""]
    assert len(blank_rows) == 1
    assert blank_rows[0]["description"] == "Blank part review"
    assert verification["summary"] == {
        "total": 3,
        "match": 1,
        "qty_mismatch": 0,
        "excel_only": 0,
        "pdf_only": 0,
        "review": 2,
    }


def test_verify_boms_excludes_removed_pdf_rows_from_current_bom():
    excel_data = {
        "file": "excel.xlsx",
        "sheet_names": ["Main"],
        "sheets": {
            "Main": {
                "records": [
                    {"Part Number": "RM-01", "QTY": 2, "Description": "Still in Excel", "REV": "A", "_excel_row": 2},
                ]
            }
        },
    }
    pdf_data = {
        "file": "bom.pdf",
        "page_count": 1,
        "pages": [],
        "bom_rows": [],
        "removed_rows": [
            {"part_number": "RM-01", "quantity": 2, "description": "Removed", "pdf_page": 1, "type": "bom", "class": "c", "location": "L1", "section": "removed", "raw_line": "RM-01"},
        ],
    }

    verification = verify_boms(excel_data, pdf_data)
    assert verification["results"][0]["status"] == "EXCEL_ONLY"
    assert verification["summary"]["excel_only"] == 1
