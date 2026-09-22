from app.core import batch_processor as batch_processor_module
from app.core.batch_processor import BatchProcessor
from app.pairing.file_pairer import FilePair


class FakeExtractor:
    def __init__(self, payload):
        self.payload = payload

    def extract(self):
        return self.payload


def test_batch_processor_retains_extracted_snapshots_and_summary(monkeypatch):
    excel_data = {
        "file": "excel.xlsx",
        "sheet_names": ["Main"],
        "sheets": {
            "Main": {
                "records": [
                    {"Part Number": "001", "QTY": 1, "Description": "Excel desc", "REV": "A", "_excel_row": 2},
                    {"Part Number": "", "QTY": 0, "Description": "Review row", "REV": "", "_excel_row": 3},
                ]
            }
        },
    }
    pdf_data = {
        "file": "bom.pdf",
        "page_count": 1,
        "pages": [],
        "bom_rows": [
            {"part_number": "001", "quantity": 1, "description": "PDF desc", "pdf_page": 1, "type": "bom", "class": "c", "location": "L1", "section": "current", "raw_line": "001"},
        ],
        "removed_rows": [
            {"part_number": "REMOVED", "quantity": 5, "description": "Removed", "pdf_page": 1, "type": "bom", "class": "c", "location": "L1", "section": "removed", "raw_line": "REMOVED"},
        ],
    }

    monkeypatch.setattr(batch_processor_module, "ExcelExtractor", lambda path: FakeExtractor(excel_data))
    monkeypatch.setattr(batch_processor_module, "PDFExtractor", lambda path: FakeExtractor(pdf_data))

    processor = BatchProcessor()
    processor.pairer = type(
        "FakePairer",
        (),
        {
            "pair_folders": lambda self, excel_folder, pdf_folder: [
                FilePair(
                    excel_file=None,
                    pdf_file=None,
                    key="lonely-pdf",
                    status="PDF_ONLY_FILE",
                    message="No Excel match",
                ),
                FilePair(
                    excel_file="excel.xlsx",
                    pdf_file="bom.pdf",
                    key="paired",
                    status="PAIRED",
                    message="paired",
                ),
            ]
        },
    )()

    result = processor.process_folders("excel", "pdf")

    assert result["pair_count"] == 2
    paired = result["pairs"][1]
    assert paired["status"] == "VERIFIED"
    assert paired["excel_data"]["sheets"]["Main"]["records"][0]["Part Number"] == "001"
    assert paired["pdf_data"]["removed_rows"][0]["part_number"] == "REMOVED"
    assert paired["excel_records"][1]["description"] == "Review row"
    assert paired["pdf_records"][0]["description"] == "PDF desc"
    assert result["summary"]["verified"] == 1
    assert result["summary"]["pdf_only"] == 1
    assert result["summary"]["match"] == 1
    assert result["summary"]["bom_review"] == 1
