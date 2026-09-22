from typing import Any

from app.matching.exact_matcher import (
    aggregate_excel_quantity,
    aggregate_pdf_quantity,
    build_excel_bom_index,
    build_pdf_bom_index,
    get_pdf_pages,
)


def verify_boms(
    excel_data: dict[str, Any],
    pdf_data: dict[str, Any],
) -> dict[str, Any]:

    excel_index = build_excel_bom_index(
        excel_data
    )

    pdf_index = build_pdf_bom_index(
        pdf_data
    )

    all_part_numbers = sorted(
        set(excel_index.keys())
        |
        set(pdf_index.keys())
    )

    results = []

    for part_number in all_part_numbers:

        excel_rows = excel_index.get(
            part_number,
            [],
        )

        pdf_rows = pdf_index.get(
            part_number,
            [],
        )

        excel_quantity = aggregate_excel_quantity(
            excel_rows
        )

        pdf_quantity = aggregate_pdf_quantity(
            pdf_rows
        )

        description = ""

        if excel_rows:
            description = (
                excel_rows[0].get(
                    "description"
                )
                or ""
            )

        if not description and pdf_rows:
            description = (
                pdf_rows[0].get(
                    "description"
                )
                or ""
            )

        excel_row_numbers = [
            row.get("excel_row")
            for row in excel_rows
            if row.get("excel_row") is not None
        ]

        pdf_pages = get_pdf_pages(
            pdf_rows
        )

        if excel_rows and pdf_rows:

            if excel_quantity is None:

                status = "REVIEW"
                remark = (
                    "Excel quantity is missing or invalid"
                )
                difference = None

            elif pdf_quantity is None:

                status = "REVIEW"
                remark = (
                    "PDF quantity is missing or invalid"
                )
                difference = None

            elif excel_quantity == pdf_quantity:

                status = "MATCH"
                remark = ""
                difference = 0

            else:

                status = "QTY_MISMATCH"
                remark = (
                    "Part number matched but "
                    "quantity differs"
                )
                difference = (
                    pdf_quantity
                    - excel_quantity
                )

        elif excel_rows:

            status = "EXCEL_ONLY"
            remark = (
                "Part number found in Excel "
                "but not in current PDF BOM"
            )
            difference = None

        else:

            status = "PDF_ONLY"
            remark = (
                "Part number found in current "
                "PDF BOM but not in Excel"
            )
            difference = None

        results.append(
            {
                "part_number": part_number,
                "description": description,
                "excel_quantity": excel_quantity,
                "pdf_quantity": pdf_quantity,
                "difference": difference,
                "status": status,
                "excel_rows": excel_row_numbers,
                "pdf_pages": pdf_pages,
                "remark": remark,
            }
        )

    # Blank Part Number rows in Excel
    for sheet_name, sheet_data in (
        excel_data["sheets"].items()
    ):

        for record in sheet_data["records"]:

            raw_part_number = record.get(
                "Part Number"
            )

            is_blank = (
                raw_part_number is None
                or not str(
                    raw_part_number
                ).strip()
            )

            if not is_blank:
                continue

            results.append(
                {
                    "part_number": "",
                    "description": record.get(
                        "Description"
                    ),
                    "excel_quantity": record.get(
                        "QTY"
                    ),
                    "pdf_quantity": None,
                    "difference": None,
                    "status": "REVIEW",
                    "excel_rows": [
                        record.get("_excel_row")
                    ],
                    "pdf_pages": [],
                    "remark": (
                        "Blank Part Number in Excel"
                    ),
                }
            )

    summary = _build_summary(
        results
    )

    return {
        "results": results,
        "summary": summary,
    }


def _build_summary(
    results: list[dict[str, Any]],
) -> dict[str, int]:

    summary = {
        "total": len(results),
        "match": 0,
        "qty_mismatch": 0,
        "excel_only": 0,
        "pdf_only": 0,
        "review": 0,
    }

    for result in results:

        status = result.get(
            "status"
        )

        if status == "MATCH":
            summary["match"] += 1

        elif status == "QTY_MISMATCH":
            summary["qty_mismatch"] += 1

        elif status == "EXCEL_ONLY":
            summary["excel_only"] += 1

        elif status == "PDF_ONLY":
            summary["pdf_only"] += 1

        elif status == "REVIEW":
            summary["review"] += 1

    return summary