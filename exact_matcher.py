from collections import defaultdict
from typing import Any

from app.matching.normalizer import normalize_part_number


def build_excel_bom_index(
    excel_data: dict[str, Any],
) -> dict[str, list[dict[str, Any]]]:
    """
    Build an index of Excel BOM rows by normalized part number.

    Multiple Excel rows with the same part number are retained
    because they may need to be aggregated.
    """

    index: dict[str, list[dict[str, Any]]] = defaultdict(list)

    for sheet_name, sheet_data in excel_data["sheets"].items():

        for record in sheet_data["records"]:

            raw_part_number = record.get("Part Number")
            part_number = normalize_part_number(raw_part_number)

            quantity = _parse_quantity(record.get("QTY"))

            description = record.get("Description")
            revision = record.get("REV")
            excel_row = record.get("_excel_row")

            # Blank part numbers are kept out of the comparison index.
            # They will be reported separately as REVIEW.
            if not part_number:
                continue

            index[part_number].append(
                {
                    "part_number": part_number,
                    "quantity": quantity,
                    "description": description,
                    "revision": revision,
                    "sheet": sheet_name,
                    "excel_row": excel_row,
                }
            )

    return dict(index)


def build_pdf_bom_index(
    pdf_data: dict[str, Any],
) -> dict[str, list[dict[str, Any]]]:
    """
    Build an index of CURRENT PDF BOM rows by normalized part number.

    Removed BOM rows are intentionally excluded.
    """

    index: dict[str, list[dict[str, Any]]] = defaultdict(list)

    for record in pdf_data["bom_rows"]:

        part_number = normalize_part_number(
            record.get("part_number")
        )

        if not part_number:
            continue

        index[part_number].append(
            {
                "part_number": part_number,
                "quantity": record.get("quantity"),
                "description": record.get("description"),
                "pdf_page": record.get("pdf_page"),
                "type": record.get("type"),
                "class": record.get("class"),
                "location": record.get("location"),
                "section": record.get("section"),
                "raw_line": record.get("raw_line"),
            }
        )

    return dict(index)


def aggregate_excel_quantity(
    rows: list[dict[str, Any]],
) -> int | float | None:
    """Sum quantities for duplicate Excel part numbers."""

    quantities = [
        row["quantity"]
        for row in rows
        if row.get("quantity") is not None
    ]

    if not quantities:
        return None

    total = sum(quantities)

    if float(total).is_integer():
        return int(total)

    return total


def aggregate_pdf_quantity(
    rows: list[dict[str, Any]],
) -> int | float | None:
    """Sum quantities for duplicate PDF part numbers."""

    quantities = [
        row["quantity"]
        for row in rows
        if row.get("quantity") is not None
    ]

    if not quantities:
        return None

    total = sum(quantities)

    if float(total).is_integer():
        return int(total)

    return total


def get_pdf_pages(
    rows: list[dict[str, Any]],
) -> list[int]:
    """Return unique PDF pages containing a part number."""

    pages = {
        row["pdf_page"]
        for row in rows
        if row.get("pdf_page") is not None
    }

    return sorted(pages)


def _parse_quantity(value):
    """Convert a quantity into a numeric value."""

    if value is None:
        return None

    if isinstance(value, bool):
        return None

    try:
        number = float(value)

        if number.is_integer():
            return int(number)

        return number

    except (ValueError, TypeError):
        return None