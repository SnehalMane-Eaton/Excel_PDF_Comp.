from pathlib import Path
from typing import Any
import re

import pymupdf


class PDFExtractor:
    """
    Extract BOM information from Eaton PDFs.

    Supported formats:
    1. Eaton structured / exploded BOM PDFs.
    2. Eaton engineering drawing PDFs containing:
       - PART NO / PART NUMBER
       - MODEL FILENAME
       - .ipt model filenames
       - one drawing per page
    """

    def __init__(self, file_path: str | Path):
        self.file_path = Path(file_path)

    def extract(self) -> dict[str, Any]:
        if not self.file_path.exists():
            raise FileNotFoundError(
                f"PDF file not found: {self.file_path}"
            )

        document = pymupdf.open(self.file_path)

        result: dict[str, Any] = {
            "file": str(self.file_path),
            "page_count": len(document),
            "pages": [],
            "bom_rows": [],
            "removed_rows": [],
        }

        try:
            # Detect the PDF format from the first pages.
            is_structured_bom = self._detect_structured_bom(document)

            if is_structured_bom:
                self._extract_structured_bom(document, result)
            else:
                self._extract_engineering_drawings(document, result)

        finally:
            document.close()

        return result

    # ------------------------------------------------------------------
    # FORMAT DETECTION
    # ------------------------------------------------------------------

    def _detect_structured_bom(self, document) -> bool:
        """
        Detect the older Eaton Exploded Structured BOM format.
        """

        pages_to_check = min(5, len(document))

        for index in range(pages_to_check):
            text = document[index].get_text("text")

            normalized = re.sub(r"\s+", " ", text).lower()

            if (
                "style number" in normalized
                and "description" in normalized
                and "qty" in normalized
            ):
                return True

            if "current product structure" in normalized:
                return True

            if "removed from product structure" in normalized:
                return True

        return False

    # ------------------------------------------------------------------
    # OLD STRUCTURED BOM FORMAT
    # ------------------------------------------------------------------

    def _extract_structured_bom(
        self,
        document,
        result: dict[str, Any],
    ) -> None:

        in_bom_table = False
        section_status = "unknown"

        for page_number, page in enumerate(document, start=1):
            text = page.get_text("text")

            result["pages"].append(
                {
                    "page": page_number,
                    "text": text,
                }
            )

            for line in text.splitlines():
                raw_line = line
                clean_line = self._clean_line(line)

                if not clean_line:
                    continue

                if self._is_bom_header(clean_line):
                    in_bom_table = True
                    continue

                if "Current product structure" in clean_line:
                    in_bom_table = True
                    section_status = "current"
                    continue

                if "Removed from product structure" in clean_line:
                    in_bom_table = True
                    section_status = "removed"
                    continue

                if "Removed from ATS" in clean_line:
                    section_status = "removed"
                    continue

                if not in_bom_table:
                    continue

                if section_status not in {"current", "removed"}:
                    continue

                if self._is_deleted_row(raw_line):
                    row = self._parse_structured_bom_line(
                        clean_line,
                        page_number,
                        "removed",
                    )

                    if row is not None:
                        result["removed_rows"].append(row)

                    continue

                row = self._parse_structured_bom_line(
                    clean_line,
                    page_number,
                    section_status,
                )

                if row is None:
                    continue

                if section_status == "removed":
                    result["removed_rows"].append(row)
                else:
                    result["bom_rows"].append(row)

    @staticmethod
    def _clean_line(line: str) -> str:
        line = line.strip()

        if not line:
            return ""

        line = re.sub(r"^b\s+", "", line)

        return line

    @staticmethod
    def _is_deleted_row(line: str) -> bool:
        return bool(
            re.match(
                r"^\s*d\s*\(\d+\)",
                line,
                flags=re.IGNORECASE,
            )
        )

    @staticmethod
    def _is_bom_header(line: str) -> bool:
        normalized = re.sub(r"\s+", " ", line).lower()

        return (
            "style number" in normalized
            and "description" in normalized
            and "qty" in normalized
        )

    def _parse_structured_bom_line(
        self,
        line: str,
        page_number: int,
        section_status: str,
    ) -> dict[str, Any] | None:

        if ":" not in line:
            return None

        part_number, remaining = line.split(":", 1)

        part_number = self._clean_part_number_prefix(part_number)
        remaining = remaining.strip()

        if not self._looks_like_part_number(part_number):
            return None

        parsed = self._parse_structured_fields(remaining)

        if parsed is None:
            return None

        (
            description,
            quantity,
            bom_type,
            bom_class,
            location,
        ) = parsed

        return {
            "pdf_page": page_number,
            "part_number": part_number,
            "description": description,
            "quantity": quantity,
            "type": bom_type,
            "class": bom_class,
            "location": location,
            "section": section_status,
            "raw_line": line,
        }

    @staticmethod
    def _clean_part_number_prefix(value: str) -> str:
        value = value.strip()

        value = re.sub(
            r"^[^\w]*[A-Za-z○●oOdD]?\s*\(\d+\)\s*",
            "",
            value,
        )

        value = re.sub(
            r"^[^A-Za-z0-9]+",
            "",
            value,
        )

        return value.strip()

    @staticmethod
    def _looks_like_part_number(part_number: str) -> bool:
        if not part_number:
            return False

        if " " in part_number:
            return False

        if len(part_number) < 3:
            return False

        if not re.search(r"[A-Za-z0-9]", part_number):
            return False

        rejected = {
            "ITEM",
            "QUANTITY",
            "LOGIC",
            "CONTACTOR",
            "POLES",
            "RATING",
            "VOLTAGE",
            "FREQUENCY",
            "PHASES",
            "NUMBER",
        }

        if part_number.upper() in rejected:
            return False

        return True

    @staticmethod
    def _parse_structured_fields(
        text: str,
    ) -> tuple[str, int | float, str, str, str] | None:

        tokens = text.split()

        if len(tokens) < 5:
            return None

        location = tokens[-1]
        bom_class = tokens[-2]
        bom_type = tokens[-3]
        quantity_token = tokens[-4]

        quantity = PDFExtractor._parse_number(quantity_token)

        if quantity is None:
            return None

        description_tokens = tokens[:-4]

        if not description_tokens:
            return None

        description = " ".join(description_tokens).strip()

        return (
            description,
            quantity,
            bom_type,
            bom_class,
            location,
        )

    # ------------------------------------------------------------------
    # ENGINEERING DRAWING FORMAT
    # ------------------------------------------------------------------

    def _extract_engineering_drawings(
        self,
        document,
        result: dict[str, Any],
    ) -> None:
        """
        Parse Eaton engineering drawing PDFs.

        The drawing format contains information similar to:

            PART NO
            PART NUMBER
            ...
            66A8370H08
            ...
            MODEL FILENAME
            ...
            66A8370H08.ipt
            6
            H08

        The model filename is treated as the drawing's part number.

        The numeric value immediately following the model filename
        is used as the drawing quantity when available.
        """

        for page_number, page in enumerate(document, start=1):
            text = page.get_text("text")

            result["pages"].append(
                {
                    "page": page_number,
                    "text": text,
                }
            )

            row = self._parse_engineering_drawing_page(
                text=text,
                page_number=page_number,
            )

            if row is not None:
                result["bom_rows"].append(row)

    def _parse_engineering_drawing_page(
        self,
        text: str,
        page_number: int,
    ) -> dict[str, Any] | None:

        lines = [
            line.strip()
            for line in text.splitlines()
            if line.strip()
        ]

        if not lines:
            return None

        # --------------------------------------------------------------
        # Find MODEL FILENAME / .ipt
        # --------------------------------------------------------------

        model_index = self._find_line_index(
            lines,
            "MODEL FILENAME",
        )

        model_filename = self._find_ipt_filename(lines)

        model_part_number = ""

        if model_filename:
            model_part_number = re.sub(
                r"\.ipt$",
                "",
                model_filename,
                flags=re.IGNORECASE,
            ).strip()

        # --------------------------------------------------------------
        # Fallback to PART NUMBER if model filename was not found.
        # --------------------------------------------------------------

        if not model_part_number:
            model_part_number = self._extract_drawing_part_number(lines)

        if not model_part_number:
            return None

        model_part_number = self._clean_engineering_part_number(
            model_part_number
        )

        if not self._looks_like_engineering_part_number(
            model_part_number
        ):
            return None

        # --------------------------------------------------------------
        # Quantity
        # --------------------------------------------------------------

        quantity = self._find_engineering_quantity(
            lines,
            model_filename=model_filename,
            model_index=model_index,
        )

        # --------------------------------------------------------------
        # Description
        # --------------------------------------------------------------

        description = self._extract_engineering_description(
            lines,
            model_part_number=model_part_number,
        )

        # --------------------------------------------------------------
        # Revision
        # --------------------------------------------------------------

        revision = self._extract_revision(lines)

        return {
            "pdf_page": page_number,
            "part_number": model_part_number,
            "description": description,
            "quantity": quantity,
            "type": "ENGINEERING_DRAWING",
            "class": "",
            "location": "",
            "section": "current",
            "raw_line": model_filename or model_part_number,
            "model_filename": model_filename,
            "revision": revision,
        }

    @staticmethod
    def _find_line_index(
        lines: list[str],
        target: str,
    ) -> int | None:

        target_normalized = re.sub(
            r"\s+",
            " ",
            target.strip().upper(),
        )

        for index, line in enumerate(lines):
            normalized = re.sub(
                r"\s+",
                " ",
                line.upper(),
            )

            if normalized == target_normalized:
                return index

        return None

    @staticmethod
    def _find_ipt_filename(lines: list[str]) -> str | None:

        pattern = re.compile(
            r"\b[A-Za-z0-9][A-Za-z0-9._-]*\.ipt\b",
            flags=re.IGNORECASE,
        )

        for line in lines:
            match = pattern.search(line)

            if match:
                return match.group(0)

        return None

    @staticmethod
    def _extract_drawing_part_number(
        lines: list[str],
    ) -> str | None:

        # Prefer an explicit PART NUMBER area.
        part_index = None

        for index, line in enumerate(lines):
            normalized = re.sub(
                r"\s+",
                " ",
                line.upper(),
            )

            if normalized in {
                "PART NUMBER",
                "PART NO",
            }:
                part_index = index
                break

        if part_index is not None:
            candidate_pattern = re.compile(
                r"^[A-Za-z0-9][A-Za-z0-9._/-]{2,}$"
            )

            # Search a short window after the header.
            for candidate in lines[
                part_index + 1 : part_index + 12
            ]:
                candidate = candidate.strip()

                if candidate_pattern.fullmatch(candidate):
                    if not PDFExtractor._is_generic_drawing_token(
                        candidate
                    ):
                        return candidate

        # Final fallback: look for typical Eaton-style part numbers.
        for line in lines:
            candidate = line.strip()

            if PDFExtractor._looks_like_engineering_part_number(
                candidate
            ):
                if re.search(
                    r"[A-Za-z]",
                    candidate,
                ):
                    return candidate

        return None

    @staticmethod
    def _looks_like_engineering_part_number(
        value: str,
    ) -> bool:

        if not value:
            return False

        value = value.strip()

        if len(value) < 4:
            return False

        if " " in value:
            return False

        if not re.search(r"[A-Za-z]", value):
            return False

        if not re.search(r"\d", value):
            return False

        if not re.fullmatch(
            r"[A-Za-z0-9._/-]+",
            value,
        ):
            return False

        if PDFExtractor._is_generic_drawing_token(value):
            return False

        return True

    @staticmethod
    def _is_generic_drawing_token(
        value: str,
    ) -> bool:

        normalized = value.strip().upper()

        rejected = {
            "PART",
            "PARTNO",
            "PARTNUMBER",
            "NUMBER",
            "REV",
            "REVISION",
            "MODEL",
            "FILENAME",
            "TITLE",
            "DWG",
            "DRAWING",
            "SHEET",
            "RAW",
            "MATERIAL",
            "LIST",
            "FINISH",
            "STEEL",
            "PAINT",
            "WHITE",
            "BLACK",
            "RMS1",
            "RMS2",
            "NTS",
            "ECN",
        }

        if normalized in rejected:
            return True

        if normalized.startswith("ECN"):
            return True

        if normalized.endswith(".IPT"):
            return True

        return False

    @staticmethod
    def _clean_engineering_part_number(
        value: str,
    ) -> str:

        value = value.strip()

        value = re.sub(
            r"\.ipt$",
            "",
            value,
            flags=re.IGNORECASE,
        )

        value = re.sub(
            r"\s+",
            "",
            value,
        )

        return value.upper()

    @staticmethod
    def _find_engineering_quantity(
        lines: list[str],
        model_filename: str | None,
        model_index: int | None,
    ) -> int | float | None:

        if not model_filename:
            return None

        # Locate the exact .ipt line.
        filename_index = None

        for index, line in enumerate(lines):
            if model_filename.lower() in line.lower():
                filename_index = index
                break

        if filename_index is None:
            return None

        # In this drawing format the quantity is normally the next
        # numeric value after the model filename.
        #
        # Search only a very small window so that dimensions,
        # revision history and drawing annotations are not mistaken
        # for quantity.

        for index in range(
            filename_index + 1,
            min(filename_index + 5, len(lines)),
        ):
            candidate = lines[index].strip()

            number = PDFExtractor._parse_number(
                candidate
            )

            if number is None:
                continue

            # Quantity must be positive.
            if number <= 0:
                continue

            return number

        return None

    @staticmethod
    def _extract_engineering_description(
        lines: list[str],
        model_part_number: str,
    ) -> str:

        # Common Eaton drawing title candidates.
        ignored = {
            "RAW MATERIAL LIST",
            "PART NO",
            "PART NUMBER",
            "MODEL FILENAME",
            "DWG NO",
            "MODEL REV",
            "SHEET OF",
            "TITLE",
            "ALL REV. NOTES",
            "LISTED ON SHT. 1",
            "REV",
            "T.BLK",
            "REV-4",
            "REV - 4",
            "THIRD ANGLE",
            "PROJECTION",
        }

        # Search the title area for useful descriptive text.
        #
        # The title in the supplied drawing appears before the model
        # filename and after the drawing metadata.

        for line in lines:
            candidate = line.strip()

            if not candidate:
                continue

            normalized = re.sub(
                r"\s+",
                " ",
                candidate.upper(),
            )

            if normalized in ignored:
                continue

            if candidate.upper() == model_part_number.upper():
                continue

            if candidate.lower().endswith(".ipt"):
                continue

            if PDFExtractor._looks_like_number_only(
                candidate
            ):
                continue

            if re.fullmatch(
                r"[A-Z0-9._/-]+",
                candidate.upper(),
            ):
                if (
                    len(candidate) <= 3
                    or candidate.upper()
                    in {
                        "H01",
                        "H02",
                        "H03",
                        "H04",
                        "H05",
                        "H06",
                        "H07",
                        "H08",
                        "H09",
                        "H10",
                        "H11",
                        "H12",
                        "H13",
                        "H14",
                        "H15",
                        "H16",
                        "H17",
                        "H18",
                        "H19",
                        "H20",
                        "H21",
                        "H22",
                        "H23",
                        "H24",
                        "H25",
                        "H26",
                        "H27",
                        "H28",
                        "H29",
                    }
                ):
                    continue

            # Exclude obvious drawing / dimension / metadata text.
            if re.fullmatch(
                r"[\d.()+\-P]+",
                candidate.upper(),
            ):
                continue

            if candidate.startswith("ECN"):
                continue

            if "EATON CORPORATION" in normalized:
                continue

            if len(candidate) < 4:
                continue

            # Avoid long revision-history sentences.
            if len(candidate) > 80:
                continue

            return candidate

        return ""

    @staticmethod
    def _extract_revision(
        lines: list[str],
    ) -> str:

        for index, line in enumerate(lines):
            normalized = line.strip().upper()

            if normalized == "MODEL REV":
                # Search nearby for a simple revision token.
                for candidate in lines[
                    index + 1 : index + 8
                ]:
                    candidate = candidate.strip()

                    if re.fullmatch(
                        r"[A-Z0-9]{1,4}",
                        candidate,
                    ):
                        if candidate not in {
                            "DWG",
                            "NO",
                            "REV",
                            "SHEET",
                            "OF",
                        }:
                            return candidate

        return ""

    @staticmethod
    def _looks_like_number_only(
        value: str,
    ) -> bool:

        value = value.strip()

        if not value:
            return True

        return bool(
            re.fullmatch(
                r"[\d.()+\-P/]+",
                value.upper(),
            )
        )

    # ------------------------------------------------------------------
    # GENERAL NUMBER PARSER
    # ------------------------------------------------------------------

    @staticmethod
    def _parse_number(
        value: str,
    ) -> int | float | None:

        value = value.strip().replace(",", "")

        try:
            number = float(value)
        except (ValueError, TypeError):
            return None

        if number.is_integer():
            return int(number)

        return number