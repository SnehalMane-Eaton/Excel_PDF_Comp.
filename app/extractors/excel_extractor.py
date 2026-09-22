from pathlib import Path
from typing import Any

import pandas as pd


class ExcelExtractor:
    """
    Robust Excel BOM extractor.

    Supports:
    - Any worksheet name
    - Header row not necessarily on row 1
    - Different column ordering
    - Extra title/header rows above the BOM
    - Excel .xlsx / .xls / .xlsm files
    - Exact Part Number and QTY extraction
    """

    REQUIRED_PART_COLUMNS = {
        "part number",
        "partnumber",
        "part no",
        "part no.",
        "part #",
        "part",
        "style number",
        "style no",
        "style no.",
    }

    QUANTITY_COLUMNS = {
        "qty",
        "quantity",
        "qty.",
        "qnty",
        "q'ty",
        "unit qty",
    }

    DESCRIPTION_COLUMNS = {
        "description",
        "part description",
        "item description",
        "desc",
    }

    REVISION_COLUMNS = {
        "rev",
        "revision",
        "rev.",
    }

    def __init__(
        self,
        file_path: str | Path,
    ):
        self.file_path = Path(
            file_path
        )

    # =========================================================
    # SHEET NAMES
    # =========================================================

    def get_sheet_names(self) -> list[str]:

        if not self.file_path.exists():

            raise FileNotFoundError(
                f"Excel file not found: "
                f"{self.file_path}"
            )

        workbook = pd.ExcelFile(
            self.file_path,
            engine="openpyxl",
        )

        return workbook.sheet_names

    # =========================================================
    # MAIN EXTRACTION
    # =========================================================

    def extract(self) -> dict[str, Any]:

        if not self.file_path.exists():

            raise FileNotFoundError(
                f"Excel file not found: "
                f"{self.file_path}"
            )

        workbook = pd.ExcelFile(
            self.file_path,
            engine="openpyxl",
        )

        result: dict[str, Any] = {
            "file": str(
                self.file_path
            ),
            "sheet_names": workbook.sheet_names,
            "sheets": {},
        }

        for sheet_name in workbook.sheet_names:

            sheet_result = self._extract_sheet(
                sheet_name
            )

            result["sheets"][
                sheet_name
            ] = sheet_result

        return result

    # =========================================================
    # SHEET EXTRACTION
    # =========================================================

    def _extract_sheet(
        self,
        sheet_name: str,
    ) -> dict[str, Any]:

        # Read without assuming where the header is.
        raw_dataframe = pd.read_excel(
            self.file_path,
            sheet_name=sheet_name,
            engine="openpyxl",
            header=None,
            dtype=object,
        )

        if raw_dataframe.empty:

            return {
                "columns": [],
                "row_count": 0,
                "records": [],
                "header_row": None,
            }

        header_row = self._find_header_row(
            raw_dataframe
        )

        if header_row is None:

            # Still return the sheet, but don't
            # incorrectly interpret random rows as BOM data.

            return {
                "columns": [],
                "row_count": 0,
                "records": [],
                "header_row": None,
            }

        headers = self._build_headers(
            raw_dataframe.iloc[
                header_row
            ].tolist()
        )

        dataframe = raw_dataframe.iloc[
            header_row + 1:
        ].copy()

        dataframe.columns = headers

        dataframe = dataframe.reset_index(
            drop=True
        )

        # Remove completely blank rows.
        dataframe = dataframe.dropna(
            how="all"
        )

        # Excel row number:
        # pandas data row after header + actual
        # Excel row numbering.
        excel_rows = (
            dataframe.index
            + header_row
            + 2
        )

        dataframe.insert(
            0,
            "_excel_row",
            excel_rows,
        )

        dataframe = dataframe.where(
            pd.notna(dataframe),
            None,
        )

        records = dataframe.to_dict(
            orient="records"
        )

        # -----------------------------------------------------
        # Ensure standard columns exist.
        # -----------------------------------------------------

        records = [
            self._standardize_record(
                record
            )
            for record in records
        ]

        # -----------------------------------------------------
        # Keep rows which actually contain BOM information.
        #
        # We do NOT require Part Number to be populated because
        # blank Part Number rows must be reported as REVIEW.
        # -----------------------------------------------------

        bom_records = []

        for record in records:

            if self._is_meaningful_bom_row(
                record
            ):
                bom_records.append(
                    record
                )

        return {
            "columns": list(
                dataframe.columns
            ),
            "row_count": len(
                bom_records
            ),
            "records": bom_records,
            "header_row": header_row + 1,
        }

    # =========================================================
    # HEADER DETECTION
    # =========================================================

    def _find_header_row(
        self,
        dataframe: pd.DataFrame,
    ) -> int | None:

        best_row = None
        best_score = 0

        # Only inspect a reasonable number of top rows.
        # BOM headers are normally near the beginning.
        max_rows = min(
            len(dataframe),
            40,
        )

        for row_index in range(
            max_rows
        ):

            values = dataframe.iloc[
                row_index
            ].tolist()

            normalized_values = {
                self._normalize_header(
                    value
                )
                for value in values
                if value is not None
                and str(value).strip()
            }

            if not normalized_values:
                continue

            score = 0

            # Part number column is the most important.
            if normalized_values & self.REQUIRED_PART_COLUMNS:
                score += 5

            if normalized_values & self.QUANTITY_COLUMNS:
                score += 3

            if normalized_values & self.DESCRIPTION_COLUMNS:
                score += 2

            if normalized_values & self.REVISION_COLUMNS:
                score += 1

            # Strong preference for a BOM-like header.
            if (
                normalized_values
                & self.REQUIRED_PART_COLUMNS
                and normalized_values
                & self.QUANTITY_COLUMNS
            ):
                score += 5

            if score > best_score:

                best_score = score
                best_row = row_index

        if best_score >= 5:
            return best_row

        return None

    # =========================================================
    # HEADER BUILDING
    # =========================================================

    @staticmethod
    def _build_headers(
        values: list[Any],
    ) -> list[str]:

        headers = []

        used: dict[str, int] = {}

        for index, value in enumerate(
            values
        ):

            if value is None:

                header = (
                    f"Column_{index + 1}"
                )

            else:

                header = str(
                    value
                ).strip()

                if not header:

                    header = (
                        f"Column_{index + 1}"
                    )

            # Make duplicate headers unique.
            base_header = header

            if header in used:

                used[header] += 1

                header = (
                    f"{base_header}_"
                    f"{used[base_header]}"
                )

            else:

                used[header] = 1

            headers.append(
                header
            )

        return headers

    # =========================================================
    # RECORD STANDARDIZATION
    # =========================================================

    def _standardize_record(
        self,
        record: dict[str, Any],
    ) -> dict[str, Any]:

        standardized = dict(
            record
        )

        part_number = self._get_column_value(
            record,
            self.REQUIRED_PART_COLUMNS,
        )

        quantity = self._get_column_value(
            record,
            self.QUANTITY_COLUMNS,
        )

        description = self._get_column_value(
            record,
            self.DESCRIPTION_COLUMNS,
        )

        revision = self._get_column_value(
            record,
            self.REVISION_COLUMNS,
        )

        standardized[
            "Part Number"
        ] = self._clean_value(
            part_number
        )

        standardized[
            "QTY"
        ] = self._parse_quantity(
            quantity
        )

        standardized[
            "Description"
        ] = self._clean_value(
            description
        )

        standardized[
            "REV"
        ] = self._clean_value(
            revision
        )

        return standardized

    # =========================================================
    # COLUMN LOOKUP
    # =========================================================

    def _get_column_value(
        self,
        record: dict[str, Any],
        possible_names: set[str],
    ):

        for column_name, value in record.items():

            if column_name == "_excel_row":
                continue

            normalized = (
                self._normalize_header(
                    column_name
                )
            )

            if normalized in possible_names:

                return value

        return None

    # =========================================================
    # MEANINGFUL BOM ROW
    # =========================================================

    @staticmethod
    def _is_meaningful_bom_row(
        record: dict[str, Any],
    ) -> bool:

        part_number = record.get(
            "Part Number"
        )

        description = record.get(
            "Description"
        )

        quantity = record.get(
            "QTY"
        )

        # Normal BOM row.
        if (
            part_number is not None
            and str(part_number).strip()
        ):
            return True

        # Keep blank Part Number rows if there is
        # actual BOM information. These become REVIEW.
        if (
            description is not None
            and str(description).strip()
        ):
            return True

        if quantity is not None:
            return True

        return False

    # =========================================================
    # HEADER NORMALIZATION
    # =========================================================

    @staticmethod
    def _normalize_header(
        value: Any,
    ) -> str:

        if value is None:
            return ""

        value = str(
            value
        ).strip().lower()

        value = (
            value
            .replace("\n", " ")
            .replace("\r", " ")
        )

        value = " ".join(
            value.split()
        )

        return value

    # =========================================================
    # VALUE CLEANING
    # =========================================================

    @staticmethod
    def _clean_value(
        value: Any,
    ):

        if value is None:
            return None

        if isinstance(
            value,
            float,
        ):

            if pd.isna(value):
                return None

            if value.is_integer():
                return int(value)

        if pd.isna(value):
            return None

        text = str(
            value
        ).strip()

        if not text:
            return None

        return text

    # =========================================================
    # QUANTITY PARSING
    # =========================================================

    @staticmethod
    def _parse_quantity(
        value: Any,
    ):

        if value is None:
            return None

        if isinstance(
            value,
            bool,
        ):
            return None

        if isinstance(
            value,
            (int, float),
        ):

            try:

                if pd.isna(value):
                    return None

            except Exception:
                pass

            number = float(
                value
            )

            if number.is_integer():
                return int(number)

            return number

        text = str(
            value
        ).strip()

        if not text:
            return None

        # Remove commas from values such as 1,000.
        text = text.replace(
            ",",
            "",
        )

        try:

            number = float(
                text
            )

        except ValueError:

            return None

        if number.is_integer():
            return int(number)

        return number