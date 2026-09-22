from dataclasses import dataclass
from pathlib import Path
import re


@dataclass
class FilePair:
    """Represents one Excel/PDF file pairing."""

    excel_file: Path | None
    pdf_file: Path | None
    key: str
    status: str
    message: str
    match_method: str = ""
    match_score: float | None = None


class FilePairer:
    """
    Generic Excel/PDF BOM file pairing.

    The user may select ANY local folders.

    Pairing strategy:

    1. Discover Excel/PDF files recursively.
    2. If exactly one Excel and one PDF exist, pair them directly.
    3. Try exact normalized filename matching.
    4. For remaining files, compare exact BOM part-number content.
    5. If only one Excel/PDF combination remains, pair it.
    6. For multiple files, require an unambiguous content relationship.
    7. Never use fuzzy engineering part-number matching.
    8. Never silently guess between ambiguous files.
    """

    EXCEL_EXTENSIONS = {
        ".xlsx",
        ".xls",
        ".xlsm",
    }

    PDF_EXTENSIONS = {
        ".pdf",
    }

    # For content-based matching, one exact common BOM
    # part number is enough to create a candidate.
    MIN_COMMON_PARTS = 1

    # Minimum normalized content score.
    MIN_SCORE = 0.01

    # Difference between first and second candidate.
    # This prevents ambiguous automatic pairing.
    MIN_SCORE_MARGIN = 0.02

    def pair_folders(
        self,
        excel_folder: str | Path,
        pdf_folder: str | Path,
    ) -> list[FilePair]:

        excel_folder = Path(
            excel_folder
        )

        pdf_folder = Path(
            pdf_folder
        )

        self._validate_folder(
            excel_folder,
            "Excel",
        )

        self._validate_folder(
            pdf_folder,
            "PDF",
        )

        excel_files = self._discover_files(
            excel_folder,
            self.EXCEL_EXTENSIONS,
        )

        pdf_files = self._discover_files(
            pdf_folder,
            self.PDF_EXTENSIONS,
        )

        if not excel_files and not pdf_files:
            return []

        # -----------------------------------------------------
        # IMPORTANT:
        # If exactly one Excel and one PDF are selected,
        # they are the intended pair regardless of filename.
        # -----------------------------------------------------

        if (
            len(excel_files) == 1
            and len(pdf_files) == 1
        ):

            excel_file = excel_files[0]
            pdf_file = pdf_files[0]

            return [
                FilePair(
                    excel_file=excel_file,
                    pdf_file=pdf_file,
                    key=self.normalize_filename(
                        excel_file.stem
                    ),
                    status="PAIRED",
                    message=(
                        "Excel and PDF paired because "
                        "one Excel and one PDF file were "
                        "found in the selected folders."
                    ),
                    match_method="single-file-pair",
                    match_score=1.0,
                )
            ]

        pairs: list[FilePair] = []

        used_excel: set[Path] = set()
        used_pdf: set[Path] = set()

        # -----------------------------------------------------
        # STEP 1: Exact normalized filename matching
        # -----------------------------------------------------

        filename_pairs = self._pair_by_filename(
            excel_files,
            pdf_files,
        )

        for pair in filename_pairs:

            pairs.append(pair)

            if pair.status == "PAIRED":

                if pair.excel_file is not None:
                    used_excel.add(
                        pair.excel_file
                    )

                if pair.pdf_file is not None:
                    used_pdf.add(
                        pair.pdf_file
                    )

        remaining_excel = [
            path
            for path in excel_files
            if path not in used_excel
        ]

        remaining_pdf = [
            path
            for path in pdf_files
            if path not in used_pdf
        ]

        # -----------------------------------------------------
        # STEP 2: Content-based pairing
        # -----------------------------------------------------

        content_pairs = self._pair_by_content(
            remaining_excel,
            remaining_pdf,
        )

        for pair in content_pairs:

            pairs.append(pair)

            if pair.status == "PAIRED":

                if pair.excel_file is not None:
                    used_excel.add(
                        pair.excel_file
                    )

                if pair.pdf_file is not None:
                    used_pdf.add(
                        pair.pdf_file
                    )

        # -----------------------------------------------------
        # STEP 3: Remaining files
        # -----------------------------------------------------

        remaining_excel = [
            path
            for path in excel_files
            if path not in used_excel
        ]

        remaining_pdf = [
            path
            for path in pdf_files
            if path not in used_pdf
        ]

        for excel_file in remaining_excel:

            pairs.append(
                FilePair(
                    excel_file=excel_file,
                    pdf_file=None,
                    key=self.normalize_filename(
                        excel_file.stem
                    ),
                    status="EXCEL_ONLY_FILE",
                    message=(
                        "No matching PDF could be established "
                        "using filename or BOM part-number content."
                    ),
                    match_method="none",
                )
            )

        for pdf_file in remaining_pdf:

            pairs.append(
                FilePair(
                    excel_file=None,
                    pdf_file=pdf_file,
                    key=self.normalize_filename(
                        pdf_file.stem
                    ),
                    status="PDF_ONLY_FILE",
                    message=(
                        "No matching Excel could be established "
                        "using filename or BOM part-number content."
                    ),
                    match_method="none",
                )
            )

        return self._sort_pairs(
            pairs
        )

    # =========================================================
    # FILENAME PAIRING
    # =========================================================

    def _pair_by_filename(
        self,
        excel_files: list[Path],
        pdf_files: list[Path],
    ) -> list[FilePair]:

        excel_groups = self._group_by_key(
            excel_files
        )

        pdf_groups = self._group_by_key(
            pdf_files
        )

        pairs: list[FilePair] = []

        all_keys = sorted(
            set(excel_groups.keys())
            |
            set(pdf_groups.keys())
        )

        for key in all_keys:

            excel_matches = excel_groups.get(
                key,
                [],
            )

            pdf_matches = pdf_groups.get(
                key,
                [],
            )

            if (
                len(excel_matches) == 1
                and len(pdf_matches) == 1
            ):

                excel_file = excel_matches[0]
                pdf_file = pdf_matches[0]

                pairs.append(
                    FilePair(
                        excel_file=excel_file,
                        pdf_file=pdf_file,
                        key=key,
                        status="PAIRED",
                        message=(
                            "Excel and PDF paired by "
                            "normalized filename."
                        ),
                        match_method="filename",
                        match_score=1.0,
                    )
                )

            elif (
                len(excel_matches) > 1
                or len(pdf_matches) > 1
            ):

                # Do not automatically choose among duplicate
                # filenames.
                pairs.append(
                    FilePair(
                        excel_file=None,
                        pdf_file=None,
                        key=key,
                        status="DUPLICATE_FILES",
                        message=(
                            "Multiple Excel or PDF files have "
                            "the same normalized filename. "
                            "Content matching will be attempted "
                            "after this stage."
                        ),
                        match_method="filename",
                    )
                )

        return pairs

    # =========================================================
    # CONTENT PAIRING
    # =========================================================

    def _pair_by_content(
        self,
        excel_files: list[Path],
        pdf_files: list[Path],
    ) -> list[FilePair]:

        if not excel_files or not pdf_files:
            return []

        excel_signatures = {
            path: self._extract_excel_signature(path)
            for path in excel_files
        }

        pdf_signatures = {
            path: self._extract_pdf_signature(path)
            for path in pdf_files
        }

        # -----------------------------------------------------
        # Build every valid Excel/PDF candidate.
        # -----------------------------------------------------

        excel_candidates: dict[
            Path,
            list[tuple[float, Path, int]],
        ] = {
            path: []
            for path in excel_files
        }

        pdf_candidates: dict[
            Path,
            list[tuple[float, Path, int]],
        ] = {
            path: []
            for path in pdf_files
        }

        for excel_file in excel_files:

            excel_parts = excel_signatures.get(
                excel_file,
                set(),
            )

            for pdf_file in pdf_files:

                pdf_parts = pdf_signatures.get(
                    pdf_file,
                    set(),
                )

                score, common_count = (
                    self._content_score(
                        excel_parts,
                        pdf_parts,
                    )
                )

                if (
                    common_count
                    < self.MIN_COMMON_PARTS
                ):
                    continue

                if score < self.MIN_SCORE:
                    continue

                excel_candidates[
                    excel_file
                ].append(
                    (
                        score,
                        pdf_file,
                        common_count,
                    )
                )

                pdf_candidates[
                    pdf_file
                ].append(
                    (
                        score,
                        excel_file,
                        common_count,
                    )
                )

        # -----------------------------------------------------
        # Sort candidates.
        # -----------------------------------------------------

        for candidates in excel_candidates.values():

            candidates.sort(
                key=lambda item: (
                    item[0],
                    item[2],
                ),
                reverse=True,
            )

        for candidates in pdf_candidates.values():

            candidates.sort(
                key=lambda item: (
                    item[0],
                    item[2],
                ),
                reverse=True,
            )

        pairs: list[FilePair] = []

        used_excel: set[Path] = set()
        used_pdf: set[Path] = set()

        # -----------------------------------------------------
        # Iteratively accept only unambiguous mutual matches.
        # -----------------------------------------------------

        changed = True

        while changed:

            changed = False

            for excel_file in excel_files:

                if excel_file in used_excel:
                    continue

                candidates = [
                    candidate
                    for candidate
                    in excel_candidates.get(
                        excel_file,
                        [],
                    )
                    if candidate[1]
                    not in used_pdf
                ]

                if not candidates:
                    continue

                best_score, pdf_file, common_count = (
                    candidates[0]
                )

                # Check whether there is another PDF with a
                # similar score.
                if len(candidates) > 1:

                    second_score = candidates[1][0]

                    if (
                        best_score - second_score
                        < self.MIN_SCORE_MARGIN
                    ):
                        continue

                reverse_candidates = [
                    candidate
                    for candidate
                    in pdf_candidates.get(
                        pdf_file,
                        []
                    )
                    if candidate[1]
                    not in used_excel
                ]

                if not reverse_candidates:
                    continue

                reverse_best_score = (
                    reverse_candidates[0][0]
                )

                reverse_best_excel = (
                    reverse_candidates[0][1]
                )

                # Require the relationship to be mutual.
                if reverse_best_excel != excel_file:
                    continue

                if (
                    abs(
                        best_score
                        - reverse_best_score
                    )
                    > 0.000001
                ):
                    continue

                pairs.append(
                    FilePair(
                        excel_file=excel_file,
                        pdf_file=pdf_file,
                        key=self.normalize_filename(
                            excel_file.stem
                        ),
                        status="PAIRED",
                        message=(
                            "Excel and PDF paired using "
                            "exact BOM part-number content."
                        ),
                        match_method="content",
                        match_score=best_score,
                    )
                )

                used_excel.add(
                    excel_file
                )

                used_pdf.add(
                    pdf_file
                )

                changed = True

        # -----------------------------------------------------
        # If exactly one Excel and one PDF remain and their
        # content has at least one common part, pair them.
        # -----------------------------------------------------

        remaining_excel = [
            path
            for path in excel_files
            if path not in used_excel
        ]

        remaining_pdf = [
            path
            for path in pdf_files
            if path not in used_pdf
        ]

        if (
            len(remaining_excel) == 1
            and len(remaining_pdf) == 1
        ):

            excel_file = remaining_excel[0]
            pdf_file = remaining_pdf[0]

            excel_parts = excel_signatures.get(
                excel_file,
                set(),
            )

            pdf_parts = pdf_signatures.get(
                pdf_file,
                set(),
            )

            score, common_count = (
                self._content_score(
                    excel_parts,
                    pdf_parts,
                )
            )

            if common_count >= 1:

                pairs.append(
                    FilePair(
                        excel_file=excel_file,
                        pdf_file=pdf_file,
                        key=self.normalize_filename(
                            excel_file.stem
                        ),
                        status="PAIRED",
                        message=(
                            "Excel and PDF paired because "
                            "one unmatched Excel and one unmatched "
                            "PDF remained and their BOM content "
                            "contains common part numbers."
                        ),
                        match_method="content-single-remaining",
                        match_score=score,
                    )
                )

        return pairs

    # =========================================================
    # EXCEL SIGNATURE
    # =========================================================

    @staticmethod
    def _extract_excel_signature(
        file_path: Path,
    ) -> set[str]:

        try:

            from app.extractors.excel_extractor import (
                ExcelExtractor,
            )

            from app.matching.normalizer import (
                normalize_part_number,
            )

            data = ExcelExtractor(
                file_path
            ).extract()

            parts: set[str] = set()

            for sheet_data in data[
                "sheets"
            ].values():

                for record in sheet_data[
                    "records"
                ]:

                    part_number = (
                        normalize_part_number(
                            record.get(
                                "Part Number"
                            )
                        )
                    )

                    if part_number:
                        parts.add(
                            part_number
                        )

            return parts

        except Exception:

            return set()

    # =========================================================
    # PDF SIGNATURE
    # =========================================================

    @staticmethod
    def _extract_pdf_signature(
        file_path: Path,
    ) -> set[str]:

        try:

            from app.extractors.pdf_extractor import (
                PDFExtractor,
            )

            from app.matching.normalizer import (
                normalize_part_number,
            )

            data = PDFExtractor(
                file_path
            ).extract()

            parts: set[str] = set()

            for record in data[
                "bom_rows"
            ]:

                part_number = (
                    normalize_part_number(
                        record.get(
                            "part_number"
                        )
                    )
                )

                if part_number:
                    parts.add(
                        part_number
                    )

            return parts

        except Exception:

            return set()

    # =========================================================
    # CONTENT SCORE
    # =========================================================

    @staticmethod
    def _content_score(
        excel_parts: set[str],
        pdf_parts: set[str],
    ) -> tuple[float, int]:

        if (
            not excel_parts
            or not pdf_parts
        ):
            return 0.0, 0

        common_parts = (
            excel_parts
            &
            pdf_parts
        )

        if not common_parts:
            return 0.0, 0

        # Use the smaller BOM as denominator.
        #
        # This is useful because a PDF may contain many more
        # detailed parts than the Excel BOM.
        denominator = min(
            len(excel_parts),
            len(pdf_parts),
        )

        if denominator == 0:
            return 0.0, 0

        score = (
            len(common_parts)
            /
            denominator
        )

        return score, len(common_parts)

    # =========================================================
    # FOLDER VALIDATION
    # =========================================================

    @staticmethod
    def _validate_folder(
        folder: Path,
        folder_type: str,
    ):

        if not folder.exists():

            raise FileNotFoundError(
                f"{folder_type} folder does not exist: "
                f"{folder}"
            )

        if not folder.is_dir():

            raise NotADirectoryError(
                f"{folder_type} path is not a folder: "
                f"{folder}"
            )

    # =========================================================
    # FILE DISCOVERY
    # =========================================================

    @staticmethod
    def _discover_files(
        folder: Path,
        extensions: set[str],
    ) -> list[Path]:

        files: list[Path] = []

        # rglob makes the application work even when the
        # selected folder contains Excel/PDF files inside
        # subfolders.

        for path in folder.rglob("*"):

            if not path.is_file():
                continue

            if path.suffix.lower() not in extensions:
                continue

            files.append(
                path
            )

        return sorted(
            files,
            key=lambda path: str(
                path
            ).lower(),
        )

    # =========================================================
    # GROUP BY NORMALIZED FILENAME
    # =========================================================

    @classmethod
    def _group_by_key(
        cls,
        files: list[Path],
    ) -> dict[str, list[Path]]:

        groups: dict[
            str,
            list[Path],
        ] = {}

        for file_path in files:

            key = cls.normalize_filename(
                file_path.stem
            )

            groups.setdefault(
                key,
                [],
            ).append(
                file_path
            )

        return groups

    # =========================================================
    # NORMALIZE FILENAME
    # =========================================================

    @staticmethod
    def normalize_filename(
        value: str,
    ) -> str:

        value = str(
            value
        ).strip().lower()

        # Remove common copy numbering:
        #
        # file (1)
        # file (2)
        # file (10)

        value = re.sub(
            r"\s*\(\d+\)\s*$",
            "",
            value,
        )

        # Remove all non-alphanumeric characters.

        value = re.sub(
            r"[^a-z0-9]+",
            "",
            value,
        )

        return value

    # =========================================================
    # SORT
    # =========================================================

    @staticmethod
    def _sort_pairs(
        pairs: list[FilePair],
    ) -> list[FilePair]:

        return sorted(
            pairs,
            key=lambda pair: (
                pair.key,
                pair.status,
                (
                    pair.excel_file.name.lower()
                    if pair.excel_file
                    else ""
                ),
                (
                    pair.pdf_file.name.lower()
                    if pair.pdf_file
                    else ""
                ),
            ),
        )


def pair_files(
    excel_folder: str | Path,
    pdf_folder: str | Path,
) -> list[FilePair]:
    """Backward-compatible helper for older direct callers/tests."""
    return FilePairer().pair_folders(excel_folder, pdf_folder)