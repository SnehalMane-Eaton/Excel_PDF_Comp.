# Excel PDF BOM Verifier

Local Windows/Tkinter utility for comparing Excel BOM data against PDF BOM data and exporting an `.xlsx` verification report.

## Important behavior

- **All BOM processing stays local.** No APIs, cloud processing, telemetry, uploads, or external AI services are used by the application flow.
- Use only **company-approved non-synced folders** for real BOM inputs and saved reports.
- The application does **not** attempt to detect every sync provider automatically.
- PDF support remains limited to the current extractor logic for:
  - structured/current-vs-removed BOM PDFs
  - supported engineering drawing PDFs
- Scanned/image-only PDFs are **not** newly supported by this fix.
- Removed PDF rows stay in `removed_rows` and are **not** treated as current BOM rows.

## Canonical source layout

The runtime code now lives under the intended package structure:

- `app/core/batch_processor.py`
- `app/core/config_manager.py`
- `app/core/models.py`
- `app/extractors/excel_extractor.py`
- `app/extractors/pdf_extractor.py`
- `app/matching/exact_matcher.py`
- `app/matching/fuzzy_matcher.py`
- `app/matching/normalizer.py`
- `app/matching/verification_engine.py`
- `app/pairing/file_pairer.py`
- `app/reporting/report_formatter.py`
- `app/reporting/report_generator.py`
- `app/ui/file_selection.py`
- `app/ui/main_window.py`
- `app/ui/pairing_view.py`
- `app/ui/progress_view.py`
- `tests/test_batch_processor.py`
- `tests/test_main_window_logic.py`
- `tests/test_report_generator.py`
- `tests/test_verification_engine.py`

The flat root modules are retained only as lightweight compatibility wrappers that import from `app/...`.

## Report output

Each paired comparison sheet contains exactly these visible columns:

1. `BOM1-PART`
2. `BOM1-QTY`
3. `BOM1-DESCRIPTION`
4. `BOM1-RESULT`
5. `QTY RESULT`
6. `BOM2-PART`
7. `BOM2-QTY`
8. `BOM2-DESCRIPTION`
9. `BOM2-RESULT`

The workbook also includes:

- `SUMMARY` sheet for file-level status/warnings
- `VERIFICATION_AUDIT` sheet for authoritative engine results/remarks
- one comparison sheet per processed pair
- hidden per-pair helper sheets used by comparison formulas

## Default local folders

By default the UI creates local folders under:

- `EXCEL_PDF_VERIFIER_HOME/input/excel` and `.../pdf` when `EXCEL_PDF_VERIFIER_HOME` is set
- otherwise `<current working directory>/input/excel` and `/input/pdf` when an `input` folder already exists
- otherwise `~/ExcelPDFVerifier/input/excel` and `/pdf`

Change folders in the UI as needed before processing real BOMs.

## Run from source

```bash
python main.py
```

## Run tests

```bash
pytest -q tests
```

## Build Windows executable

```bash
python -m PyInstaller --clean --noconfirm ExcelPDFVerifier.spec
```

`PyInstaller` is listed in `requirements.txt`, but the Windows `.exe` build was **not** executed in this Linux sandbox.

## Replacement / download guidance

If you are manually updating another checkout, replace the canonical files listed above in the same repository-relative paths. The `app/...` files are the authoritative implementations after this fix.
