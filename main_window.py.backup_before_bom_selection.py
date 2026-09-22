from __future__ import annotations

import os
import shutil
import tempfile
import threading
import tkinter as tk

from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from typing import Any

from app.core.batch_processor import BatchProcessor
from app.reporting.report_generator import ReportGenerator


class DesktopFolderManager:
    """
    Manages the standard employee folder structure.

    Default structure:

        Desktop/
            files/
                excel/
                pdf/

    The application also supports manually selecting any local folders.
    """

    STANDARD_ROOT = "files"
    EXCEL_FOLDER = "excel"
    PDF_FOLDER = "pdf"

    def __init__(self) -> None:
        self.desktop_path = self._find_desktop()

        self.files_path = (
            self.desktop_path / self.STANDARD_ROOT
        )

        self.excel_path = (
            self.files_path / self.EXCEL_FOLDER
        )

        self.pdf_path = (
            self.files_path / self.PDF_FOLDER
        )

    def prepare_default_folders(self) -> tuple[Path, Path]:
        """
        Create the standard folders if they do not exist.
        """
        self.excel_path.mkdir(
            parents=True,
            exist_ok=True,
        )

        self.pdf_path.mkdir(
            parents=True,
            exist_ok=True,
        )

        return (
            self.excel_path,
            self.pdf_path,
        )

    @staticmethod
    def _find_desktop() -> Path:
        """
        Find the user's Windows Desktop.

        Priority:
            1. OneDrive Desktop
            2. USERPROFILE Desktop
            3. Home/Desktop
        """

        candidates: list[Path] = []

        one_drive = os.environ.get("OneDrive")

        if one_drive:
            candidates.append(
                Path(one_drive) / "Desktop"
            )

        user_profile = os.environ.get("USERPROFILE")

        if user_profile:
            candidates.append(
                Path(user_profile) / "Desktop"
            )

        candidates.append(
            Path.home() / "Desktop"
        )

        for path in candidates:
            if path.exists() and path.is_dir():
                return path

        return Path.home() / "Desktop"


class FolderSelectionService:
    """
    Handles folder selection and validation.
    """

    def __init__(self, parent: tk.Misc) -> None:
        self.parent = parent

    def select_folder(
        self,
        title: str,
        initial_directory: str,
    ) -> str:
        """
        Open the Windows folder-selection dialog.
        """

        selected = filedialog.askdirectory(
            parent=self.parent,
            title=title,
            initialdir=initial_directory,
            mustexist=True,
        )

        return selected.strip() if selected else ""

    @staticmethod
    def validate_folder(
        path: str,
        label: str,
    ) -> None:
        """
        Validate that the selected path exists and is a directory.
        """

        if not path:
            raise ValueError(
                f"Please select the {label}."
            )

        folder = Path(path)

        if not folder.exists():
            raise ValueError(
                f"{label} does not exist:\n\n{folder}"
            )

        if not folder.is_dir():
            raise ValueError(
                f"The selected {label} is not a folder:\n\n{folder}"
            )


class ReportFileManager:
    """
    Manages temporary report files and saving the final report.
    """

    def __init__(self) -> None:
        self.temp_directory: Path | None = None
        self.temp_report_path: Path | None = None

    def create_temp_report_path(self) -> Path:
        """
        Create a temporary location for the generated report.
        """

        self.clear_temp_report()

        self.temp_directory = Path(
            tempfile.mkdtemp(
                prefix="excel_pdf_bom_verifier_"
            )
        )

        self.temp_report_path = (
            self.temp_directory
            / "BOM_Verification_Report.xlsx"
        )

        return self.temp_report_path

    def save_report(
        self,
        parent: tk.Misc,
    ) -> bool:
        """
        Ask the user where to save the report.
        """

        if self.temp_report_path is None:
            raise ValueError(
                "No verification report is available."
            )

        if not self.temp_report_path.exists():
            raise FileNotFoundError(
                "The generated report could not be found."
            )

        destination = filedialog.asksaveasfilename(
            parent=parent,
            title="Save BOM Verification Report",
            initialfile="BOM_Verification_Report.xlsx",
            defaultextension=".xlsx",
            filetypes=[
                (
                    "Excel Workbook",
                    "*.xlsx",
                ),
                (
                    "All Files",
                    "*.*",
                ),
            ],
        )

        if not destination:
            return False

        destination_path = Path(destination)

        shutil.copy2(
            self.temp_report_path,
            destination_path,
        )

        return True

    def clear_temp_report(self) -> None:
        """
        Delete the temporary report directory.
        """

        if self.temp_directory is None:
            return

        try:
            shutil.rmtree(
                self.temp_directory,
                ignore_errors=True,
            )
        finally:
            self.temp_directory = None
            self.temp_report_path = None


class VerificationController:
    """
    Coordinates BOM verification and report generation.

    Business logic remains inside:
        BatchProcessor
        ReportGenerator
    """

    def __init__(self) -> None:
        self.batch_processor = BatchProcessor()
        self.report_generator = ReportGenerator()

    def verify(
        self,
        excel_folder: str,
        pdf_folder: str,
        report_path: Path,
    ) -> dict[str, Any]:
        """
        Process the selected folders and generate the report.
        """

        batch_result = (
            self.batch_processor.process_folders(
                excel_folder,
                pdf_folder,
            )
        )

        self.report_generator.generate(
            batch_result,
            str(report_path),
        )

        return batch_result


class MainWindow:
    """
    Main application window.

    Responsibilities:
        - Folder selection
        - Verification
        - Progress display
        - Result display
        - Report saving
        - Resetting the interface
    """

    WINDOW_TITLE = "Excel / PDF BOM Verifier"

    WINDOW_WIDTH = 1150
    WINDOW_HEIGHT = 900

    def __init__(
        self,
        root: tk.Tk,
    ) -> None:

        self.root = root

        self.folder_manager = (
            DesktopFolderManager()
        )

        self.folder_service = (
            FolderSelectionService(root)
        )

        self.verification_controller = (
            VerificationController()
        )

        self.report_manager = (
            ReportFileManager()
        )

        self.excel_folder_var = tk.StringVar()
        self.pdf_folder_var = tk.StringVar()

        self.progress_var = tk.DoubleVar(
            value=0
        )

        self.status_var = tk.StringVar(
            value=(
                "Ready. Select the BOM folders "
                "and click VERIFY BOM."
            )
        )

        self.total_pairs_var = tk.StringVar(
            value="0"
        )

        self.verified_pairs_var = tk.StringVar(
            value="0"
        )

        self.matched_var = tk.StringVar(
            value="0"
        )

        self.qty_mismatch_var = tk.StringVar(
            value="0"
        )

        self.pdf_only_var = tk.StringVar(
            value="0"
        )

        self.excel_only_var = tk.StringVar(
            value="0"
        )

        self.review_var = tk.StringVar(
            value="0"
        )

        self.errors_var = tk.StringVar(
            value="0"
        )

        self.verify_button: ttk.Button | None = None
        self.save_button: ttk.Button | None = None
        self.clear_button: ttk.Button | None = None
        self.progress_bar: ttk.Progressbar | None = None

        self.verification_running = False

        self._configure_window()
        self._build_interface()
        self._load_default_folders()

        self.root.protocol(
            "WM_DELETE_WINDOW",
            self._on_close,
        )

    # ================================================================
    # WINDOW
    # ================================================================

    def _configure_window(self) -> None:
        self.root.title(
            self.WINDOW_TITLE
        )

        self.root.geometry(
            f"{self.WINDOW_WIDTH}x{self.WINDOW_HEIGHT}"
        )

        self.root.minsize(
            1050,
            760,
        )

        self.root.columnconfigure(
            0,
            weight=1,
        )

        self.root.rowconfigure(
            0,
            weight=1,
        )

    # ================================================================
    # INTERFACE
    # ================================================================

    def _build_interface(self) -> None:
        main_frame = ttk.Frame(
            self.root,
            padding=20,
        )

        main_frame.grid(
            row=0,
            column=0,
            sticky="nsew",
        )

        main_frame.columnconfigure(
            0,
            weight=1,
        )

        self._build_header(
            main_frame
        )

        self._build_folder_section(
            main_frame
        )

        self._build_progress_section(
            main_frame
        )

        self._build_results_section(
            main_frame
        )

        self._build_action_section(
            main_frame
        )

    def _build_header(
        self,
        parent: ttk.Frame,
    ) -> None:

        header = ttk.Frame(parent)

        header.grid(
            row=0,
            column=0,
            sticky="ew",
            pady=(0, 18),
        )

        header.columnconfigure(
            0,
            weight=1,
        )

        title = tk.Label(
            header,
            text="Excel / PDF BOM Verifier",
            font=(
                "Segoe UI",
                26,
                "bold",
            ),
        )

        title.grid(
            row=0,
            column=0,
            pady=(0, 6),
        )

        subtitle = tk.Label(
            header,
            text=(
                "Compare Excel BOM data against "
                "PDF BOM data quickly and accurately"
            ),
            font=(
                "Segoe UI",
                13,
            ),
        )

        subtitle.grid(
            row=1,
            column=0,
        )

    def _build_folder_section(
        self,
        parent: ttk.Frame,
    ) -> None:

        section = ttk.LabelFrame(
            parent,
            text="1. Select Input Folders",
            padding=15,
        )

        section.grid(
            row=1,
            column=0,
            sticky="ew",
            pady=(0, 16),
        )

        section.columnconfigure(
            1,
            weight=1,
        )

        # ------------------------------------------------------------
        # EXCEL
        # ------------------------------------------------------------

        ttk.Label(
            section,
            text="Excel BOM Folder:",
            font=(
                "Segoe UI",
                11,
                "bold",
            ),
        ).grid(
            row=0,
            column=0,
            sticky="w",
            padx=(0, 12),
            pady=8,
        )

        ttk.Entry(
            section,
            textvariable=self.excel_folder_var,
            font=(
                "Segoe UI",
                10,
            ),
        ).grid(
            row=0,
            column=1,
            sticky="ew",
            pady=8,
        )

        ttk.Button(
            section,
            text="Select Folder",
            command=self._select_excel_folder,
            width=15,
        ).grid(
            row=0,
            column=2,
            padx=(12, 0),
            pady=8,
        )

        # ------------------------------------------------------------
        # PDF
        # ------------------------------------------------------------

        ttk.Label(
            section,
            text="PDF BOM Folder:",
            font=(
                "Segoe UI",
                11,
                "bold",
            ),
        ).grid(
            row=1,
            column=0,
            sticky="w",
            padx=(0, 12),
            pady=8,
        )

        ttk.Entry(
            section,
            textvariable=self.pdf_folder_var,
            font=(
                "Segoe UI",
                10,
            ),
        ).grid(
            row=1,
            column=1,
            sticky="ew",
            pady=8,
        )

        ttk.Button(
            section,
            text="Select Folder",
            command=self._select_pdf_folder,
            width=15,
        ).grid(
            row=1,
            column=2,
            padx=(12, 0),
            pady=8,
        )

        # ------------------------------------------------------------
        # INFORMATION
        # ------------------------------------------------------------

        ttk.Label(
            section,
            text=(
                "Default: Desktop\\files\\excel and "
                "Desktop\\files\\pdf. "
                "You can also select any local folders."
            ),
            font=(
                "Segoe UI",
                9,
            ),
        ).grid(
            row=2,
            column=0,
            columnspan=3,
            sticky="w",
            pady=(8, 0),
        )

    def _build_progress_section(
        self,
        parent: ttk.Frame,
    ) -> None:

        section = ttk.LabelFrame(
            parent,
            text="2. Verification Progress",
            padding=15,
        )

        section.grid(
            row=2,
            column=0,
            sticky="ew",
            pady=(0, 16),
        )

        section.columnconfigure(
            0,
            weight=1,
        )

        self.progress_bar = ttk.Progressbar(
            section,
            variable=self.progress_var,
            maximum=100,
            mode="determinate",
        )

        self.progress_bar.grid(
            row=0,
            column=0,
            sticky="ew",
            ipady=4,
        )

        ttk.Label(
            section,
            textvariable=self.status_var,
            font=(
                "Segoe UI",
                10,
            ),
        ).grid(
            row=1,
            column=0,
            sticky="w",
            pady=(10, 0),
        )

    def _build_results_section(
        self,
        parent: ttk.Frame,
    ) -> None:

        section = ttk.LabelFrame(
            parent,
            text="3. Verification Results",
            padding=20,
        )

        section.grid(
            row=3,
            column=0,
            sticky="ew",
            pady=(0, 16),
        )

        section.columnconfigure(
            1,
            weight=1,
        )

        section.columnconfigure(
            3,
            weight=1,
        )

        self._create_result_row(
            section,
            0,
            "Total File Pairs",
            self.total_pairs_var,
            "Verified Pairs",
            self.verified_pairs_var,
        )

        self._create_result_row(
            section,
            1,
            "Matched Parts",
            self.matched_var,
            "Quantity Mismatch",
            self.qty_mismatch_var,
        )

        self._create_result_row(
            section,
            2,
            "PDF Only",
            self.pdf_only_var,
            "Excel Only",
            self.excel_only_var,
        )

        self._create_result_row(
            section,
            3,
            "Review",
            self.review_var,
            "Processing Errors",
            self.errors_var,
        )

    def _create_result_row(
        self,
        parent: ttk.LabelFrame,
        row: int,
        left_label: str,
        left_variable: tk.StringVar,
        right_label: str,
        right_variable: tk.StringVar,
    ) -> None:

        ttk.Label(
            parent,
            text=left_label,
            font=(
                "Segoe UI",
                11,
            ),
        ).grid(
            row=row,
            column=0,
            sticky="w",
            padx=(0, 20),
            pady=10,
        )

        ttk.Label(
            parent,
            textvariable=left_variable,
            font=(
                "Segoe UI",
                12,
                "bold",
            ),
        ).grid(
            row=row,
            column=1,
            sticky="e",
            padx=(0, 50),
            pady=10,
        )

        ttk.Label(
            parent,
            text=right_label,
            font=(
                "Segoe UI",
                11,
            ),
        ).grid(
            row=row,
            column=2,
            sticky="w",
            padx=(20, 20),
            pady=10,
        )

        ttk.Label(
            parent,
            textvariable=right_variable,
            font=(
                "Segoe UI",
                12,
                "bold",
            ),
        ).grid(
            row=row,
            column=3,
            sticky="e",
            pady=10,
        )

    def _build_action_section(
        self,
        parent: ttk.Frame,
    ) -> None:

        frame = ttk.Frame(parent)

        frame.grid(
            row=4,
            column=0,
            sticky="ew",
        )

        frame.columnconfigure(
            1,
            weight=1,
        )

        self.clear_button = ttk.Button(
            frame,
            text="Clear",
            command=self._clear,
            width=15,
        )

        self.clear_button.grid(
            row=0,
            column=0,
            sticky="w",
        )

        self.verify_button = ttk.Button(
            frame,
            text="VERIFY BOM",
            command=self._start_verification,
            width=20,
        )

        self.verify_button.grid(
            row=0,
            column=1,
        )

        self.save_button = ttk.Button(
            frame,
            text="SAVE REPORT",
            command=self._save_report,
            width=20,
            state="disabled",
        )

        self.save_button.grid(
            row=0,
            column=2,
            sticky="e",
        )

    # ================================================================
    # DEFAULT FOLDERS
    # ================================================================

    def _load_default_folders(self) -> None:
        """
        Automatically prepare:

            Desktop/files/excel
            Desktop/files/pdf
        """

        try:
            excel_folder, pdf_folder = (
                self.folder_manager.prepare_default_folders()
            )

            self.excel_folder_var.set(
                str(excel_folder)
            )

            self.pdf_folder_var.set(
                str(pdf_folder)
            )

            self.status_var.set(
                "Ready. Default Desktop\\files folders are selected."
            )

        except Exception:
            self.excel_folder_var.set("")
            self.pdf_folder_var.set("")

            self.status_var.set(
                "Ready. Please select the Excel and PDF folders."
            )

    # ================================================================
    # FOLDER SELECTION
    # ================================================================

    def _select_excel_folder(self) -> None:

        current = self.excel_folder_var.get()

        if current and Path(current).exists():
            initial_directory = current
        else:
            initial_directory = str(
                self.folder_manager.desktop_path
            )

        selected = (
            self.folder_service.select_folder(
                "Select Excel BOM Folder",
                initial_directory,
            )
        )

        if selected:
            self.excel_folder_var.set(
                selected
            )

            self._reset_report_state()

    def _select_pdf_folder(self) -> None:

        current = self.pdf_folder_var.get()

        if current and Path(current).exists():
            initial_directory = current
        else:
            initial_directory = str(
                self.folder_manager.desktop_path
            )

        selected = (
            self.folder_service.select_folder(
                "Select PDF BOM Folder",
                initial_directory,
            )
        )

        if selected:
            self.pdf_folder_var.set(
                selected
            )

            self._reset_report_state()

    # ================================================================
    # VERIFICATION
    # ================================================================

    def _start_verification(self) -> None:

        if self.verification_running:
            return

        excel_folder = (
            self.excel_folder_var.get().strip()
        )

        pdf_folder = (
            self.pdf_folder_var.get().strip()
        )

        try:
            self.folder_service.validate_folder(
                excel_folder,
                "Excel BOM folder",
            )

            self.folder_service.validate_folder(
                pdf_folder,
                "PDF BOM folder",
            )

        except ValueError as error:
            messagebox.showerror(
                "Invalid Folder",
                str(error),
                parent=self.root,
            )

            return

        self._reset_results()

        self.verification_running = True

        self._set_controls(
            running=True
        )

        self._set_progress(
            5,
            "Preparing verification..."
        )

        worker = threading.Thread(
            target=self._verification_worker,
            args=(
                excel_folder,
                pdf_folder,
            ),
            daemon=True,
            name="BOMVerificationWorker",
        )

        worker.start()

    def _verification_worker(
        self,
        excel_folder: str,
        pdf_folder: str,
    ) -> None:
        """
        Runs verification outside the Tkinter event handler.
        """

        try:
            self._queue_progress(
                10,
                "Scanning Excel and PDF folders..."
            )

            report_path = (
                self.report_manager
                .create_temp_report_path()
            )

            self._queue_progress(
                25,
                "Pairing Excel and PDF BOM files..."
            )

            result = (
                self.verification_controller.verify(
                    excel_folder=excel_folder,
                    pdf_folder=pdf_folder,
                    report_path=report_path,
                )
            )

            self._queue_progress(
                75,
                "Verifying BOM quantities and generating report..."
            )

            self.root.after(
                0,
                self._verification_completed,
                result,
            )

        except Exception as error:

            self.root.after(
                0,
                self._verification_failed,
                error,
            )

    # ================================================================
    # VERIFICATION RESULT
    # ================================================================

    def _verification_completed(
        self,
        result: dict[str, Any],
    ) -> None:

        try:
            summary = result.get(
                "summary",
                {},
            )

            self.total_pairs_var.set(
                str(
                    self._summary_value(
                        summary,
                        "total",
                        "total_pairs",
                    )
                )
            )

            self.verified_pairs_var.set(
                str(
                    self._summary_value(
                        summary,
                        "verified_pairs",
                        "verified",
                    )
                )
            )

            self.matched_var.set(
                str(
                    self._summary_value(
                        summary,
                        "match",
                        "matched",
                    )
                )
            )

            self.qty_mismatch_var.set(
                str(
                    self._summary_value(
                        summary,
                        "qty_mismatch",
                    )
                )
            )

            self.pdf_only_var.set(
                str(
                    self._summary_value(
                        summary,
                        "pdf_only",
                    )
                )
            )

            self.excel_only_var.set(
                str(
                    self._summary_value(
                        summary,
                        "excel_only",
                    )
                )
            )

            self.review_var.set(
                str(
                    self._summary_value(
                        summary,
                        "review",
                    )
                )
            )

            self.errors_var.set(
                str(
                    self._summary_value(
                        summary,
                        "processing_errors",
                        "errors",
                    )
                )
            )

            self._set_progress(
                100,
                "Verification completed successfully. 100%"
            )

            self.verification_running = False

            self._set_controls(
                running=False
            )

            if self.report_manager.temp_report_path:
                self.save_button.config(
                    state="normal"
                )

            verified_pairs = (
                self._summary_value(
                    summary,
                    "verified_pairs",
                    "verified",
                )
            )

            if verified_pairs == 0:

                messagebox.showwarning(
                    "Verification Completed",
                    (
                        "Verification completed, but no file pairs "
                        "were successfully verified.\n\n"
                        "Please check that the selected folders contain "
                        "the intended BOM files."
                    ),
                    parent=self.root,
                )

        except Exception as error:
            self._verification_failed(
                error
            )

    @staticmethod
    def _summary_value(
        summary: dict[str, Any],
        *keys: str,
    ) -> Any:

        for key in keys:
            if key in summary:
                return summary[key]

        return 0

    def _verification_failed(
        self,
        error: Exception,
    ) -> None:

        self.verification_running = False

        self._set_controls(
            running=False
        )

        self._set_progress(
            0,
            "Verification failed."
        )

        messagebox.showerror(
            "Verification Error",
            (
                "The BOM verification could not be completed.\n\n"
                f"{error}"
            ),
            parent=self.root,
        )

    # ================================================================
    # SAVE REPORT
    # ================================================================

    def _save_report(self) -> None:

        if self.verification_running:
            return

        if (
            self.report_manager.temp_report_path is None
        ):
            messagebox.showwarning(
                "No Report",
                (
                    "Please run VERIFY BOM first "
                    "to generate a report."
                ),
                parent=self.root,
            )

            return

        try:

            saved = (
                self.report_manager.save_report(
                    self.root
                )
            )

            if saved:

                messagebox.showinfo(
                    "Report Saved",
                    (
                        "BOM verification report "
                        "saved successfully."
                    ),
                    parent=self.root,
                )

        except Exception as error:

            messagebox.showerror(
                "Save Error",
                (
                    "The report could not be saved.\n\n"
                    f"{error}"
                ),
                parent=self.root,
            )

    # ================================================================
    # RESET
    # ================================================================

    def _clear(self) -> None:

        if self.verification_running:

            messagebox.showwarning(
                "Verification Running",
                (
                    "Please wait until the current verification "
                    "is completed."
                ),
                parent=self.root,
            )

            return

        self.report_manager.clear_temp_report()

        self._reset_results()

        self._set_progress(
            0,
            "Ready. Select the BOM folders and click VERIFY BOM."
        )

        self._set_controls(
            running=False
        )

    def _reset_report_state(self) -> None:

        if self.verification_running:
            return

        self.report_manager.clear_temp_report()

        self._reset_results()

        self._set_progress(
            0,
            "Folders changed. Click VERIFY BOM to start verification."
        )

    def _reset_results(self) -> None:

        self.total_pairs_var.set("0")
        self.verified_pairs_var.set("0")
        self.matched_var.set("0")
        self.qty_mismatch_var.set("0")
        self.pdf_only_var.set("0")
        self.excel_only_var.set("0")
        self.review_var.set("0")
        self.errors_var.set("0")

        if self.save_button is not None:
            self.save_button.config(
                state="disabled"
            )

    # ================================================================
    # PROGRESS
    # ================================================================

    def _queue_progress(
        self,
        value: float,
        message: str,
    ) -> None:

        self.root.after(
            0,
            self._set_progress,
            value,
            message,
        )

    def _set_progress(
        self,
        value: float,
        message: str,
    ) -> None:

        self.progress_var.set(
            value
        )

        self.status_var.set(
            message
        )

        if self.progress_bar is not None:
            self.progress_bar.update_idletasks()

    # ================================================================
    # BUTTON STATES
    # ================================================================

    def _set_controls(
        self,
        running: bool,
    ) -> None:

        if self.verify_button is not None:

            self.verify_button.config(
                state=(
                    "disabled"
                    if running
                    else "normal"
                )
            )

        if self.clear_button is not None:

            self.clear_button.config(
                state=(
                    "disabled"
                    if running
                    else "normal"
                )
            )

    # ================================================================
    # CLOSE
    # ================================================================

    def _on_close(self) -> None:

        if self.verification_running:

            messagebox.showwarning(
                "Verification Running",
                (
                    "Verification is still running.\n\n"
                    "Please wait until it is completed."
                ),
                parent=self.root,
            )

            return

        self.report_manager.clear_temp_report()

        self.root.destroy()


# ====================================================================
# APPLICATION ENTRY POINTS
# ====================================================================

def create_application() -> tk.Tk:
    """
    Create and configure the application.
    """

    root = tk.Tk()

    MainWindow(root)

    return root


def launch() -> None:
    """
    Backward-compatible application launcher.

    This is kept because the existing main.py imports launch().
    """

    root = create_application()

    root.mainloop()


def main() -> None:
    """
    Standard application entry point.
    """

    launch()


if __name__ == "__main__":
    launch()