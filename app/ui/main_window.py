from __future__ import annotations

import os
import queue
import shutil
import tempfile
import threading
from pathlib import Path
from typing import Any

try:
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk
except ImportError:  # pragma: no cover - exercised in headless CI
    tk = None
    filedialog = None
    messagebox = None
    ttk = None

from app.core.batch_processor import BatchProcessor
from app.reporting.report_generator import ReportGenerator


class DesktopFolderManager:
    """Manage default local folders without preferring cloud-synced desktops."""

    ENV_ROOT = "EXCEL_PDF_VERIFIER_HOME"

    def __init__(self) -> None:
        self.root_path = self._find_local_root()
        self.files_path = self.root_path / "input"
        self.excel_path = self.files_path / "excel"
        self.pdf_path = self.files_path / "pdf"

    def prepare_default_folders(self) -> tuple[Path, Path]:
        self.excel_path.mkdir(parents=True, exist_ok=True)
        self.pdf_path.mkdir(parents=True, exist_ok=True)
        return self.excel_path, self.pdf_path

    @classmethod
    def _find_local_root(cls) -> Path:
        configured_root = os.environ.get(cls.ENV_ROOT)
        if configured_root:
            return Path(configured_root).expanduser()

        cwd = Path.cwd()
        if (cwd / "input").exists():
            return cwd

        return Path.home() / "ExcelPDFVerifier"


class FolderSelectionService:
    def __init__(self, parent: Any | None) -> None:
        self.parent = parent

    def select_folder(self, title: str, initial_directory: str) -> str:
        if filedialog is None:
            raise RuntimeError("Tkinter file dialogs are unavailable in this environment.")

        selected = filedialog.askdirectory(
            parent=self.parent,
            title=title,
            initialdir=initial_directory,
            mustexist=True,
        )
        return selected.strip() if selected else ""

    @staticmethod
    def validate_folder(path: str, label: str) -> None:
        if not path:
            raise ValueError(f"Please select the {label}.")

        folder = Path(path)
        if not folder.exists():
            raise ValueError(f"{label} does not exist:\n\n{folder}")
        if not folder.is_dir():
            raise ValueError(f"The selected {label} is not a folder:\n\n{folder}")


class ReportFileManager:
    def __init__(self) -> None:
        self.temp_directory: Path | None = None
        self.temp_report_path: Path | None = None

    def create_temp_report_path(self) -> Path:
        self.clear_temp_report()
        self.temp_directory = Path(tempfile.mkdtemp(prefix="excel_pdf_bom_verifier_"))
        self.temp_report_path = self.temp_directory / "BOM_Verification_Report.xlsx"
        return self.temp_report_path

    def save_report(self, parent: Any | None) -> bool:
        if self.temp_report_path is None:
            raise ValueError("No verification report is available.")
        if not self.temp_report_path.exists():
            raise FileNotFoundError("The generated report could not be found.")
        if filedialog is None:
            raise RuntimeError("Tkinter file dialogs are unavailable in this environment.")

        destination = filedialog.asksaveasfilename(
            parent=parent,
            title="Save BOM Verification Report",
            initialfile="BOM_Verification_Report.xlsx",
            defaultextension=".xlsx",
            filetypes=[("Excel Workbook", "*.xlsx"), ("All Files", "*.*")],
        )
        if not destination:
            return False

        destination_path = Path(destination)
        destination_path.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(self.temp_report_path, destination_path)
        return True

    def clear_temp_report(self) -> None:
        if self.temp_directory is None:
            return
        try:
            shutil.rmtree(self.temp_directory, ignore_errors=True)
        finally:
            self.temp_directory = None
            self.temp_report_path = None


class VerificationController:
    def __init__(self) -> None:
        self.batch_processor = BatchProcessor()
        self.report_generator = ReportGenerator()

    def verify(
        self,
        excel_folder: str,
        pdf_folder: str,
        report_path: Path,
        bom1_type: str = "Excel",
        bom2_type: str = "PDF",
    ) -> dict[str, Any]:
        batch_result = self.batch_processor.process_folders(excel_folder, pdf_folder)
        self.report_generator.generate(
            batch_result,
            str(report_path),
            bom1_type=bom1_type,
            bom2_type=bom2_type,
        )
        return batch_result


def normalize_bom_selection(bom1_value: str) -> tuple[str, str]:
    bom1 = "PDF" if str(bom1_value).strip().upper() == "PDF" else "Excel"
    bom2 = "Excel" if bom1 == "PDF" else "PDF"
    return bom1, bom2


def summary_display_values(summary: dict[str, Any]) -> dict[str, str]:
    file_issue_count = (
        int(summary.get("processing_error", 0))
        + int(summary.get("review", 0))
        + int(summary.get("excel_only", 0))
        + int(summary.get("pdf_only", 0))
        + int(summary.get("duplicate_files", 0))
        + int(summary.get("paired_not_verified", 0))
    )
    return {
        "total_pairs": str(summary.get("total_pairs", 0)),
        "verified_pairs": str(summary.get("verified", 0)),
        "matched": str(summary.get("match", 0)),
        "qty_mismatch": str(summary.get("qty_mismatch", 0)),
        "pdf_only": str(summary.get("bom_pdf_only", 0)),
        "excel_only": str(summary.get("bom_excel_only", 0)),
        "review": str(summary.get("bom_review", 0)),
        "errors": str(summary.get("processing_error", 0)),
        "file_issues": str(file_issue_count),
    }


if tk is not None:
    class MainWindow:
        WINDOW_TITLE = "Excel / PDF BOM Verifier"
        WINDOW_WIDTH = 1150
        WINDOW_HEIGHT = 900

        def __init__(self, root: tk.Tk) -> None:
            self.root = root
            self.folder_manager = DesktopFolderManager()
            self.folder_service = FolderSelectionService(root)
            self.verification_controller = VerificationController()
            self.report_manager = ReportFileManager()
            self.verification_running = False
            self._worker_queue: queue.Queue[tuple[str, Any]] = queue.Queue()
            self._worker: threading.Thread | None = None
            self._suspend_change_tracking = False

            self.excel_folder_var = tk.StringVar()
            self.pdf_folder_var = tk.StringVar()
            self.bom1_type = tk.StringVar(value="Excel")
            self.bom2_type = tk.StringVar(value="PDF")
            self.progress_var = tk.DoubleVar(value=0)
            self.status_var = tk.StringVar(
                value="Ready. Select the BOM folders and click VERIFY BOM."
            )
            self.total_pairs_var = tk.StringVar(value="0")
            self.verified_pairs_var = tk.StringVar(value="0")
            self.matched_var = tk.StringVar(value="0")
            self.qty_mismatch_var = tk.StringVar(value="0")
            self.pdf_only_var = tk.StringVar(value="0")
            self.excel_only_var = tk.StringVar(value="0")
            self.review_var = tk.StringVar(value="0")
            self.errors_var = tk.StringVar(value="0")

            self.verify_button: ttk.Button | None = None
            self.save_button: ttk.Button | None = None
            self.clear_button: ttk.Button | None = None
            self.progress_bar: ttk.Progressbar | None = None
            self.bom1_combo: ttk.Combobox | None = None
            self.bom2_combo: ttk.Combobox | None = None
            self._input_controls: list[Any] = []
            self._control_states: dict[Any, str] = {}

            self._configure_window()
            self._build_interface()
            self._load_default_folders()
            self._bind_change_tracking()
            self.root.protocol("WM_DELETE_WINDOW", self._on_close)

        def _configure_window(self) -> None:
            self.root.title(self.WINDOW_TITLE)
            self.root.geometry(f"{self.WINDOW_WIDTH}x{self.WINDOW_HEIGHT}")
            self.root.minsize(1050, 760)
            self.root.columnconfigure(0, weight=1)
            self.root.rowconfigure(0, weight=1)

        def _build_interface(self) -> None:
            main_frame = ttk.Frame(self.root, padding=20)
            main_frame.grid(row=0, column=0, sticky="nsew")
            main_frame.columnconfigure(0, weight=1)
            self._build_header(main_frame)
            self._build_folder_section(main_frame)
            self._build_progress_section(main_frame)
            self._build_results_section(main_frame)
            self._build_action_section(main_frame)

        def _build_header(self, parent: ttk.Frame) -> None:
            header = ttk.Frame(parent)
            header.grid(row=0, column=0, sticky="ew", pady=(0, 18))
            header.columnconfigure(0, weight=1)
            tk.Label(header, text=self.WINDOW_TITLE, font=("Segoe UI", 26, "bold")).grid(row=0, column=0, pady=(0, 6))
            tk.Label(
                header,
                text="Compare local Excel BOM data against local PDF BOM data.",
                font=("Segoe UI", 13),
            ).grid(row=1, column=0)

        def _build_folder_section(self, parent: ttk.Frame) -> None:
            section = ttk.LabelFrame(parent, text="1. Select Input Folders", padding=15)
            section.grid(row=1, column=0, sticky="ew", pady=(0, 16))
            section.columnconfigure(1, weight=1)

            ttk.Label(section, text="Excel BOM Folder:", font=("Segoe UI", 11, "bold")).grid(row=0, column=0, sticky="w", padx=(0, 12), pady=8)
            excel_entry = ttk.Entry(section, textvariable=self.excel_folder_var, font=("Segoe UI", 10))
            excel_entry.grid(row=0, column=1, sticky="ew", pady=8)
            excel_button = ttk.Button(section, text="Select Folder", command=self._select_excel_folder, width=15)
            excel_button.grid(row=0, column=2, padx=(12, 0), pady=8)

            ttk.Label(section, text="PDF BOM Folder:", font=("Segoe UI", 11, "bold")).grid(row=1, column=0, sticky="w", padx=(0, 12), pady=8)
            pdf_entry = ttk.Entry(section, textvariable=self.pdf_folder_var, font=("Segoe UI", 10))
            pdf_entry.grid(row=1, column=1, sticky="ew", pady=8)
            pdf_button = ttk.Button(section, text="Select Folder", command=self._select_pdf_folder, width=15)
            pdf_button.grid(row=1, column=2, padx=(12, 0), pady=8)

            ttk.Label(
                section,
                text=(
                    "Default folders are created locally under the configured application path. "
                    "Use company-approved non-synced folders for actual BOM data."
                ),
                font=("Segoe UI", 9),
            ).grid(row=2, column=0, columnspan=3, sticky="w", pady=(8, 0))

            bom_order_frame = ttk.LabelFrame(section, text="BOM1 / BOM2 ORDER", padding=6)
            bom_order_frame.grid(row=3, column=0, columnspan=3, sticky="ew", padx=5, pady=(8, 3))
            ttk.Label(bom_order_frame, text="BOM1:").grid(row=0, column=0, padx=(5, 4), pady=4, sticky="w")
            self.bom1_combo = ttk.Combobox(
                bom_order_frame,
                textvariable=self.bom1_type,
                values=("Excel", "PDF"),
                state="readonly",
                width=12,
            )
            self.bom1_combo.grid(row=0, column=1, padx=(0, 15), pady=4)
            self.bom1_combo.bind("<<ComboboxSelected>>", self._on_bom1_changed)
            ttk.Label(bom_order_frame, text="BOM2:").grid(row=0, column=2, padx=(5, 4), pady=4, sticky="w")
            self.bom2_combo = ttk.Combobox(
                bom_order_frame,
                textvariable=self.bom2_type,
                values=("Excel", "PDF"),
                state="disabled",
                width=12,
            )
            self.bom2_combo.grid(row=0, column=3, padx=(0, 5), pady=4)

            self._input_controls.extend([excel_entry, excel_button, pdf_entry, pdf_button, self.bom1_combo])
            self._control_states.update(
                {
                    excel_entry: str(excel_entry.cget("state")),
                    excel_button: str(excel_button.cget("state")),
                    pdf_entry: str(pdf_entry.cget("state")),
                    pdf_button: str(pdf_button.cget("state")),
                    self.bom1_combo: str(self.bom1_combo.cget("state")),
                }
            )

        def _build_progress_section(self, parent: ttk.Frame) -> None:
            section = ttk.LabelFrame(parent, text="2. Verification Progress", padding=15)
            section.grid(row=2, column=0, sticky="ew", pady=(0, 16))
            section.columnconfigure(0, weight=1)
            self.progress_bar = ttk.Progressbar(section, variable=self.progress_var, maximum=100, mode="determinate")
            self.progress_bar.grid(row=0, column=0, sticky="ew", ipady=4)
            ttk.Label(section, textvariable=self.status_var, font=("Segoe UI", 10)).grid(row=1, column=0, sticky="w", pady=(10, 0))

        def _build_results_section(self, parent: ttk.Frame) -> None:
            section = ttk.LabelFrame(parent, text="3. Verification Results", padding=20)
            section.grid(row=3, column=0, sticky="ew", pady=(0, 16))
            section.columnconfigure(1, weight=1)
            section.columnconfigure(3, weight=1)
            self._create_result_row(section, 0, "Total File Pairs", self.total_pairs_var, "Verified Pairs", self.verified_pairs_var)
            self._create_result_row(section, 1, "Matched Parts", self.matched_var, "Quantity Mismatch", self.qty_mismatch_var)
            self._create_result_row(section, 2, "PDF Only", self.pdf_only_var, "Excel Only", self.excel_only_var)
            self._create_result_row(section, 3, "Review", self.review_var, "Processing Errors", self.errors_var)

        def _create_result_row(
            self,
            parent: ttk.LabelFrame,
            row: int,
            left_label: str,
            left_variable: tk.StringVar,
            right_label: str,
            right_variable: tk.StringVar,
        ) -> None:
            ttk.Label(parent, text=left_label, font=("Segoe UI", 11)).grid(row=row, column=0, sticky="w", padx=(0, 20), pady=10)
            ttk.Label(parent, textvariable=left_variable, font=("Segoe UI", 12, "bold")).grid(row=row, column=1, sticky="e", padx=(0, 50), pady=10)
            ttk.Label(parent, text=right_label, font=("Segoe UI", 11)).grid(row=row, column=2, sticky="w", padx=(20, 20), pady=10)
            ttk.Label(parent, textvariable=right_variable, font=("Segoe UI", 12, "bold")).grid(row=row, column=3, sticky="e", pady=10)

        def _build_action_section(self, parent: ttk.Frame) -> None:
            frame = ttk.Frame(parent)
            frame.grid(row=4, column=0, sticky="ew")
            frame.columnconfigure(1, weight=1)
            self.clear_button = ttk.Button(frame, text="Clear", command=self._clear, width=15)
            self.clear_button.grid(row=0, column=0, sticky="w")
            self.verify_button = ttk.Button(frame, text="VERIFY BOM", command=self._start_verification, width=20)
            self.verify_button.grid(row=0, column=1)
            self.save_button = ttk.Button(frame, text="SAVE REPORT", command=self._save_report, width=20, state="disabled")
            self.save_button.grid(row=0, column=2, sticky="e")

        def _bind_change_tracking(self) -> None:
            for variable in (self.excel_folder_var, self.pdf_folder_var):
                variable.trace_add("write", self._on_input_changed)

        def _load_default_folders(self) -> None:
            self._suspend_change_tracking = True
            try:
                excel_folder, pdf_folder = self.folder_manager.prepare_default_folders()
                self.excel_folder_var.set(str(excel_folder))
                self.pdf_folder_var.set(str(pdf_folder))
                bom1, bom2 = normalize_bom_selection(self.bom1_type.get())
                self.bom1_type.set(bom1)
                self.bom2_type.set(bom2)
            finally:
                self._suspend_change_tracking = False

        def _on_input_changed(self, *_args: Any) -> None:
            if self._suspend_change_tracking or self.verification_running:
                return
            self._reset_report_state("Inputs changed. Click VERIFY BOM to start verification.")

        def _select_excel_folder(self) -> None:
            initial_directory = self.excel_folder_var.get().strip() or str(self.folder_manager.root_path)
            selected = self.folder_service.select_folder("Select Excel BOM Folder", initial_directory)
            if selected:
                self.excel_folder_var.set(selected)

        def _select_pdf_folder(self) -> None:
            initial_directory = self.pdf_folder_var.get().strip() or str(self.folder_manager.root_path)
            selected = self.folder_service.select_folder("Select PDF BOM Folder", initial_directory)
            if selected:
                self.pdf_folder_var.set(selected)

        def _on_bom1_changed(self, _event: Any = None) -> None:
            bom1, bom2 = normalize_bom_selection(self.bom1_type.get())
            self._suspend_change_tracking = True
            try:
                self.bom1_type.set(bom1)
                self.bom2_type.set(bom2)
            finally:
                self._suspend_change_tracking = False
            if not self.verification_running:
                self._reset_report_state("BOM order changed. Click VERIFY BOM to regenerate the report.")

        def _start_verification(self) -> None:
            if self.verification_running:
                return

            excel_folder = self.excel_folder_var.get().strip()
            pdf_folder = self.pdf_folder_var.get().strip()
            bom1_type, bom2_type = normalize_bom_selection(self.bom1_type.get())

            try:
                self.folder_service.validate_folder(excel_folder, "Excel BOM folder")
                self.folder_service.validate_folder(pdf_folder, "PDF BOM folder")
            except ValueError as error:
                messagebox.showerror("Invalid Folder", str(error), parent=self.root)
                return

            self.report_manager.clear_temp_report()
            report_path = self.report_manager.create_temp_report_path()
            self._reset_results()
            self._set_running_state(True)
            self.status_var.set("Verification started. Processing selected BOM folders locally...")
            self.progress_var.set(0)

            self._worker = threading.Thread(
                target=self._verification_worker,
                args=(excel_folder, pdf_folder, bom1_type, bom2_type, report_path),
                daemon=True,
                name="BOMVerificationWorker",
            )
            self._worker.start()
            self._poll_worker_queue()

        def _verification_worker(
            self,
            excel_folder: str,
            pdf_folder: str,
            bom1_type: str,
            bom2_type: str,
            report_path: Path,
        ) -> None:
            try:
                self._worker_queue.put(("status", "Extracting, pairing, verifying, and writing the report..."))
                result = self.verification_controller.verify(
                    excel_folder=excel_folder,
                    pdf_folder=pdf_folder,
                    report_path=report_path,
                    bom1_type=bom1_type,
                    bom2_type=bom2_type,
                )
                self._worker_queue.put(("completed", result))
            except Exception as error:  # pragma: no cover - GUI pathway
                self._worker_queue.put(("failed", error))

        def _poll_worker_queue(self) -> None:
            while True:
                try:
                    event, payload = self._worker_queue.get_nowait()
                except queue.Empty:
                    break

                if event == "status":
                    self.status_var.set(str(payload))
                elif event == "completed":
                    self._verification_completed(payload)
                    return
                elif event == "failed":
                    self._verification_failed(payload)
                    return

            if self.verification_running:
                self.root.after(100, self._poll_worker_queue)

        def _verification_completed(self, result: dict[str, Any]) -> None:
            summary = result.get("summary", {})
            display = summary_display_values(summary)
            self.total_pairs_var.set(display["total_pairs"])
            self.verified_pairs_var.set(display["verified_pairs"])
            self.matched_var.set(display["matched"])
            self.qty_mismatch_var.set(display["qty_mismatch"])
            self.pdf_only_var.set(display["pdf_only"])
            self.excel_only_var.set(display["excel_only"])
            self.review_var.set(display["review"])
            self.errors_var.set(display["errors"])

            verified_pairs = int(display["verified_pairs"])
            file_issues = int(display["file_issues"])
            report_available = bool(
                verified_pairs > 0
                and self.report_manager.temp_report_path
                and self.report_manager.temp_report_path.exists()
            )
            self._set_running_state(False, report_available=report_available)
            self.progress_var.set(
                100 if verified_pairs > 0 else 0
            )
            workbook_exists = bool(
                self.report_manager.temp_report_path
                and self.report_manager.temp_report_path.exists()
            )

            if verified_pairs == 0:
                self.status_var.set("Verification finished, but no file pairs were successfully verified.")
                messagebox.showwarning(
                    "Verification Completed",
                    "Verification finished, but no file pairs were successfully verified. Review the selected folders and summary sheet.",
                    parent=self.root,
                )
            elif file_issues:
                self.status_var.set(f"Verification completed with {file_issues} file-level issue(s).")
            elif workbook_exists:
                self.status_var.set("Verification completed. Report is ready to save.")
            else:
                self.status_var.set("Verification completed.")

        def _verification_failed(self, error: Exception) -> None:
            self.report_manager.clear_temp_report()
            self._set_running_state(False, report_available=False)
            self.progress_var.set(0)
            self.status_var.set("Verification failed.")
            messagebox.showerror(
                "Verification Error",
                f"The BOM verification could not be completed.\n\n{error}",
                parent=self.root,
            )

        def _set_running_state(self, running: bool, report_available: bool = False) -> None:
            self.verification_running = running
            for control in self._input_controls:
                control.configure(
                    state="disabled"
                    if running
                    else self._control_states.get(control, "normal")
                )
            if self.clear_button is not None:
                self.clear_button.configure(state="disabled" if running else "normal")
            if self.verify_button is not None:
                self.verify_button.configure(state="disabled" if running else "normal")
            if self.save_button is not None:
                self.save_button.configure(state="disabled" if running or not report_available else "normal")
            if self.progress_bar is not None:
                if running:
                    self.progress_bar.configure(mode="indeterminate")
                    self.progress_bar.start(10)
                else:
                    self.progress_bar.stop()
                    self.progress_bar.configure(mode="determinate")

        def _save_report(self) -> None:
            if self.verification_running:
                return
            if self.report_manager.temp_report_path is None:
                messagebox.showwarning("No Report", "Please run VERIFY BOM first to generate a report.", parent=self.root)
                return
            try:
                saved = self.report_manager.save_report(self.root)
                if saved:
                    messagebox.showinfo("Report Saved", "BOM verification report saved successfully.", parent=self.root)
            except Exception as error:
                messagebox.showerror("Save Error", f"The report could not be saved.\n\n{error}", parent=self.root)

        def _clear(self) -> None:
            if self.verification_running:
                messagebox.showwarning("Verification Running", "Please wait until the current verification is completed.", parent=self.root)
                return
            self.report_manager.clear_temp_report()
            self._suspend_change_tracking = True
            try:
                self._load_default_folders()
                bom1, bom2 = normalize_bom_selection("Excel")
                self.bom1_type.set(bom1)
                self.bom2_type.set(bom2)
            finally:
                self._suspend_change_tracking = False
            self._reset_results()
            self.progress_var.set(0)
            self.status_var.set("Ready. Select the BOM folders and click VERIFY BOM.")
            self._set_running_state(False, report_available=False)

        def _reset_report_state(self, message: str) -> None:
            if self.verification_running:
                return
            self.report_manager.clear_temp_report()
            self._reset_results()
            self.progress_var.set(0)
            self.status_var.set(message)
            self._set_running_state(False, report_available=False)

        def _reset_results(self) -> None:
            self.total_pairs_var.set("0")
            self.verified_pairs_var.set("0")
            self.matched_var.set("0")
            self.qty_mismatch_var.set("0")
            self.pdf_only_var.set("0")
            self.excel_only_var.set("0")
            self.review_var.set("0")
            self.errors_var.set("0")

        def _on_close(self) -> None:
            if self.verification_running:
                messagebox.showwarning(
                    "Verification Running",
                    "Verification is still running. Please wait until it is completed.",
                    parent=self.root,
                )
                return
            self.report_manager.clear_temp_report()
            self.root.destroy()
else:
    class MainWindow:  # pragma: no cover - only used when tkinter is unavailable
        def __init__(self, _root: Any) -> None:
            raise RuntimeError("Tkinter is unavailable in this environment.")



def create_application() -> Any:
    if tk is None:
        raise RuntimeError("Tkinter is unavailable in this environment.")
    root = tk.Tk()
    MainWindow(root)
    return root



def launch() -> None:
    root = create_application()
    root.mainloop()



def main() -> None:
    launch()


if __name__ == "__main__":
    launch()
