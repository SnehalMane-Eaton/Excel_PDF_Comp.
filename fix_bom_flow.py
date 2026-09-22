from pathlib import Path

p = Path("app/ui/main_window.py")
s = p.read_text(encoding="utf-8")

old_signature = '''    def verify(
        self,
        excel_folder: str,
        pdf_folder: str,
        report_path: Path,
    ) -> dict[str, Any]:'''

new_signature = '''    def verify(
        self,
        excel_folder: str,
        pdf_folder: str,
        report_path: Path,
        bom1_type: str = "Excel",
        bom2_type: str = "PDF",
    ) -> dict[str, Any]:'''

if old_signature not in s:
    raise SystemExit("ERROR: verify() signature not found.")

s = s.replace(old_signature, new_signature, 1)

old_report_call = '''        self.report_generator.generate(
            batch_result,
            str(report_path),
        
                bom1_type=self.bom1_type.get(),
                bom2_type=self.bom2_type.get()
)'''

new_report_call = '''        self.report_generator.generate(
            batch_result,
            str(report_path),
            bom1_type=bom1_type,
            bom2_type=bom2_type,
        )'''

if old_report_call not in s:
    raise SystemExit("ERROR: report generator call not found.")

s = s.replace(old_report_call, new_report_call, 1)

# Find the verification call in _verification_worker and pass the UI selections.
old_verify_call = '''self.controller.verify(
                excel_folder,
                pdf_folder,
                report_path,
            )'''

new_verify_call = '''self.controller.verify(
                excel_folder,
                pdf_folder,
                report_path,
                bom1_type=self.bom1_type.get(),
                bom2_type=self.bom2_type.get(),
            )'''

if old_verify_call not in s:
    raise SystemExit("ERROR: controller.verify() call not found.")

s = s.replace(old_verify_call, new_verify_call, 1)

compile(s, str(p), "exec")
p.write_text(s, encoding="utf-8")

print("SUCCESS")
print("Controller now receives BOM order from the UI.")
print("Tkinter .get() is used only in MainWindow.")
