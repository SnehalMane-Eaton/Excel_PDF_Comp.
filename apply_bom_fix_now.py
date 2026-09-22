from pathlib import Path
import ast

p = Path("app/ui/main_window.py")
s = p.read_text(encoding="utf-8")

# 1. Remove incorrect BOM string attributes from VerificationController
s = s.replace('        self.bom1_type = "Excel"\n        self.bom2_type = "PDF"\n', '', 1)

# 2. Ensure verify() accepts BOM order
old_sig = '''    def verify(
        self,
        excel_folder: str,
        pdf_folder: str,
        report_path: Path,
    ) -> dict[str, Any]:'''

new_sig = '''    def verify(
        self,
        excel_folder: str,
        pdf_folder: str,
        report_path: Path,
        bom1_type: str = "Excel",
        bom2_type: str = "PDF",
    ) -> dict[str, Any]:'''

if old_sig in s:
    s = s.replace(old_sig, new_sig, 1)

# 3. Replace the controller's incorrect StringVar access
s = s.replace(
'''                bom1_type=self.bom1_type.get(),
                bom2_type=self.bom2_type.get()
''',
'''                bom1_type=bom1_type,
                bom2_type=bom2_type,
''',
1
)

# 4. Pass the UI StringVar values from MainWindow to controller.
# Replace the first controller.verify call that does not already contain BOM args.
old_call = '''self.controller.verify(
                excel_folder,
                pdf_folder,
                report_path,
            )'''

new_call = '''self.controller.verify(
                excel_folder,
                pdf_folder,
                report_path,
                bom1_type=self.bom1_type.get(),
                bom2_type=self.bom2_type.get(),
            )'''

if old_call in s:
    s = s.replace(old_call, new_call, 1)

# Validate syntax before writing.
compile(s, str(p), "exec")

# Validate the Controller itself.
tree = ast.parse(s)
controller = next(
    x for x in tree.body
    if isinstance(x, ast.ClassDef) and x.name == "VerificationController"
)

controller_text = ast.get_source_segment(s, controller)

if "self.bom1_type = " in controller_text:
    raise RuntimeError("Controller still contains incorrect bom1_type assignment.")

if "self.bom2_type = " in controller_text:
    raise RuntimeError("Controller still contains incorrect bom2_type assignment.")

if "self.bom1_type.get()" in controller_text:
    raise RuntimeError("Controller still contains self.bom1_type.get().")

if "self.bom2_type.get()" in controller_text:
    raise RuntimeError("Controller still contains self.bom2_type.get().")

p.write_text(s, encoding="utf-8")

print("========================================")
print("BOM FLOW FIXED SUCCESSFULLY")
print("========================================")
print("Controller: receives BOM strings")
print("MainWindow: owns Tkinter StringVars")
print("Controller: no .get() on BOM variables")
print("Syntax: OK")
