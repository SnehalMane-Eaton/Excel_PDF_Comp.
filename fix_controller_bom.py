from pathlib import Path
import ast

path = Path("app/ui/main_window.py")
text = path.read_text(encoding="utf-8")

tree = ast.parse(text)
controller = next(
    n for n in tree.body
    if isinstance(n, ast.ClassDef) and n.name == "VerificationController"
)

init_method = next(
    n for n in controller.body
    if isinstance(n, ast.FunctionDef) and n.name == "__init__"
)

lines = text.splitlines(True)

# Find the first statement inside VerificationController.__init__
insert_at = init_method.body[0].lineno - 1

indent = " " * (init_method.col_offset + 4)

new_lines = [
    f'{indent}self.bom1_type = "Excel"\n',
    f'{indent}self.bom2_type = "PDF"\n',
]

# Avoid duplicate insertion
method_text = "".join(lines[init_method.lineno - 1:init_method.end_lineno])
if 'self.bom1_type = "Excel"' not in method_text:
    lines[insert_at:insert_at] = new_lines

path.write_text("".join(lines), encoding="utf-8")

compile(path.read_text(encoding="utf-8"), str(path), "exec")
print("SUCCESS: VerificationController BOM defaults added.")
