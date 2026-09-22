from pathlib import Path
import ast

path = Path("app/ui/main_window.py")
text = path.read_text(encoding="utf-8")

if "self.bom1_type = tk.StringVar(value=\"Excel\")" not in text:
    tree = ast.parse(text)

    main_cls = next(
        n for n in tree.body
        if isinstance(n, ast.ClassDef) and n.name == "MainWindow"
    )

    build_interface = next(
        n for n in main_cls.body
        if isinstance(n, ast.FunctionDef)
        and n.name == "_build_interface"
    )

    lines = text.splitlines(True)

    indent = " " * (build_interface.col_offset + 4)

    block = (
        f'{indent}# BOM order selection\n'
        f'{indent}self.bom1_type = tk.StringVar(value="Excel")\n'
        f'{indent}self.bom2_type = tk.StringVar(value="PDF")\n'
    )

    # Insert immediately after the _build_interface definition line.
    lines.insert(build_interface.lineno, block)

    text = "".join(lines)

    ast.parse(text)
    path.write_text(text, encoding="utf-8")

    print("SUCCESS: BOM variables initialized.")
else:
    print("BOM variables already present.")
