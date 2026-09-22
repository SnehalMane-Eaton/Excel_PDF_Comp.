from pathlib import Path
import ast
import shutil

path = Path("app/ui/main_window.py")
backup = Path("app/ui/main_window.py.backup_before_bom_selection.py")

# Always restore the clean original first
if backup.exists():
    shutil.copy2(backup, path)

text = path.read_text(encoding="utf-8")
tree = ast.parse(text)


def get_main_window(tree):
    return next(
        n for n in tree.body
        if isinstance(n, ast.ClassDef) and n.name == "MainWindow"
    )


def get_methods(cls):
    return {
        n.name: n
        for n in cls.body
        if isinstance(n, (ast.FunctionDef, ast.AsyncFunctionDef))
    }


# =========================================================
# 1. Add BOM variables inside __init__
# =========================================================
if "self.bom1_type = tk.StringVar" not in text:
    cls = get_main_window(tree)
    methods = get_methods(cls)
    init = methods["__init__"]

    lines = text.splitlines(True)

    indent = " " * (init.col_offset + 4)

    block = (
        f"{indent}# BOM order selection\n"
        f'{indent}self.bom1_type = tk.StringVar(value="Excel")\n'
        f'{indent}self.bom2_type = tk.StringVar(value="PDF")\n'
    )

    # IMPORTANT: insert AFTER the complete method,
    # not inside a multiline statement.
    lines.insert(init.end_lineno, block)

    text = "".join(lines)


# =========================================================
# 2. Add BOM selector inside _build_folder_section
# =========================================================
tree = ast.parse(text)

if "BOM1 / BOM2 ORDER" not in text:
    cls = get_main_window(tree)
    methods = get_methods(cls)
    method = methods["_build_folder_section"]

    lines = text.splitlines(True)

    indent = " " * (method.col_offset + 4)

    block = (
        f"\n"
        f"{indent}# BOM1 / BOM2 ORDER\n"
        f"{indent}bom_order_frame = ttk.LabelFrame(\n"
        f"{indent}    section,\n"
        f'{indent}    text="BOM1 / BOM2 ORDER",\n'
        f"{indent}    padding=6,\n"
        f"{indent})\n"
        f"{indent}bom_order_frame.grid(\n"
        f"{indent}    row=2,\n"
        f"{indent}    column=0,\n"
        f"{indent}    columnspan=3,\n"
        f"{indent}    sticky=\"ew\",\n"
        f"{indent}    padx=5,\n"
        f"{indent}    pady=(8, 3),\n"
        f"{indent})\n"
        f"\n"
        f"{indent}ttk.Label(\n"
        f"{indent}    bom_order_frame,\n"
        f'{indent}    text="BOM1:",\n'
        f"{indent}).grid(row=0, column=0, padx=(5, 4), pady=4, sticky=\"w\")\n"
        f"\n"
        f"{indent}self.bom1_combo = ttk.Combobox(\n"
        f"{indent}    bom_order_frame,\n"
        f"{indent}    textvariable=self.bom1_type,\n"
        f'{indent}    values=("Excel", "PDF"),\n'
        f'{indent}    state="readonly",\n'
        f"{indent}    width=12,\n"
        f"{indent})\n"
        f"{indent}self.bom1_combo.grid(\n"
        f"{indent}    row=0, column=1, padx=(0, 15), pady=4\n"
        f"{indent})\n"
        f"{indent}self.bom1_combo.bind(\n"
        f'{indent}    "<<ComboboxSelected>>",\n'
        f"{indent}    self._on_bom1_changed,\n"
        f"{indent})\n"
        f"\n"
        f"{indent}ttk.Label(\n"
        f"{indent}    bom_order_frame,\n"
        f'{indent}    text="BOM2:",\n'
        f"{indent}).grid(row=0, column=2, padx=(5, 4), pady=4, sticky=\"w\")\n"
        f"\n"
        f"{indent}self.bom2_combo = ttk.Combobox(\n"
        f"{indent}    bom_order_frame,\n"
        f"{indent}    textvariable=self.bom2_type,\n"
        f'{indent}    values=("Excel", "PDF"),\n'
        f'{indent}    state="disabled",\n'
        f"{indent}    width=12,\n"
        f"{indent})\n"
        f"{indent}self.bom2_combo.grid(\n"
        f"{indent}    row=0, column=3, padx=(0, 5), pady=4\n"
        f"{indent})\n"
    )

    # Insert AFTER the complete method.
    lines.insert(method.end_lineno, block)

    text = "".join(lines)


# =========================================================
# 3. Add BOM1 -> BOM2 automatic switching method
# =========================================================
tree = ast.parse(text)

if "def _on_bom1_changed" not in text:
    cls = get_main_window(tree)
    methods = get_methods(cls)
    start_method = methods["_start_verification"]

    lines = text.splitlines(True)
    indent = " " * start_method.col_offset

    block = (
        f"{indent}def _on_bom1_changed(self, _event=None):\n"
        f'{indent}    """Automatically make BOM2 the opposite of BOM1."""\n'
        f'{indent}    if self.bom1_type.get() == "PDF":\n'
        f'{indent}        self.bom2_type.set("Excel")\n'
        f"{indent}    else:\n"
        f'{indent}        self.bom1_type.set("Excel")\n'
        f'{indent}        self.bom2_type.set("PDF")\n'
        f"\n"
    )

    # Insert immediately before _start_verification.
    lines.insert(start_method.lineno - 1, block)

    text = "".join(lines)


# =========================================================
# 4. Pass BOM1/BOM2 to report generator
# =========================================================
tree = ast.parse(text)

if "bom1_type=self.bom1_type.get()" not in text:

    calls = []

    for node in ast.walk(tree):
        if isinstance(node, ast.Call):
            if isinstance(node.func, ast.Attribute):
                if node.func.attr == "generate":
                    if (
                        isinstance(node.func.value, ast.Attribute)
                        and node.func.value.attr == "report_generator"
                    ):
                        calls.append(node)

    if not calls:
        raise RuntimeError(
            "Could not find self.report_generator.generate(...)"
        )

    node = calls[0]

    lines = text.splitlines(True)

    start = node.lineno - 1
    end = node.end_lineno

    source = "".join(lines[start:end])

    close = source.rfind(")")
    if close == -1:
        raise RuntimeError("Could not locate generate() closing parenthesis.")

    addition = (
        ",\n"
        "                bom1_type=self.bom1_type.get(),\n"
        "                bom2_type=self.bom2_type.get()"
    )

    source = source[:close] + addition + source[close:]

    lines[start:end] = [source]

    text = "".join(lines)


# =========================================================
# 5. FINAL SYNTAX VALIDATION
# =========================================================
ast.parse(text)

# =========================================================
# 6. Save
# =========================================================
path.write_text(text, encoding="utf-8")

print()
print("==========================================")
print("SUCCESS - BOM selection added")
print("==========================================")
print("BOM1 default : Excel")
print("BOM2 default : PDF")
print("If BOM1 = PDF -> BOM2 = Excel")
print("If BOM1 = Excel -> BOM2 = PDF")
print("Backup       :", backup)
print("==========================================")
