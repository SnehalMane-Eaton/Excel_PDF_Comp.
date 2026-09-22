from pathlib import Path
import ast
import shutil

path = Path("app/ui/main_window.py")
backup = Path("app/ui/main_window.py.backup_before_bom_selection.py")

# Always start from the clean original
if backup.exists():
    shutil.copy2(backup, path)

text = path.read_text(encoding="utf-8")
tree = ast.parse(text)

main_cls = next(
    n for n in tree.body
    if isinstance(n, ast.ClassDef) and n.name == "MainWindow"
)

methods = {
    n.name: n for n in main_cls.body
    if isinstance(n, (ast.FunctionDef, ast.AsyncFunctionDef))
}

# ---------------------------------------------------------
# 1. Add BOM variables to __init__
# ---------------------------------------------------------
if "self.bom1_type" not in text:
    init = methods["__init__"]
    lines = text.splitlines(True)

    # Insert immediately before __init__ ends.
    insert_at = init.end_lineno - 1

    indent = " " * (init.col_offset + 4)

    block = (
        f"{indent}# BOM order selection\n"
        f"{indent}# Default: BOM1 = Excel, BOM2 = PDF\n"
        f'{indent}self.bom1_type = tk.StringVar(value="Excel")\n'
        f'{indent}self.bom2_type = tk.StringVar(value="PDF")\n'
    )

    lines.insert(insert_at, block)
    text = "".join(lines)

# Reparse after modification
tree = ast.parse(text)

main_cls = next(
    n for n in tree.body
    if isinstance(n, ast.ClassDef) and n.name == "MainWindow"
)

methods = {
    n.name: n for n in main_cls.body
    if isinstance(n, (ast.FunctionDef, ast.AsyncFunctionDef))
}

# ---------------------------------------------------------
# 2. Add BOM selection UI at the END of folder section
# ---------------------------------------------------------
if "BOM1 / BOM2 ORDER" not in text:
    method = methods["_build_folder_section"]
    lines = text.splitlines(True)

    # Insert before the next MainWindow method.
    next_method = next(
        (
            n for n in main_cls.body
            if isinstance(n, (ast.FunctionDef, ast.AsyncFunctionDef))
            and n.lineno > method.lineno
        ),
        None,
    )

    if next_method is None:
        raise RuntimeError("Could not locate next MainWindow method.")

    insert_at = next_method.lineno - 1
    indent = " " * (method.col_offset + 4)

    block = f"""
{indent}# -------------------------------------------------
{indent}# BOM1 / BOM2 ORDER
{indent}# -------------------------------------------------
{indent}bom_order_frame = ttk.LabelFrame(
{indent}    section,
{indent}    text="BOM1 / BOM2 ORDER",
{indent}    padding=6,
{indent})
{indent}bom_order_frame.grid(
{indent}    row=2,
{indent}    column=0,
{indent}    columnspan=3,
{indent}    sticky="ew",
{indent}    padx=5,
{indent}    pady=(8, 3),
{indent})

{indent}ttk.Label(
{indent}    bom_order_frame,
{indent}    text="BOM1:"
{indent}).grid(row=0, column=0, padx=(5, 4), pady=4, sticky="w")

{indent}self.bom1_combo = ttk.Combobox(
{indent}    bom_order_frame,
{indent}    textvariable=self.bom1_type,
{indent}    values=("Excel", "PDF"),
{indent}    state="readonly",
{indent}    width=12,
{indent})
{indent}self.bom1_combo.grid(row=0, column=1, padx=(0, 15), pady=4)

{indent}ttk.Label(
{indent}    bom_order_frame,
{indent}    text="BOM2:"
{indent}).grid(row=0, column=2, padx=(5, 4), pady=4, sticky="w")

{indent}self.bom2_combo = ttk.Combobox(
{indent}    bom_order_frame,
{indent}    textvariable=self.bom2_type,
{indent}    values=("Excel", "PDF"),
{indent}    state="disabled",
{indent}    width=12,
{indent})
{indent}self.bom2_combo.grid(row=0, column=3, padx=(0, 5), pady=4)

"""
    lines.insert(insert_at, block)
    text = "".join(lines)

# ---------------------------------------------------------
# 3. Add BOM1 change handler
# ---------------------------------------------------------
if "def _on_bom1_changed" not in text:
    tree = ast.parse(text)

    main_cls = next(
        n for n in tree.body
        if isinstance(n, ast.ClassDef) and n.name == "MainWindow"
    )

    methods = {
        n.name: n for n in main_cls.body
        if isinstance(n, (ast.FunctionDef, ast.AsyncFunctionDef))
    }

    start_method = methods["_start_verification"]
    lines = text.splitlines(True)

    insert_at = start_method.lineno - 1
    indent = " " * start_method.col_offset

    handler = (
        f"{indent}def _on_bom1_changed(self, _event=None):\n"
        f'{indent}    """Automatically keep BOM2 opposite to BOM1."""\n'
        f'        selected = self.bom1_type.get().strip()\n'
        f'        if selected == "PDF":\n'
        f'            self.bom2_type.set("Excel")\n'
        f'        else:\n'
        f'            self.bom1_type.set("Excel")\n'
        f'            self.bom2_type.set("PDF")\n'
        f"\n"
    )

    lines.insert(insert_at, handler)
    text = "".join(lines)

# ---------------------------------------------------------
# 4. Bind the combobox
# ---------------------------------------------------------
if 'self.bom1_combo.bind("<<ComboboxSelected>>"' not in text:
    marker = 'self.bom1_combo.grid(row=0, column=1, padx=(0, 15), pady=4)'
    replacement = (
        marker
        + '\n'
        + '        self.bom1_combo.bind('
        + '"<<ComboboxSelected>>", self._on_bom1_changed)'
    )

    if marker not in text:
        raise RuntimeError("Could not find BOM1 combobox.")
    text = text.replace(marker, replacement, 1)

# ---------------------------------------------------------
# 5. Add BOM order to ReportGenerator.generate()
# ---------------------------------------------------------
tree = ast.parse(text)

generate_calls = [
    n for n in ast.walk(tree)
    if isinstance(n, ast.Call)
    and isinstance(n.func, ast.Attribute)
    and n.func.attr == "generate"
    and isinstance(n.func.value, ast.Attribute)
    and n.func.value.attr == "report_generator"
]

if not generate_calls:
    raise RuntimeError("Could not find ReportGenerator.generate() call.")

if "bom1_type=self.bom1_type.get()" not in text:
    node = generate_calls[0]

    lines = text.splitlines(True)

    start = node.lineno - 1
    end = node.end_lineno

    source = "".join(lines[start:end])

    closing = source.rfind(")")
    if closing < 0:
        raise RuntimeError("Could not locate generate() closing parenthesis.")

    addition = (
        ",\n"
        "                bom1_type=self.bom1_type.get(),\n"
        "                bom2_type=self.bom2_type.get()"
    )

    source = source[:closing] + addition + source[closing:]
    lines[start:end] = [source]

    text = "".join(lines)

# ---------------------------------------------------------
# 6. Validate before writing
# ---------------------------------------------------------
ast.parse(text)

# ---------------------------------------------------------
# 7. Write modified file
# ---------------------------------------------------------
path.write_text(text, encoding="utf-8")

print("SUCCESS")
print("main_window.py updated.")
print("BOM1 default = Excel")
print("BOM2 default = PDF")
print("BOM1 PDF -> BOM2 Excel")
print("Backup preserved:", backup)
