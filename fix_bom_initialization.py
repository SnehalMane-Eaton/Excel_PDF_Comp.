from pathlib import Path

path = Path("app/ui/main_window.py")
text = path.read_text(encoding="utf-8")

old = '''        self._build_interface()
'''

new = '''        # BOM order selection
        self.bom1_type = tk.StringVar(value="Excel")
        self.bom2_type = tk.StringVar(value="PDF")

        self._build_interface()
'''

if old not in text:
    raise RuntimeError("Could not find self._build_interface() in __init__.")

# Remove existing BOM initialization wherever it currently occurs.
lines = text.splitlines(True)
filtered = [
    line for line in lines
    if 'self.bom1_type = tk.StringVar(value="Excel")' not in line
    and 'self.bom2_type = tk.StringVar(value="PDF")' not in line
]

text = "".join(filtered)

# Insert initialization immediately before _build_interface().
old = '''        self._build_interface()
'''
new = '''        # BOM order selection
        self.bom1_type = tk.StringVar(value="Excel")
        self.bom2_type = tk.StringVar(value="PDF")

        self._build_interface()
'''

if old not in text:
    raise RuntimeError("Could not find _build_interface() after cleanup.")

text = text.replace(old, new, 1)

path.write_text(text, encoding="utf-8")

print("SUCCESS: BOM variables moved before _build_interface().")
