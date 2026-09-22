from pathlib import Path

p = Path("app/reporting/report_generator.py")
s = p.read_text(encoding="utf-8")

start = s.find("        # BatchProcessor returns:")
end = s.find("        return generate_report(", start)

if start == -1 or end == -1:
    raise SystemExit("Could not find the temporary report connection change.")

s = s[:start] + s[end:]

p.write_text(s, encoding="utf-8")
compile(s, str(p), "exec")

print("RESTORED REPORT INPUT")
