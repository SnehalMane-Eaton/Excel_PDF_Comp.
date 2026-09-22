from pathlib import Path
p=Path("app/reporting/report_generator.py")
s=p.read_text(encoding="utf-8")
s=s.replace(
'''        return generate_report(
            verification_data=verification_data,
            output_path=output_path,
            bom1_type=bom1_type,
            bom2_type=bom2_type,
        )''',
'''        # BatchProcessor returns:
        # {pairs: [{excel_file, pdf_file, verification: {results, summary}}]}
        # The report must consume the actual verification results.
        if isinstance(verification_data, dict):
            pairs = verification_data.get("pairs", [])
            report_data = []
            for pair in pairs:
                verification = pair.get("verification", {}) if isinstance(pair, dict) else {}
                results = verification.get("results", []) if isinstance(verification, dict) else []
                report_data.extend(results)
        else:
            report_data = verification_data

        return generate_report(
            verification_data=report_data,
            output_path=output_path,
            bom1_type=bom1_type,
            bom2_type=bom2_type,
        )''',
1
)
p.write_text(s,encoding="utf-8")
compile(s,str(p),"exec")
print("REPORT CONNECTION FIXED")
