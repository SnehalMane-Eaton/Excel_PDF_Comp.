from app.extractors.pdf_extractor import PDFExtractor

pdf_file = "input/pdf/W10057W-P 1.pdf"

data = PDFExtractor(pdf_file).extract()

target_pages = [5, 8, 9, 14, 15, 16]

for page in data["pages"]:
    if page["page"] not in target_pages:
        continue

    print("\n" + "=" * 100)
    print(f"PAGE {page['page']}")
    print("=" * 100)

    lines = page["text"].splitlines()

    for i, line in enumerate(lines):

        if any(
            part in line
            for part in [
                "221C247G01",
                "W10057W-D127",
                "1A85322G01",
                "4ABV-750",
            ]
        ):

            print("\n" + "-" * 80)
            print(f"MATCH AT PDF LINE {i + 1}")
            print("-" * 80)

            start = max(0, i - 12)
            end = min(len(lines), i + 3)

            for j in range(start, end):
                print(f"{j + 1:03}: {lines[j]}")