import re


def normalize_part_number(value) -> str:
    """
    Normalize a part number for exact comparison.

    Rules:
    - Convert to string
    - Remove leading/trailing whitespace
    - Convert to uppercase
    - Collapse internal whitespace
    - Do NOT remove meaningful engineering characters
    """

    if value is None:
        return ""

    value = str(value).strip().upper()

    # Treat whitespace-only values as blank.
    if not value:
        return ""

    # Collapse accidental multiple spaces.
    value = re.sub(r"\s+", " ", value)

    return value