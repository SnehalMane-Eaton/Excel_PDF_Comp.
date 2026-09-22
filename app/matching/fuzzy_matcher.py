from difflib import SequenceMatcher

from app.matching.normalizer import normalize_part_number


def best_match(expected: str, actual: list[str]) -> tuple[str | None, float]:
    target = normalize_part_number(expected)
    candidates = [
        (
            value,
            SequenceMatcher(
                None,
                target,
                normalize_part_number(value),
            ).ratio(),
        )
        for value in actual
    ]
    return max(candidates, key=lambda item: item[1], default=(None, 0.0))
