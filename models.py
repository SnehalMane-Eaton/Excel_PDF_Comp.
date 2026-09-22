from dataclasses import dataclass, field
from pathlib import Path
from typing import Any


@dataclass(frozen=True)
class SourceField:
    name: str
    value: str
    source: str = ""


@dataclass
class VerificationResult:
    excel_file: Path
    pdf_file: Path
    matched: list[SourceField] = field(default_factory=list)
    missing: list[SourceField] = field(default_factory=list)
    uncertain: list[tuple[SourceField, str, float]] = field(default_factory=list)
    errors: list[str] = field(default_factory=list)

    @property
    def passed(self) -> bool:
        return not self.missing and not self.uncertain and not self.errors


@dataclass(frozen=True)
class FilePair:
    excel: Path
    pdf: Path
    key: str


@dataclass
class BatchSummary:
    results: list[VerificationResult] = field(default_factory=list)

    @property
    def passed(self) -> int:
        return sum(result.passed for result in self.results)

    @property
    def failed(self) -> int:
        return len(self.results) - self.passed
