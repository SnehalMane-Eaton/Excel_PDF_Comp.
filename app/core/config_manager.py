import json
from dataclasses import dataclass, asdict
from pathlib import Path


@dataclass
class Config:
    fuzzy_threshold: float = 0.88
    excel_sheet: str | None = None
    recursive_input: bool = False


class ConfigManager:
    def __init__(self, path: Path):
        self.path = path

    def load(self) -> Config:
        if not self.path.exists():
            return Config()
        with self.path.open(encoding="utf-8") as handle:
            values = json.load(handle)
        return Config(**{key: value for key, value in values.items() if key in asdict(Config())})

    def save(self, config: Config) -> None:
        self.path.parent.mkdir(parents=True, exist_ok=True)
        with self.path.open("w", encoding="utf-8") as handle:
            json.dump(asdict(config), handle, indent=2)
            handle.write("\n")
