from __future__ import annotations

from pathlib import Path

from services.json_storage import InvalidJsonFileError, atomic_write_json, load_json_file


def load_pins(path: str | Path) -> list[dict]:
    payload = load_json_file(path, default=list, create_if_missing=False, backup_invalid=True)
    if not isinstance(payload, list):
        raise InvalidJsonFileError(f"Pins file has unexpected structure: {path}")
    return [item for item in payload if isinstance(item, dict)]


def save_pins(path: str | Path, pins: list[dict]) -> None:
    atomic_write_json(path, [item for item in pins if isinstance(item, dict)])
