from __future__ import annotations

import copy
import json
import os
import shutil
import tempfile
from datetime import datetime
from pathlib import Path


class JsonStorageError(Exception):
    """Base exception for JSON file access."""


class InvalidJsonFileError(JsonStorageError):
    """Raised when a JSON file exists but cannot be decoded."""


def _clone_default(default):
    value = default() if callable(default) else default
    return copy.deepcopy(value)


def atomic_write_json(path: str | Path, payload, *, indent: int = 2) -> None:
    target = Path(path)
    target.parent.mkdir(parents=True, exist_ok=True)

    tmp_path = None
    try:
        with tempfile.NamedTemporaryFile(
            "w",
            encoding="utf-8",
            delete=False,
            dir=str(target.parent),
            suffix=".tmp",
        ) as handle:
            json.dump(payload, handle, ensure_ascii=False, indent=indent)
            handle.flush()
            os.fsync(handle.fileno())
            tmp_path = Path(handle.name)
        os.replace(tmp_path, target)
    finally:
        if tmp_path and tmp_path.exists():
            try:
                tmp_path.unlink()
            except OSError:
                pass


def backup_corrupt_file(path: str | Path) -> Path | None:
    source = Path(path)
    if not source.exists():
        return None
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_path = source.with_name(f"{source.stem}.corrupt_{timestamp}{source.suffix}")
    shutil.copyfile(source, backup_path)
    return backup_path


def load_json_file(
    path: str | Path,
    *,
    default,
    create_if_missing: bool = False,
    backup_invalid: bool = False,
):
    source = Path(path)
    fallback = _clone_default(default)

    if not source.exists():
        if create_if_missing:
            atomic_write_json(source, fallback)
        return fallback

    try:
        with source.open("r", encoding="utf-8") as handle:
            return json.load(handle)
    except json.JSONDecodeError as exc:
        if backup_invalid:
            backup_corrupt_file(source)
        raise InvalidJsonFileError(f"Invalid JSON in {source}") from exc
    except OSError as exc:
        raise JsonStorageError(f"Could not read JSON file {source}") from exc
