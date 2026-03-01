from __future__ import annotations

import fnmatch
import hashlib
import json
import os
import shutil
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path


class BackupManager:
    BACKUP_VERSION = 1
    EXCLUDE_GLOBS = ["*.key", "*token*", "secrets.json"]
    ROOT_MUTABLE_FILES = {
        "settings.json": "config/settings.json",
        "pins.json": "data_root/pins.json",
        "tours.json": "data_root/tours.json",
        "geocode_cache.json": "data_root/geocode_cache.json",
        "config.json": "data_root/config.json",
    }
    RESTORE_LABELS = {
        "orders": "Aufträge & Adressen",
        "tours": "Liefertouren",
        "employees": "Mitarbeiter",
        "vehicles": "Fahrzeuge",
        "settings": "Einstellungen",
        "misc": "Zusatzdaten",
        "other_data": "Weitere Daten",
    }

    def __init__(self, app_name: str, config_dir: Path, data_dir: Path, backup_dir: Path):
        self.app_name = str(app_name or "App").strip() or "App"
        self.config_dir = Path(config_dir)
        self.data_dir = Path(data_dir)
        self.backup_dir = Path(backup_dir)
        self.backup_dir.mkdir(parents=True, exist_ok=True)

    def create_backup(self, mode: str) -> Path:
        mode = str(mode or "full").strip().lower()
        if mode == "incremental":
            return self.create_incremental_backup()
        return self.create_full_backup()

    def create_full_backup(self) -> Path:
        self.backup_dir.mkdir(parents=True, exist_ok=True)
        timestamp = datetime.now().strftime("%Y-%m-%d_%H%M%S")
        target = self.backup_dir / f"{self.app_name}_backup_FULL_{timestamp}.bak"
        file_index, file_map, skipped = self.compute_file_index()
        manifest = self.build_manifest(file_index=file_index, mode="full", base_backup=None, deleted_paths=None)

        with tempfile.NamedTemporaryFile(delete=False, suffix=".bak", dir=str(self.backup_dir)) as tmp:
            tmp_path = Path(tmp.name)

        try:
            with zipfile.ZipFile(tmp_path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
                self._write_snapshot_entries(archive=archive, file_map=file_map)
                archive.writestr("manifest.json", json.dumps(manifest, ensure_ascii=False, indent=2))
                archive.writestr("meta/log.txt", self._build_log(skipped))
            os.replace(tmp_path, target)
        finally:
            if tmp_path.exists():
                try:
                    tmp_path.unlink()
                except Exception:
                    pass
        return target

    def create_incremental_backup(self) -> Path:
        latest = self.find_latest_backup()
        if latest is None:
            return self.create_full_backup()

        try:
            previous_manifest = self._read_manifest(latest)
            previous_index = {
                entry["path"]: entry
                for entry in previous_manifest.get("file_index", [])
                if isinstance(entry, dict) and entry.get("path")
            }
        except Exception:
            return self.create_full_backup()

        file_index, file_map, skipped = self.compute_file_index()
        current_index = {entry["path"]: entry for entry in file_index}

        changed_paths = {
            path for path, entry in current_index.items() if previous_index.get(path, {}).get("sha256") != entry.get("sha256")
        }
        deleted_paths = sorted(path for path in previous_index if path not in current_index)

        with tempfile.NamedTemporaryFile(delete=False, suffix=".bak", dir=str(self.backup_dir)) as tmp:
            tmp_path = Path(tmp.name)

        try:
            with zipfile.ZipFile(latest, "r") as source_archive, zipfile.ZipFile(
                tmp_path, "w", compression=zipfile.ZIP_DEFLATED
            ) as target_archive:
                self._copy_zip_without_meta(source_archive, target_archive, skip_paths=changed_paths | {"manifest.json", "meta/log.txt"})
                self._write_snapshot_entries(target_archive, file_map, include_paths=changed_paths)
                manifest = self.build_manifest(
                    file_index=file_index,
                    mode="incremental",
                    base_backup=latest.name,
                    deleted_paths=deleted_paths,
                )
                target_archive.writestr("manifest.json", json.dumps(manifest, ensure_ascii=False, indent=2))
                target_archive.writestr("meta/log.txt", self._build_log(skipped))
            os.replace(tmp_path, latest)
        finally:
            if tmp_path.exists():
                try:
                    tmp_path.unlink()
                except Exception:
                    pass
        return latest

    def find_latest_backup(self) -> Path | None:
        backups = sorted(self.backup_dir.glob("*.bak"), key=lambda item: item.stat().st_mtime, reverse=True)
        return backups[0] if backups else None

    def cleanup_old_backups(self, retention_days: int) -> None:
        try:
            retention_days = int(retention_days)
        except Exception:
            return
        cutoff = datetime.now(timezone.utc).timestamp() - (retention_days * 86400)
        for bak_file in self.backup_dir.glob("*.bak"):
            try:
                if bak_file.stat().st_mtime < cutoff:
                    bak_file.unlink()
            except Exception:
                continue

    def build_manifest(self, file_index, mode, base_backup=None, deleted_paths=None) -> dict:
        return {
            "backup_version": self.BACKUP_VERSION,
            "created_at_iso": datetime.now(timezone.utc).isoformat(),
            "backup_type": mode,
            "app_name": self.app_name,
            "source_paths": {
                "config_dir": str(self.config_dir),
                "data_dir": str(self.data_dir),
            },
            "file_index": list(file_index or []),
            "deleted_paths": list(deleted_paths or []),
            "incremental_base": base_backup,
        }

    def compute_file_index(self):
        file_index = []
        file_map = {}
        skipped = []
        for archive_path, fs_path in self.scan_files():
            try:
                stat = fs_path.stat()
                digest = self._sha256(fs_path)
            except PermissionError:
                skipped.append(f"Permission denied: {fs_path}")
                continue
            except Exception as exc:
                skipped.append(f"Skipped {fs_path}: {exc}")
                continue

            file_index.append(
                {
                    "path": archive_path,
                    "size": int(stat.st_size),
                    "mtime": float(stat.st_mtime),
                    "sha256": digest,
                }
            )
            file_map[archive_path] = fs_path

        file_index.sort(key=lambda item: item["path"])
        return file_index, file_map, skipped

    def scan_files(self):
        # Back up mutable root-level app files as well because pins/tours live next to main.py.
        for file_name, archive_path in self.ROOT_MUTABLE_FILES.items():
            source_path = self.config_dir / file_name
            if source_path.exists() and self._is_included(source_path.name):
                yield archive_path, source_path

        if self.data_dir.exists():
            for path in sorted(self.data_dir.rglob("*")):
                if not path.is_file():
                    continue
                relative = path.relative_to(self.data_dir).as_posix()
                if not self._is_included(relative):
                    continue
                yield f"data/{relative}", path

    def restore_backup(self, bak_path: Path, target_data_dir: Path, target_config_dir: Path, selected_groups=None) -> None:
        bak_path = Path(bak_path)
        target_data_dir = Path(target_data_dir)
        target_config_dir = Path(target_config_dir)
        target_data_dir.mkdir(parents=True, exist_ok=True)
        target_config_dir.mkdir(parents=True, exist_ok=True)
        allowed_groups = None if not selected_groups or "all" in selected_groups else set(selected_groups)

        with zipfile.ZipFile(bak_path, "r") as archive:
            manifest = self._read_manifest(bak_path)
            for member in archive.namelist():
                if member in {"manifest.json", "meta/log.txt"} or member.endswith("/"):
                    continue
                if allowed_groups is not None and self._classify_archive_member(member) not in allowed_groups:
                    continue
                if member == "config/settings.json":
                    destination = target_config_dir / "settings.json"
                elif member.startswith("data_root/"):
                    destination = target_config_dir / Path(member).relative_to("data_root")
                elif member.startswith("data/"):
                    destination = target_data_dir / Path(member).relative_to("data")
                else:
                    continue
                destination.parent.mkdir(parents=True, exist_ok=True)
                with archive.open(member, "r") as source, destination.open("wb") as target:
                    shutil.copyfileobj(source, target)

            for deleted in manifest.get("deleted_paths", []) or []:
                if allowed_groups is not None and self._classify_archive_member(str(deleted)) not in allowed_groups:
                    continue
                if str(deleted).startswith("data_root/"):
                    destination = target_config_dir / Path(deleted).relative_to("data_root")
                elif str(deleted).startswith("data/"):
                    destination = target_data_dir / Path(deleted).relative_to("data")
                else:
                    continue
                if destination.exists():
                    try:
                        destination.unlink()
                    except Exception:
                        pass

    def _copy_zip_without_meta(self, source_archive, target_archive, skip_paths):
        skipped = set(skip_paths or set())
        copied = set()
        for info in source_archive.infolist():
            if info.filename in skipped or info.filename in copied:
                continue
            copied.add(info.filename)
            target_archive.writestr(info, source_archive.read(info.filename))

    def _write_snapshot_entries(self, archive, file_map, include_paths=None):
        include = set(include_paths or file_map.keys())
        for archive_path in sorted(include):
            source_path = file_map.get(archive_path)
            if source_path is None or not source_path.exists():
                continue
            archive.write(source_path, arcname=archive_path)

    def _read_manifest(self, bak_path: Path) -> dict:
        with zipfile.ZipFile(bak_path, "r") as archive:
            with archive.open("manifest.json", "r") as handle:
                return json.load(handle)

    def _build_log(self, skipped_messages):
        lines = [
            f"Backup created at {datetime.now(timezone.utc).isoformat()}",
        ]
        for message in skipped_messages or []:
            lines.append(message)
        return "\n".join(lines) + "\n"

    def _is_included(self, relative_path: str) -> bool:
        name = str(relative_path or "")
        return not any(fnmatch.fnmatch(name, pattern) for pattern in self.EXCLUDE_GLOBS)

    def _classify_archive_member(self, member: str) -> str:
        path = str(member or "").replace("\\", "/")
        if path == "config/settings.json":
            return "settings"
        if path == "data_root/pins.json":
            return "orders"
        if path == "data_root/tours.json":
            return "tours"
        if path == "data_root/geocode_cache.json" or path == "data_root/config.json":
            return "misc"
        if path == "data/employees.json":
            return "employees"
        if path == "data/vehicles.json":
            return "vehicles"
        if path.startswith("data/"):
            return "other_data"
        return "misc"

    def _sha256(self, path: Path) -> str:
        digest = hashlib.sha256()
        with path.open("rb") as handle:
            while True:
                chunk = handle.read(1024 * 1024)
                if not chunk:
                    break
                digest.update(chunk)
        return digest.hexdigest()
