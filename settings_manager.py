from __future__ import annotations

from pathlib import Path

from services.json_storage import InvalidJsonFileError, atomic_write_json, load_json_file


class SettingsManager:
    DEFAULT_QUICK_ACCESS = [
        "action:export_route",
        "action:import_folder",
        "",
        "",
    ]
    DEFAULTS = {
        "xml_folder": "",
        "appearance_mode": "System",
        "quick_access_items": list(DEFAULT_QUICK_ACCESS),
        "backups_enabled": False,
        # Default choice: backups live next to the app config to keep them outside mutable data snapshots.
        "backup_dir": "",
        "backup_mode_default": "full",
        "backup_retention_days": 30,
        "auto_backup_enabled": False,
        "auto_backup_interval_days": 7,
        "last_backup_iso": "",
    }

    VALID_APPEARANCE_MODES = {"System", "Light", "Dark"}
    VALID_BACKUP_MODES = {"full", "incremental"}

    def __init__(self, config_dir: Path):
        self.config_dir = Path(config_dir)
        self.config_dir.mkdir(parents=True, exist_ok=True)
        self.settings_file = self.config_dir / "settings.json"

    def default_backup_dir(self) -> Path:
        return self.config_dir / "backups"

    def load(self) -> dict:
        if not self.settings_file.exists():
            settings = dict(self.DEFAULTS)
            settings["backup_dir"] = str(self.default_backup_dir())
            self.save(settings)
            return settings

        try:
            loaded = load_json_file(self.settings_file, default=dict, backup_invalid=True) or {}
        except (InvalidJsonFileError, OSError):
            loaded = {}

        settings = dict(self.DEFAULTS)
        settings.update(loaded if isinstance(loaded, dict) else {})
        settings["backup_dir"] = str(
            Path(settings.get("backup_dir") or self.default_backup_dir()).expanduser()
        )
        return self.validate(settings)

    def save(self, settings: dict) -> dict:
        existing = {}
        if self.settings_file.exists():
            try:
                existing = load_json_file(self.settings_file, default=dict, backup_invalid=True) or {}
            except (InvalidJsonFileError, OSError):
                existing = {}
        merged = dict(existing if isinstance(existing, dict) else {})
        merged.update(settings or {})
        validated = self.validate(merged)
        self.config_dir.mkdir(parents=True, exist_ok=True)
        atomic_write_json(self.settings_file, validated)
        return validated

    def validate(self, settings: dict) -> dict:
        settings = dict(self.DEFAULTS) | dict(settings or {})

        xml_folder = str(settings.get("xml_folder") or "").strip()
        appearance_mode = str(settings.get("appearance_mode") or "System").title()
        raw_quick_access = settings.get("quick_access_items", self.DEFAULT_QUICK_ACCESS)
        quick_access_items = []
        for value in raw_quick_access if isinstance(raw_quick_access, list) else self.DEFAULT_QUICK_ACCESS:
            text = str(value or "").strip()
            if text in quick_access_items and text:
                continue
            quick_access_items.append(text)
        while len(quick_access_items) < len(self.DEFAULT_QUICK_ACCESS):
            quick_access_items.append("")
        quick_access_items = quick_access_items[: len(self.DEFAULT_QUICK_ACCESS)]
        backup_dir = Path(settings.get("backup_dir") or self.default_backup_dir()).expanduser()
        backup_mode_default = str(settings.get("backup_mode_default") or "full").strip().lower()

        try:
            backup_retention_days = int(settings.get("backup_retention_days", 30))
        except Exception as exc:
            raise ValueError("backup_retention_days must be an integer.") from exc
        try:
            auto_backup_interval_days = int(settings.get("auto_backup_interval_days", 7))
        except Exception as exc:
            raise ValueError("auto_backup_interval_days must be an integer.") from exc

        if appearance_mode not in self.VALID_APPEARANCE_MODES:
            appearance_mode = "System"

        if backup_mode_default not in self.VALID_BACKUP_MODES:
            backup_mode_default = "full"

        if backup_retention_days < 1 or backup_retention_days > 365:
            raise ValueError("backup_retention_days must be between 1 and 365.")
        if auto_backup_interval_days < 1 or auto_backup_interval_days > 365:
            raise ValueError("auto_backup_interval_days must be between 1 and 365.")

        try:
            backup_dir.mkdir(parents=True, exist_ok=True)
        except Exception as exc:
            raise ValueError(f"backup_dir is not writable: {backup_dir}") from exc

        return {
            "xml_folder": xml_folder,
            "appearance_mode": appearance_mode,
            "quick_access_items": quick_access_items,
            "backups_enabled": bool(settings.get("backups_enabled", False)),
            "backup_dir": str(backup_dir),
            "backup_mode_default": backup_mode_default,
            "backup_retention_days": backup_retention_days,
            "auto_backup_enabled": bool(settings.get("auto_backup_enabled", False)),
            "auto_backup_interval_days": auto_backup_interval_days,
            "last_backup_iso": str(settings.get("last_backup_iso") or "").strip(),
        }
