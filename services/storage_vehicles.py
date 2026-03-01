import logging
import os
import shutil
from datetime import datetime
from uuid import uuid4

from services.json_storage import InvalidJsonFileError, atomic_write_json, load_json_file


DEFAULT_VEHICLE_DATA = {"vehicles": [], "trailers": []}
VEHICLE_TYPES = {"truck", "van", "car", "other"}
logger = logging.getLogger(__name__)


def _ensure_parent_dir(path: str):
    parent = os.path.dirname(path)
    if parent:
        os.makedirs(parent, exist_ok=True)


def _timestamp() -> str:
    return datetime.now().replace(microsecond=0).isoformat()


def _write_file(path: str, payload: dict):
    _ensure_parent_dir(path)
    atomic_write_json(path, payload)


def _coerce_int(value, field_name: str) -> int:
    if value in ("", None):
        return 0
    try:
        result = int(value)
    except Exception as exc:
        raise ValueError(f"{field_name} muss eine ganze Zahl sein.") from exc
    if result < 0:
        raise ValueError(f"{field_name} darf nicht negativ sein.")
    return result


def _normalize_text(value) -> str:
    return str(value or "").strip()


def _normalize_dimensions(value) -> dict | None:
    if not isinstance(value, dict):
        return None
    length = _coerce_int(value.get("length_cm", 0), "Ladefläche Länge")
    width = _coerce_int(value.get("width_cm", 0), "Ladefläche Breite")
    height = _coerce_int(value.get("height_cm", 0), "Ladefläche Höhe")
    if not any([length, width, height]):
        return None
    return {
        "length_cm": length,
        "width_cm": width,
        "height_cm": height,
    }


def normalize_vehicle(vehicle: dict) -> dict:
    vehicle = vehicle if isinstance(vehicle, dict) else {}
    created_at = _normalize_text(vehicle.get("created_at")) or _timestamp()
    vehicle_type = _normalize_text(vehicle.get("type")).lower() or "other"
    if vehicle_type not in VEHICLE_TYPES:
        vehicle_type = "other"

    normalized = {
        "id": _normalize_text(vehicle.get("id")) or str(uuid4()),
        "type": vehicle_type,
        "name": _normalize_text(vehicle.get("name")),
        "license_plate": _normalize_text(vehicle.get("license_plate")).upper(),
        "max_payload_kg": _coerce_int(vehicle.get("max_payload_kg", 0), "Nutzlast"),
        "max_trailer_load_kg": _coerce_int(vehicle.get("max_trailer_load_kg", 0), "Anhängelast"),
        "active": bool(vehicle.get("active", True)),
        "notes": _normalize_text(vehicle.get("notes")),
        "volume_m3": _coerce_int(vehicle.get("volume_m3", 0), "Volumen"),
        "loading_area": _normalize_dimensions(vehicle.get("loading_area")),
        "created_at": created_at,
        "updated_at": _normalize_text(vehicle.get("updated_at")) or _timestamp(),
    }
    return normalized


def normalize_trailer(trailer: dict) -> dict:
    trailer = trailer if isinstance(trailer, dict) else {}
    created_at = _normalize_text(trailer.get("created_at")) or _timestamp()
    normalized = {
        "id": _normalize_text(trailer.get("id")) or str(uuid4()),
        "name": _normalize_text(trailer.get("name")),
        "license_plate": _normalize_text(trailer.get("license_plate")).upper(),
        "max_payload_kg": _coerce_int(trailer.get("max_payload_kg", 0), "Nutzlast"),
        "active": bool(trailer.get("active", True)),
        "notes": _normalize_text(trailer.get("notes")),
        "volume_m3": _coerce_int(trailer.get("volume_m3", 0), "Volumen"),
        "loading_area": _normalize_dimensions(trailer.get("loading_area")),
        "created_at": created_at,
        "updated_at": _normalize_text(trailer.get("updated_at")) or _timestamp(),
    }
    return normalized


def _deduplicate(items: list, normalizer, label: str) -> list:
    cleaned = []
    seen_ids = set()
    seen_names = set()
    seen_plates = set()
    for raw in items if isinstance(items, list) else []:
        try:
            item = normalizer(raw)
        except ValueError:
            continue
        if not item.get("name"):
            continue
        item_id = item["id"]
        if item_id in seen_ids:
            item["id"] = str(uuid4())
            item_id = item["id"]
        seen_ids.add(item_id)

        name_key = item["name"].casefold()
        plate_key = item["license_plate"].replace(" ", "").casefold()
        if name_key in seen_names:
            continue
        if plate_key and plate_key in seen_plates:
            continue
        seen_names.add(name_key)
        if plate_key:
            seen_plates.add(plate_key)
        cleaned.append(item)

    cleaned.sort(key=lambda entry: (not entry.get("active", True), entry.get("name", "").lower()))
    return cleaned


def _normalize_payload(payload) -> dict:
    payload = payload if isinstance(payload, dict) else {}
    vehicles = _deduplicate(payload.get("vehicles", []), normalize_vehicle, "Fahrzeug")
    trailers = _deduplicate(payload.get("trailers", []), normalize_trailer, "Anhänger")
    return {"vehicles": vehicles, "trailers": trailers}


def load_vehicles(path: str) -> dict:
    if not os.path.exists(path):
        _write_file(path, DEFAULT_VEHICLE_DATA)
        return dict(DEFAULT_VEHICLE_DATA)

    try:
        raw = load_json_file(path, default=lambda: dict(DEFAULT_VEHICLE_DATA), backup_invalid=True)
    except InvalidJsonFileError:
        backup_path = f"{path}.bak"
        try:
            shutil.copyfile(path, backup_path)
        except OSError:
            logger.exception("Vehicle file backup copy failed: %s", path)
        logger.warning("Vehicle file is invalid and was backed up: %s", path)
        return dict(DEFAULT_VEHICLE_DATA)
    except OSError:
        logger.exception("Vehicle file could not be read: %s", path)
        return dict(DEFAULT_VEHICLE_DATA)

    normalized = _normalize_payload(raw)
    try:
        _write_file(path, normalized)
    except Exception:
        pass
    return normalized


def save_vehicles(path: str, payload: dict) -> dict:
    normalized = _normalize_payload(payload)
    _write_file(path, normalized)
    return normalized


def _assert_unique(items: list, item_id: str, name: str, license_plate: str, label: str):
    name_key = name.casefold()
    plate_key = license_plate.replace(" ", "").casefold()
    for item in items:
        current_id = str(item.get("id"))
        if current_id == item_id:
            continue
        if name_key and name_key == str(item.get("name", "")).casefold():
            raise ValueError(f"{label}-Name existiert bereits.")
        existing_plate_key = str(item.get("license_plate", "")).replace(" ", "").casefold()
        if plate_key and plate_key == existing_plate_key:
            raise ValueError(f"{label}-Kennzeichen existiert bereits.")


def upsert_vehicle(path: str, vehicle: dict) -> dict:
    payload = load_vehicles(path)
    normalized = normalize_vehicle(vehicle)
    if not normalized["name"]:
        raise ValueError("Der Fahrzeugname ist erforderlich.")
    _assert_unique(payload["vehicles"], normalized["id"], normalized["name"], normalized["license_plate"], "Fahrzeug")

    for index, item in enumerate(payload["vehicles"]):
        if str(item.get("id")) == normalized["id"]:
            normalized["created_at"] = item.get("created_at") or normalized["created_at"]
            payload["vehicles"][index] = normalized
            break
    else:
        payload["vehicles"].append(normalized)

    return save_vehicles(path, payload)


def delete_vehicle(path: str, vehicle_id: str) -> dict:
    payload = load_vehicles(path)
    key = _normalize_text(vehicle_id)
    payload["vehicles"] = [item for item in payload["vehicles"] if _normalize_text(item.get("id")) != key]
    return save_vehicles(path, payload)


def upsert_trailer(path: str, trailer: dict) -> dict:
    payload = load_vehicles(path)
    normalized = normalize_trailer(trailer)
    if not normalized["name"]:
        raise ValueError("Der Anhängername ist erforderlich.")
    _assert_unique(payload["trailers"], normalized["id"], normalized["name"], normalized["license_plate"], "Anhänger")

    for index, item in enumerate(payload["trailers"]):
        if str(item.get("id")) == normalized["id"]:
            normalized["created_at"] = item.get("created_at") or normalized["created_at"]
            payload["trailers"][index] = normalized
            break
    else:
        payload["trailers"].append(normalized)

    return save_vehicles(path, payload)


def delete_trailer(path: str, trailer_id: str) -> dict:
    payload = load_vehicles(path)
    key = _normalize_text(trailer_id)
    payload["trailers"] = [item for item in payload["trailers"] if _normalize_text(item.get("id")) != key]
    return save_vehicles(path, payload)
