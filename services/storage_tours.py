import logging
import os
from datetime import datetime
from uuid import uuid4

from services.json_storage import InvalidJsonFileError, atomic_write_json, load_json_file


logger = logging.getLogger(__name__)


def _ensure_parent_dir(path: str):
    parent = os.path.dirname(path)
    if parent:
        os.makedirs(parent, exist_ok=True)


def _write_default_file(path: str, default_data):
    _ensure_parent_dir(path)
    atomic_write_json(path, default_data)


def parse_date(value):
    text = str(value or "").strip()
    if not text:
        return None

    for fmt in ("%Y-%m-%d", "%d-%m-%Y", "%d.%m.%Y", "%d/%m/%Y"):
        try:
            return datetime.strptime(text, fmt).date()
        except Exception:
            pass
    return None


def format_date(value) -> str:
    if value is None:
        return ""
    return value.strftime("%d-%m-%Y")


def normalize_date_string(value) -> str:
    parsed = parse_date(value)
    return format_date(parsed) if parsed else ""


def _stop_id_from_legacy(stop: dict) -> str:
    auftrag = str((stop or {}).get("auftragsnummer") or "").strip()
    if auftrag:
        return f"auftrag:{auftrag}"
    lat = (stop or {}).get("lat")
    lon = (stop or {}).get("lon", (stop or {}).get("lng"))
    if lat is not None and lon is not None:
        try:
            return f"coord:{round(float(lat), 6)}:{round(float(lon), 6)}"
        except Exception:
            pass
    return str(uuid4())


def normalize_stop(stop: dict, order: int = 0) -> dict:
    stop = stop if isinstance(stop, dict) else {}
    lat = stop.get("lat")
    lon = stop.get("lon", stop.get("lng"))

    name = str(stop.get("name") or "").strip()
    address = str(stop.get("address") or "").strip()
    if not name:
        name = str(stop.get("Name") or "").strip()
    if not address:
        street = str(stop.get("Strasse") or "").strip()
        plz = str(stop.get("PLZ") or "").strip()
        ort = str(stop.get("Ort") or "").strip()
        address = ", ".join([value for value in [street, " ".join([plz, ort]).strip()] if value])

    service_minutes = stop.get("service_minutes", 0)
    try:
        service_minutes = int(service_minutes or 0)
    except Exception:
        service_minutes = 0

    normalized = dict(stop)
    normalized["id"] = str(stop.get("id") or _stop_id_from_legacy(stop))
    normalized["name"] = name
    normalized["address"] = address
    normalized["lat"] = lat
    normalized["lon"] = lon
    normalized["lng"] = lon
    normalized["order"] = int(stop.get("order") or order)
    normalized["time_window_start"] = str(stop.get("time_window_start") or "").strip()
    normalized["time_window_end"] = str(stop.get("time_window_end") or "").strip()
    normalized["service_minutes"] = service_minutes
    normalized["planned_arrival"] = str(stop.get("planned_arrival") or "").strip()
    normalized["planned_departure"] = str(stop.get("planned_departure") or "").strip()
    return normalized


def filter_tours_by_date(tours: list, date_str: str) -> list:
    target = normalize_date_string(date_str)
    if not target:
        return list(tours or [])
    return [tour for tour in (tours or []) if normalize_date_string((tour or {}).get("date")) == target]


def filter_tours_by_range(tours: list, start_str: str, end_str: str) -> list:
    start = parse_date(start_str)
    end = parse_date(end_str)
    if start and end and start > end:
        start, end = end, start

    filtered = []
    for tour in tours or []:
        current = parse_date((tour or {}).get("date"))
        if current is None:
            continue
        if start and current < start:
            continue
        if end and current > end:
            continue
        filtered.append(tour)
    return filtered


def normalize_tour(tour: dict) -> dict:
    tour = tour if isinstance(tour, dict) else {}
    employee_ids = []
    for value in tour.get("employee_ids", []) or []:
        text = str(value).strip()
        if text and text not in employee_ids:
            employee_ids.append(text)

    stops = tour.get("stops", [])
    if not isinstance(stops, list):
        stops = []
    normalized_stops = [normalize_stop(item, order=index + 1) for index, item in enumerate(stops)]

    travel_time_cache = tour.get("travel_time_cache", {})
    if not isinstance(travel_time_cache, dict):
        travel_time_cache = {}
    cleaned_cache = {}
    for key, value in travel_time_cache.items():
        try:
            cleaned_cache[str(key)] = int(round(float(value)))
        except Exception:
            continue

    normalized = dict(tour)
    normalized["id"] = tour.get("id")
    normalized["date"] = normalize_date_string(tour.get("date"))
    normalized["name"] = str(tour.get("name") or "").strip()
    normalized["stops"] = normalized_stops
    normalized["employee_ids"] = employee_ids[:2]
    normalized["start_time"] = str(tour.get("start_time") or "08:00").strip() or "08:00"
    normalized["route_mode"] = str(tour.get("route_mode") or "car").strip() or "car"
    # Tour-Datenmodell wurde um Fahrzeug-/Anhänger-Zuordnung erweitert.
    # Legacy-Touren ohne diese Keys bleiben kompatibel und laden mit None-Defaults.
    normalized["vehicle_id"] = str(tour.get("vehicle_id") or "").strip() or None
    normalized["trailer_id"] = str(tour.get("trailer_id") or "").strip() or None
    normalized["travel_time_cache"] = cleaned_cache
    return normalized


def load_tours(path: str) -> list:
    if not os.path.exists(path):
        _write_default_file(path, [])
        return []

    try:
        data = load_json_file(path, default=list, backup_invalid=True)
    except InvalidJsonFileError:
        logger.warning("Tour file is invalid and was backed up: %s", path)
        return []
    except OSError:
        logger.exception("Tour file could not be read: %s", path)
        return []

    if not isinstance(data, list):
        logger.warning("Tour file has unexpected structure: %s", path)
        return []

    tours = [normalize_tour(item) for item in data if isinstance(item, dict)]
    try:
        _write_default_file(path, tours)
    except Exception:
        pass
    return tours


def save_tours(path: str, tours: list):
    cleaned = [normalize_tour(item) for item in tours if isinstance(item, dict)]
    _write_default_file(path, cleaned)
    return cleaned


def tour_assignment_count(tour: dict) -> int:
    employee_ids = [str(value).strip() for value in (tour.get("employee_ids", []) or []) if str(value).strip()]
    if not employee_ids:
        return 1
    return min(2, len(employee_ids))
