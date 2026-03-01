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


def normalize_employee(employee: dict) -> dict:
    employee = employee if isinstance(employee, dict) else {}
    name = str(employee.get("name") or "").strip()
    short = str(employee.get("short") or "").strip()
    phone = str(employee.get("phone") or "").strip()
    active = bool(employee.get("active", True))
    created_at = str(employee.get("created_at") or "").strip() or datetime.now().replace(microsecond=0).isoformat()
    employee_id = str(employee.get("id") or uuid4())
    return {
        "id": employee_id,
        "name": name,
        "short": short,
        "phone": phone,
        "active": active,
        "created_at": created_at,
    }


def load_employees(path: str) -> list:
    if not os.path.exists(path):
        _write_default_file(path, [])
        return []

    try:
        data = load_json_file(path, default=list, backup_invalid=True)
    except InvalidJsonFileError:
        logger.warning("Employees file is invalid and was backed up: %s", path)
        return []
    except OSError:
        logger.exception("Employees file could not be read: %s", path)
        return []

    if not isinstance(data, list):
        logger.warning("Employees file has unexpected structure: %s", path)
        return []

    employees = []
    seen_ids = set()
    for raw in data:
        employee = normalize_employee(raw)
        if not employee["name"]:
            continue
        if employee["id"] in seen_ids:
            employee["id"] = str(uuid4())
        seen_ids.add(employee["id"])
        employees.append(employee)

    employees.sort(key=lambda item: (not item.get("active", True), item.get("name", "").lower()))

    try:
        _write_default_file(path, employees)
    except Exception:
        pass

    return employees


def save_employees(path: str, employees: list):
    cleaned = []
    seen_ids = set()
    for raw in employees if isinstance(employees, list) else []:
        employee = normalize_employee(raw)
        if not employee["name"]:
            continue
        if employee["id"] in seen_ids:
            employee["id"] = str(uuid4())
        seen_ids.add(employee["id"])
        cleaned.append(employee)

    cleaned.sort(key=lambda item: (not item.get("active", True), item.get("name", "").lower()))
    _write_default_file(path, cleaned)
    return cleaned
