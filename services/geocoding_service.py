from __future__ import annotations

import logging
import threading
import time
from pathlib import Path

from geopy.geocoders import Nominatim

from services.json_storage import InvalidJsonFileError, atomic_write_json, load_json_file


logger = logging.getLogger(__name__)


class GeocodingService:
    def __init__(
        self,
        cache_file: str | Path,
        *,
        user_agent: str,
        timeout: int = 10,
        fair_use_delay_seconds: float = 0.25,
    ):
        self.cache_file = Path(cache_file)
        self.fair_use_delay_seconds = max(0.0, float(fair_use_delay_seconds))
        self._lock = threading.Lock()
        self._dirty = False
        self._cache = self._load_cache()
        self._geolocator = Nominatim(user_agent=user_agent, timeout=timeout)

    def _load_cache(self) -> dict:
        try:
            payload = load_json_file(self.cache_file, default=dict, create_if_missing=False, backup_invalid=True)
        except InvalidJsonFileError:
            logger.warning("Geocode cache is invalid and was backed up: %s", self.cache_file)
            return {}
        except OSError:
            logger.exception("Geocode cache could not be read: %s", self.cache_file)
            return {}
        if not isinstance(payload, dict):
            logger.warning("Geocode cache has unexpected structure: %s", self.cache_file)
            return {}
        return payload

    def lookup(self, address: str):
        key = str(address or "").strip().lower()
        if not key:
            return None

        with self._lock:
            cached_value = self._cache.get(key)
        if isinstance(cached_value, dict) and "lat" in cached_value and "lng" in cached_value:
            try:
                return float(cached_value["lat"]), float(cached_value["lng"])
            except (TypeError, ValueError):
                logger.warning("Ignoring malformed geocode cache entry for %s", key)

        if self.fair_use_delay_seconds:
            time.sleep(self.fair_use_delay_seconds)

        location = self._geolocator.geocode(address)
        if not location:
            return None

        latlng = (float(location.latitude), float(location.longitude))
        with self._lock:
            self._cache[key] = {"lat": latlng[0], "lng": latlng[1]}
            self._dirty = True
        return latlng

    def save_cache(self) -> None:
        with self._lock:
            if not self._dirty:
                return
            snapshot = dict(self._cache)

        atomic_write_json(self.cache_file, snapshot)
        with self._lock:
            self._dirty = False
