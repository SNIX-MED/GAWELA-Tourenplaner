from __future__ import annotations

import logging
import threading
import time
from pathlib import Path

from geopy.geocoders import Nominatim
from geopy.exc import GeocoderServiceError, GeocoderTimedOut, GeocoderUnavailable

from services.json_storage import InvalidJsonFileError, atomic_write_json, load_json_file


logger = logging.getLogger(__name__)


class GeocodingService:
    def __init__(
        self,
        cache_file: str | Path,
        *,
        user_agent: str,
        timeout: int = 10,
        fair_use_delay_seconds: float = 1.1,
        retry_attempts: int = 3,
    ):
        self.cache_file = Path(cache_file)
        # Nominatim fair use: around max. 1 request per second.
        self.fair_use_delay_seconds = max(1.0, float(fair_use_delay_seconds))
        self.retry_attempts = max(1, int(retry_attempts))
        self._lock = threading.Lock()
        self._dirty = False
        self._last_request_monotonic = 0.0
        self._cache = self._load_cache()
        self._geolocator = Nominatim(user_agent=user_agent, timeout=timeout)

    def _sleep_for_rate_limit(self) -> None:
        if not self.fair_use_delay_seconds:
            return
        now = time.monotonic()
        wait_for = self.fair_use_delay_seconds - (now - self._last_request_monotonic)
        if wait_for > 0:
            time.sleep(wait_for)

    def _geocode_with_retry(self, address: str):
        for attempt in range(1, self.retry_attempts + 1):
            try:
                self._sleep_for_rate_limit()
                location = self._geolocator.geocode(address)
                self._last_request_monotonic = time.monotonic()
                return location
            except (GeocoderTimedOut, GeocoderUnavailable, GeocoderServiceError):
                self._last_request_monotonic = time.monotonic()
                if attempt >= self.retry_attempts:
                    raise
                # Backoff for temporary limits/outages.
                time.sleep(min(4.0, 0.8 * attempt))
        return None

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
        location = self._geocode_with_retry(address)
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
