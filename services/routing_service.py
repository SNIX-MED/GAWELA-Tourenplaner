import math
from urllib.parse import quote

import requests


OSRM_BASE_URL = "https://router.project-osrm.org"
OSRM_ROUTE_MODE = "driving"
OSRM_TIMEOUT_SECONDS = 12


def build_cache_key(stop_a: dict, stop_b: dict) -> str:
    return f"{stop_a.get('id', '')}->{stop_b.get('id', '')}"


def get_travel_segment(
    stop_a: dict,
    stop_b: dict,
    cache: dict | None = None,
    route_mode: str = "car",
    timeout_seconds: int = OSRM_TIMEOUT_SECONDS,
):
    """Lädt die Fahrzeit für ein Segment über OSRM und nutzt dabei einen einfachen Cache.

    Der Cache wird pro Segment unter `stopA_id->stopB_id` abgelegt. Basis-URL und Timeout
    sind als Konstanten gekapselt, damit ein späterer Wechsel des Routers zentral möglich ist.
    """
    cache = cache if isinstance(cache, dict) else {}
    cache_key = build_cache_key(stop_a, stop_b)
    cached_minutes = cache.get(cache_key)
    if cached_minutes is not None:
        try:
            return {
                "cache_key": cache_key,
                "minutes": int(round(float(cached_minutes))),
                "distance_km": None,
                "cached": True,
                "error": "",
            }
        except Exception:
            pass

    lat_a = stop_a.get("lat")
    lon_a = stop_a.get("lon", stop_a.get("lng"))
    lat_b = stop_b.get("lat")
    lon_b = stop_b.get("lon", stop_b.get("lng"))
    if None in (lat_a, lon_a, lat_b, lon_b):
        return {"cache_key": cache_key, "minutes": None, "distance_km": None, "cached": False, "error": "Koordinaten fehlen"}

    profile = OSRM_ROUTE_MODE if route_mode == "car" else route_mode
    coords = f"{float(lon_a)},{float(lat_a)};{float(lon_b)},{float(lat_b)}"
    url = f"{OSRM_BASE_URL}/route/v1/{quote(profile)}/{coords}?overview=false"

    try:
        response = requests.get(url, timeout=timeout_seconds)
        response.raise_for_status()
        data = response.json()
        route = (data.get("routes") or [{}])[0]
        duration_seconds = route.get("duration")
        distance_meters = route.get("distance")
        if duration_seconds is None:
            raise ValueError("Keine Dauer in Routing-Antwort")
        minutes = max(1, int(round(float(duration_seconds) / 60.0)))
        distance_km = None
        if distance_meters is not None:
            distance_km = round(float(distance_meters) / 1000.0, 1)
        cache[cache_key] = minutes
        return {
            "cache_key": cache_key,
            "minutes": minutes,
            "distance_km": distance_km,
            "cached": False,
            "error": "",
        }
    except Exception as exc:
        return {
            "cache_key": cache_key,
            "minutes": None,
            "distance_km": None,
            "cached": False,
            "error": str(exc),
        }


def estimate_distance_km(stop_a: dict, stop_b: dict):
    lat_a = stop_a.get("lat")
    lon_a = stop_a.get("lon", stop_a.get("lng"))
    lat_b = stop_b.get("lat")
    lon_b = stop_b.get("lon", stop_b.get("lng"))
    if None in (lat_a, lon_a, lat_b, lon_b):
        return None

    radius_km = 6371.0
    d_lat = math.radians(float(lat_b) - float(lat_a))
    d_lon = math.radians(float(lon_b) - float(lon_a))
    a = (
        math.sin(d_lat / 2) ** 2
        + math.cos(math.radians(float(lat_a)))
        * math.cos(math.radians(float(lat_b)))
        * math.sin(d_lon / 2) ** 2
    )
    c = 2 * math.atan2(math.sqrt(a), math.sqrt(1 - a))
    return round(radius_km * c, 1)
