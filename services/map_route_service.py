from __future__ import annotations

from urllib.parse import quote

import requests


class RouteServiceError(RuntimeError):
    """Raised when a route path cannot be loaded."""


def fetch_route_path(route_points: list[tuple[float, float]], *, profile: str = "driving", timeout_seconds: int = 15):
    if len(route_points) < 2:
        return []

    coords_str = ";".join(f"{lng},{lat}" for lat, lng in route_points)
    url = (
        f"https://router.project-osrm.org/route/v1/{quote(profile)}/{coords_str}"
        "?overview=full&geometries=geojson"
    )

    try:
        response = requests.get(url, timeout=timeout_seconds)
        response.raise_for_status()
        data = response.json()
    except requests.RequestException as exc:
        raise RouteServiceError(str(exc)) from exc

    if data.get("code") != "Ok":
        raise RouteServiceError("Routing-Dienst hat keine gültige Route geliefert.")

    try:
        geometry = data["routes"][0]["geometry"]["coordinates"]
    except (KeyError, IndexError, TypeError) as exc:
        raise RouteServiceError("Routing-Antwort war unvollständig.") from exc

    return [(point[1], point[0]) for point in geometry]
