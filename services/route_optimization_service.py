from __future__ import annotations

import math

from services.time_utils import time_to_minutes


DEFAULT_AVG_SPEED_KMH = 42.0
DEFAULT_MIN_LEG_MINUTES = 3
DEFAULT_LATE_PENALTY = 8.0
DEFAULT_WAIT_WEIGHT = 0.35


def _coord(node: dict) -> tuple[float, float] | None:
    lat = node.get("lat")
    lon = node.get("lon", node.get("lng"))
    if lat is None or lon is None:
        return None
    try:
        return float(lat), float(lon)
    except Exception:
        return None


def _distance_km(a: dict, b: dict) -> float:
    ca = _coord(a)
    cb = _coord(b)
    if not ca or not cb:
        return 0.0

    lat_a, lon_a = ca
    lat_b, lon_b = cb
    radius_km = 6371.0
    d_lat = math.radians(lat_b - lat_a)
    d_lon = math.radians(lon_b - lon_a)
    h = (
        math.sin(d_lat / 2.0) ** 2
        + math.cos(math.radians(lat_a))
        * math.cos(math.radians(lat_b))
        * math.sin(d_lon / 2.0) ** 2
    )
    return radius_km * (2.0 * math.atan2(math.sqrt(h), math.sqrt(max(0.0, 1.0 - h))))


def _estimate_minutes(a: dict, b: dict, cache: dict[str, int] | None) -> int:
    cache = cache if isinstance(cache, dict) else {}
    cache_key = f"{a.get('id', '')}->{b.get('id', '')}"
    cached = cache.get(cache_key)
    if cached is not None:
        try:
            return max(1, int(round(float(cached))))
        except Exception:
            pass

    dist = _distance_km(a, b)
    if dist <= 0:
        return DEFAULT_MIN_LEG_MINUTES
    estimate = (dist / DEFAULT_AVG_SPEED_KMH) * 60.0
    return max(DEFAULT_MIN_LEG_MINUTES, int(round(estimate)))


def _time_window(stop: dict) -> tuple[int | None, int | None]:
    start = time_to_minutes((stop or {}).get("time_window_start"))
    end = time_to_minutes((stop or {}).get("time_window_end"))
    return start, end


def _service_minutes(stop: dict) -> int:
    try:
        return max(0, int((stop or {}).get("service_minutes") or 0))
    except Exception:
        return 0


def _simulate_route(
    start_node: dict,
    ordered_stops: list[dict],
    end_node: dict,
    *,
    start_time_minutes: int,
    travel_time_cache: dict[str, int] | None = None,
) -> dict:
    now = int(start_time_minutes)
    drive_total = 0
    wait_total = 0
    late_total = 0
    prev = start_node

    for stop in ordered_stops:
        travel = _estimate_minutes(prev, stop, travel_time_cache)
        drive_total += travel
        now += travel

        tw_start, tw_end = _time_window(stop)
        if tw_start is not None and now < tw_start:
            wait = tw_start - now
            wait_total += wait
            now = tw_start
        if tw_end is not None and now > tw_end:
            late_total += now - tw_end

        now += _service_minutes(stop)
        prev = stop

    if end_node:
        travel_back = _estimate_minutes(prev, end_node, travel_time_cache)
        drive_total += travel_back
        now += travel_back

    objective = drive_total + (wait_total * DEFAULT_WAIT_WEIGHT) + (late_total * DEFAULT_LATE_PENALTY)
    return {
        "objective": float(objective),
        "drive_minutes": int(drive_total),
        "wait_minutes": int(wait_total),
        "late_minutes": int(late_total),
        "end_minutes": int(now),
    }


def _nearest_neighbor_seed(
    start_node: dict,
    stops: list[dict],
    *,
    start_time_minutes: int,
    travel_time_cache: dict[str, int] | None,
) -> list[dict]:
    remaining = [dict(stop) for stop in stops]
    ordered: list[dict] = []
    now = int(start_time_minutes)
    prev = start_node

    while remaining:
        best_idx = 0
        best_score = None
        best_arrival_state = None

        for idx, candidate in enumerate(remaining):
            travel = _estimate_minutes(prev, candidate, travel_time_cache)
            arrival = now + travel
            tw_start, tw_end = _time_window(candidate)

            wait = 0
            if tw_start is not None and arrival < tw_start:
                wait = tw_start - arrival
                arrival = tw_start

            late = 0
            if tw_end is not None and arrival > tw_end:
                late = arrival - tw_end

            score = travel + (wait * DEFAULT_WAIT_WEIGHT) + (late * DEFAULT_LATE_PENALTY)
            if best_score is None or score < best_score:
                best_score = score
                best_idx = idx
                best_arrival_state = (arrival, wait, late)

        chosen = remaining.pop(best_idx)
        ordered.append(chosen)
        arrival, _, _ = best_arrival_state if best_arrival_state else (now, 0, 0)
        now = arrival + _service_minutes(chosen)
        prev = chosen

    return ordered


def _improve_with_2opt(
    start_node: dict,
    seed_order: list[dict],
    end_node: dict,
    *,
    start_time_minutes: int,
    travel_time_cache: dict[str, int] | None,
) -> tuple[list[dict], dict]:
    best = [dict(stop) for stop in seed_order]
    best_metrics = _simulate_route(
        start_node,
        best,
        end_node,
        start_time_minutes=start_time_minutes,
        travel_time_cache=travel_time_cache,
    )
    n = len(best)
    if n < 4:
        return best, best_metrics

    improved = True
    iterations = 0
    while improved and iterations < 4:
        improved = False
        iterations += 1
        for i in range(0, n - 1):
            for j in range(i + 1, n):
                candidate = best[:i] + list(reversed(best[i : j + 1])) + best[j + 1 :]
                metrics = _simulate_route(
                    start_node,
                    candidate,
                    end_node,
                    start_time_minutes=start_time_minutes,
                    travel_time_cache=travel_time_cache,
                )
                if metrics["objective"] + 1e-9 < best_metrics["objective"]:
                    best = candidate
                    best_metrics = metrics
                    improved = True
        n = len(best)

    return best, best_metrics


def optimize_stop_order(
    start_node: dict,
    stops: list[dict],
    end_node: dict,
    *,
    start_time: str = "08:00",
    travel_time_cache: dict[str, int] | None = None,
) -> dict:
    """Optimiert die Reihenfolge von Stopps mit Nearest-Neighbor + 2-opt.

    Die Zielfunktion minimiert primär Fahrzeit und berücksichtigt Zeitfenster
    als weiche Restriktionen über Strafkosten.
    """
    regular_stops = [dict(stop) for stop in stops if isinstance(stop, dict)]
    if len(regular_stops) < 2:
        metrics = _simulate_route(
            start_node,
            regular_stops,
            end_node,
            start_time_minutes=time_to_minutes(start_time) or (8 * 60),
            travel_time_cache=travel_time_cache,
        )
        return {"stops": regular_stops, "metrics": metrics}

    start_minutes = time_to_minutes(start_time)
    if start_minutes is None:
        start_minutes = 8 * 60

    baseline_metrics = _simulate_route(
        start_node,
        regular_stops,
        end_node,
        start_time_minutes=start_minutes,
        travel_time_cache=travel_time_cache,
    )
    seed = _nearest_neighbor_seed(
        start_node,
        regular_stops,
        start_time_minutes=start_minutes,
        travel_time_cache=travel_time_cache,
    )
    best_stops, best_metrics = _improve_with_2opt(
        start_node,
        seed,
        end_node,
        start_time_minutes=start_minutes,
        travel_time_cache=travel_time_cache,
    )

    return {
        "stops": best_stops,
        "metrics": best_metrics,
        "baseline_metrics": baseline_metrics,
    }
