from services.time_utils import format_time, minutes_to_time, time_to_minutes


def compute_schedule(stops: list, segment_minutes: list, start_time: str):
    """Berechnet ETA/ETD entlang der Stoppreihenfolge.

    Start ist die Tour-Startzeit. Für jeden Stopp gilt:
    Ankunft = vorige Abfahrt + Fahrzeit. Liegt die Ankunft vor dem Zeitfenster,
    wartet der Plan bis zum Zeitfenster-Start. Liegt sie nach dem Zeitfenster-Ende,
    wird ein Konflikt markiert. Abfahrt = effektive Ankunft + Aufenthaltszeit.

    `segment_minutes` enthält die Fahrzeit vom Depot zum ersten Stopp, dann zwischen
    Stopps und optional zuletzt die Rückfahrt zum Depot. Für ETA/ETD der Stopps werden
    nur die ersten `len(stops)` Segmente benötigt.
    """

    result_stops = []
    segment_details = []
    current_minutes = time_to_minutes(start_time)
    if current_minutes is None:
        current_minutes = 8 * 60

    total_travel_minutes = 0
    total_service_minutes = 0
    total_wait_minutes = 0
    schedule_blocked = False

    for index, stop in enumerate(stops):
        segment_value = segment_minutes[index] if index < len(segment_minutes) else None
        if segment_value is not None:
            total_travel_minutes += segment_value

        arrival_minutes = None if schedule_blocked or segment_value is None else current_minutes + segment_value
        arrival_text = format_time(minutes_to_time(arrival_minutes)) if arrival_minutes is not None else ""

        window_start = time_to_minutes(stop.get("time_window_start"))
        window_end = time_to_minutes(stop.get("time_window_end"))
        wait_minutes = 0
        conflict_text = ""
        conflict = False

        effective_arrival = arrival_minutes
        if effective_arrival is not None and window_start is not None and effective_arrival < window_start:
            wait_minutes = window_start - effective_arrival
            effective_arrival = window_start
            total_wait_minutes += wait_minutes

        if effective_arrival is not None and window_end is not None and effective_arrival > window_end:
            conflict = True
            conflict_text = f"Ankunft {format_time(minutes_to_time(effective_arrival))} > Fenster Ende {stop.get('time_window_end')}"

        service_minutes = int(stop.get("service_minutes") or 0)
        departure_minutes = None if effective_arrival is None else effective_arrival + service_minutes
        if departure_minutes is not None:
            total_service_minutes += service_minutes

        if effective_arrival is None:
            schedule_blocked = True
        else:
            current_minutes = departure_minutes

        updated = dict(stop)
        updated["planned_arrival"] = arrival_text
        updated["planned_departure"] = format_time(minutes_to_time(departure_minutes)) if departure_minutes is not None else ""
        updated["schedule_conflict"] = conflict
        updated["schedule_conflict_text"] = conflict_text
        updated["wait_minutes"] = wait_minutes
        result_stops.append(updated)

        segment_details.append(
            {
                "index": index,
                "travel_minutes": segment_value,
                "arrival": updated["planned_arrival"],
                "departure": updated["planned_departure"],
                "wait_minutes": wait_minutes,
            }
        )

    final_segment = segment_minutes[len(stops)] if len(segment_minutes) > len(stops) else None
    if final_segment is not None:
        total_travel_minutes += final_segment
        end_minutes = None if schedule_blocked else current_minutes + final_segment
    else:
        end_minutes = None if schedule_blocked else current_minutes

    return {
        "stops": result_stops,
        "total_travel_minutes": total_travel_minutes,
        "total_service_minutes": total_service_minutes,
        "total_wait_minutes": total_wait_minutes,
        "end_time": format_time(minutes_to_time(end_minutes)) if end_minutes is not None else "",
        "segment_details": segment_details,
        "has_conflicts": any(bool(stop.get("schedule_conflict")) for stop in result_stops),
    }
