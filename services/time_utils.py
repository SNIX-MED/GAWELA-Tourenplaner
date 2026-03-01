from datetime import datetime


TIME_FORMAT = "%H:%M"


def parse_time(value):
    text = str(value or "").strip()
    if not text:
        return None
    try:
        return datetime.strptime(text, TIME_FORMAT).time()
    except Exception:
        return None


def format_time(value) -> str:
    if value is None:
        return ""
    try:
        return value.strftime(TIME_FORMAT)
    except Exception:
        return ""


def time_to_minutes(value):
    parsed = value if hasattr(value, "hour") and hasattr(value, "minute") else parse_time(value)
    if parsed is None:
        return None
    return parsed.hour * 60 + parsed.minute


def minutes_to_time(total_minutes):
    if total_minutes is None:
        return None
    minutes = int(round(float(total_minutes)))
    minutes = max(0, minutes)
    hour = (minutes // 60) % 24
    minute = minutes % 60
    return datetime.strptime(f"{hour:02d}:{minute:02d}", TIME_FORMAT).time()


def is_valid_time_string(value) -> bool:
    return parse_time(value) is not None


def validate_time_window(start_value, end_value):
    start_minutes = time_to_minutes(start_value)
    end_minutes = time_to_minutes(end_value)

    if start_value and start_minutes is None:
        return False, "Zeitfenster Start ist ungültig. Bitte HH:MM verwenden."
    if end_value and end_minutes is None:
        return False, "Zeitfenster Ende ist ungültig. Bitte HH:MM verwenden."
    if start_minutes is not None and end_minutes is not None and end_minutes < start_minutes:
        return False, "Zeitfenster Ende muss nach oder gleich Start liegen."
    return True, ""
