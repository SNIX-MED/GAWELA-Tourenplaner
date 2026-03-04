using System.Globalization;

namespace GAWELA.Tourenplaner2.Services;

public static class TimeUtils
{
    private const string Format = "HH:mm";

    public static TimeOnly? ParseTime(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return null;
        }

        return TimeOnly.TryParseExact(value.Trim(), Format, CultureInfo.InvariantCulture, DateTimeStyles.None, out var parsed)
            ? parsed
            : null;
    }

    public static string FormatTime(TimeOnly? value) => value?.ToString(Format) ?? string.Empty;

    public static int? TimeToMinutes(string? value)
    {
        var parsed = ParseTime(value);
        return parsed.HasValue ? parsed.Value.Hour * 60 + parsed.Value.Minute : null;
    }

    public static TimeOnly? MinutesToTime(int? totalMinutes)
    {
        if (totalMinutes is null)
        {
            return null;
        }

        var minutes = Math.Max(0, totalMinutes.Value);
        return new TimeOnly((minutes / 60) % 24, minutes % 60);
    }

    public static (bool IsValid, string Error) ValidateTimeWindow(string? start, string? end)
    {
        var startMinutes = TimeToMinutes(start);
        var endMinutes = TimeToMinutes(end);

        if (!string.IsNullOrWhiteSpace(start) && startMinutes is null)
        {
            return (false, "Zeitfenster Start ist ungültig. Bitte HH:MM verwenden.");
        }

        if (!string.IsNullOrWhiteSpace(end) && endMinutes is null)
        {
            return (false, "Zeitfenster Ende ist ungültig. Bitte HH:MM verwenden.");
        }

        if (startMinutes.HasValue && endMinutes.HasValue && endMinutes.Value < startMinutes.Value)
        {
            return (false, "Zeitfenster Ende muss nach oder gleich Start liegen.");
        }

        return (true, string.Empty);
    }
}
