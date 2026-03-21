namespace Tourenplaner.CSharp.Domain.Services;

public static class TimeParser
{
    public static int? ToMinutes(string? hhmm)
    {
        if (string.IsNullOrWhiteSpace(hhmm))
        {
            return null;
        }

        var parts = hhmm.Trim().Split(':', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        if (parts.Length != 2 || !int.TryParse(parts[0], out var hours) || !int.TryParse(parts[1], out var minutes))
        {
            return null;
        }

        if (hours is < 0 or > 23 || minutes is < 0 or > 59)
        {
            return null;
        }

        return (hours * 60) + minutes;
    }

    public static string FromMinutes(int? totalMinutes)
    {
        if (totalMinutes is null)
        {
            return string.Empty;
        }

        var minutes = Math.Max(0, totalMinutes.Value);
        var hours = (minutes / 60) % 24;
        var remaining = minutes % 60;
        return $"{hours:00}:{remaining:00}";
    }
}
