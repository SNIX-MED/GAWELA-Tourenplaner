using System.Globalization;
using GAWELA.Tourenplaner2.Models;

namespace GAWELA.Tourenplaner2.Services;

public static class TourStorageService
{
    private static readonly string[] DateFormats = ["yyyy-MM-dd", "dd-MM-yyyy", "dd.MM.yyyy", "dd/MM/yyyy"];

    public static async Task<List<Tour>> LoadToursAsync(string path)
    {
        try
        {
            var tours = await JsonStorageService.LoadJsonFileAsync(path, () => new List<Tour>(), createIfMissing: true);
            return tours.Select(NormalizeTour).ToList();
        }
        catch (InvalidJsonFileException)
        {
            return [];
        }
    }

    public static async Task<List<Tour>> SaveToursAsync(string path, List<Tour> tours)
    {
        var cleaned = tours.Select(NormalizeTour).ToList();
        await JsonStorageService.AtomicWriteJsonAsync(path, cleaned);
        return cleaned;
    }

    public static DateOnly? ParseDate(string? value)
    {
        if (string.IsNullOrWhiteSpace(value)) return null;
        foreach (var fmt in DateFormats)
        {
            if (DateOnly.TryParseExact(value.Trim(), fmt, CultureInfo.InvariantCulture, DateTimeStyles.None, out var parsed))
            {
                return parsed;
            }
        }
        return null;
    }

    public static string FormatDate(DateOnly? value) => value?.ToString("dd-MM-yyyy") ?? string.Empty;

    public static string NormalizeDateString(string? value) => FormatDate(ParseDate(value));

    public static List<Tour> FilterToursByDate(List<Tour> tours, string date) =>
        tours.Where(t => NormalizeDateString(t.Date) == NormalizeDateString(date)).ToList();

    public static List<Tour> FilterToursByRange(List<Tour> tours, string start, string end)
    {
        var startDate = ParseDate(start);
        var endDate = ParseDate(end);
        if (startDate.HasValue && endDate.HasValue && startDate > endDate)
        {
            (startDate, endDate) = (endDate, startDate);
        }

        return tours.Where(t =>
        {
            var current = ParseDate(t.Date);
            if (!current.HasValue) return false;
            if (startDate.HasValue && current.Value < startDate.Value) return false;
            if (endDate.HasValue && current.Value > endDate.Value) return false;
            return true;
        }).ToList();
    }

    public static int TourAssignmentCount(Tour tour)
    {
        var count = tour.EmployeeIds.Where(s => !string.IsNullOrWhiteSpace(s)).Distinct().Count();
        return count == 0 ? 1 : Math.Min(2, count);
    }

    private static Tour NormalizeTour(Tour? tour)
    {
        tour ??= new Tour();
        var employeeIds = tour.EmployeeIds.Where(s => !string.IsNullOrWhiteSpace(s)).Select(s => s.Trim()).Distinct().Take(2).ToList();

        return new Tour
        {
            Id = string.IsNullOrWhiteSpace(tour.Id) ? Guid.NewGuid().ToString() : tour.Id,
            Date = NormalizeDateString(tour.Date),
            Name = tour.Name?.Trim() ?? string.Empty,
            Stops = (tour.Stops ?? []).Select((s, idx) => NormalizeStop(s, idx + 1)).ToList(),
            EmployeeIds = employeeIds,
            StartTime = string.IsNullOrWhiteSpace(tour.StartTime) ? "08:00" : tour.StartTime.Trim(),
            RouteMode = string.IsNullOrWhiteSpace(tour.RouteMode) ? "car" : tour.RouteMode.Trim(),
            VehicleId = string.IsNullOrWhiteSpace(tour.VehicleId) ? null : tour.VehicleId,
            TrailerId = string.IsNullOrWhiteSpace(tour.TrailerId) ? null : tour.TrailerId,
            TravelTimeCache = tour.TravelTimeCache ?? [],
        };
    }

    private static TourStop NormalizeStop(TourStop? stop, int order)
    {
        stop ??= new TourStop();
        return new TourStop
        {
            Id = string.IsNullOrWhiteSpace(stop.Id) ? Guid.NewGuid().ToString() : stop.Id,
            Name = stop.Name?.Trim() ?? string.Empty,
            Address = stop.Address?.Trim() ?? string.Empty,
            Lat = stop.Lat,
            Lon = stop.Lon,
            Order = stop.Order == 0 ? order : stop.Order,
            TimeWindowStart = stop.TimeWindowStart?.Trim() ?? string.Empty,
            TimeWindowEnd = stop.TimeWindowEnd?.Trim() ?? string.Empty,
            ServiceMinutes = Math.Max(0, stop.ServiceMinutes),
            PlannedArrival = stop.PlannedArrival?.Trim() ?? string.Empty,
            PlannedDeparture = stop.PlannedDeparture?.Trim() ?? string.Empty,
        };
    }
}
