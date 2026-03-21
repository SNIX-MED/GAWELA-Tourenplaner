namespace Tourenplaner.CSharp.Domain.Entities;

public sealed record TourStop
{
    public string Id { get; init; } = string.Empty;
    public string Name { get; init; } = string.Empty;
    public string Address { get; init; } = string.Empty;
    public double? Latitude { get; init; }
    public double? Longitude { get; init; }
    public int Order { get; init; }
    public string OrderNumber { get; init; } = string.Empty;
    public string Weight { get; init; } = string.Empty;
    public string TimeWindowStart { get; init; } = string.Empty;
    public string TimeWindowEnd { get; init; } = string.Empty;
    public int ServiceMinutes { get; init; }
    public string PlannedArrival { get; init; } = string.Empty;
    public string PlannedDeparture { get; init; } = string.Empty;
    public bool ScheduleConflict { get; init; }
    public string ScheduleConflictText { get; init; } = string.Empty;
    public int WaitMinutes { get; init; }
}
