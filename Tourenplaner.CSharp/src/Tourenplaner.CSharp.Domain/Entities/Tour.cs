namespace Tourenplaner.CSharp.Domain.Entities;

public sealed record Tour
{
    public int Id { get; init; }
    public string Date { get; init; } = string.Empty;
    public string Name { get; init; } = string.Empty;
    public IReadOnlyList<TourStop> Stops { get; init; } = [];
    public IReadOnlyList<string> EmployeeIds { get; init; } = [];
    public string VehicleId { get; init; } = string.Empty;
    public string? TrailerId { get; init; }
    public string StartTime { get; init; } = "08:00";
    public string RouteMode { get; init; } = "car";
    public IReadOnlyDictionary<string, int> TravelTimeCache { get; init; } = new Dictionary<string, int>();
}
