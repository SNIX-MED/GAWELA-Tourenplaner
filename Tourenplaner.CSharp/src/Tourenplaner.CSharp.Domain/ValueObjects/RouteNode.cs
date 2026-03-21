namespace Tourenplaner.CSharp.Domain.ValueObjects;

public sealed record RouteNode(
    string Id,
    string Name,
    double? Latitude,
    double? Longitude,
    string TimeWindowStart,
    string TimeWindowEnd,
    int ServiceMinutes);
