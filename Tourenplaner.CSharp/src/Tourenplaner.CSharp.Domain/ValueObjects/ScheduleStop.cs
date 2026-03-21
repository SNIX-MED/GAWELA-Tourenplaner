namespace Tourenplaner.CSharp.Domain.ValueObjects;

public sealed record ScheduleStop(
    string Name,
    string TimeWindowStart,
    string TimeWindowEnd,
    int ServiceMinutes,
    string Weight,
    string OrderNumber,
    string Address);
