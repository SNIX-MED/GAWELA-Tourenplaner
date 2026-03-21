namespace Tourenplaner.CSharp.Domain.ValueObjects;

public sealed record ScheduledStop(
    string Name,
    string PlannedArrival,
    string PlannedDeparture,
    bool ScheduleConflict,
    string ScheduleConflictText,
    int WaitMinutes,
    string TimeWindowStart,
    string TimeWindowEnd,
    int ServiceMinutes,
    string Weight,
    string OrderNumber,
    string Address);
