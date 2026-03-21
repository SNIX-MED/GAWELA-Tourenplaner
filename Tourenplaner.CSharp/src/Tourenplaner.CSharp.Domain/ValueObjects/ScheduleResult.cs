namespace Tourenplaner.CSharp.Domain.ValueObjects;

public sealed record ScheduleResult(
    IReadOnlyList<ScheduledStop> Stops,
    int TotalTravelMinutes,
    int TotalServiceMinutes,
    int TotalWaitMinutes,
    string EndTime,
    bool HasConflicts);
