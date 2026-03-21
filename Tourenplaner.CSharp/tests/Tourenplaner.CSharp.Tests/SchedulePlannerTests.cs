using Tourenplaner.CSharp.Domain.Services;
using Tourenplaner.CSharp.Domain.ValueObjects;
using Xunit;

namespace Tourenplaner.CSharp.Tests;

public sealed class SchedulePlannerTests
{
    [Fact]
    public void Compute_ShouldWaitForTimeWindowAndSetDeparture()
    {
        var stops = new List<ScheduleStop>
        {
            new("Stop A", "09:00", "10:00", 15, "150kg", "1001", "A-Strasse 1"),
        };

        var result = SchedulePlanner.Compute(stops, new List<int?> { 30, 30 }, "08:00");

        Assert.Single(result.Stops);
        Assert.Equal("08:30", result.Stops[0].PlannedArrival);
        Assert.Equal("09:15", result.Stops[0].PlannedDeparture);
        Assert.Equal(30, result.TotalWaitMinutes);
        Assert.Equal("09:45", result.EndTime);
    }

    [Fact]
    public void Compute_ShouldFlagLateArrivalConflict()
    {
        var stops = new List<ScheduleStop>
        {
            new("Stop B", "08:00", "08:15", 0, "90kg", "1002", "B-Strasse 2"),
        };

        var result = SchedulePlanner.Compute(stops, new List<int?> { 25 }, "08:00");

        Assert.True(result.HasConflicts);
        Assert.True(result.Stops[0].ScheduleConflict);
        Assert.Contains("Fenster Ende 08:15", result.Stops[0].ScheduleConflictText);
    }
}
