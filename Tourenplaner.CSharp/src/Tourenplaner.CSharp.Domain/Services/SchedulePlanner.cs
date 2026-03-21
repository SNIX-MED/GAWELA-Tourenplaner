using Tourenplaner.CSharp.Domain.ValueObjects;

namespace Tourenplaner.CSharp.Domain.Services;

public static class SchedulePlanner
{
    public static ScheduleResult Compute(IReadOnlyList<ScheduleStop> stops, IReadOnlyList<int?> segmentMinutes, string startTime)
    {
        var resultStops = new List<ScheduledStop>();
        var currentMinutes = TimeParser.ToMinutes(startTime) ?? (8 * 60);
        var totalTravelMinutes = 0;
        var totalServiceMinutes = 0;
        var totalWaitMinutes = 0;
        var blocked = false;

        for (var index = 0; index < stops.Count; index++)
        {
            var stop = stops[index];
            var segmentValue = index < segmentMinutes.Count ? segmentMinutes[index] : null;
            if (segmentValue is not null)
            {
                totalTravelMinutes += segmentValue.Value;
            }

            int? arrivalMinutes = blocked || segmentValue is null ? null : currentMinutes + segmentValue.Value;
            var windowStart = TimeParser.ToMinutes(stop.TimeWindowStart);
            var windowEnd = TimeParser.ToMinutes(stop.TimeWindowEnd);
            var effectiveArrival = arrivalMinutes;
            var waitMinutes = 0;
            var conflict = false;
            var conflictText = string.Empty;

            if (effectiveArrival is not null && windowStart is not null && effectiveArrival < windowStart)
            {
                waitMinutes = windowStart.Value - effectiveArrival.Value;
                totalWaitMinutes += waitMinutes;
                effectiveArrival = windowStart.Value;
            }

            if (effectiveArrival is not null && windowEnd is not null && effectiveArrival > windowEnd)
            {
                conflict = true;
                conflictText = $"Ankunft {TimeParser.FromMinutes(effectiveArrival)} > Fenster Ende {stop.TimeWindowEnd}";
            }

            int? departureMinutes = effectiveArrival is null ? null : effectiveArrival.Value + Math.Max(0, stop.ServiceMinutes);
            if (departureMinutes is not null)
            {
                totalServiceMinutes += Math.Max(0, stop.ServiceMinutes);
                currentMinutes = departureMinutes.Value;
            }
            else
            {
                blocked = true;
            }

            resultStops.Add(new ScheduledStop(
                stop.Name,
                TimeParser.FromMinutes(arrivalMinutes),
                TimeParser.FromMinutes(departureMinutes),
                conflict,
                conflictText,
                waitMinutes,
                stop.TimeWindowStart,
                stop.TimeWindowEnd,
                stop.ServiceMinutes,
                stop.Weight,
                stop.OrderNumber,
                stop.Address));
        }

        var finalSegment = stops.Count < segmentMinutes.Count ? segmentMinutes[stops.Count] : null;
        if (finalSegment is not null)
        {
            totalTravelMinutes += finalSegment.Value;
        }

        var endMinutes = blocked ? null : currentMinutes + (finalSegment ?? 0);
        return new ScheduleResult(
            resultStops,
            totalTravelMinutes,
            totalServiceMinutes,
            totalWaitMinutes,
            TimeParser.FromMinutes(endMinutes),
            resultStops.Any(stop => stop.ScheduleConflict));
    }
}
