using GAWELA.Tourenplaner2.Models;

namespace GAWELA.Tourenplaner2.Services;

public static class SchedulePlanner
{
    public static ScheduleResult ComputeSchedule(List<TourStop> stops, List<int?> segmentMinutes, string startTime)
    {
        var resultStops = new List<TourStop>();
        var segmentDetails = new List<SegmentDetail>();
        var currentMinutes = TimeUtils.TimeToMinutes(startTime) ?? 8 * 60;

        var totalTravelMinutes = 0;
        var totalServiceMinutes = 0;
        var totalWaitMinutes = 0;
        var scheduleBlocked = false;

        for (var index = 0; index < stops.Count; index++)
        {
            var stop = stops[index];
            var segment = index < segmentMinutes.Count ? segmentMinutes[index] : null;
            if (segment.HasValue)
            {
                totalTravelMinutes += segment.Value;
            }

            var arrivalMinutes = (!scheduleBlocked && segment.HasValue) ? currentMinutes + segment.Value : null;
            var arrivalText = TimeUtils.FormatTime(TimeUtils.MinutesToTime(arrivalMinutes));
            var windowStart = TimeUtils.TimeToMinutes(stop.TimeWindowStart);
            var windowEnd = TimeUtils.TimeToMinutes(stop.TimeWindowEnd);

            var waitMinutes = 0;
            var conflict = false;
            var conflictText = string.Empty;
            var effectiveArrival = arrivalMinutes;

            if (effectiveArrival.HasValue && windowStart.HasValue && effectiveArrival.Value < windowStart.Value)
            {
                waitMinutes = windowStart.Value - effectiveArrival.Value;
                effectiveArrival = windowStart.Value;
                totalWaitMinutes += waitMinutes;
            }

            if (effectiveArrival.HasValue && windowEnd.HasValue && effectiveArrival.Value > windowEnd.Value)
            {
                conflict = true;
                conflictText = $"Ankunft {TimeUtils.FormatTime(TimeUtils.MinutesToTime(effectiveArrival))} > Fenster Ende {stop.TimeWindowEnd}";
            }

            var serviceMinutes = Math.Max(0, stop.ServiceMinutes);
            var departureMinutes = effectiveArrival.HasValue ? effectiveArrival.Value + serviceMinutes : null;
            if (departureMinutes.HasValue)
            {
                totalServiceMinutes += serviceMinutes;
                currentMinutes = departureMinutes.Value;
            }
            else
            {
                scheduleBlocked = true;
            }

            var updated = new TourStop
            {
                Id = stop.Id,
                Name = stop.Name,
                Address = stop.Address,
                Lat = stop.Lat,
                Lon = stop.Lon,
                Order = stop.Order,
                TimeWindowStart = stop.TimeWindowStart,
                TimeWindowEnd = stop.TimeWindowEnd,
                ServiceMinutes = stop.ServiceMinutes,
                PlannedArrival = arrivalText,
                PlannedDeparture = TimeUtils.FormatTime(TimeUtils.MinutesToTime(departureMinutes)),
                ScheduleConflict = conflict,
                ScheduleConflictText = conflictText,
                WaitMinutes = waitMinutes,
            };

            resultStops.Add(updated);
            segmentDetails.Add(new SegmentDetail
            {
                Index = index,
                TravelMinutes = segment,
                Arrival = updated.PlannedArrival,
                Departure = updated.PlannedDeparture,
                WaitMinutes = waitMinutes,
            });
        }

        var finalSegment = stops.Count < segmentMinutes.Count ? segmentMinutes[stops.Count] : null;
        int? endMinutes;
        if (finalSegment.HasValue)
        {
            totalTravelMinutes += finalSegment.Value;
            endMinutes = scheduleBlocked ? null : currentMinutes + finalSegment.Value;
        }
        else
        {
            endMinutes = scheduleBlocked ? null : currentMinutes;
        }

        return new ScheduleResult
        {
            Stops = resultStops,
            TotalTravelMinutes = totalTravelMinutes,
            TotalServiceMinutes = totalServiceMinutes,
            TotalWaitMinutes = totalWaitMinutes,
            EndTime = TimeUtils.FormatTime(TimeUtils.MinutesToTime(endMinutes)),
            SegmentDetails = segmentDetails,
            HasConflicts = resultStops.Any(s => s.ScheduleConflict),
        };
    }
}
