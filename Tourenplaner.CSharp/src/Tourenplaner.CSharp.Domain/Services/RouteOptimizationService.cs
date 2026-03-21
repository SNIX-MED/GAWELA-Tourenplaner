using Tourenplaner.CSharp.Domain.ValueObjects;

namespace Tourenplaner.CSharp.Domain.Services;

public static class RouteOptimizationService
{
    private const double DefaultAverageSpeedKmh = 42.0;
    private const int DefaultMinLegMinutes = 3;
    private const double LatePenalty = 8.0;
    private const double WaitWeight = 0.35;

    public static OptimizedRouteResult Optimize(RouteNode startNode, IReadOnlyList<RouteNode> stops, RouteNode endNode, string startTime, IReadOnlyDictionary<string, int>? travelTimeCache = null)
    {
        var regularStops = stops.ToList();
        var startMinutes = TimeParser.ToMinutes(startTime) ?? (8 * 60);
        if (regularStops.Count < 2)
        {
            var metrics = Simulate(startNode, regularStops, endNode, startMinutes, travelTimeCache);
            return new OptimizedRouteResult(regularStops, metrics);
        }

        var seed = NearestNeighbor(startNode, regularStops, startMinutes, travelTimeCache);
        var bestStops = seed.ToList();
        var bestMetrics = Simulate(startNode, bestStops, endNode, startMinutes, travelTimeCache);

        var improved = true;
        var iterations = 0;
        while (improved && iterations < 4)
        {
            iterations++;
            improved = false;
            for (var i = 0; i < bestStops.Count - 1; i++)
            {
                for (var j = i + 1; j < bestStops.Count; j++)
                {
                    var candidate = bestStops.Take(i)
                        .Concat(bestStops.Skip(i).Take(j - i + 1).Reverse())
                        .Concat(bestStops.Skip(j + 1))
                        .ToList();

                    var metrics = Simulate(startNode, candidate, endNode, startMinutes, travelTimeCache);
                    if (metrics.Objective + 1e-9 < bestMetrics.Objective)
                    {
                        bestStops = candidate;
                        bestMetrics = metrics;
                        improved = true;
                    }
                }
            }
        }

        return new OptimizedRouteResult(bestStops, bestMetrics);
    }

    private static IReadOnlyList<RouteNode> NearestNeighbor(RouteNode startNode, IReadOnlyList<RouteNode> stops, int startMinutes, IReadOnlyDictionary<string, int>? cache)
    {
        var remaining = stops.ToList();
        var ordered = new List<RouteNode>();
        var now = startMinutes;
        var previous = startNode;

        while (remaining.Count > 0)
        {
            var bestIndex = 0;
            double? bestScore = null;
            var bestArrival = now;

            for (var index = 0; index < remaining.Count; index++)
            {
                var candidate = remaining[index];
                var travel = EstimateMinutes(previous, candidate, cache);
                var arrival = now + travel;
                var windowStart = TimeParser.ToMinutes(candidate.TimeWindowStart);
                var windowEnd = TimeParser.ToMinutes(candidate.TimeWindowEnd);
                var wait = 0;
                if (windowStart is not null && arrival < windowStart)
                {
                    wait = windowStart.Value - arrival;
                    arrival = windowStart.Value;
                }

                var late = 0;
                if (windowEnd is not null && arrival > windowEnd)
                {
                    late = arrival - windowEnd.Value;
                }

                var score = travel + (wait * WaitWeight) + (late * LatePenalty);
                if (bestScore is null || score < bestScore.Value)
                {
                    bestScore = score;
                    bestIndex = index;
                    bestArrival = arrival;
                }
            }

            var chosen = remaining[bestIndex];
            remaining.RemoveAt(bestIndex);
            ordered.Add(chosen);
            now = bestArrival + Math.Max(0, chosen.ServiceMinutes);
            previous = chosen;
        }

        return ordered;
    }

    private static OptimizedRouteMetrics Simulate(RouteNode startNode, IReadOnlyList<RouteNode> orderedStops, RouteNode endNode, int startMinutes, IReadOnlyDictionary<string, int>? cache)
    {
        var now = startMinutes;
        var driveTotal = 0;
        var waitTotal = 0;
        var lateTotal = 0;
        var previous = startNode;

        foreach (var stop in orderedStops)
        {
            var travel = EstimateMinutes(previous, stop, cache);
            driveTotal += travel;
            now += travel;
            var windowStart = TimeParser.ToMinutes(stop.TimeWindowStart);
            var windowEnd = TimeParser.ToMinutes(stop.TimeWindowEnd);
            if (windowStart is not null && now < windowStart)
            {
                var wait = windowStart.Value - now;
                waitTotal += wait;
                now = windowStart.Value;
            }

            if (windowEnd is not null && now > windowEnd)
            {
                lateTotal += now - windowEnd.Value;
            }

            now += Math.Max(0, stop.ServiceMinutes);
            previous = stop;
        }

        driveTotal += EstimateMinutes(previous, endNode, cache);
        now += EstimateMinutes(previous, endNode, cache);
        var objective = driveTotal + (waitTotal * WaitWeight) + (lateTotal * LatePenalty);
        return new OptimizedRouteMetrics(objective, driveTotal, waitTotal, lateTotal, now);
    }

    private static int EstimateMinutes(RouteNode from, RouteNode to, IReadOnlyDictionary<string, int>? cache)
    {
        var cacheKey = $"{from.Id}->{to.Id}";
        if (cache is not null && cache.TryGetValue(cacheKey, out var cachedMinutes))
        {
            return Math.Max(1, cachedMinutes);
        }

        var distance = EstimateDistanceKm(from, to);
        if (distance <= 0)
        {
            return DefaultMinLegMinutes;
        }

        return Math.Max(DefaultMinLegMinutes, (int)Math.Round((distance / DefaultAverageSpeedKmh) * 60.0));
    }

    private static double EstimateDistanceKm(RouteNode a, RouteNode b)
    {
        if (a.Latitude is null || a.Longitude is null || b.Latitude is null || b.Longitude is null)
        {
            return 0;
        }

        const double radius = 6371.0;
        var dLat = DegreesToRadians(b.Latitude.Value - a.Latitude.Value);
        var dLon = DegreesToRadians(b.Longitude.Value - a.Longitude.Value);
        var h = Math.Pow(Math.Sin(dLat / 2.0), 2)
            + Math.Cos(DegreesToRadians(a.Latitude.Value))
            * Math.Cos(DegreesToRadians(b.Latitude.Value))
            * Math.Pow(Math.Sin(dLon / 2.0), 2);

        return radius * (2.0 * Math.Atan2(Math.Sqrt(h), Math.Sqrt(Math.Max(0.0, 1.0 - h))));
    }

    private static double DegreesToRadians(double value) => value * Math.PI / 180.0;
}
