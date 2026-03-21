namespace Tourenplaner.CSharp.Domain.ValueObjects;

public sealed record OptimizedRouteMetrics(double Objective, int DriveMinutes, int WaitMinutes, int LateMinutes, int EndMinutes);

public sealed record OptimizedRouteResult(IReadOnlyList<RouteNode> Stops, OptimizedRouteMetrics Metrics);
