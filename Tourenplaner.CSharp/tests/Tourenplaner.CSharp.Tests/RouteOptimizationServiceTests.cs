using Tourenplaner.CSharp.Domain.Services;
using Tourenplaner.CSharp.Domain.ValueObjects;
using Xunit;

namespace Tourenplaner.CSharp.Tests;

public sealed class RouteOptimizationServiceTests
{
    [Fact]
    public void Optimize_ShouldPreferCheaperSequenceWhenCacheIsPresent()
    {
        var start = new RouteNode("depot_start", "Depot", 47.0, 8.0, string.Empty, string.Empty, 0);
        var end = new RouteNode("depot_end", "Depot", 47.0, 8.0, string.Empty, string.Empty, 0);
        var stops = new List<RouteNode>
        {
            new("A", "A", 47.1, 8.1, string.Empty, string.Empty, 0),
            new("B", "B", 47.2, 8.2, string.Empty, string.Empty, 0),
            new("C", "C", 47.3, 8.3, string.Empty, string.Empty, 0),
        };

        var cache = new Dictionary<string, int>
        {
            ["depot_start->A"] = 5,
            ["A->B"] = 5,
            ["B->C"] = 5,
            ["C->depot_end"] = 5,
            ["depot_start->C"] = 25,
            ["C->B"] = 25,
            ["B->A"] = 25,
            ["A->depot_end"] = 25,
        };

        var result = RouteOptimizationService.Optimize(start, stops, end, "08:00", cache);

        Assert.Equal(new[] { "A", "B", "C" }, result.Stops.Select(stop => stop.Id).ToArray());
    }
}
