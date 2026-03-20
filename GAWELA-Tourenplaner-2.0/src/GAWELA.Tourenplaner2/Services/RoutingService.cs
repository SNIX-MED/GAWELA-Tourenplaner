using System.Net.Http.Json;
using System.Text.Json.Serialization;

namespace GAWELA.Tourenplaner2.Services;

public sealed class TravelSegment
{
    public string CacheKey { get; set; } = string.Empty;
    public int? Minutes { get; set; }
    public double? DistanceKm { get; set; }
    public bool Cached { get; set; }
    public string Error { get; set; } = string.Empty;
}

public static class RoutingService
{
    private static readonly HttpClient HttpClient = new() { Timeout = TimeSpan.FromSeconds(12) };
    private const string BaseUrl = "https://router.project-osrm.org";

    public static string BuildCacheKey(string stopAId, string stopBId) => $"{stopAId}->{stopBId}";

    public static async Task<TravelSegment> GetTravelSegmentAsync(
        string stopAId,
        double? latA,
        double? lonA,
        string stopBId,
        double? latB,
        double? lonB,
        Dictionary<string, int> cache,
        string routeMode = "car")
    {
        var cacheKey = BuildCacheKey(stopAId, stopBId);
        if (cache.TryGetValue(cacheKey, out var cachedMinutes))
        {
            return new TravelSegment { CacheKey = cacheKey, Minutes = cachedMinutes, Cached = true };
        }

        if (!latA.HasValue || !lonA.HasValue || !latB.HasValue || !lonB.HasValue)
        {
            return new TravelSegment { CacheKey = cacheKey, Error = "Koordinaten fehlen" };
        }

        try
        {
            var profile = routeMode == "car" ? "driving" : routeMode;
            var coords = $"{lonA.Value},{latA.Value};{lonB.Value},{latB.Value}";
            var url = $"{BaseUrl}/route/v1/{Uri.EscapeDataString(profile)}/{coords}?overview=false";
            var response = await HttpClient.GetFromJsonAsync<OsrmResponse>(url);

            var route = response?.Routes?.FirstOrDefault();
            if (route?.Duration is null)
            {
                throw new InvalidOperationException("Keine Dauer in Routing-Antwort");
            }

            var minutes = Math.Max(1, (int)Math.Round(route.Duration.Value / 60.0));
            var distanceKm = route.Distance.HasValue ? Math.Round(route.Distance.Value / 1000.0, 1) : null;
            cache[cacheKey] = minutes;

            return new TravelSegment { CacheKey = cacheKey, Minutes = minutes, DistanceKm = distanceKm, Cached = false };
        }
        catch (Exception ex)
        {
            return new TravelSegment { CacheKey = cacheKey, Error = ex.Message, Cached = false };
        }
    }

    public static double? EstimateDistanceKm(double? latA, double? lonA, double? latB, double? lonB)
    {
        if (!latA.HasValue || !lonA.HasValue || !latB.HasValue || !lonB.HasValue)
        {
            return null;
        }

        const double radiusKm = 6371.0;
        var dLat = DegreesToRadians(latB.Value - latA.Value);
        var dLon = DegreesToRadians(lonB.Value - lonA.Value);

        var a = Math.Sin(dLat / 2) * Math.Sin(dLat / 2) +
                Math.Cos(DegreesToRadians(latA.Value)) * Math.Cos(DegreesToRadians(latB.Value)) *
                Math.Sin(dLon / 2) * Math.Sin(dLon / 2);

        var c = 2 * Math.Atan2(Math.Sqrt(a), Math.Sqrt(1 - a));
        return Math.Round(radiusKm * c, 1);
    }

    private static double DegreesToRadians(double degrees) => degrees * Math.PI / 180.0;

    private sealed class OsrmResponse
    {
        [JsonPropertyName("routes")]
        public List<OsrmRoute>? Routes { get; set; }
    }

    private sealed class OsrmRoute
    {
        [JsonPropertyName("duration")]
        public double? Duration { get; set; }

        [JsonPropertyName("distance")]
        public double? Distance { get; set; }
    }
}
