namespace Tourenplaner.CSharp.Domain.Entities;

public sealed record SqlImportWorkspace(
    IReadOnlyList<SqlImportOrder> PendingOrders,
    IReadOnlyList<SqlImportOrder> NonMapOrders,
    int GeocodeCacheEntries,
    string LatestGeocodeFailureReport,
    DateTimeOffset? LatestGeocodeFailureAt)
{
    public static SqlImportWorkspace Empty { get; } = new(
        Array.Empty<SqlImportOrder>(),
        Array.Empty<SqlImportOrder>(),
        0,
        string.Empty,
        null);
}
