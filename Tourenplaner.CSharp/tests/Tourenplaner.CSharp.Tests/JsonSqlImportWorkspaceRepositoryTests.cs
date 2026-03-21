using Tourenplaner.CSharp.Infrastructure.Json;
using Tourenplaner.CSharp.Infrastructure.Repositories;
using Xunit;

namespace Tourenplaner.CSharp.Tests;

public sealed class JsonSqlImportWorkspaceRepositoryTests
{
    [Fact]
    public async Task LoadAsync_ShouldReadPendingNonMapCacheAndLatestReport()
    {
        var root = Directory.CreateTempSubdirectory();
        try
        {
            var dataDir = Directory.CreateDirectory(Path.Combine(root.FullName, "data"));
            var logsDir = Directory.CreateDirectory(Path.Combine(root.FullName, "logs"));
            await File.WriteAllTextAsync(Path.Combine(dataDir.FullName, "pending_sql_orders.json"), """
[
  {
    "ImportID": "100",
    "Auftragsnummer": "A100",
    "Name": "Musterkunde",
    "Strasse": "Hauptstrasse 1",
    "PLZ": "8000",
    "Ort": "Zürich",
    "Status": "offen",
    "Lieferart": "Standard"
  }
]
""");
            await File.WriteAllTextAsync(Path.Combine(dataDir.FullName, "non_map_sql_orders.json"), """
[
  {
    "ImportID": "200",
    "Auftragsnummer": "A200",
    "Name": "Abholung",
    "Lieferart": "Abholung",
    "NichtKarteKategorie": "Abholung"
  }
]
""");
            await File.WriteAllTextAsync(Path.Combine(root.FullName, "geocode_cache.json"), """
{
  "a": [47.0, 8.0],
  "b": [48.0, 9.0]
}
""");

            var olderReport = Path.Combine(logsDir.FullName, "sql_import_geocode_failed_20260320-100000.txt");
            var newerReport = Path.Combine(logsDir.FullName, "sql_import_geocode_failed_20260321-100000.txt");
            await File.WriteAllTextAsync(olderReport, "older");
            await File.WriteAllTextAsync(newerReport, "newer");
            File.SetLastWriteTimeUtc(olderReport, new DateTime(2026, 3, 20, 10, 0, 0, DateTimeKind.Utc));
            File.SetLastWriteTimeUtc(newerReport, new DateTime(2026, 3, 21, 10, 0, 0, DateTimeKind.Utc));

            var repository = new JsonSqlImportWorkspaceRepository(
                new JsonFileStore(),
                Path.Combine(dataDir.FullName, "pending_sql_orders.json"),
                Path.Combine(dataDir.FullName, "non_map_sql_orders.json"),
                Path.Combine(root.FullName, "geocode_cache.json"),
                logsDir.FullName);

            var result = await repository.LoadAsync();

            Assert.Single(result.PendingOrders);
            Assert.Equal("A100", result.PendingOrders[0].OrderNumber);
            Assert.Equal("Hauptstrasse 1, 8000 Zürich", result.PendingOrders[0].Address);
            Assert.Single(result.NonMapOrders);
            Assert.Equal("Abholung", result.NonMapOrders[0].NonMapCategory);
            Assert.Equal(2, result.GeocodeCacheEntries);
            Assert.Equal(newerReport, result.LatestGeocodeFailureReport);
            Assert.Equal(new DateTimeOffset(2026, 3, 21, 10, 0, 0, TimeSpan.Zero), result.LatestGeocodeFailureAt);
        }
        finally
        {
            root.Delete(recursive: true);
        }
    }
}
