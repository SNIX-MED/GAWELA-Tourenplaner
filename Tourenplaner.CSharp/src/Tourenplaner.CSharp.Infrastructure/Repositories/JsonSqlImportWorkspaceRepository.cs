using System.Globalization;
using System.Text.Json;
using Tourenplaner.CSharp.Domain.Entities;
using Tourenplaner.CSharp.Infrastructure.Json;

namespace Tourenplaner.CSharp.Infrastructure.Repositories;

public sealed class JsonSqlImportWorkspaceRepository
{
    private readonly JsonFileStore _fileStore;
    private readonly string _pendingOrdersPath;
    private readonly string _nonMapOrdersPath;
    private readonly string _geocodeCachePath;
    private readonly string _logsDirectory;

    public JsonSqlImportWorkspaceRepository(
        JsonFileStore fileStore,
        string pendingOrdersPath,
        string nonMapOrdersPath,
        string geocodeCachePath,
        string logsDirectory)
    {
        _fileStore = fileStore;
        _pendingOrdersPath = pendingOrdersPath;
        _nonMapOrdersPath = nonMapOrdersPath;
        _geocodeCachePath = geocodeCachePath;
        _logsDirectory = logsDirectory;
    }

    public async Task<SqlImportWorkspace> LoadAsync(CancellationToken cancellationToken = default)
    {
        var pendingTask = _fileStore.LoadAsync(_pendingOrdersPath, Array.Empty<JsonElement>(), cancellationToken);
        var nonMapTask = _fileStore.LoadAsync(_nonMapOrdersPath, Array.Empty<JsonElement>(), cancellationToken);
        var geocodeTask = _fileStore.LoadAsync(_geocodeCachePath, default(JsonElement), cancellationToken);

        await Task.WhenAll(pendingTask, nonMapTask, geocodeTask);

        var latestReport = FindLatestGeocodeFailureReport();
        return new SqlImportWorkspace(
            MapOrders(await pendingTask),
            MapOrders(await nonMapTask),
            CountGeocodeCacheEntries(await geocodeTask),
            latestReport?.FullName ?? string.Empty,
            latestReport is null ? null : new DateTimeOffset(latestReport.LastWriteTimeUtc, TimeSpan.Zero));
    }

    private static IReadOnlyList<SqlImportOrder> MapOrders(IEnumerable<JsonElement> rows)
        => rows
            .Where(row => row.ValueKind == JsonValueKind.Object)
            .Select(MapOrder)
            .Where(order => !string.IsNullOrWhiteSpace(order.Identity))
            .OrderBy(order => order.OrderNumber, StringComparer.OrdinalIgnoreCase)
            .ThenBy(order => order.ImportId, StringComparer.OrdinalIgnoreCase)
            .ToList();

    private static SqlImportOrder MapOrder(JsonElement row)
    {
        var order = new SqlImportOrder
        {
            ImportId = ReadString(row, "ImportID"),
            OrderNumber = ReadString(row, "Auftragsnummer"),
            Name = ReadString(row, "Name"),
            Street = FirstNonEmpty(
                ReadString(row, "LieferadresseStrasse"),
                ReadString(row, "Strasse")),
            PostalCode = FirstNonEmpty(
                ReadString(row, "LieferadressePLZ"),
                ReadString(row, "PLZ")),
            City = FirstNonEmpty(
                ReadString(row, "LieferadresseOrt"),
                ReadString(row, "Ort")),
            Country = ReadString(row, "Land"),
            DeliveryType = ReadString(row, "Lieferart"),
            NonMapCategory = FirstNonEmpty(ReadString(row, "NichtKarteKategorie"), ReadString(row, "Lieferart")),
            Status = FirstNonEmpty(ReadString(row, "Status"), "nicht festgelegt"),
            Weight = FirstNonEmpty(ReadString(row, "Auftragsgewicht"), ReadString(row, "Gewicht")),
            Phone = ReadString(row, "Telefon"),
            Email = ReadString(row, "Email"),
            Notes = FirstNonEmpty(ReadString(row, "Notizen"), ReadString(row, "Produkte")),
            Latitude = ReadDouble(row, "lat"),
            Longitude = ReadDouble(row, "lng"),
        };

        return order;
    }

    private static int CountGeocodeCacheEntries(JsonElement cacheRoot)
    {
        if (cacheRoot.ValueKind != JsonValueKind.Object)
        {
            return 0;
        }

        return cacheRoot.EnumerateObject().Count();
    }

    private FileInfo? FindLatestGeocodeFailureReport()
    {
        if (!Directory.Exists(_logsDirectory))
        {
            return null;
        }

        return new DirectoryInfo(_logsDirectory)
            .EnumerateFiles("sql_import_geocode_failed_*.txt", SearchOption.TopDirectoryOnly)
            .OrderByDescending(file => file.LastWriteTimeUtc)
            .FirstOrDefault();
    }

    private static string ReadString(JsonElement row, string propertyName)
    {
        if (!row.TryGetProperty(propertyName, out var value))
        {
            return string.Empty;
        }

        return value.ValueKind switch
        {
            JsonValueKind.String => value.GetString()?.Trim() ?? string.Empty,
            JsonValueKind.Number => value.GetRawText(),
            JsonValueKind.True => bool.TrueString,
            JsonValueKind.False => bool.FalseString,
            _ => string.Empty,
        };
    }

    private static double? ReadDouble(JsonElement row, string propertyName)
    {
        if (!row.TryGetProperty(propertyName, out var value))
        {
            return null;
        }

        if (value.ValueKind == JsonValueKind.Number && value.TryGetDouble(out var parsed))
        {
            return parsed;
        }

        if (value.ValueKind == JsonValueKind.String
            && double.TryParse(value.GetString(), NumberStyles.Float, CultureInfo.InvariantCulture, out parsed))
        {
            return parsed;
        }

        return null;
    }

    private static string FirstNonEmpty(params string[] values)
        => values.FirstOrDefault(value => !string.IsNullOrWhiteSpace(value)) ?? string.Empty;
}
