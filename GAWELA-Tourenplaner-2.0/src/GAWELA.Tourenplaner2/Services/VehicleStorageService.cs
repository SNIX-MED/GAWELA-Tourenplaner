using GAWELA.Tourenplaner2.Models;

namespace GAWELA.Tourenplaner2.Services;

public static class VehicleStorageService
{
    private static readonly HashSet<string> VehicleTypes = ["truck", "van", "car", "other"];

    public static async Task<VehiclePayload> LoadVehiclesAsync(string path)
    {
        try
        {
            var payload = await JsonStorageService.LoadJsonFileAsync(path, () => new VehiclePayload(), createIfMissing: true);
            return NormalizePayload(payload);
        }
        catch (InvalidJsonFileException)
        {
            return new VehiclePayload();
        }
    }

    public static async Task<VehiclePayload> SaveVehiclesAsync(string path, VehiclePayload payload)
    {
        var normalized = NormalizePayload(payload);
        await JsonStorageService.AtomicWriteJsonAsync(path, normalized);
        return normalized;
    }

    public static Vehicle NormalizeVehicle(Vehicle? vehicle)
    {
        var v = vehicle ?? new Vehicle();
        var vehicleType = (v.Type ?? "other").Trim().ToLowerInvariant();
        if (!VehicleTypes.Contains(vehicleType)) vehicleType = "other";

        return new Vehicle
        {
            Id = string.IsNullOrWhiteSpace(v.Id) ? Guid.NewGuid().ToString() : v.Id,
            Type = vehicleType,
            Name = v.Name?.Trim() ?? string.Empty,
            LicensePlate = (v.LicensePlate ?? string.Empty).Trim().ToUpperInvariant(),
            MaxPayloadKg = Math.Max(0, v.MaxPayloadKg),
            MaxTrailerLoadKg = Math.Max(0, v.MaxTrailerLoadKg),
            Active = v.Active,
            Notes = v.Notes?.Trim() ?? string.Empty,
            VolumeM3 = Math.Max(0, v.VolumeM3),
            LoadingArea = NormalizeArea(v.LoadingArea),
            CreatedAt = string.IsNullOrWhiteSpace(v.CreatedAt) ? DateTime.Now.ToString("s") : v.CreatedAt,
            UpdatedAt = DateTime.Now.ToString("s"),
        };
    }

    public static Trailer NormalizeTrailer(Trailer? trailer)
    {
        var t = trailer ?? new Trailer();
        return new Trailer
        {
            Id = string.IsNullOrWhiteSpace(t.Id) ? Guid.NewGuid().ToString() : t.Id,
            Name = t.Name?.Trim() ?? string.Empty,
            LicensePlate = (t.LicensePlate ?? string.Empty).Trim().ToUpperInvariant(),
            MaxPayloadKg = Math.Max(0, t.MaxPayloadKg),
            Active = t.Active,
            Notes = t.Notes?.Trim() ?? string.Empty,
            VolumeM3 = Math.Max(0, t.VolumeM3),
            LoadingArea = NormalizeArea(t.LoadingArea),
            CreatedAt = string.IsNullOrWhiteSpace(t.CreatedAt) ? DateTime.Now.ToString("s") : t.CreatedAt,
            UpdatedAt = DateTime.Now.ToString("s"),
        };
    }

    private static VehiclePayload NormalizePayload(VehiclePayload? payload)
    {
        payload ??= new VehiclePayload();
        return new VehiclePayload
        {
            Vehicles = DeduplicateVehicles(payload.Vehicles.Select(NormalizeVehicle).Where(v => !string.IsNullOrWhiteSpace(v.Name)).ToList()),
            Trailers = DeduplicateTrailers(payload.Trailers.Select(NormalizeTrailer).Where(t => !string.IsNullOrWhiteSpace(t.Name)).ToList()),
        };
    }

    private static List<Vehicle> DeduplicateVehicles(List<Vehicle> items)
    {
        var result = new List<Vehicle>();
        var seenId = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var seenName = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var seenPlate = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        foreach (var item in items)
        {
            if (!seenId.Add(item.Id)) item.Id = Guid.NewGuid().ToString();
            if (!seenName.Add(item.Name)) continue;
            var plate = item.LicensePlate.Replace(" ", string.Empty);
            if (!string.IsNullOrWhiteSpace(plate) && !seenPlate.Add(plate)) continue;
            result.Add(item);
        }

        return result.OrderBy(v => !v.Active).ThenBy(v => v.Name, StringComparer.OrdinalIgnoreCase).ToList();
    }

    private static List<Trailer> DeduplicateTrailers(List<Trailer> items)
    {
        var result = new List<Trailer>();
        var seenId = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var seenName = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var seenPlate = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        foreach (var item in items)
        {
            if (!seenId.Add(item.Id)) item.Id = Guid.NewGuid().ToString();
            if (!seenName.Add(item.Name)) continue;
            var plate = item.LicensePlate.Replace(" ", string.Empty);
            if (!string.IsNullOrWhiteSpace(plate) && !seenPlate.Add(plate)) continue;
            result.Add(item);
        }

        return result.OrderBy(v => !v.Active).ThenBy(v => v.Name, StringComparer.OrdinalIgnoreCase).ToList();
    }

    private static LoadingArea? NormalizeArea(LoadingArea? area)
    {
        if (area is null) return null;
        var normalized = new LoadingArea
        {
            LengthCm = Math.Max(0, area.LengthCm),
            WidthCm = Math.Max(0, area.WidthCm),
            HeightCm = Math.Max(0, area.HeightCm),
        };
        return (normalized.LengthCm + normalized.WidthCm + normalized.HeightCm) == 0 ? null : normalized;
    }
}
