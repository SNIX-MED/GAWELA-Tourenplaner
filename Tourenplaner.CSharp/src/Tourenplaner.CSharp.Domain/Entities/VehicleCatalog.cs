namespace Tourenplaner.CSharp.Domain.Entities;

public sealed record VehicleCatalog(IReadOnlyList<Vehicle> Vehicles, IReadOnlyList<Trailer> Trailers)
{
    public static VehicleCatalog Empty { get; } = new([], []);
}
