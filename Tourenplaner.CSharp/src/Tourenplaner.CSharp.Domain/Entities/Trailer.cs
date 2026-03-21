namespace Tourenplaner.CSharp.Domain.Entities;

public sealed record Trailer(
    string Id,
    string Name,
    string LicensePlate,
    int MaxPayloadKg,
    bool Active,
    string Notes,
    int VolumeM3,
    LoadingArea? LoadingArea,
    DateTime CreatedAt,
    DateTime? UpdatedAt)
    : VehicleBase(Id, Name, LicensePlate, MaxPayloadKg, Active, Notes, VolumeM3, LoadingArea, CreatedAt, UpdatedAt);
