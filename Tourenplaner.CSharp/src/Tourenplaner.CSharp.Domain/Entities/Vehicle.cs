namespace Tourenplaner.CSharp.Domain.Entities;

public sealed record Vehicle(
    string Id,
    string Type,
    string Name,
    string LicensePlate,
    int MaxPayloadKg,
    int MaxTrailerLoadKg,
    bool Active,
    string Notes,
    int VolumeM3,
    LoadingArea? LoadingArea,
    DateTime CreatedAt,
    DateTime? UpdatedAt)
    : VehicleBase(Id, Name, LicensePlate, MaxPayloadKg, Active, Notes, VolumeM3, LoadingArea, CreatedAt, UpdatedAt);
