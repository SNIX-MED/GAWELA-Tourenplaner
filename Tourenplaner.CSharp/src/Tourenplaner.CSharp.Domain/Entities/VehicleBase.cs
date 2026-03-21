namespace Tourenplaner.CSharp.Domain.Entities;

public abstract record VehicleBase(
    string Id,
    string Name,
    string LicensePlate,
    int MaxPayloadKg,
    bool Active,
    string Notes,
    int VolumeM3,
    LoadingArea? LoadingArea,
    DateTime CreatedAt,
    DateTime? UpdatedAt);
