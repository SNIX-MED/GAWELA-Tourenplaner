namespace Tourenplaner.CSharp.Domain.Entities;

public sealed record OrderPin(double Latitude, double Longitude, string Status, IReadOnlyDictionary<string, string?> Data);
