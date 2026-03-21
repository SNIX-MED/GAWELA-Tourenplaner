namespace Tourenplaner.CSharp.Domain.Entities;

public sealed record PinRecord
{
    public double Latitude { get; init; }
    public double Longitude { get; init; }
    public string Status { get; init; } = "nicht festgelegt";
    public IReadOnlyDictionary<string, string?> Data { get; init; } = new Dictionary<string, string?>();
}
