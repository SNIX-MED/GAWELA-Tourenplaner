namespace Tourenplaner.CSharp.Domain.Entities;

public sealed record SqlImportOrder
{
    public string ImportId { get; init; } = string.Empty;
    public string OrderNumber { get; init; } = string.Empty;
    public string Name { get; init; } = string.Empty;
    public string Street { get; init; } = string.Empty;
    public string PostalCode { get; init; } = string.Empty;
    public string City { get; init; } = string.Empty;
    public string Country { get; init; } = string.Empty;
    public string DeliveryType { get; init; } = string.Empty;
    public string NonMapCategory { get; init; } = string.Empty;
    public string Status { get; init; } = "nicht festgelegt";
    public string Weight { get; init; } = string.Empty;
    public string Phone { get; init; } = string.Empty;
    public string Email { get; init; } = string.Empty;
    public string Notes { get; init; } = string.Empty;
    public double? Latitude { get; init; }
    public double? Longitude { get; init; }

    public string Address
        => string.Join(", ", new[]
        {
            Street.Trim(),
            string.Join(" ", new[] { PostalCode.Trim(), City.Trim() }.Where(value => !string.IsNullOrWhiteSpace(value))),
            Country.Trim(),
        }.Where(value => !string.IsNullOrWhiteSpace(value)));

    public string Identity
        => !string.IsNullOrWhiteSpace(ImportId)
            ? ImportId.Trim()
            : OrderNumber.Trim();
}
