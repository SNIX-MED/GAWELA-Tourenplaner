namespace Tourenplaner.CSharp.Domain.Entities;

public sealed record Employee(
    string Id,
    string Name,
    string Short,
    string Phone,
    bool Active,
    DateTime CreatedAt);
