using Tourenplaner.CSharp.Domain.Entities;

namespace Tourenplaner.CSharp.Application.Models;

public sealed record AppSnapshot(
    AppSettings Settings,
    IReadOnlyList<Employee> Employees,
    VehicleCatalog VehicleCatalog,
    IReadOnlyList<Tour> Tours,
    IReadOnlyList<PinRecord> Pins,
    SqlImportWorkspace SqlImportWorkspace);
