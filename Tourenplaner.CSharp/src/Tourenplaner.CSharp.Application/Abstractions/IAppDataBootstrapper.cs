using Tourenplaner.CSharp.Domain.Entities;

namespace Tourenplaner.CSharp.Application.Abstractions;

public interface IAppDataBootstrapper
{
    Task<AppSettings> LoadSettingsAsync(CancellationToken cancellationToken = default);
    Task<IReadOnlyList<Employee>> LoadEmployeesAsync(CancellationToken cancellationToken = default);
    Task<VehicleCatalog> LoadVehicleCatalogAsync(CancellationToken cancellationToken = default);
    Task<IReadOnlyList<Tour>> LoadToursAsync(CancellationToken cancellationToken = default);
    Task<IReadOnlyList<PinRecord>> LoadPinsAsync(CancellationToken cancellationToken = default);
    Task<SqlImportWorkspace> LoadSqlImportWorkspaceAsync(CancellationToken cancellationToken = default);
}
