using Tourenplaner.CSharp.Application.Abstractions;
using Tourenplaner.CSharp.Domain.Entities;
using Tourenplaner.CSharp.Infrastructure.Json;

namespace Tourenplaner.CSharp.Infrastructure.Repositories;

public sealed class JsonAppDataBootstrapper : IAppDataBootstrapper
{
    private readonly JsonSettingsRepository _settingsRepository;
    private readonly JsonEmployeesRepository _employeesRepository;
    private readonly JsonVehicleCatalogRepository _vehicleCatalogRepository;
    private readonly JsonToursRepository _toursRepository;
    private readonly JsonPinsRepository _pinsRepository;
    private readonly JsonSqlImportWorkspaceRepository _sqlImportWorkspaceRepository;

    public JsonAppDataBootstrapper(IPathProvider paths)
    {
        var fileStore = new JsonFileStore();
        _settingsRepository = new JsonSettingsRepository(fileStore, paths.SettingsPath, Path.Combine(paths.ConfigDirectory, "backups"));
        _employeesRepository = new JsonEmployeesRepository(fileStore, paths.EmployeesPath);
        _vehicleCatalogRepository = new JsonVehicleCatalogRepository(fileStore, paths.VehiclesPath);
        _toursRepository = new JsonToursRepository(fileStore, paths.ToursPath);
        _pinsRepository = new JsonPinsRepository(fileStore, paths.PinsPath);
        _sqlImportWorkspaceRepository = new JsonSqlImportWorkspaceRepository(
            fileStore,
            paths.PendingSqlOrdersPath,
            paths.NonMapSqlOrdersPath,
            paths.GeocodeCachePath,
            paths.LogsDirectory);
    }

    public Task<AppSettings> LoadSettingsAsync(CancellationToken cancellationToken = default) => _settingsRepository.LoadAsync(cancellationToken);
    public Task<IReadOnlyList<Employee>> LoadEmployeesAsync(CancellationToken cancellationToken = default) => _employeesRepository.LoadAsync(cancellationToken);
    public Task<VehicleCatalog> LoadVehicleCatalogAsync(CancellationToken cancellationToken = default) => _vehicleCatalogRepository.LoadAsync(cancellationToken);
    public Task<IReadOnlyList<Tour>> LoadToursAsync(CancellationToken cancellationToken = default) => _toursRepository.LoadAsync(cancellationToken);
    public Task<IReadOnlyList<PinRecord>> LoadPinsAsync(CancellationToken cancellationToken = default) => _pinsRepository.LoadAsync(cancellationToken);
    public Task<SqlImportWorkspace> LoadSqlImportWorkspaceAsync(CancellationToken cancellationToken = default) => _sqlImportWorkspaceRepository.LoadAsync(cancellationToken);
}
