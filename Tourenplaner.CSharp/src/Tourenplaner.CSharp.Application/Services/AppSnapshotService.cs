using Tourenplaner.CSharp.Application.Abstractions;
using Tourenplaner.CSharp.Application.Models;

namespace Tourenplaner.CSharp.Application.Services;

public sealed class AppSnapshotService
{
    private readonly IAppDataBootstrapper _bootstrapper;

    public AppSnapshotService(IAppDataBootstrapper bootstrapper)
    {
        _bootstrapper = bootstrapper;
    }

    public async Task<AppSnapshot> LoadAsync(CancellationToken cancellationToken = default)
    {
        var settingsTask = _bootstrapper.LoadSettingsAsync(cancellationToken);
        var employeesTask = _bootstrapper.LoadEmployeesAsync(cancellationToken);
        var vehiclesTask = _bootstrapper.LoadVehicleCatalogAsync(cancellationToken);
        var toursTask = _bootstrapper.LoadToursAsync(cancellationToken);
        var pinsTask = _bootstrapper.LoadPinsAsync(cancellationToken);
        var sqlWorkspaceTask = _bootstrapper.LoadSqlImportWorkspaceAsync(cancellationToken);

        await Task.WhenAll(settingsTask, employeesTask, vehiclesTask, toursTask, pinsTask, sqlWorkspaceTask);

        return new AppSnapshot(
            await settingsTask,
            await employeesTask,
            await vehiclesTask,
            await toursTask,
            await pinsTask,
            await sqlWorkspaceTask);
    }
}
