using Tourenplaner.CSharp.Application.Abstractions;
using Tourenplaner.CSharp.Application.Services;
using Tourenplaner.CSharp.Domain.Entities;
using Xunit;

namespace Tourenplaner.CSharp.Tests;

public sealed class AppSnapshotServiceTests
{
    [Fact]
    public async Task LoadAsync_ShouldAggregateAllDatasets()
    {
        var bootstrapper = new FakeBootstrapper();
        var service = new AppSnapshotService(bootstrapper);

        var snapshot = await service.LoadAsync();

        Assert.Equal("DemoDb", snapshot.Settings.SqlDatabase);
        Assert.Single(snapshot.Employees);
        Assert.Single(snapshot.VehicleCatalog.Vehicles);
        Assert.Single(snapshot.Tours);
        Assert.Single(snapshot.Pins);
        Assert.Single(snapshot.SqlImportWorkspace.PendingOrders);
        Assert.Equal(4, snapshot.SqlImportWorkspace.GeocodeCacheEntries);
    }

    private sealed class FakeBootstrapper : IAppDataBootstrapper
    {
        public Task<AppSettings> LoadSettingsAsync(CancellationToken cancellationToken = default) => Task.FromResult(new AppSettings { SqlDatabase = "DemoDb" });
        public Task<IReadOnlyList<Employee>> LoadEmployeesAsync(CancellationToken cancellationToken = default) => Task.FromResult<IReadOnlyList<Employee>>(new[] { new Employee("1", "Max", "M", "123", true, DateTime.UnixEpoch) });
        public Task<VehicleCatalog> LoadVehicleCatalogAsync(CancellationToken cancellationToken = default) => Task.FromResult(new VehicleCatalog(new[] { new Vehicle("1", "other", "Bus", "TG1", 1000, 1500, true, string.Empty, 0, null, DateTime.UnixEpoch, null) }, Array.Empty<Trailer>()));
        public Task<IReadOnlyList<Tour>> LoadToursAsync(CancellationToken cancellationToken = default) => Task.FromResult<IReadOnlyList<Tour>>(new[] { new Tour { Id = 1, Date = "21-03-2026", Name = "Demo" } });
        public Task<IReadOnlyList<PinRecord>> LoadPinsAsync(CancellationToken cancellationToken = default) => Task.FromResult<IReadOnlyList<PinRecord>>(new[] { new PinRecord { Latitude = 47.0, Longitude = 8.0 } });
        public Task<SqlImportWorkspace> LoadSqlImportWorkspaceAsync(CancellationToken cancellationToken = default)
            => Task.FromResult(new SqlImportWorkspace(
                new[] { new SqlImportOrder { OrderNumber = "A-1", Name = "Pending" } },
                Array.Empty<SqlImportOrder>(),
                4,
                string.Empty,
                null));
    }
}
