using Tourenplaner.CSharp.Application.Abstractions;

namespace Tourenplaner.CSharp.Infrastructure.Repositories;

public sealed class AppPathProvider : IPathProvider
{
    public AppPathProvider(string baseDirectory)
    {
        BaseDirectory = Path.GetFullPath(baseDirectory);
        DataDirectory = Path.Combine(BaseDirectory, "data");
        ConfigDirectory = BaseDirectory;
    }

    public string BaseDirectory { get; }
    public string DataDirectory { get; }
    public string ConfigDirectory { get; }
    public string LogsDirectory => Path.Combine(ConfigDirectory, "logs");
    public string EmployeesPath => Path.Combine(DataDirectory, "employees.json");
    public string VehiclesPath => Path.Combine(DataDirectory, "vehicles.json");
    public string ToursPath => Path.Combine(BaseDirectory, "tours.json");
    public string PinsPath => Path.Combine(BaseDirectory, "pins.json");
    public string SettingsPath => Path.Combine(BaseDirectory, "settings.json");
    public string PendingSqlOrdersPath => Path.Combine(DataDirectory, "pending_sql_orders.json");
    public string NonMapSqlOrdersPath => Path.Combine(DataDirectory, "non_map_sql_orders.json");
    public string GeocodeCachePath => Path.Combine(ConfigDirectory, "geocode_cache.json");
}
