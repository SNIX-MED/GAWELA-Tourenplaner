using Tourenplaner.CSharp.Infrastructure.Repositories;
using Xunit;

namespace Tourenplaner.CSharp.Tests;

public sealed class AppPathProviderTests
{
    [Fact]
    public void Constructor_ShouldMapPythonCompatibleFilePaths()
    {
        var provider = new AppPathProvider("/workspace/app-root");

        Assert.Equal("/workspace/app-root", provider.BaseDirectory.Replace('\\', '/'));
        Assert.Equal("/workspace/app-root/data/employees.json", provider.EmployeesPath.Replace('\\', '/'));
        Assert.Equal("/workspace/app-root/data/vehicles.json", provider.VehiclesPath.Replace('\\', '/'));
        Assert.Equal("/workspace/app-root/tours.json", provider.ToursPath.Replace('\\', '/'));
        Assert.Equal("/workspace/app-root/pins.json", provider.PinsPath.Replace('\\', '/'));
        Assert.Equal("/workspace/app-root/settings.json", provider.SettingsPath.Replace('\\', '/'));
    }
}
