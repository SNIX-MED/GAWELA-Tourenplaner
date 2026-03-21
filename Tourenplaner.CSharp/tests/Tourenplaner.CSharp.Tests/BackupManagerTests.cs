using Tourenplaner.CSharp.Domain.Enums;
using Tourenplaner.CSharp.Infrastructure.Backups;
using Xunit;

namespace Tourenplaner.CSharp.Tests;

public sealed class BackupManagerTests
{
    [Fact]
    public async Task CreateBackupAndRestore_ShouldRoundTripSelectedGroups()
    {
        var root = Directory.CreateTempSubdirectory();
        var configDir = Directory.CreateDirectory(Path.Combine(root.FullName, "config"));
        var dataDir = Directory.CreateDirectory(Path.Combine(root.FullName, "data"));
        var backupDir = Directory.CreateDirectory(Path.Combine(root.FullName, "backups"));
        var restoreDataDir = Directory.CreateDirectory(Path.Combine(root.FullName, "restore-data"));
        var restoreConfigDir = Directory.CreateDirectory(Path.Combine(root.FullName, "restore-config"));

        try
        {
            await File.WriteAllTextAsync(Path.Combine(configDir.FullName, "settings.json"), "{\"appearance_mode\":\"Dark\"}");
            await File.WriteAllTextAsync(Path.Combine(configDir.FullName, "pins.json"), "[{\"id\":1}]");
            await File.WriteAllTextAsync(Path.Combine(configDir.FullName, "tours.json"), "[{\"id\":2}]");
            await File.WriteAllTextAsync(Path.Combine(dataDir.FullName, "employees.json"), "[{\"id\":\"e1\"}]");
            await File.WriteAllTextAsync(Path.Combine(dataDir.FullName, "vehicles.json"), "{\"vehicles\":[],\"trailers\":[]}");
            await File.WriteAllTextAsync(Path.Combine(dataDir.FullName, "notes.txt"), "misc");

            var manager = new BackupManager("GAWELA", configDir.FullName, dataDir.FullName, backupDir.FullName);
            var backupPath = await manager.CreateBackupAsync(BackupMode.Full);

            Assert.True(File.Exists(backupPath));

            await File.WriteAllTextAsync(Path.Combine(restoreConfigDir.FullName, "settings.json"), "old");
            await manager.RestoreBackupAsync(backupPath, restoreDataDir.FullName, restoreConfigDir.FullName, new[] { "settings", "employees" });

            Assert.Equal("{\"appearance_mode\":\"Dark\"}", await File.ReadAllTextAsync(Path.Combine(restoreConfigDir.FullName, "settings.json")));
            Assert.Equal("[{\"id\":\"e1\"}]", await File.ReadAllTextAsync(Path.Combine(restoreDataDir.FullName, "employees.json")));
            Assert.False(File.Exists(Path.Combine(restoreDataDir.FullName, "vehicles.json")));
        }
        finally
        {
            root.Delete(recursive: true);
        }
    }
}
