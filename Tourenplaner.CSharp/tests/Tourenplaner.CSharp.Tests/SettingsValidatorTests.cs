using Tourenplaner.CSharp.Domain.Entities;
using Tourenplaner.CSharp.Domain.Enums;
using Tourenplaner.CSharp.Domain.Services;
using Xunit;

namespace Tourenplaner.CSharp.Tests;

public sealed class SettingsValidatorTests
{
    [Fact]
    public void Validate_ShouldClampRangesAndDeduplicateQuickAccess()
    {
        var settings = new AppSettings
        {
            BackupRetentionDays = 999,
            AutoBackupIntervalDays = 0,
            BackupDir = "",
            QuickAccessItems = new[] { "page:map", "page:map", "action:export_route", "" },
            BackupModeDefault = BackupMode.Incremental,
        };

        var result = SettingsValidator.Validate(settings, "C:/Backups");

        Assert.Equal(365, result.BackupRetentionDays);
        Assert.Equal(1, result.AutoBackupIntervalDays);
        Assert.Equal("C:/Backups", result.BackupDir);
        Assert.Equal(new[] { "page:map", "action:export_route", "", "" }, result.QuickAccessItems);
        Assert.Equal(BackupMode.Incremental, result.BackupModeDefault);
    }
}
