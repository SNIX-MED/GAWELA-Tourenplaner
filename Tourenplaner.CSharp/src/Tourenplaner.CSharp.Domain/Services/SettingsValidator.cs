using Tourenplaner.CSharp.Domain.Entities;
using Tourenplaner.CSharp.Domain.Enums;

namespace Tourenplaner.CSharp.Domain.Services;

public static class SettingsValidator
{
    public static AppSettings Validate(AppSettings? settings, string defaultBackupDirectory)
    {
        var source = settings ?? new AppSettings();
        var quickAccess = (source.QuickAccessItems ?? [])
            .Select(item => item?.Trim() ?? string.Empty)
            .Where((item, index) => string.IsNullOrWhiteSpace(item) || !source.QuickAccessItems.Take(index).Contains(item, StringComparer.OrdinalIgnoreCase))
            .ToList();

        while (quickAccess.Count < 4)
        {
            quickAccess.Add(string.Empty);
        }

        quickAccess = quickAccess.Take(4).ToList();

        var retentionDays = Math.Clamp(source.BackupRetentionDays, 1, 365);
        var autoBackupIntervalDays = Math.Clamp(source.AutoBackupIntervalDays, 1, 365);
        var backupDir = string.IsNullOrWhiteSpace(source.BackupDir) ? defaultBackupDirectory : source.BackupDir.Trim();

        var appearanceMode = Enum.IsDefined(source.AppearanceMode) ? source.AppearanceMode : AppearanceMode.System;
        var backupMode = Enum.IsDefined(source.BackupModeDefault) ? source.BackupModeDefault : BackupMode.Full;

        return source with
        {
            SqlDataDir = string.IsNullOrWhiteSpace(source.SqlDataDir) ? new AppSettings().SqlDataDir : source.SqlDataDir.Trim(),
            SqlServerInstance = string.IsNullOrWhiteSpace(source.SqlServerInstance) ? @".\SQLEXPRESS" : source.SqlServerInstance.Trim(),
            SqlDatabase = source.SqlDatabase?.Trim() ?? string.Empty,
            AppearanceMode = appearanceMode,
            QuickAccessItems = quickAccess,
            BackupDir = backupDir,
            BackupModeDefault = backupMode,
            BackupRetentionDays = retentionDays,
            AutoBackupIntervalDays = autoBackupIntervalDays,
            LastBackupIso = source.LastBackupIso?.Trim() ?? string.Empty,
        };
    }
}
