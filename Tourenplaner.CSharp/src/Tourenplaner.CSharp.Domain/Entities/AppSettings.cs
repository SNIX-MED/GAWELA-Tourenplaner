using Tourenplaner.CSharp.Domain.Enums;

namespace Tourenplaner.CSharp.Domain.Entities;

public sealed record AppSettings
{
    public string SqlDataDir { get; init; } = @"C:\Program Files\Microsoft SQL Server\MSSQL15.SQLEXPRESS\MSSQL\DATA";
    public string SqlServerInstance { get; init; } = @".\SQLEXPRESS";
    public string SqlDatabase { get; init; } = string.Empty;
    public AppearanceMode AppearanceMode { get; init; } = AppearanceMode.System;
    public IReadOnlyList<string> QuickAccessItems { get; init; } = ["action:export_route", "action:import_sql", string.Empty, string.Empty];
    public bool BackupsEnabled { get; init; }
    public string BackupDir { get; init; } = string.Empty;
    public BackupMode BackupModeDefault { get; init; } = BackupMode.Full;
    public int BackupRetentionDays { get; init; } = 30;
    public bool AutoBackupEnabled { get; init; }
    public int AutoBackupIntervalDays { get; init; } = 7;
    public string LastBackupIso { get; init; } = string.Empty;
}
