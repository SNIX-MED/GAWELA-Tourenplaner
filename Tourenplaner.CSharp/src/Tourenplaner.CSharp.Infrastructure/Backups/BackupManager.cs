using System.IO.Compression;
using System.Security.Cryptography;
using System.Text.Json;
using System.Text.Json.Serialization;
using Tourenplaner.CSharp.Domain.Enums;

namespace Tourenplaner.CSharp.Infrastructure.Backups;

public sealed class BackupManager
{
    private const int BackupVersion = 1;
    private static readonly string[] ExcludeGlobs = ["*.key", "*token*", "secrets.json"];

    private static readonly IReadOnlyDictionary<string, string> RootMutableFiles = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
    {
        ["settings.json"] = "config/settings.json",
        ["pins.json"] = "data_root/pins.json",
        ["tours.json"] = "data_root/tours.json",
        ["geocode_cache.json"] = "data_root/geocode_cache.json",
        ["config.json"] = "data_root/config.json",
    };

    public static IReadOnlyDictionary<string, string> RestoreLabels { get; } = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
    {
        ["orders"] = "Aufträge & Adressen",
        ["tours"] = "Liefertouren",
        ["employees"] = "Mitarbeiter",
        ["vehicles"] = "Fahrzeuge",
        ["settings"] = "Einstellungen",
        ["misc"] = "Zusatzdaten",
        ["other_data"] = "Weitere Daten",
    };

    private readonly string _appName;
    private readonly DirectoryInfo _configDirectory;
    private readonly DirectoryInfo _dataDirectory;
    private readonly DirectoryInfo _backupDirectory;

    public BackupManager(string appName, string configDirectory, string dataDirectory, string backupDirectory)
    {
        _appName = string.IsNullOrWhiteSpace(appName) ? "App" : appName.Trim();
        _configDirectory = new DirectoryInfo(configDirectory);
        _dataDirectory = new DirectoryInfo(dataDirectory);
        _backupDirectory = new DirectoryInfo(backupDirectory);
        _backupDirectory.Create();
    }

    public Task<string> CreateBackupAsync(BackupMode mode, CancellationToken cancellationToken = default)
        => Task.Run(() => CreateBackup(mode, cancellationToken), cancellationToken);

    public string CreateBackup(BackupMode mode, CancellationToken cancellationToken = default)
        => mode == BackupMode.Incremental ? CreateIncrementalBackup(cancellationToken) : CreateFullBackup(cancellationToken);

    public string? FindLatestBackup()
        => _backupDirectory.Exists
            ? _backupDirectory.EnumerateFiles("*.bak", SearchOption.TopDirectoryOnly).OrderByDescending(file => file.LastWriteTimeUtc).FirstOrDefault()?.FullName
            : null;

    public void CleanupOldBackups(int retentionDays)
    {
        var cutoff = DateTimeOffset.UtcNow.AddDays(-Math.Max(1, retentionDays));
        if (!_backupDirectory.Exists)
        {
            return;
        }

        foreach (var file in _backupDirectory.EnumerateFiles("*.bak", SearchOption.TopDirectoryOnly))
        {
            if (file.LastWriteTimeUtc < cutoff.UtcDateTime)
            {
                file.Delete();
            }
        }
    }

    public Task RestoreBackupAsync(string backupPath, string targetDataDirectory, string targetConfigDirectory, IReadOnlyCollection<string>? selectedGroups, CancellationToken cancellationToken = default)
        => Task.Run(() => RestoreBackup(backupPath, targetDataDirectory, targetConfigDirectory, selectedGroups, cancellationToken), cancellationToken);

    public void RestoreBackup(string backupPath, string targetDataDirectory, string targetConfigDirectory, IReadOnlyCollection<string>? selectedGroups, CancellationToken cancellationToken = default)
    {
        var allowedGroups = selectedGroups is null || selectedGroups.Count == 0 || selectedGroups.Contains("all", StringComparer.OrdinalIgnoreCase)
            ? null
            : new HashSet<string>(selectedGroups, StringComparer.OrdinalIgnoreCase);

        var targetData = new DirectoryInfo(targetDataDirectory);
        var targetConfig = new DirectoryInfo(targetConfigDirectory);
        targetData.Create();
        targetConfig.Create();

        using var archive = ZipFile.OpenRead(backupPath);
        var manifest = ReadManifest(archive);
        foreach (var entry in archive.Entries)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (string.Equals(entry.FullName, "manifest.json", StringComparison.OrdinalIgnoreCase)
                || string.Equals(entry.FullName, "meta/log.txt", StringComparison.OrdinalIgnoreCase)
                || string.IsNullOrWhiteSpace(entry.Name))
            {
                continue;
            }

            if (allowedGroups is not null && !allowedGroups.Contains(ClassifyArchiveMember(entry.FullName)))
            {
                continue;
            }

            var destination = ResolveRestorePath(entry.FullName, targetData.FullName, targetConfig.FullName);
            if (destination is null)
            {
                continue;
            }

            Directory.CreateDirectory(Path.GetDirectoryName(destination)!);
            entry.ExtractToFile(destination, overwrite: true);
        }

        foreach (var deletedPath in manifest.DeletedPaths)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (allowedGroups is not null && !allowedGroups.Contains(ClassifyArchiveMember(deletedPath)))
            {
                continue;
            }

            var destination = ResolveRestorePath(deletedPath, targetData.FullName, targetConfig.FullName);
            if (!string.IsNullOrWhiteSpace(destination) && File.Exists(destination))
            {
                File.Delete(destination);
            }
        }
    }

    private string CreateFullBackup(CancellationToken cancellationToken)
    {
        _backupDirectory.Create();
        var timestamp = DateTime.Now.ToString("yyyy-MM-dd_HHmmss");
        var target = Path.Combine(_backupDirectory.FullName, $"{_appName}_backup_FULL_{timestamp}.bak");
        var snapshot = ComputeFileSnapshot(cancellationToken);
        var manifest = BuildManifest(snapshot.Index, "full", null, []);

        using var archive = ZipFile.Open(target, ZipArchiveMode.Create);
        WriteSnapshotEntries(archive, snapshot.Map, snapshot.Map.Keys, cancellationToken);
        WriteTextEntry(archive, "manifest.json", JsonSerializer.Serialize(manifest, JsonOptions));
        WriteTextEntry(archive, "meta/log.txt", BuildLog(snapshot.Skipped));
        return target;
    }

    private string CreateIncrementalBackup(CancellationToken cancellationToken)
    {
        var latest = FindLatestBackup();
        if (string.IsNullOrWhiteSpace(latest))
        {
            return CreateFullBackup(cancellationToken);
        }

        ManifestModel previousManifest;
        try
        {
            using var latestArchive = ZipFile.OpenRead(latest);
            previousManifest = ReadManifest(latestArchive);
        }
        catch
        {
            return CreateFullBackup(cancellationToken);
        }

        var previousIndex = previousManifest.FileIndex.ToDictionary(entry => entry.Path, StringComparer.OrdinalIgnoreCase);
        var snapshot = ComputeFileSnapshot(cancellationToken);
        var currentIndex = snapshot.Index.ToDictionary(entry => entry.Path, StringComparer.OrdinalIgnoreCase);
        var changedPaths = currentIndex
            .Where(pair => !previousIndex.TryGetValue(pair.Key, out var previous) || !string.Equals(previous.Sha256, pair.Value.Sha256, StringComparison.OrdinalIgnoreCase))
            .Select(pair => pair.Key)
            .ToHashSet(StringComparer.OrdinalIgnoreCase);
        var deletedPaths = previousIndex.Keys.Where(path => !currentIndex.ContainsKey(path)).OrderBy(path => path, StringComparer.OrdinalIgnoreCase).ToList();

        var tempPath = Path.Combine(_backupDirectory.FullName, $"{Guid.NewGuid():N}.bak");
        using (var sourceArchive = ZipFile.OpenRead(latest))
        using (var targetArchive = ZipFile.Open(tempPath, ZipArchiveMode.Create))
        {
            CopyArchiveEntries(sourceArchive, targetArchive, changedPaths, cancellationToken);
            WriteSnapshotEntries(targetArchive, snapshot.Map, changedPaths, cancellationToken);
            var manifest = BuildManifest(snapshot.Index, "incremental", Path.GetFileName(latest), deletedPaths);
            WriteTextEntry(targetArchive, "manifest.json", JsonSerializer.Serialize(manifest, JsonOptions));
            WriteTextEntry(targetArchive, "meta/log.txt", BuildLog(snapshot.Skipped));
        }

        File.Copy(tempPath, latest, overwrite: true);
        File.Delete(tempPath);
        return latest;
    }

    private Snapshot ComputeFileSnapshot(CancellationToken cancellationToken)
    {
        var index = new List<FileIndexEntry>();
        var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        var skipped = new List<string>();

        foreach (var (archivePath, sourcePath) in ScanFiles())
        {
            cancellationToken.ThrowIfCancellationRequested();
            try
            {
                var fileInfo = new FileInfo(sourcePath);
                var hash = ComputeSha256(sourcePath);
                index.Add(new FileIndexEntry(archivePath, fileInfo.Length, (fileInfo.LastWriteTimeUtc - DateTime.UnixEpoch).TotalSeconds, hash));
                map[archivePath] = sourcePath;
            }
            catch (UnauthorizedAccessException)
            {
                skipped.Add($"Permission denied: {sourcePath}");
            }
            catch (Exception exc)
            {
                skipped.Add($"Skipped {sourcePath}: {exc.Message}");
            }
        }

        index.Sort((left, right) => StringComparer.OrdinalIgnoreCase.Compare(left.Path, right.Path));
        return new Snapshot(index, map, skipped);
    }

    private IEnumerable<(string ArchivePath, string SourcePath)> ScanFiles()
    {
        foreach (var pair in RootMutableFiles)
        {
            var sourcePath = Path.Combine(_configDirectory.FullName, pair.Key);
            if (File.Exists(sourcePath) && IsIncluded(Path.GetFileName(sourcePath)))
            {
                yield return (pair.Value, sourcePath);
            }
        }

        if (!_dataDirectory.Exists)
        {
            yield break;
        }

        foreach (var path in _dataDirectory.EnumerateFiles("*", SearchOption.AllDirectories).OrderBy(file => file.FullName, StringComparer.OrdinalIgnoreCase))
        {
            var relative = Path.GetRelativePath(_dataDirectory.FullName, path.FullName).Replace('\\', '/');
            if (IsIncluded(relative))
            {
                yield return ($"data/{relative}", path.FullName);
            }
        }
    }

    private static void CopyArchiveEntries(ZipArchive sourceArchive, ZipArchive targetArchive, ISet<string> changedPaths, CancellationToken cancellationToken)
    {
        var skipped = new HashSet<string>(changedPaths, StringComparer.OrdinalIgnoreCase) { "manifest.json", "meta/log.txt" };
        foreach (var entry in sourceArchive.Entries)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (skipped.Contains(entry.FullName) || string.IsNullOrWhiteSpace(entry.Name))
            {
                continue;
            }

            var targetEntry = targetArchive.CreateEntry(entry.FullName, CompressionLevel.Optimal);
            using var input = entry.Open();
            using var output = targetEntry.Open();
            input.CopyTo(output);
        }
    }

    private static void WriteSnapshotEntries(ZipArchive archive, IReadOnlyDictionary<string, string> fileMap, IEnumerable<string> paths, CancellationToken cancellationToken)
    {
        foreach (var archivePath in paths.OrderBy(path => path, StringComparer.OrdinalIgnoreCase))
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (!fileMap.TryGetValue(archivePath, out var sourcePath) || !File.Exists(sourcePath))
            {
                continue;
            }

            archive.CreateEntryFromFile(sourcePath, archivePath, CompressionLevel.Optimal);
        }
    }

    private static void WriteTextEntry(ZipArchive archive, string path, string content)
    {
        var entry = archive.CreateEntry(path, CompressionLevel.Optimal);
        using var writer = new StreamWriter(entry.Open());
        writer.Write(content);
    }

    private ManifestModel BuildManifest(IReadOnlyList<FileIndexEntry> index, string mode, string? baseBackup, IReadOnlyList<string> deletedPaths)
        => new(
            BackupVersion,
            DateTimeOffset.UtcNow,
            mode,
            _appName,
            new SourcePaths(_configDirectory.FullName, _dataDirectory.FullName),
            index,
            deletedPaths,
            baseBackup);

    private static ManifestModel ReadManifest(ZipArchive archive)
    {
        var entry = archive.GetEntry("manifest.json") ?? throw new InvalidDataException("manifest.json fehlt im Backup.");
        using var reader = new StreamReader(entry.Open());
        return JsonSerializer.Deserialize<ManifestModel>(reader.ReadToEnd(), JsonOptions) ?? throw new InvalidDataException("manifest.json ist ungültig.");
    }

    private static string? ResolveRestorePath(string archiveMember, string targetDataDirectory, string targetConfigDirectory)
    {
        var normalized = archiveMember.Replace('\\', '/');
        if (string.Equals(normalized, "config/settings.json", StringComparison.OrdinalIgnoreCase))
        {
            return Path.Combine(targetConfigDirectory, "settings.json");
        }

        if (normalized.StartsWith("data_root/", StringComparison.OrdinalIgnoreCase))
        {
            return Path.Combine(targetConfigDirectory, normalized["data_root/".Length..].Replace('/', Path.DirectorySeparatorChar));
        }

        if (normalized.StartsWith("data/", StringComparison.OrdinalIgnoreCase))
        {
            return Path.Combine(targetDataDirectory, normalized["data/".Length..].Replace('/', Path.DirectorySeparatorChar));
        }

        return null;
    }

    public static string ClassifyArchiveMember(string member)
    {
        var path = (member ?? string.Empty).Replace('\\', '/');
        if (string.Equals(path, "config/settings.json", StringComparison.OrdinalIgnoreCase)) return "settings";
        if (string.Equals(path, "data_root/pins.json", StringComparison.OrdinalIgnoreCase)) return "orders";
        if (string.Equals(path, "data_root/tours.json", StringComparison.OrdinalIgnoreCase)) return "tours";
        if (string.Equals(path, "data_root/geocode_cache.json", StringComparison.OrdinalIgnoreCase)
            || string.Equals(path, "data_root/config.json", StringComparison.OrdinalIgnoreCase)) return "misc";
        if (string.Equals(path, "data/employees.json", StringComparison.OrdinalIgnoreCase)) return "employees";
        if (string.Equals(path, "data/vehicles.json", StringComparison.OrdinalIgnoreCase)) return "vehicles";
        return path.StartsWith("data/", StringComparison.OrdinalIgnoreCase) ? "other_data" : "misc";
    }

    private static string ComputeSha256(string path)
    {
        using var stream = File.OpenRead(path);
        using var sha = SHA256.Create();
        return Convert.ToHexString(sha.ComputeHash(stream));
    }

    private static bool IsIncluded(string relativePath)
        => !ExcludeGlobs.Any(pattern => MatchesGlob(relativePath, pattern));

    private static bool MatchesGlob(string path, string pattern)
    {
        var regex = "^" + System.Text.RegularExpressions.Regex.Escape(pattern)
            .Replace(@"\*", ".*")
            .Replace(@"\?", ".") + "$";
        return System.Text.RegularExpressions.Regex.IsMatch(path, regex, System.Text.RegularExpressions.RegexOptions.IgnoreCase);
    }

    private static string BuildLog(IReadOnlyList<string> skipped)
    {
        var lines = new List<string> { $"Backup created at {DateTimeOffset.UtcNow:O}" };
        lines.AddRange(skipped);
        return string.Join(Environment.NewLine, lines) + Environment.NewLine;
    }

    private static readonly JsonSerializerOptions JsonOptions = new() { WriteIndented = true };

    private sealed record Snapshot(IReadOnlyList<FileIndexEntry> Index, IReadOnlyDictionary<string, string> Map, IReadOnlyList<string> Skipped);

    private sealed record ManifestModel(
        [property: JsonPropertyName("backup_version")] int BackupVersion,
        [property: JsonPropertyName("created_at_iso")] DateTimeOffset CreatedAtIso,
        [property: JsonPropertyName("backup_type")] string BackupType,
        [property: JsonPropertyName("app_name")] string AppName,
        [property: JsonPropertyName("source_paths")] SourcePaths SourcePaths,
        [property: JsonPropertyName("file_index")] IReadOnlyList<FileIndexEntry> FileIndex,
        [property: JsonPropertyName("deleted_paths")] IReadOnlyList<string> DeletedPaths,
        [property: JsonPropertyName("incremental_base")] string? IncrementalBase);

    private sealed record SourcePaths(
        [property: JsonPropertyName("config_dir")] string ConfigDirectory,
        [property: JsonPropertyName("data_dir")] string DataDirectory);

    private sealed record FileIndexEntry(
        [property: JsonPropertyName("path")] string Path,
        [property: JsonPropertyName("size")] long Size,
        [property: JsonPropertyName("mtime")] double Mtime,
        [property: JsonPropertyName("sha256")] string Sha256);
}
