using System.IO;

namespace Tourenplaner.CSharp.Domain.Services;

public static class SqlDatabaseNameInference
{
    private static readonly HashSet<string> SystemDatabaseFiles = new(StringComparer.OrdinalIgnoreCase)
    {
        "master",
        "model",
        "msdb",
        "tempdb",
    };

    public static string InferFromDataDirectory(string? dataDirectory)
    {
        var path = new DirectoryInfo(string.IsNullOrWhiteSpace(dataDirectory) ? new Entities.AppSettings().SqlDataDir : dataDirectory.Trim());
        if (!path.Exists)
        {
            return string.Empty;
        }

        return path.EnumerateFiles("*.mdf", SearchOption.TopDirectoryOnly)
            .Where(file => !SystemDatabaseFiles.Contains(Path.GetFileNameWithoutExtension(file.Name)))
            .OrderByDescending(file => file.Length)
            .Select(file => Path.GetFileNameWithoutExtension(file.Name))
            .FirstOrDefault() ?? string.Empty;
    }
}
