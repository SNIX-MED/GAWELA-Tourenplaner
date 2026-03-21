namespace Tourenplaner.CSharp.Infrastructure.Logging;

public sealed class FileLogger
{
    private readonly string _logFilePath;

    public FileLogger(string logFilePath)
    {
        _logFilePath = logFilePath;
    }

    public async Task LogAsync(string level, string message, CancellationToken cancellationToken = default)
    {
        var directory = Path.GetDirectoryName(_logFilePath);
        if (!string.IsNullOrWhiteSpace(directory))
        {
            Directory.CreateDirectory(directory);
        }

        var line = $"{DateTimeOffset.UtcNow:O} [{level}] {message}{Environment.NewLine}";
        await File.AppendAllTextAsync(_logFilePath, line, cancellationToken);
    }
}
