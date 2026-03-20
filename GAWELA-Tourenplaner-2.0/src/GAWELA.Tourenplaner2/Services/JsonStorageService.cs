using System.Text.Json;

namespace GAWELA.Tourenplaner2.Services;

public sealed class InvalidJsonFileException(string message, Exception? inner = null) : Exception(message, inner);

public static class JsonStorageService
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true,
        PropertyNamingPolicy = JsonNamingPolicy.SnakeCaseLower,
    };

    public static async Task AtomicWriteJsonAsync<T>(string path, T payload)
    {
        var directory = Path.GetDirectoryName(path);
        if (!string.IsNullOrWhiteSpace(directory))
        {
            Directory.CreateDirectory(directory);
        }

        var tempFile = Path.Combine(directory ?? ".", $".{Path.GetFileName(path)}.{Guid.NewGuid():N}.tmp");
        await File.WriteAllTextAsync(tempFile, JsonSerializer.Serialize(payload, JsonOptions));
        File.Move(tempFile, path, true);
    }

    public static async Task<T> LoadJsonFileAsync<T>(string path, Func<T> defaultFactory, bool createIfMissing = false)
    {
        if (!File.Exists(path))
        {
            var fallback = defaultFactory();
            if (createIfMissing)
            {
                await AtomicWriteJsonAsync(path, fallback);
            }

            return fallback;
        }

        try
        {
            var raw = await File.ReadAllTextAsync(path);
            var parsed = JsonSerializer.Deserialize<T>(raw, JsonOptions);
            return parsed ?? defaultFactory();
        }
        catch (JsonException ex)
        {
            throw new InvalidJsonFileException($"Invalid JSON in {path}", ex);
        }
    }
}
