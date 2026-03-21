using Tourenplaner.CSharp.Application.Abstractions;
using Tourenplaner.CSharp.Domain.Entities;
using Tourenplaner.CSharp.Domain.Services;
using Tourenplaner.CSharp.Infrastructure.Json;

namespace Tourenplaner.CSharp.Infrastructure.Repositories;

public sealed class JsonSettingsRepository : IJsonRepository<AppSettings>, IAppBootstrapper
{
    private readonly JsonFileStore _fileStore;
    private readonly string _settingsPath;
    private readonly string _defaultBackupDirectory;

    public JsonSettingsRepository(JsonFileStore fileStore, string settingsPath, string defaultBackupDirectory)
    {
        _fileStore = fileStore;
        _settingsPath = settingsPath;
        _defaultBackupDirectory = defaultBackupDirectory;
    }

    public async Task<AppSettings> LoadAsync(CancellationToken cancellationToken = default)
    {
        var loaded = await _fileStore.LoadAsync(_settingsPath, new AppSettings(), cancellationToken);
        return SettingsValidator.Validate(loaded, _defaultBackupDirectory);
    }

    public Task SaveAsync(AppSettings value, CancellationToken cancellationToken = default)
    {
        var validated = SettingsValidator.Validate(value, _defaultBackupDirectory);
        return _fileStore.SaveAsync(_settingsPath, validated, cancellationToken);
    }

    public Task<AppSettings> LoadSettingsAsync(CancellationToken cancellationToken = default) => LoadAsync(cancellationToken);
}
