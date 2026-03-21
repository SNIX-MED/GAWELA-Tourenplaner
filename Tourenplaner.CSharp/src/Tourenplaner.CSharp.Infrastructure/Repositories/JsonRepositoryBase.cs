using Tourenplaner.CSharp.Application.Abstractions;
using Tourenplaner.CSharp.Infrastructure.Json;

namespace Tourenplaner.CSharp.Infrastructure.Repositories;

public abstract class JsonRepositoryBase<T> : IJsonRepository<T>
{
    private readonly JsonFileStore _fileStore;
    private readonly string _path;
    private readonly T _fallback;

    protected JsonRepositoryBase(JsonFileStore fileStore, string path, T fallback)
    {
        _fileStore = fileStore;
        _path = path;
        _fallback = fallback;
    }

    public virtual Task<T> LoadAsync(CancellationToken cancellationToken = default) => _fileStore.LoadAsync(_path, _fallback, cancellationToken);

    public virtual Task SaveAsync(T value, CancellationToken cancellationToken = default) => _fileStore.SaveAsync(_path, value, cancellationToken);
}
