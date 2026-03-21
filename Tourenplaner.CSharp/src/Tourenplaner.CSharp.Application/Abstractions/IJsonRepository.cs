namespace Tourenplaner.CSharp.Application.Abstractions;

public interface IJsonRepository<T>
{
    Task<T> LoadAsync(CancellationToken cancellationToken = default);
    Task SaveAsync(T value, CancellationToken cancellationToken = default);
}
