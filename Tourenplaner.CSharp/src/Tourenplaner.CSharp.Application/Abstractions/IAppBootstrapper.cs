using Tourenplaner.CSharp.Domain.Entities;

namespace Tourenplaner.CSharp.Application.Abstractions;

public interface IAppBootstrapper
{
    Task<AppSettings> LoadSettingsAsync(CancellationToken cancellationToken = default);
}
