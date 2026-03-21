using Tourenplaner.CSharp.Domain.Entities;
using Tourenplaner.CSharp.Infrastructure.Json;

namespace Tourenplaner.CSharp.Infrastructure.Repositories;

public sealed class JsonPinsRepository : JsonRepositoryBase<IReadOnlyList<PinRecord>>
{
    public JsonPinsRepository(JsonFileStore fileStore, string path)
        : base(fileStore, path, Array.Empty<PinRecord>())
    {
    }
}
