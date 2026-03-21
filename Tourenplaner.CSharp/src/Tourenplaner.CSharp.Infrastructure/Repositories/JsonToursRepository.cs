using Tourenplaner.CSharp.Domain.Entities;
using Tourenplaner.CSharp.Infrastructure.Json;

namespace Tourenplaner.CSharp.Infrastructure.Repositories;

public sealed class JsonToursRepository : JsonRepositoryBase<IReadOnlyList<Tour>>
{
    public JsonToursRepository(JsonFileStore fileStore, string path)
        : base(fileStore, path, Array.Empty<Tour>())
    {
    }
}
