using Tourenplaner.CSharp.Domain.Entities;
using Tourenplaner.CSharp.Infrastructure.Json;

namespace Tourenplaner.CSharp.Infrastructure.Repositories;

public sealed class JsonEmployeesRepository : JsonRepositoryBase<IReadOnlyList<Employee>>
{
    public JsonEmployeesRepository(JsonFileStore fileStore, string path)
        : base(fileStore, path, Array.Empty<Employee>())
    {
    }
}
