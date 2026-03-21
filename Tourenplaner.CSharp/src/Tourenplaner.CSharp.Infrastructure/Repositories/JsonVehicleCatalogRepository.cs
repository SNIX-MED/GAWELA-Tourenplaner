using Tourenplaner.CSharp.Domain.Entities;
using Tourenplaner.CSharp.Infrastructure.Json;

namespace Tourenplaner.CSharp.Infrastructure.Repositories;

public sealed class JsonVehicleCatalogRepository : JsonRepositoryBase<VehicleCatalog>
{
    public JsonVehicleCatalogRepository(JsonFileStore fileStore, string path)
        : base(fileStore, path, VehicleCatalog.Empty)
    {
    }
}
