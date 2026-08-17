using Smartstore.Engine.Modularity;

namespace Smartstore.SpecificationAttributeExchange;

internal sealed class Module : ModuleBase, IConfigurable
{
    public RouteInfo GetConfigurationRoute()
        => new("Configure", "SpecificationAttributeExchange", new { area = "Admin" });
}
