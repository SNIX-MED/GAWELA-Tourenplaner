using Tourenplaner.CSharp.Infrastructure.Repositories;
using Xunit;

namespace Tourenplaner.CSharp.Tests;

public sealed class RepositoryRootLocatorTests
{
    [Fact]
    public void LooksLikeRepositoryRoot_ShouldDetectPythonProjectLayout()
    {
        var temp = Path.Combine(Path.GetTempPath(), $"tourenplaner-root-{Guid.NewGuid():N}");
        Directory.CreateDirectory(temp);
        Directory.CreateDirectory(Path.Combine(temp, "data"));
        File.WriteAllText(Path.Combine(temp, "pins.json"), "[]");
        File.WriteAllText(Path.Combine(temp, "tours.json"), "[]");

        try
        {
            Assert.True(RepositoryRootLocator.LooksLikeRepositoryRoot(temp));
            Assert.Equal(temp, RepositoryRootLocator.Find(Path.Combine(temp, "data")));
        }
        finally
        {
            Directory.Delete(temp, recursive: true);
        }
    }
}
