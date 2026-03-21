using Tourenplaner.CSharp.Domain.Services;
using Xunit;

namespace Tourenplaner.CSharp.Tests;

public sealed class SqlDatabaseNameInferenceTests
{
    [Fact]
    public void InferFromDataDirectory_ShouldPreferLargestNonSystemDatabase()
    {
        var directory = Directory.CreateTempSubdirectory();
        try
        {
            File.WriteAllBytes(Path.Combine(directory.FullName, "master.mdf"), new byte[10]);
            File.WriteAllBytes(Path.Combine(directory.FullName, "Touren_A.mdf"), new byte[50]);
            File.WriteAllBytes(Path.Combine(directory.FullName, "Touren_B.mdf"), new byte[100]);

            var result = SqlDatabaseNameInference.InferFromDataDirectory(directory.FullName);

            Assert.Equal("Touren_B", result);
        }
        finally
        {
            directory.Delete(recursive: true);
        }
    }
}
