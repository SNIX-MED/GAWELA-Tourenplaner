namespace Tourenplaner.CSharp.Infrastructure.Repositories;

public static class RepositoryRootLocator
{
    public static string Find(string startDirectory)
    {
        var current = new DirectoryInfo(Path.GetFullPath(startDirectory));
        while (current is not null)
        {
            if (LooksLikeRepositoryRoot(current.FullName))
            {
                return current.FullName;
            }

            current = current.Parent;
        }

        return Path.GetFullPath(startDirectory);
    }

    public static bool LooksLikeRepositoryRoot(string path)
    {
        var pins = Path.Combine(path, "pins.json");
        var tours = Path.Combine(path, "tours.json");
        var data = Path.Combine(path, "data");
        return File.Exists(pins) && File.Exists(tours) && Directory.Exists(data);
    }
}
