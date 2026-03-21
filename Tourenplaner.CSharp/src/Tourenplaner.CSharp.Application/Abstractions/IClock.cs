namespace Tourenplaner.CSharp.Application.Abstractions;

public interface IClock
{
    DateTimeOffset UtcNow { get; }
}
