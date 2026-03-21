using Tourenplaner.CSharp.Application.Abstractions;

namespace Tourenplaner.CSharp.Application.Services;

public sealed class SystemClock : IClock
{
    public DateTimeOffset UtcNow => DateTimeOffset.UtcNow;
}
