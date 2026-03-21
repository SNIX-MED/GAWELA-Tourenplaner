namespace Tourenplaner.CSharp.Application.Abstractions;

public interface IPathProvider
{
    string BaseDirectory { get; }
    string DataDirectory { get; }
    string ConfigDirectory { get; }
    string LogsDirectory { get; }
    string EmployeesPath { get; }
    string VehiclesPath { get; }
    string ToursPath { get; }
    string PinsPath { get; }
    string SettingsPath { get; }
    string PendingSqlOrdersPath { get; }
    string NonMapSqlOrdersPath { get; }
    string GeocodeCachePath { get; }
}
